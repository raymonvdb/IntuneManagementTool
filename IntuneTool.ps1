#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Management Tool - read-only overview of Intune configuration and assignments.

.DESCRIPTION
    Sign in with your Microsoft account and browse, in an Intune-like interface:
      - configuration profiles, compliance, Settings catalog, scripts, remediations
      - endpoint security (incl. ASR rules with conflict detection)
      - apps, app protection and app configuration
      - Autopilot, enrollment and Windows updates
      - groups, assignment filters and devices
    For every item you see the include and exclude assignments. You can search by
    group or user, view details and settings, and export to Excel, CSV and HTML.

    The tool only reads; nothing is changed in Intune.

.NOTES
    Structure:
      IntuneTool.ps1      this file: startup and wiring up the events
      lib\Graph.ps1       Microsoft Graph, sources, groups and filters
      lib\Settings.ps1    reading settings, scripts and ASR rules
      lib\Analysis.ps1    dashboard, attention points, search
      lib\Export.ps1      xlsx, csv and html
      lib\Ui.ps1          everything for the interface
      ui\*.xaml           layout and styles
    Start: double-click Start-IntuneTool.bat
#>

# WPF (the GUI) requires an STA thread. If the script starts in an MTA shell,
# it restarts itself with -STA.
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath
    exit
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

try {
    # -------------------------------------------------------------------------
    # 1. Load the components
    # -------------------------------------------------------------------------
    $script:Root = $PSScriptRoot
    foreach ($lib in 'Graph', 'Settings', 'Analysis', 'Export', 'Ui') {
        . (Join-Path $PSScriptRoot ('lib\' + $lib + '.ps1'))
    }

    # Microsoft Graph permissions (all read-only)
    $script:Scopes = @(
        'DeviceManagementConfiguration.Read.All',   # profiles, compliance, scripts, endpoint security, filters
        'DeviceManagementApps.Read.All',            # apps, app protection, app configuration
        'DeviceManagementServiceConfig.Read.All',   # Autopilot and enrollment
        'DeviceManagementManagedDevices.Read.All',  # devices and their policy status
        'Group.Read.All',                           # group names
        'GroupMember.Read.All',                     # member counts and memberships
        'User.Read.All'                             # searching for users
    )

    # Check the Graph module
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $answer = [System.Windows.MessageBox]::Show(
            "The module Microsoft.Graph.Authentication is required.`n`nInstall it now for your own account?",
            'Module missing', 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return }
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Authentication

    # -------------------------------------------------------------------------
    # 2. Build the window from XAML
    # -------------------------------------------------------------------------
    $script:ResourceXaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui\Resources.xaml'))
    $mainXaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui\MainWindow.xaml')).Replace('@@RESOURCES@@', $script:ResourceXaml)
    $window = [System.Windows.Markup.XamlReader]::Parse($mainXaml)

    # Make named elements available as variables ($Grid, $NavPanel, ...)
    $elementNames = @(
        'TxtAccount', 'BtnLogin', 'BtnLoad', 'NavPanel',
        'PanelDashboard', 'TxtDashHint', 'CardPanel', 'DashTypeGrid', 'DashAttentionGrid', 'BtnDashExport',
        'PanelTable', 'TxtViewTitle', 'TxtViewHint', 'TxtSearch', 'CmbPlatform', 'ChkUnassigned', 'BtnExtra',
        'TxtCount', 'BtnExport', 'Grid', 'TxtEmpty',
        'PanelSearch', 'TxtDirQuery', 'BtnSearch', 'ChkParents', 'ChkAll', 'BtnSearchExport',
        'LstCandidates', 'TxtSearchInfo', 'SearchGrid',
        'TxtStatus', 'Progress', 'BtnCancel'
    )
    foreach ($name in $elementNames) {
        $el = $window.FindName($name)
        if ($null -eq $el) { throw ('Element "' + $name + '" is missing in MainWindow.xaml') }
        Set-Variable -Name $name -Scope Script -Value $el
    }

    # -------------------------------------------------------------------------
    # 3. Build the left navigation (like Intune)
    # -------------------------------------------------------------------------
    $lastGroup = ''
    foreach ($view in $script:Views) {
        if ($view.Group -ne $lastGroup) {
            $header = New-Object System.Windows.Controls.TextBlock
            $header.Text  = ([string]$view.Group).ToUpper()
            $header.Style = $window.FindResource('NavHeader')
            [void]$NavPanel.Children.Add($header)
            $lastGroup = $view.Group
        }
        $radio = New-Object System.Windows.Controls.RadioButton
        $radio.Content   = [string]$view.Label
        $radio.Tag       = [string]$view.Key
        $radio.GroupName = 'nav'
        $radio.Style     = $window.FindResource('NavButton')
        $radio.Add_Checked({ param($sender, $e) Switch-View ([string]$sender.Tag) })
        [void]$NavPanel.Children.Add($radio)
        $script:NavButtons[[string]$view.Key] = $radio
    }

    foreach ($choice in $script:PlatformChoices) { [void]$CmbPlatform.Items.Add($choice) }
    $CmbPlatform.SelectedIndex = 0

    # -------------------------------------------------------------------------
    # 4. Wire up the events (the logic itself lives in lib\Ui.ps1)
    # -------------------------------------------------------------------------
    $BtnLogin.Add_Click({ Invoke-Login })
    $BtnLoad.Add_Click({ Invoke-FullLoad })
    $BtnCancel.Add_Click({ $script:Cancel = $true; Set-Status 'Cancelling...' })

    # Table view
    $TxtSearch.Add_TextChanged({ Update-Table })
    $CmbPlatform.Add_SelectionChanged({ Update-Table })
    $ChkUnassigned.Add_Click({ Update-Table })
    $BtnExtra.Add_Click({ Invoke-ViewAction })
    $BtnExport.Add_Click({ Export-CurrentTable })
    $Grid.Add_MouseDoubleClick({
        param($sender, $e)
        $row = Get-ClickedRow $e.OriginalSource
        if ($null -ne $row -and $script:CurrentView.Open) { & $script:CurrentView.Open $row }
    })

    # Dashboard
    $BtnDashExport.Add_Click({ Export-FullReport })
    $DashAttentionGrid.Add_MouseDoubleClick({
        param($sender, $e)
        $row = Get-ClickedRow $e.OriginalSource
        if ($null -ne $row) { Open-ItemById ([string]$row.ItemId) }
    })

    # Search by group/user
    $BtnSearch.Add_Click({ Invoke-DirectorySearch })
    $TxtDirQuery.Add_KeyDown({ param($sender, $e) if ($e.Key -eq 'Return') { Invoke-DirectorySearch } })
    $LstCandidates.Add_SelectionChanged({ if ($LstCandidates.SelectedItem) { Invoke-ReverseSearch $LstCandidates.SelectedItem } })
    $ChkParents.Add_Click({ if ($LstCandidates.SelectedItem) { Invoke-ReverseSearch $LstCandidates.SelectedItem } })
    $ChkAll.Add_Click({ if ($LstCandidates.SelectedItem) { Invoke-ReverseSearch $LstCandidates.SelectedItem } })
    $SearchGrid.Add_MouseDoubleClick({
        param($sender, $e)
        $row = Get-ClickedRow $e.OriginalSource
        if ($null -ne $row -and $row.Item) { Show-ItemDetails $row.Item }
    })
    $BtnSearchExport.Add_Click({
        $data = @($SearchGrid.ItemsSource | Where-Object { $null -ne $_ })
        Invoke-ExportDialog -Cols $script:ColSets.search -Rows $data -Title 'Search results'
    })

    # Report unexpected errors in handlers nicely instead of crashing
    $window.Dispatcher.Add_UnhandledException({
        param($sender, $e)
        Show-Error ('Unexpected error: ' + $e.Exception.Message)
        $e.Handled = $true
    })

    # -------------------------------------------------------------------------
    # 5. Start
    # -------------------------------------------------------------------------
    $script:NavButtons['dashboard'].IsChecked = $true
    Set-Status 'Sign in to get started.'
    [void]$window.ShowDialog()
}
catch {
    [System.Windows.MessageBox]::Show(('The tool could not start:' + [Environment]::NewLine + $_.Exception.Message), 'Intune Management Tool', 'OK', 'Error') | Out-Null
}
