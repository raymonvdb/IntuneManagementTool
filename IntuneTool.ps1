#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Management Tool - alleen-lezen overzicht van Intune-configuratie en assignments.

.DESCRIPTION
    Meld je aan met je Microsoft-account en bekijk in een Intune-achtige interface:
      - configuration profiles, compliance, Settings catalog, scripts, remediations
      - endpoint security (incl. ASR-regels met conflictdetectie)
      - apps, app protection en app configuration
      - Autopilot, enrollment en Windows updates
      - groepen, assignment filters en apparaten
    Per item zie je de include- en exclude-assignments. Je kunt zoeken per groep of
    gebruiker, details en instellingen bekijken en exporteren naar Excel, CSV en HTML.

    De tool leest alleen; er wordt niets in Intune gewijzigd.

.NOTES
    Opbouw:
      IntuneTool.ps1      dit bestand: opstarten en events koppelen
      lib\Graph.ps1       Microsoft Graph, bronnen, groepen en filters
      lib\Settings.ps1    instellingen, scripts en ASR-regels uitlezen
      lib\Analysis.ps1    dashboard, aandachtspunten, zoeken
      lib\Export.ps1      xlsx, csv en html
      lib\Ui.ps1          alles voor de interface
      ui\*.xaml           layout en stijlen
    Starten: dubbelklik op Start-IntuneTool.bat
#>

# WPF (de GUI) vereist een STA-thread. Als het script in een MTA-shell start,
# herstarten we zichzelf met -STA.
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath
    exit
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

try {
    # -------------------------------------------------------------------------
    # 1. Onderdelen laden
    # -------------------------------------------------------------------------
    $script:Root = $PSScriptRoot
    foreach ($lib in 'Graph', 'Settings', 'Analysis', 'Export', 'Ui') {
        . (Join-Path $PSScriptRoot ('lib\' + $lib + '.ps1'))
    }

    # Microsoft Graph-rechten (allemaal alleen-lezen)
    $script:Scopes = @(
        'DeviceManagementConfiguration.Read.All',   # profielen, compliance, scripts, endpoint security, filters
        'DeviceManagementApps.Read.All',            # apps, app protection, app configuration
        'DeviceManagementServiceConfig.Read.All',   # Autopilot en enrollment
        'DeviceManagementManagedDevices.Read.All',  # apparaten en hun policy-status
        'Group.Read.All',                           # groepsnamen
        'GroupMember.Read.All',                     # ledenaantal en lidmaatschappen
        'User.Read.All'                             # gebruikers zoeken
    )

    # De Graph-module controleren
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $answer = [System.Windows.MessageBox]::Show(
            "De module Microsoft.Graph.Authentication is nodig.`n`nNu installeren voor je eigen account?",
            'Module ontbreekt', 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return }
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Authentication

    # -------------------------------------------------------------------------
    # 2. Venster opbouwen uit XAML
    # -------------------------------------------------------------------------
    $script:ResourceXaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui\Resources.xaml'))
    $mainXaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui\MainWindow.xaml')).Replace('@@RESOURCES@@', $script:ResourceXaml)
    $window = [System.Windows.Markup.XamlReader]::Parse($mainXaml)

    # Benoemde elementen als variabelen beschikbaar maken ($Grid, $NavPanel, ...)
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
        if ($null -eq $el) { throw ('Element "' + $name + '" ontbreekt in MainWindow.xaml') }
        Set-Variable -Name $name -Scope Script -Value $el
    }

    # -------------------------------------------------------------------------
    # 3. Linker navigatie opbouwen (zoals in Intune)
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
    # 4. Events koppelen (de logica zelf staat in lib\Ui.ps1)
    # -------------------------------------------------------------------------
    $BtnLogin.Add_Click({ Invoke-Login })
    $BtnLoad.Add_Click({ Invoke-FullLoad })
    $BtnCancel.Add_Click({ $script:Cancel = $true; Set-Status 'Annuleren...' })

    # Tabelweergave
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

    # Zoeken per groep/gebruiker
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
        Invoke-ExportDialog -Cols $script:ColSets.search -Rows $data -Title 'Zoekresultaat'
    })

    # Onverwachte fouten in handlers netjes melden in plaats van crashen
    $window.Dispatcher.Add_UnhandledException({
        param($sender, $e)
        Show-Error ('Onverwachte fout: ' + $e.Exception.Message)
        $e.Handled = $true
    })

    # -------------------------------------------------------------------------
    # 5. Starten
    # -------------------------------------------------------------------------
    $script:NavButtons['dashboard'].IsChecked = $true
    Set-Status 'Meld je aan om te beginnen.'
    [void]$window.ShowDialog()
}
catch {
    [System.Windows.MessageBox]::Show(('De tool kon niet starten:' + [Environment]::NewLine + $_.Exception.Message), 'Intune Management Tool', 'OK', 'Error') | Out-Null
}
