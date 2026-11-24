# =============================================================================
# Ui.ps1
# Everything related to the user interface: status bar, tables,
# navigation, dashboard, search page, detail windows and export dialogs.
#
# Note: event handlers (Add_Click etc.) live in IntuneTool.ps1 and only call
# the functions below. Handlers do NOT use closures; they get their context
# via $sender.Tag or via $script: variables.
# =============================================================================

$script:CurrentView = $null
$script:CurrentCols = @()
$script:VisibleRows = @()
$script:NavButtons  = @{}
$script:Suspend     = $false


# =============================================================================
# 1. BASICS: status, progress, messages
# =============================================================================

function Invoke-UiRefresh {
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Set-Status {
    param([string]$Text)
    $TxtStatus.Text = $Text
    Invoke-UiRefresh
}

function Set-Progress {
    param([int]$Value, [int]$Max)
    $Progress.Maximum = [math]::Max(1, $Max)
    $Progress.Value   = [math]::Min($Value, $Max)
    Invoke-UiRefresh
}

# Puts the UI in "busy" mode: wait cursor, navigation off, cancel button visible.
function Set-Busy {
    param([bool]$Busy, [bool]$ShowCancel = $true)

    if ($Busy) { $window.Cursor = [System.Windows.Input.Cursors]::Wait } else { $window.Cursor = $null }
    $NavPanel.IsEnabled   = -not $Busy
    $BtnLogin.IsEnabled   = -not $Busy
    $BtnLoad.IsEnabled    = ((-not $Busy) -and $script:Connected)
    $BtnExtra.IsEnabled   = -not $Busy
    $BtnExport.IsEnabled  = -not $Busy

    if ($Busy) {
        $Progress.Visibility = 'Visible'
        if ($ShowCancel) { $BtnCancel.Visibility = 'Visible' }
        $script:Cancel = $false
    } else {
        $Progress.Visibility  = 'Collapsed'
        $BtnCancel.Visibility = 'Collapsed'
    }
    Invoke-UiRefresh
}

function Show-Info {
    param([string]$Text, [string]$Title = 'Intune Management Tool')
    [System.Windows.MessageBox]::Show($Text, $Title, 'OK', 'Information') | Out-Null
}

function Show-Error {
    param([string]$Text)
    [System.Windows.MessageBox]::Show($Text, 'Error', 'OK', 'Error') | Out-Null
}

function New-Brush {
    param([string]$Hex)
    return (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex)
}


# =============================================================================
# 2. TABLE COLUMNS
# =============================================================================
# H = header, P = property name, W = width (number or '*'), Min = minimum
# width for '*' columns, Style = name of a cell style from Resources.xaml

$script:ColSets = @{
    items = @(
        @{ H = 'Name';                           P = 'Name';     W = '*'; Min = 200 }
        @{ H = 'Type';                           P = 'Type';     W = 125 }
        @{ H = 'Subtype';                        P = 'Subtype';  W = 150 }
        @{ H = 'Platform';                       P = 'Platform'; W = 90 }
        @{ H = 'Included groups (include)';      P = 'Include';  W = '*'; Min = 190 }
        @{ H = 'Excluded groups (exclude)';      P = 'Exclude';  W = '*'; Min = 150 }
        @{ H = 'Status';                         P = 'Status';   W = 120; Style = 'StatusCell' }
        @{ H = 'Modified';                       P = 'Modified'; W = 120 }
    )
    asr = @(
        @{ H = 'Policy';                         P = 'Policy';   W = '*'; Min = 170 }
        @{ H = 'ASR rule / setting';             P = 'Rule';     W = '*'; Min = 220 }
        @{ H = 'Mode';                           P = 'Mode';     W = 110 }
        @{ H = 'Included groups (include)';      P = 'Include';  W = '*'; Min = 150 }
        @{ H = 'Excluded groups (exclude)';      P = 'Exclude';  W = '*'; Min = 130 }
        @{ H = 'Source';                         P = 'Source';   W = 120 }
        @{ H = 'Overlap / conflict';             P = 'Conflict'; W = '*'; Min = 200 }
    )
    attention = @(
        @{ H = 'Severity'; P = 'Severity'; W = 110; Style = 'SeverityCell' }
        @{ H = 'Type';     P = 'Type';     W = 140 }
        @{ H = 'Item';     P = 'Item';     W = '*'; Min = 220 }
        @{ H = 'Issue';    P = 'Issue';    W = '*'; Min = 300 }
    )
    groups = @(
        @{ H = 'Group';                      P = 'Name';     W = '*'; Min = 240 }
        @{ H = 'Dynamic';                    P = 'Dynamic';  W = 90 }
        @{ H = 'Members (incl. nested)';     P = 'Members';  W = 150 }
        @{ H = 'Included in (count)';        P = 'IncCount'; W = 140 }
        @{ H = 'Excluded in (count)';        P = 'ExcCount'; W = 140 }
        @{ H = 'Types';                      P = 'Types';    W = '*'; Min = 200 }
    )
    filters = @(
        @{ H = 'Filter';          P = 'Name';           W = '*'; Min = 200 }
        @{ H = 'Platform';        P = 'Platform';       W = 100 }
        @{ H = 'Management type'; P = 'ManagementType'; W = 130 }
        @{ H = 'Used in';         P = 'Used';           W = 90 }
        @{ H = 'Rule';            P = 'Rule';           W = '*'; Min = 260 }
        @{ H = 'Description';     P = 'Description';    W = '*'; Min = 150 }
    )
    devices = @(
        @{ H = 'Device';          P = 'Name';       W = '*'; Min = 160 }
        @{ H = 'User';            P = 'User';       W = '*'; Min = 170 }
        @{ H = 'OS';              P = 'OS';         W = 90 }
        @{ H = 'Version';         P = 'Version';    W = 110 }
        @{ H = 'Compliance';      P = 'Compliance'; W = 110; Style = 'ComplianceCell' }
        @{ H = 'Owner';           P = 'Owner';      W = 90 }
        @{ H = 'Model';           P = 'Model';      W = '*'; Min = 140 }
        @{ H = 'Serial number';   P = 'Serial';     W = 130 }
        @{ H = 'Encrypted';       P = 'Encrypted';  W = 90 }
        @{ H = 'Last sync';       P = 'LastSync';   W = 120 }
    )
    search = @(
        @{ H = 'Name';        P = 'Name';    W = '*'; Min = 220 }
        @{ H = 'Type';        P = 'Type';    W = 125 }
        @{ H = 'Subtype';     P = 'Subtype'; W = 150 }
        @{ H = 'Via';         P = 'Via';     W = '*'; Min = 180 }
        @{ H = 'Mode';        P = 'Mode';    W = 95 }
        @{ H = 'Effect';      P = 'Effect';  W = 170 }
        @{ H = 'Details';     P = 'Detail';  W = '*'; Min = 150 }
    )
    typesummary = @(
        @{ H = 'Type';            P = 'Type';        W = '*'; Min = 160 }
        @{ H = 'Total';           P = 'Total';       W = 90 }
        @{ H = 'Assigned';        P = 'Assigned';    W = 110 }
        @{ H = 'Not assigned';    P = 'Unassigned';  W = 120 }
        @{ H = 'Exclusion only';  P = 'ExcludeOnly'; W = 130 }
        @{ H = 'All devices';     P = 'AllDevices';  W = 110 }
        @{ H = 'All users';       P = 'AllUsers';    W = 110 }
    )
    flat = @(
        @{ H = 'Type';        P = 'Type';       W = 120 }
        @{ H = 'Subtype';     P = 'Subtype';    W = 140 }
        @{ H = 'Platform';    P = 'Platform';   W = 90 }
        @{ H = 'Name';        P = 'Name';       W = '*' }
        @{ H = 'Mode';        P = 'Mode';       W = 90 }
        @{ H = 'Assigned to'; P = 'Target';     W = '*' }
        @{ H = 'Filter';      P = 'Filter';     W = 150 }
        @{ H = 'Filter mode'; P = 'FilterMode'; W = 90 }
        @{ H = 'Intent';      P = 'Intent';     W = 100 }
    )
    kv = @(
        @{ H = 'Property'; P = 'Property'; W = 320 }
        @{ H = 'Value';    P = 'Value';    W = '*' }
    )
    settings = @(
        @{ H = 'Setting'; P = 'Setting'; W = '*'; Min = 300 }
        @{ H = 'Value';   P = 'Value';   W = '*'; Min = 300 }
    )
    assignments = @(
        @{ H = 'Mode';        P = 'ModeText';   W = 100 }
        @{ H = 'Target';      P = 'Target';     W = '*'; Min = 260 }
        @{ H = 'Filter';      P = 'Filter';     W = 200 }
        @{ H = 'Filter mode'; P = 'FilterMode'; W = 100 }
        @{ H = 'Intent';      P = 'Intent';     W = 160 }
    )
    states = @(
        @{ H = 'Kind';     P = 'Kind';     W = 110 }
        @{ H = 'Policy';   P = 'Policy';   W = '*'; Min = 300 }
        @{ H = 'Platform'; P = 'Platform'; W = 120 }
        @{ H = 'Status';   P = 'State';    W = 130; Style = 'StateCell' }
    )
}

# Sets the columns of a DataGrid.
function Set-GridColumns {
    param($Target, $Cols)

    $Target.Columns.Clear()
    foreach ($c in $Cols) {
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Header  = [string]$c.H
        $col.Binding = New-Object System.Windows.Data.Binding -ArgumentList ([string]$c.P)

        if ([string]$c.W -eq '*') {
            $col.Width = New-Object System.Windows.Controls.DataGridLength -ArgumentList 1, ([System.Windows.Controls.DataGridLengthUnitType]::Star)
        } else {
            $col.Width = New-Object System.Windows.Controls.DataGridLength -ArgumentList ([double]$c.W)
        }
        if ($c.Min) { $col.MinWidth = [double]$c.Min }

        if ($c.Style) { $styleKey = [string]$c.Style } else { $styleKey = 'WrapCell' }
        $col.ElementStyle = $window.FindResource($styleKey)
        $Target.Columns.Add($col) | Out-Null
    }
}

# Searches all text/number properties of a row (case-insensitive).
function Test-RowMatch {
    param($Row, [string]$Query)

    foreach ($prop in $Row.PSObject.Properties) {
        if ($prop.Name -eq 'Raw' -or $prop.Name -eq 'Assignments' -or $prop.Name -eq 'Item') { continue }
        $v = $prop.Value
        if ($null -eq $v) { continue }
        if ($v -is [string] -or $v -is [ValueType]) {
            if (([string]$v).IndexOf($Query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        }
    }
    return $false
}


# =============================================================================
# 3. VIEWS (the items in the left navigation)
# =============================================================================
# Kind    : dashboard | search | items | rows
# Filter  : (items) scriptblock that decides which items belong here
# Rows    : (rows) scriptblock that supplies the rows
# Cols    : key in $script:ColSets
# Hide    : columns (property names) that are not shown
# Open    : scriptblock that runs on double-click of a row
# Action  : (optional) extra button; ActionLabel is its text
# Ensure  : (optional) scriptblock that retrieves the data first

$script:Views = @(
    @{ Key = 'dashboard'; Group = 'Overview'; Label = 'Dashboard'; Kind = 'dashboard' }
    @{ Key = 'attention'; Group = 'Overview'; Label = 'Attention points'; Kind = 'rows'; Cols = 'attention'
       Title = 'Attention points'
       Hint  = 'Things that usually deserve a check or a cleanup. Double-click a row for details. Use "Get member counts" to also detect empty groups.'
       Rows  = { Get-AttentionRows }
       Open  = { param($r) Open-ItemById $r.ItemId }
       Action = { Invoke-MemberCountAction }; ActionLabel = 'Get member counts' }
    @{ Key = 'all'; Group = 'Overview'; Label = 'All assignments'; Kind = 'items'; Cols = 'items'
       Title = 'All assignments'
       Hint  = 'All policies, scripts, apps and profiles with their include and exclude assignments. Double-click a row for details and settings.'
       Filter = { $true }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'search'; Group = 'Overview'; Label = 'Search by group/user'; Kind = 'search' }

    @{ Key = 'config'; Group = 'Devices'; Label = 'Configuration profiles'; Kind = 'items'; Cols = 'items'; Hide = @()
       Title = 'Configuration profiles'
       Hint  = 'Device configuration profiles, administrative templates and Settings catalog.'
       Filter = { $_.Type -eq 'Device configuration' -or $_.Type -eq 'Settings catalog' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'compliance'; Group = 'Devices'; Label = 'Compliance policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Subtype')
       Title = 'Compliance policies'; Hint = 'Device compliance policies.'
       Filter = { $_.Type -eq 'Compliance' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'scripts'; Group = 'Devices'; Label = 'Platform scripts'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Platform scripts'; Hint = 'PowerShell and shell scripts. Double-click to read the script content.'
       Filter = { $_.Type -eq 'Platform script' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'remediations'; Group = 'Devices'; Label = 'Remediations'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Remediations'; Hint = 'Proactive remediations. Double-click to read the detection and remediation scripts.'
       Filter = { $_.Type -eq 'Remediation' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'winupdate'; Group = 'Devices'; Label = 'Windows updates'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Windows updates'; Hint = 'Update rings, feature updates, quality updates and driver updates.'
       Filter = { $_.Type -eq 'Windows update' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'autopilot'; Group = 'Devices'; Label = 'Autopilot profiles'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Platform')
       Title = 'Autopilot deployment profiles'; Hint = 'Windows Autopilot deployment profiles and their assignments.'
       Filter = { $_.Type -eq 'Autopilot' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'enrollment'; Group = 'Devices'; Label = 'Enrollment (ESP, restrictions)'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Enrollment'; Hint = 'Enrollment Status Page, platform restrictions, device limits and Windows Hello for Business.'
       Filter = { $_.Type -eq 'Enrollment' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'devices'; Group = 'Devices'; Label = 'Devices'; Kind = 'rows'; Cols = 'devices'
       Title = 'Managed devices'
       Hint  = 'All managed devices (loaded when you open this page). Double-click a device to see the status of its policies.'
       Rows  = { $script:Devices }
       Ensure = { Initialize-Devices $false }
       Open  = { param($r) Show-DeviceDetails $r }
       Action = { Initialize-Devices $true }; ActionLabel = 'Refresh' }

    @{ Key = 'es-all'; Group = 'Endpoint security'; Label = 'All endpoint security'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Endpoint security'; Hint = 'All endpoint security policies (new and old style) and security baselines.'
       Filter = { $_.Type -eq 'Endpoint security' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'asr'; Group = 'Endpoint security'; Label = 'ASR rules'; Kind = 'rows'; Cols = 'asr'
       Title = 'Attack Surface Reduction rules'
       Hint  = 'Every ASR rule per policy with mode and assignments. Overlap or conflicts between policies are shown on the right. The button also scans all Settings catalog profiles.'
       Rows  = { $script:AsrRows }
       Open  = { param($r) Open-ItemById $r.PolicyId }
       Action = { Invoke-AsrScanAction }; ActionLabel = 'Also scan Settings catalog' }
    @{ Key = 'es-asrpol'; Group = 'Endpoint security'; Label = 'Attack surface reduction'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Attack surface reduction policies'; Hint = 'ASR rules, device control, exploit protection, app and browser isolation and web protection.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'ASR' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-av'; Group = 'Endpoint security'; Label = 'Antivirus'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Antivirus'; Hint = 'Microsoft Defender Antivirus policies.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Antivirus' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-fw'; Group = 'Endpoint security'; Label = 'Firewall'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Firewall'; Hint = 'Microsoft Defender Firewall policies and rules.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Firewall' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-de'; Group = 'Endpoint security'; Label = 'Disk encryption'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Disk encryption'; Hint = 'BitLocker and FileVault policies.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Disk encryption' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-edr'; Group = 'Endpoint security'; Label = 'Endpoint detection and response'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Endpoint detection and response'; Hint = 'EDR policies (Defender for Endpoint).'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'EDR' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-acc'; Group = 'Endpoint security'; Label = 'Account protection'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Account protection'; Hint = 'Windows Hello, Credential Guard and local administrators.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Account protection' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-base'; Group = 'Endpoint security'; Label = 'Security baselines'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Security baselines'; Hint = 'Security baselines for Windows, Edge, Microsoft 365 Apps and Defender for Endpoint.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Security baseline' }; Open = { param($r) Show-ItemDetails $r } }

    @{ Key = 'apps'; Group = 'Apps'; Label = 'Apps'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Apps'; Hint = 'All apps with intent (Required / Available / Uninstall) per group.'
       Filter = { $_.Type -eq 'App' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'appprot'; Group = 'Apps'; Label = 'App protection policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Subtype')
       Title = 'App protection policies'; Hint = 'MAM policies for iOS/iPadOS, Android and Windows.'
       Filter = { $_.Type -eq 'App protection' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'appconf'; Group = 'Apps'; Label = 'App configuration policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'App configuration policies'; Hint = 'App configuration for managed devices and for apps (MAM).'
       Filter = { $_.Type -eq 'App configuration' }; Open = { param($r) Show-ItemDetails $r } }

    @{ Key = 'groups'; Group = 'Groups and filters'; Label = 'Groups in use'; Kind = 'rows'; Cols = 'groups'
       Title = 'Groups in use'
       Hint  = 'All groups that occur in an assignment. Double-click a group to see everything it is linked to.'
       Rows  = { Get-GroupRows }
       Open  = { param($r) Open-GroupInSearch $r }
       Action = { Invoke-MemberCountAction }; ActionLabel = 'Get member counts' }
    @{ Key = 'filters'; Group = 'Groups and filters'; Label = 'Assignment filters'; Kind = 'rows'; Cols = 'filters'
       Title = 'Assignment filters'
       Hint  = 'All assignment filters and how often they are used. Double-click to see where a filter is used.'
       Rows  = { Get-FilterRows }
       Open  = { param($r) Open-FilterUsage $r } }
)

$script:PlatformChoices = @('All platforms', 'Windows', 'macOS', 'iOS/iPadOS', 'Android', 'Linux')

function Get-ViewColumns {
    param($V)
    $set = $script:ColSets[[string]$V.Cols]
    if ($V.Hide) { return @($set | Where-Object { $V.Hide -notcontains $_.P }) }
    return @($set)
}

function Get-ViewRows {
    param($V)
    if ($V.Kind -eq 'items') { return @($script:AllItems | Where-Object -FilterScript $V.Filter) }
    if ($V.Rows)             { return @(& $V.Rows | Where-Object { $null -ne $_ }) }
    return @()
}

# Applies the search filters and refreshes the table.
function Update-Table {
    if ($script:Suspend) { return }
    $v = $script:CurrentView
    if ($null -eq $v -or $v.Kind -eq 'dashboard' -or $v.Kind -eq 'search') { return }

    $rows = @(Get-ViewRows $v)

    if ($v.Kind -eq 'items') {
        if ($CmbPlatform.SelectedIndex -gt 0) {
            $plat = [string]$CmbPlatform.SelectedItem
            $rows = @($rows | Where-Object { $_.Platform -eq $plat })
        }
        if ($ChkUnassigned.IsChecked) { $rows = @($rows | Where-Object { $_.Status -ne 'Assigned' }) }
    }

    $q = $TxtSearch.Text.Trim()
    if ($q) { $rows = @($rows | Where-Object { Test-RowMatch $_ $q }) }

    $script:VisibleRows = $rows
    $Grid.ItemsSource   = $rows
    $TxtCount.Text      = '{0} rows' -f $rows.Count

    if ($rows.Count -eq 0) {
        if (-not $script:Connected)                                        { $TxtEmpty.Text = 'Sign in and click "Load data".' }
        elseif (@($script:AllItems).Count -eq 0 -and $v.Key -ne 'devices') { $TxtEmpty.Text = 'No data yet. Click "Load data".' }
        else                                                               { $TxtEmpty.Text = 'No results.' }
        $TxtEmpty.Visibility = 'Visible'
    } else {
        $TxtEmpty.Visibility = 'Collapsed'
    }
}

function Show-TableView {
    param($V)

    $TxtViewTitle.Text = [string]$V.Title
    $TxtViewHint.Text  = [string]$V.Hint

    $cols = Get-ViewColumns $V
    Set-GridColumns $Grid $cols
    $script:CurrentCols = $cols

    $isItems = ($V.Kind -eq 'items')
    if ($isItems) { $CmbPlatform.Visibility = 'Visible'; $ChkUnassigned.Visibility = 'Visible' }
    else          { $CmbPlatform.Visibility = 'Collapsed'; $ChkUnassigned.Visibility = 'Collapsed' }

    if ($V.Action) { $BtnExtra.Content = [string]$V.ActionLabel; $BtnExtra.Visibility = 'Visible' }
    else           { $BtnExtra.Visibility = 'Collapsed' }

    # Reset filters without every change triggering a recalculation
    $script:Suspend = $true
    $TxtSearch.Text = ''
    $CmbPlatform.SelectedIndex = 0
    $ChkUnassigned.IsChecked = $false
    $script:Suspend = $false

    if ($V.Ensure -and $script:Connected) { & $V.Ensure }
    Update-Table
}

function Switch-View {
    param([string]$Key)

    $v = $script:Views | Where-Object { $_.Key -eq $Key } | Select-Object -First 1
    if (-not $v) { return }
    $script:CurrentView = $v

    $PanelDashboard.Visibility = 'Collapsed'
    $PanelTable.Visibility     = 'Collapsed'
    $PanelSearch.Visibility    = 'Collapsed'

    if ($v.Kind -eq 'dashboard') {
        $PanelDashboard.Visibility = 'Visible'
        Update-Dashboard
    }
    elseif ($v.Kind -eq 'search') {
        $PanelSearch.Visibility = 'Visible'
        if (@($script:AllItems).Count -eq 0 -and -not $LstCandidates.ItemsSource) {
            $TxtSearchInfo.Text = 'Sign in and load the data first.'
        }
    }
    else {
        $PanelTable.Visibility = 'Visible'
        Show-TableView $v
    }
}

# Puts counters behind the navigation items.
function Update-NavCounts {
    $loaded = (@($script:AllItems).Count -gt 0)
    foreach ($v in $script:Views) {
        if ($v.Kind -eq 'dashboard' -or $v.Kind -eq 'search') { continue }
        $btn = $script:NavButtons[$v.Key]
        if (-not $loaded -or ($v.Key -eq 'devices' -and $null -eq $script:Devices)) { $btn.Content = $v.Label; continue }
        $n = @(Get-ViewRows $v).Count
        $btn.Content = '{0}  ({1})' -f $v.Label, $n
    }
}

# Finds the row that was double-clicked (and ignores clicks on headers).
function Get-ClickedRow {
    param($OriginalSource)
    $d = $OriginalSource
    $guard = 0
    while ($null -ne $d -and $guard -lt 40) {
        if ($d -is [System.Windows.Controls.DataGridRow]) { return $d.Item }
        try {
            if ($d -is [System.Windows.Media.Visual] -or $d -is [System.Windows.Media.Media3D.Visual3D]) {
                $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
            } else {
                $d = [System.Windows.LogicalTreeHelper]::GetParent($d)
            }
        } catch { return $null }
        $guard++
    }
    return $null
}

function Open-ItemById {
    param([string]$Id)
    $it = $script:AllItems | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if ($it) { Show-ItemDetails $it } else { Show-Info 'This item was not found (anymore) in the loaded data.' }
}

function Open-FilterUsage {
    param($Row)
    $script:NavButtons['all'].IsChecked = $true
    $TxtSearch.Text = 'filter: ' + $Row.Name
}

function Open-GroupInSearch {
    param($Row)
    if ($Row.GroupId -eq 'ALL_DEVICES' -or $Row.GroupId -eq 'ALL_USERS') {
        Show-Info 'A per-group search is not possible for "All devices" and "All users". Use "All assignments" and search on the name.'
        return
    }
    $script:NavButtons['search'].IsChecked = $true
    $cand = [pscustomobject]@{ Kind = 'group'; Id = [string]$Row.GroupId; Name = [string]$Row.Name; Display = ('[Group] ' + [string]$Row.Name) }
    $LstCandidates.ItemsSource = @($cand)
    $LstCandidates.SelectedIndex = 0
}


# =============================================================================
# 4. ACTIONS (extra buttons)
# =============================================================================

function Invoke-ViewAction {
    $v = $script:CurrentView
    if ($v -and $v.Action) { & $v.Action }
}

function Invoke-MemberCountAction {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Load the data first.'; return }
    try {
        Set-Busy $true
        $r = Invoke-MemberCountCheck
        Set-Status ('Member counts retrieved for {0} groups ({1} failed).' -f ($r.Total - $r.Failed), $r.Failed)
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { Set-Busy $false }
    if ($r -and $r.Failed -gt 0) {
        Show-Info ('The member count could not be retrieved for {0} groups (probably missing permission: GroupMember.Read.All).' -f $r.Failed)
    }
    Update-NavCounts
    Update-Table
}

function Invoke-AsrScanAction {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Load the data first.'; return }
    try {
        Set-Busy $true
        $script:LoadErrors.Clear()
        $script:AsrRows = @(Get-AsrRows -ScanAllCatalog $true)
        Set-Status ('{0} ASR rules found (including Settings catalog).' -f $script:AsrRows.Count)
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { Set-Busy $false }
    Update-NavCounts
    Update-Table
}

function Initialize-Devices {
    param([bool]$Force)
    if (-not $script:Connected) { return }
    if (-not $Force -and $null -ne $script:Devices) { return }
    try {
        Set-Busy $true $false
        Set-Status 'Retrieving devices...'
        $script:Devices = @(Get-DeviceRows)
        Set-Status ('{0} devices loaded.' -f $script:Devices.Count)
    }
    catch {
        $script:Devices = @()
        Show-Error ('Retrieving devices failed (permission DeviceManagementManagedDevices.Read.All is required): ' + (Get-ErrorText $_))
    }
    finally { Set-Busy $false }
    Update-NavCounts
    if ($Force) { Update-Table }
}


# =============================================================================
# 5. DASHBOARD
# =============================================================================

function New-Card {
    param([string]$Label, $Value, [string]$Accent = '#0F6CBD', [string]$Tag = '')

    $border = New-Object System.Windows.Controls.Border
    $border.Width           = 190
    $border.MinHeight       = 80
    $border.Margin          = New-Object System.Windows.Thickness -ArgumentList 0, 0, 12, 12
    $border.Padding         = New-Object System.Windows.Thickness -ArgumentList 14, 10, 14, 10
    $border.Background      = New-Brush '#FFFFFF'
    $border.BorderBrush     = New-Brush $Accent
    $border.BorderThickness = New-Object System.Windows.Thickness -ArgumentList 0, 0, 0, 3

    $stack = New-Object System.Windows.Controls.StackPanel
    $big = New-Object System.Windows.Controls.TextBlock
    $big.Text       = [string]$Value
    $big.FontSize   = 28
    $big.FontWeight = [System.Windows.FontWeights]::SemiBold
    $small = New-Object System.Windows.Controls.TextBlock
    $small.Text         = $Label
    $small.Foreground   = New-Brush '#605E5C'
    $small.TextWrapping = 'Wrap'
    [void]$stack.Children.Add($big)
    [void]$stack.Children.Add($small)
    $border.Child = $stack

    if ($Tag) {
        $border.Tag    = $Tag
        $border.Cursor = [System.Windows.Input.Cursors]::Hand
        $border.Add_MouseLeftButtonUp({ param($s, $e) Invoke-CardAction ([string]$s.Tag) })
    }
    return $border
}

# Click on a card: go to the matching view (optionally with a filter).
function Invoke-CardAction {
    param([string]$Tag)
    $parts = $Tag -split '\|'
    $script:NavButtons[$parts[0]].IsChecked = $true
    if ($parts.Count -gt 1 -and $parts[1] -eq 'unassigned') {
        $ChkUnassigned.IsChecked = $true
        Update-Table
    }
}

function Update-Dashboard {
    $CardPanel.Children.Clear()
    $items = @($script:AllItems)

    if ($items.Count -eq 0) {
        $TxtDashHint.Text = 'No data yet. Sign in and click "Load data".'
        $DashTypeGrid.ItemsSource = @()
        $DashAttentionGrid.ItemsSource = @()
        return
    }

    $attention = @(Get-AttentionRows)
    $typeRows  = @(Get-TypeSummaryRows)

    $assigned   = @($items | Where-Object { $_.Status -eq 'Assigned' }).Count
    $unassigned = @($items | Where-Object { $_.Status -eq 'Not assigned' }).Count
    $exclOnly   = @($items | Where-Object { $_.Status -eq 'Exclusion only' }).Count
    $allDev     = ($typeRows | Measure-Object -Property AllDevices -Sum).Sum
    $allUsr     = ($typeRows | Measure-Object -Property AllUsers -Sum).Sum
    $groupIds   = @($items | ForEach-Object { $_.Assignments } | Where-Object { $_.Kind -eq 'group' } | Select-Object -ExpandProperty GroupId -Unique)
    $errors     = @($attention | Where-Object { $_.Severity -eq 'Error' }).Count
    $asrRules   = @($script:AsrRows | Select-Object -ExpandProperty Rule -Unique).Count
    $asrConf    = @($script:AsrRows | Where-Object { $_.Conflict -like 'CONFLICT*' } | Select-Object -ExpandProperty DefinitionId -Unique).Count

    $TxtDashHint.Text = 'Summary of all loaded items. Click a card to go to the details.'

    [void]$CardPanel.Children.Add((New-Card 'Items total' $items.Count '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Assigned' $assigned '#107C10' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Not assigned' $unassigned '#C50F1F' 'all|unassigned'))
    [void]$CardPanel.Children.Add((New-Card 'Exclusions only' $exclOnly '#B45F06' 'all|unassigned'))
    [void]$CardPanel.Children.Add((New-Card 'Assigned to All devices' $allDev '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Assigned to All users' $allUsr '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Groups in use' $groupIds.Count '#0F6CBD' 'groups'))
    [void]$CardPanel.Children.Add((New-Card 'Attention points' $attention.Count '#B45F06' 'attention'))
    [void]$CardPanel.Children.Add((New-Card 'Errors' $errors '#C50F1F' 'attention'))
    [void]$CardPanel.Children.Add((New-Card 'ASR rules in use' $asrRules '#0F6CBD' 'asr'))
    [void]$CardPanel.Children.Add((New-Card 'ASR conflicts' $asrConf '#C50F1F' 'asr'))

    Set-GridColumns $DashTypeGrid $script:ColSets.typesummary
    $DashTypeGrid.ItemsSource = $typeRows

    Set-GridColumns $DashAttentionGrid $script:ColSets.attention
    $DashAttentionGrid.ItemsSource = @($attention | Select-Object -First 12)
}


# =============================================================================
# 6. SEARCH BY GROUP OR USER
# =============================================================================

function Invoke-DirectorySearch {
    if (-not $script:Connected) { $TxtSearchInfo.Text = 'Sign in first.'; return }
    $text = $TxtDirQuery.Text.Trim()
    if ($text.Length -lt 2) { $TxtSearchInfo.Text = 'Type at least 2 characters.'; return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $TxtSearchInfo.Text = 'Searching...'
        Invoke-UiRefresh
        $res = Search-Directory -Text $text
        $LstCandidates.ItemsSource = @($res.Results)
        $SearchGrid.ItemsSource = @()
        if (@($res.Results).Count -eq 0) {
            $msg = 'Nothing found for "' + $text + '".'
            if ($res.Error) { $msg += ' ' + $res.Error }
            $TxtSearchInfo.Text = $msg
        } else {
            $TxtSearchInfo.Text = '{0} results. Pick one in the list.' -f @($res.Results).Count
            if (@($res.Results).Count -eq 1) { $LstCandidates.SelectedIndex = 0 }
        }
    }
    catch { $TxtSearchInfo.Text = Get-ErrorText $_ }
    finally { $window.Cursor = $null }
}

function Invoke-ReverseSearch {
    param($Cand)

    if ($null -eq $Cand) { return }
    if (@($script:AllItems).Count -eq 0) {
        $TxtSearchInfo.Text = 'Load the data first (button "Load data").'
        return
    }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $map  = @{}
        $note = ''
        if ($Cand.Kind -eq 'group') { $map[[string]$Cand.Id] = [string]$Cand.Name }

        if ($Cand.Kind -eq 'user' -or $ChkParents.IsChecked) {
            try {
                foreach ($g in @(Get-ParentGroups -Kind $Cand.Kind -Id $Cand.Id)) {
                    $gid = [string]$g.id
                    if ($map.ContainsKey($gid)) { continue }
                    if ($Cand.Kind -eq 'user') { $map[$gid] = 'Member of: ' + [string]$g.displayName }
                    else                       { $map[$gid] = 'Parent group: ' + [string]$g.displayName }
                }
            }
            catch { $note = ' Memberships could not be retrieved: ' + (Get-ErrorText $_) }
        }

        $rows = @(Find-AssignmentRows -GroupMap $map -IncludeAll ([bool]$ChkAll.IsChecked))
        $SearchGrid.ItemsSource = $rows
        Set-GridColumns $SearchGrid $script:ColSets.search

        $effective = @($rows | Where-Object { $_.Effect -eq 'Applies' }).Count
        $TxtSearchInfo.Text = ('{0}: {1} assignments found via {2} group(s); {3} of them apply. Double-click for details.{4}' -f $Cand.Display, $rows.Count, $map.Count, $effective, $note)
    }
    catch { $TxtSearchInfo.Text = Get-ErrorText $_ }
    finally { $window.Cursor = $null }
}


# =============================================================================
# 7. DETAIL WINDOWS
# =============================================================================

# Creates an empty detail window with a TabControl.
function New-DetailWindow {
    param([string]$Title, [string]$Heading, [string]$Sub)

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1080" Height="740" WindowStartupLocation="CenterOwner" Background="#F3F2F1"
        FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <ResourceDictionary>
@@RESOURCES@@
    </ResourceDictionary>
  </Window.Resources>
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0">
      <TextBlock x:Name="DetTitle" FontSize="20" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="DetSub" Foreground="#605E5C" Margin="0,2,0,0" TextWrapping="Wrap"/>
    </StackPanel>
    <TabControl x:Name="DetTabs" Grid.Row="1" Margin="0,12,0,0"/>
  </Grid>
</Window>
'@
    $xaml = $xaml.Replace('@@RESOURCES@@', $script:ResourceXaml)
    $w = [System.Windows.Markup.XamlReader]::Parse($xaml)
    $w.Title = $Title
    $w.Owner = $window
    $w.FindName('DetTitle').Text = $Heading
    $w.FindName('DetSub').Text   = $Sub
    return [pscustomobject]@{ Window = $w; Tabs = $w.FindName('DetTabs') }
}

# Tab with a table. With $Searchable a search box is added above it.
function Add-GridTab {
    param($Tabs, [string]$Header, $Rows, $Cols, [bool]$Searchable = $false)

    $dg = New-Object System.Windows.Controls.DataGrid
    Set-GridColumns $dg $Cols
    $all = @($Rows)
    $dg.ItemsSource = $all

    $tab = New-Object System.Windows.Controls.TabItem
    $tab.Header = $Header

    if ($Searchable) {
        $dock = New-Object System.Windows.Controls.DockPanel
        $tb = New-Object System.Windows.Controls.TextBox
        $tb.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 8, 0, 8
        $tb.Tag    = @{ Grid = $dg; All = $all }
        [System.Windows.Controls.DockPanel]::SetDock($tb, [System.Windows.Controls.Dock]::Top)
        $tb.Add_TextChanged({ param($s, $e) Update-TabFilter $s })
        [void]$dock.Children.Add($tb)
        [void]$dock.Children.Add($dg)
        $tab.Content = $dock
    } else {
        $tab.Content = $dg
    }
    [void]$Tabs.Items.Add($tab)
}

function Update-TabFilter {
    param($Box)
    $st = $Box.Tag
    $q  = $Box.Text.Trim()
    if ($q) { $st.Grid.ItemsSource = @($st.All | Where-Object { Test-RowMatch $_ $q }) }
    else    { $st.Grid.ItemsSource = $st.All }
}

# Tab with plain text (scripts).
function Add-TextTab {
    param($Tabs, [string]$Header, [string]$Text)

    $tb = New-Object System.Windows.Controls.TextBox
    $tb.IsReadOnly = $true
    $tb.AcceptsReturn = $true
    $tb.FontFamily = New-Object System.Windows.Media.FontFamily -ArgumentList 'Consolas'
    $tb.FontSize   = 12
    $tb.TextWrapping = [System.Windows.TextWrapping]::NoWrap
    $tb.HorizontalScrollBarVisibility = 'Auto'
    $tb.VerticalScrollBarVisibility   = 'Auto'
    $tb.Text = $Text

    $tab = New-Object System.Windows.Controls.TabItem
    $tab.Header  = $Header
    $tab.Content = $tb
    [void]$Tabs.Items.Add($tab)
}

function Show-ItemDetails {
    param($Item)
    if ($null -eq $Item) { return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Invoke-UiRefresh

        $sub = (@($Item.Type, $Item.Subtype, $Item.Platform) | Where-Object { $_ }) -join '  |  '
        $dw  = New-DetailWindow -Title ('Details - ' + $Item.Name) -Heading $Item.Name -Sub $sub
        $tabs = $dw.Tabs

        # 1. Assignments
        $arows = @(foreach ($a in $Item.Assignments) {
            [pscustomobject]@{ ModeText = $a.Mode; Target = $a.Name; Filter = $a.FilterName; FilterMode = $a.FilterMode; Intent = $a.Intent }
        })
        if ($arows.Count -eq 0) { $arows = @([pscustomobject]@{ ModeText = ''; Target = '(not assigned)'; Filter = ''; FilterMode = ''; Intent = '' }) }
        Add-GridTab $tabs ('Assignments (' + @($Item.Assignments).Count + ')') $arows $script:ColSets.assignments

        # 2. General properties
        $props = @(
            [pscustomobject]@{ Property = 'Name';          Value = $Item.Name }
            [pscustomobject]@{ Property = 'Type';          Value = $Item.Type }
            [pscustomobject]@{ Property = 'Subtype';       Value = $Item.Subtype }
            [pscustomobject]@{ Property = 'Platform';      Value = $Item.Platform }
            [pscustomobject]@{ Property = 'Description';   Value = $Item.Description }
            [pscustomobject]@{ Property = 'Last modified'; Value = $Item.Modified }
            [pscustomobject]@{ Property = 'Id';            Value = $Item.Id }
            [pscustomobject]@{ Property = 'Graph source';  Value = ($Item.Root + '/' + $Item.Path) }
        )
        Add-GridTab $tabs 'Properties' $props $script:ColSets.kv

        # 3. Settings or script content (handled separately so an error does not block the rest)
        try {
            switch ($Item.SrcKey) {
                'sc' {
                    $rows = @(Get-ConfigPolicySettings -Id $Item.Id)
                    Add-GridTab $tabs ('Settings (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                'intents' {
                    $rows = @(Get-IntentSettings -Id $Item.Id)
                    Add-GridTab $tabs ('Settings (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                'gpo' {
                    $rows = @(Get-GpoSettings -Id $Item.Id)
                    Add-GridTab $tabs ('Settings (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                { $_ -eq 'scripts' -or $_ -eq 'shell' -or $_ -eq 'health' } {
                    $texts = Get-ScriptTexts -Item $Item
                    foreach ($k in $texts.Keys) { Add-TextTab $tabs $k ([string]$texts[$k]) }
                }
                default {
                    $rows = @(Get-ItemProperties -Item $Item)
                    Add-GridTab $tabs ('Settings (' + $rows.Count + ')') $rows $script:ColSets.kv $true
                }
            }
        }
        catch {
            Add-TextTab $tabs 'Retrieval error' ('The settings could not be retrieved:' + [Environment]::NewLine + (Get-ErrorText $_))
        }

        $dw.Window.Show()
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { $window.Cursor = $null }
}

function Show-DeviceDetails {
    param($Device)
    if ($null -eq $Device) { return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Invoke-UiRefresh

        $sub = (@($Device.User, $Device.OS + ' ' + $Device.Version, $Device.Model) | Where-Object { $_ }) -join '  |  '
        $dw  = New-DetailWindow -Title ('Device - ' + $Device.Name) -Heading $Device.Name -Sub $sub

        try {
            $states = @(Get-DeviceStateRows -DeviceId $Device.Id)
            Add-GridTab $dw.Tabs ('Policy status (' + $states.Count + ')') $states $script:ColSets.states $true
        }
        catch { Add-TextTab $dw.Tabs 'Policy status' ('Status could not be retrieved:' + [Environment]::NewLine + (Get-ErrorText $_)) }

        try {
            $obj = Invoke-Graph -Uri ($script:GraphBeta + '/deviceManagement/managedDevices/' + $Device.Id)
            $out = New-Object System.Collections.Generic.List[object]
            ConvertTo-FlatProperties -Obj $obj -Prefix '' -Out $out
            Add-GridTab $dw.Tabs 'Properties' @($out.ToArray()) $script:ColSets.kv $true
        }
        catch { Add-TextTab $dw.Tabs 'Properties' ('Properties could not be retrieved:' + [Environment]::NewLine + (Get-ErrorText $_)) }

        $dw.Window.Show()
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { $window.Cursor = $null }
}


# =============================================================================
# 8. EXPORT
# =============================================================================

# Asks where the file should go. Returns $null when cancelled.
function Select-ExportPath {
    param([string]$Name)

    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Excel workbook (*.xlsx)|*.xlsx|CSV, semicolon (*.csv)|*.csv|HTML report (*.html)|*.html'
    $safe = [regex]::Replace($Name, '[^A-Za-z0-9]+', '_')
    $dlg.FileName = '{0}_{1}' -f $safe, (Get-Date -Format 'yyyyMMdd_HHmm')
    $dlg.InitialDirectory = $script:Root
    if (-not $dlg.ShowDialog($window)) { return $null }
    return $dlg.FileName
}

function Open-ExportedFile {
    param([string]$Path)
    $ans = [System.Windows.MessageBox]::Show(('Saved:' + [Environment]::NewLine + $Path + [Environment]::NewLine + [Environment]::NewLine + 'Open it now?'), 'Export complete', 'YesNo', 'Question')
    if ($ans -eq 'Yes') { Start-Process -FilePath $Path }
}

# Exports one table (the visible rows).
function Invoke-ExportDialog {
    param($Cols, $Rows, [string]$Title)

    $data = @($Rows)
    if ($data.Count -eq 0) { Show-Info 'There are no rows to export.'; return }
    $path = Select-ExportPath $Title
    if (-not $path) { return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        switch ([System.IO.Path]::GetExtension($path).ToLowerInvariant()) {
            '.csv'  { Export-RowsCsv -Path $path -Cols $Cols -Rows $data }
            '.html' { Export-HtmlReport -Path $path -Title $Title -Cards $null -Sections @(@{ Title = $Title; Cols = $Cols; Rows = $data }) -Subtitle $TxtAccount.Text }
            default { Export-Xlsx -Path $path -Sheets @(@{ Name = $Title; Cols = $Cols; Rows = $data }) }
        }
    }
    catch { Show-Error ('Export failed: ' + (Get-ErrorText $_)); return }
    finally { $window.Cursor = $null }

    Set-Status ('Exported {0} rows to {1}' -f $data.Count, $path)
    Open-ExportedFile $path
}

function Export-CurrentTable {
    $v = $script:CurrentView
    if ($null -eq $v) { return }
    Invoke-ExportDialog -Cols $script:CurrentCols -Rows $script:VisibleRows -Title ([string]$v.Title)
}

# Full report from the dashboard: Excel (multiple sheets), HTML or CSV (flat).
function Export-FullReport {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Load the data first.'; return }
    $path = Select-ExportPath 'Intune_report'
    if (-not $path) { return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $typeRows  = @(Get-TypeSummaryRows)
        $attention = @(Get-AttentionRows)
        $flat      = @(Get-FlatAssignmentRows)
        $itemCols  = $script:ColSets.items

        switch ([System.IO.Path]::GetExtension($path).ToLowerInvariant()) {
            '.csv' {
                Export-RowsCsv -Path $path -Cols $script:ColSets.flat -Rows $flat
            }
            '.html' {
                $cards = @(
                    @{ Label = 'Items total';      Value = @($script:AllItems).Count }
                    @{ Label = 'Not assigned';     Value = @($script:AllItems | Where-Object { $_.Status -eq 'Not assigned' }).Count }
                    @{ Label = 'Attention points'; Value = $attention.Count }
                    @{ Label = 'ASR rules';        Value = @($script:AsrRows).Count }
                )
                $sections = @(
                    @{ Title = 'Per type';          Cols = $script:ColSets.typesummary; Rows = $typeRows }
                    @{ Title = 'Attention points';  Cols = $script:ColSets.attention;   Rows = $attention }
                    @{ Title = 'All assignments';   Cols = $itemCols;                   Rows = @($script:AllItems) }
                )
                if (@($script:AsrRows).Count -gt 0) { $sections += @{ Title = 'ASR rules'; Cols = $script:ColSets.asr; Rows = @($script:AsrRows) } }
                Export-HtmlReport -Path $path -Title 'Intune assignments - report' -Cards $cards -Sections $sections -Subtitle $TxtAccount.Text
            }
            default {
                $sheets = @(
                    @{ Name = 'Per type';             Cols = $script:ColSets.typesummary; Rows = $typeRows }
                    @{ Name = 'Attention points';     Cols = $script:ColSets.attention;   Rows = $attention }
                    @{ Name = 'All items';            Cols = $itemCols;                   Rows = @($script:AllItems) }
                    @{ Name = 'Assignments (flat)';   Cols = $script:ColSets.flat;        Rows = $flat }
                )
                if (@($script:AsrRows).Count -gt 0) { $sheets += @{ Name = 'ASR rules'; Cols = $script:ColSets.asr; Rows = @($script:AsrRows) } }
                Export-Xlsx -Path $path -Sheets $sheets
            }
        }
    }
    catch { Show-Error ('Export failed: ' + (Get-ErrorText $_)); return }
    finally { $window.Cursor = $null }

    Set-Status ('Report saved: ' + $path)
    Open-ExportedFile $path
}


# =============================================================================
# 9. SIGNING IN AND LOADING
# =============================================================================

function Invoke-Login {
    try {
        Set-Busy $true $false
        Set-Status 'Signing in... a sign-in window will open.'
        if (Get-MgContext) { try { Disconnect-MgGraph | Out-Null } catch { } }
        Connect-MgGraph -Scopes $script:Scopes -NoWelcome -ContextScope Process | Out-Null

        $ctx = Get-MgContext
        $script:Connected = $true
        $tenant = [string]$ctx.TenantId
        try {
            $org = Invoke-Graph -Uri ($script:GraphV1 + '/organization?$select=displayName')
            if (@($org.value).Count -gt 0 -and $org.value[0].displayName) { $tenant = [string]$org.value[0].displayName }
        } catch { }
        $TxtAccount.Text   = [string]$ctx.Account + '  |  ' + $tenant
        $BtnLogin.Content  = 'Switch account'

        $missing = @($script:Scopes | Where-Object { @($ctx.Scopes) -notcontains $_ })
        if ($missing.Count -gt 0) {
            Show-Info ('Note: not all requested permissions were granted. Parts that need these permissions may stay empty.' + [Environment]::NewLine + [Environment]::NewLine + ($missing -join [Environment]::NewLine))
        }
    }
    catch {
        $script:Connected = $false
        Show-Error ('Sign-in failed: ' + (Get-ErrorText $_))
        Set-Status 'Sign-in failed.'
        Set-Busy $false
        return
    }
    Set-Busy $false
    Invoke-FullLoad
}

function Invoke-FullLoad {
    if (-not $script:Connected) { return }
    try {
        Set-Busy $true
        Get-AllItems
    }
    catch { Show-Error ('Loading failed: ' + (Get-ErrorText $_)) }
    finally { Set-Busy $false }

    Update-NavCounts
    if ($script:CurrentView) { Switch-View ([string]$script:CurrentView.Key) }

    if ($script:Cancel) { Set-Status 'Loading cancelled; the data may be incomplete.' }
    if ($script:LoadErrors.Count -gt 0) {
        $first = @($script:LoadErrors | Select-Object -First 8)
        $msg = 'Some parts could not be retrieved (usually a missing permission or an unlicensed feature):' + [Environment]::NewLine + [Environment]::NewLine + ($first -join [Environment]::NewLine)
        if ($script:LoadErrors.Count -gt 8) { $msg += [Environment]::NewLine + '... and ' + ($script:LoadErrors.Count - 8) + ' more.' }
        Show-Info $msg 'Warning'
    }
}
