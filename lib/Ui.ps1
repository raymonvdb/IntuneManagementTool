# =============================================================================
# Ui.ps1
# Alles wat met de gebruikersinterface te maken heeft: statusbalk, tabellen,
# navigatie, dashboard, zoekpagina, detailvensters en exportdialogen.
#
# Let op: event handlers (Add_Click e.d.) staan in IntuneTool.ps1 en roepen
# alleen de functies hieronder aan. Handlers gebruiken GEEN closures; ze krijgen
# hun context via $sender.Tag of via $script:-variabelen.
# =============================================================================

$script:CurrentView = $null
$script:CurrentCols = @()
$script:VisibleRows = @()
$script:NavButtons  = @{}
$script:Suspend     = $false


# =============================================================================
# 1. BASIS: status, voortgang, meldingen
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

# Zet de UI in "bezig"-modus: wachtcursor, navigatie uit, annuleerknop zichtbaar.
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
    [System.Windows.MessageBox]::Show($Text, 'Fout', 'OK', 'Error') | Out-Null
}

function New-Brush {
    param([string]$Hex)
    return (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex)
}


# =============================================================================
# 2. TABELKOLOMMEN
# =============================================================================
# H = koptekst, P = eigenschapsnaam, W = breedte (getal of '*'), Min = minimale
# breedte voor '*'-kolommen, Style = naam van een celstijl uit Resources.xaml

$script:ColSets = @{
    items = @(
        @{ H = 'Naam';                           P = 'Name';     W = '*'; Min = 200 }
        @{ H = 'Type';                           P = 'Type';     W = 125 }
        @{ H = 'Subtype';                        P = 'Subtype';  W = 150 }
        @{ H = 'Platform';                       P = 'Platform'; W = 90 }
        @{ H = 'Opgenomen groepen (include)';    P = 'Include';  W = '*'; Min = 190 }
        @{ H = 'Uitgesloten groepen (exclude)';  P = 'Exclude';  W = '*'; Min = 150 }
        @{ H = 'Status';                         P = 'Status';   W = 120; Style = 'StatusCell' }
        @{ H = 'Gewijzigd';                      P = 'Modified'; W = 120 }
    )
    asr = @(
        @{ H = 'Policy';                         P = 'Policy';   W = '*'; Min = 170 }
        @{ H = 'ASR-regel / instelling';         P = 'Rule';     W = '*'; Min = 220 }
        @{ H = 'Modus';                          P = 'Mode';     W = 110 }
        @{ H = 'Opgenomen groepen (include)';    P = 'Include';  W = '*'; Min = 150 }
        @{ H = 'Uitgesloten groepen (exclude)';  P = 'Exclude';  W = '*'; Min = 130 }
        @{ H = 'Bron';                           P = 'Source';   W = 120 }
        @{ H = 'Overlap / conflict';             P = 'Conflict'; W = '*'; Min = 200 }
    )
    attention = @(
        @{ H = 'Ernst';    P = 'Severity'; W = 110; Style = 'SeverityCell' }
        @{ H = 'Type';     P = 'Type';     W = 140 }
        @{ H = 'Item';     P = 'Item';     W = '*'; Min = 220 }
        @{ H = 'Probleem'; P = 'Issue';    W = '*'; Min = 300 }
    )
    groups = @(
        @{ H = 'Groep';                      P = 'Name';     W = '*'; Min = 240 }
        @{ H = 'Dynamisch';                  P = 'Dynamic';  W = 90 }
        @{ H = 'Leden (incl. geneste)';      P = 'Members';  W = 130 }
        @{ H = 'Opgenomen in (aantal)';      P = 'IncCount'; W = 140 }
        @{ H = 'Uitgesloten in (aantal)';    P = 'ExcCount'; W = 150 }
        @{ H = 'Typen';                      P = 'Types';    W = '*'; Min = 200 }
    )
    filters = @(
        @{ H = 'Filter';          P = 'Name';           W = '*'; Min = 200 }
        @{ H = 'Platform';        P = 'Platform';       W = 100 }
        @{ H = 'Beheertype';      P = 'ManagementType'; W = 100 }
        @{ H = 'Gebruikt in';     P = 'Used';           W = 90 }
        @{ H = 'Regel';           P = 'Rule';           W = '*'; Min = 260 }
        @{ H = 'Beschrijving';    P = 'Description';    W = '*'; Min = 150 }
    )
    devices = @(
        @{ H = 'Apparaat';        P = 'Name';       W = '*'; Min = 160 }
        @{ H = 'Gebruiker';       P = 'User';       W = '*'; Min = 170 }
        @{ H = 'OS';              P = 'OS';         W = 90 }
        @{ H = 'Versie';          P = 'Version';    W = 110 }
        @{ H = 'Compliance';      P = 'Compliance'; W = 110; Style = 'ComplianceCell' }
        @{ H = 'Eigenaar';        P = 'Owner';      W = 90 }
        @{ H = 'Model';           P = 'Model';      W = '*'; Min = 140 }
        @{ H = 'Serienummer';     P = 'Serial';     W = 130 }
        @{ H = 'Versleuteld';     P = 'Encrypted';  W = 90 }
        @{ H = 'Laatste sync';    P = 'LastSync';   W = 120 }
    )
    search = @(
        @{ H = 'Naam';        P = 'Name';    W = '*'; Min = 220 }
        @{ H = 'Type';        P = 'Type';    W = 125 }
        @{ H = 'Subtype';     P = 'Subtype'; W = 150 }
        @{ H = 'Via';         P = 'Via';     W = '*'; Min = 180 }
        @{ H = 'Modus';       P = 'Mode';    W = 95 }
        @{ H = 'Effect';      P = 'Effect';  W = 170 }
        @{ H = 'Details';     P = 'Detail';  W = '*'; Min = 150 }
    )
    typesummary = @(
        @{ H = 'Type';               P = 'Type';        W = '*'; Min = 160 }
        @{ H = 'Totaal';             P = 'Total';       W = 90 }
        @{ H = 'Toegewezen';         P = 'Assigned';    W = 110 }
        @{ H = 'Niet toegewezen';    P = 'Unassigned';  W = 130 }
        @{ H = 'Alleen uitsluiting'; P = 'ExcludeOnly'; W = 140 }
        @{ H = 'Alle apparaten';     P = 'AllDevices';  W = 120 }
        @{ H = 'Alle gebruikers';    P = 'AllUsers';    W = 120 }
    )
    flat = @(
        @{ H = 'Type';        P = 'Type';       W = 120 }
        @{ H = 'Subtype';     P = 'Subtype';    W = 140 }
        @{ H = 'Platform';    P = 'Platform';   W = 90 }
        @{ H = 'Naam';        P = 'Name';       W = '*' }
        @{ H = 'Modus';       P = 'Mode';       W = 90 }
        @{ H = 'Toegewezen aan'; P = 'Target';  W = '*' }
        @{ H = 'Filter';      P = 'Filter';     W = 150 }
        @{ H = 'Filtermodus'; P = 'FilterMode'; W = 90 }
        @{ H = 'Intent';      P = 'Intent';     W = 100 }
    )
    kv = @(
        @{ H = 'Eigenschap'; P = 'Eigenschap'; W = 320 }
        @{ H = 'Waarde';     P = 'Waarde';     W = '*' }
    )
    settings = @(
        @{ H = 'Instelling'; P = 'Instelling'; W = '*'; Min = 300 }
        @{ H = 'Waarde';     P = 'Waarde';     W = '*'; Min = 300 }
    )
    assignments = @(
        @{ H = 'Modus';       P = 'ModeText';   W = 100 }
        @{ H = 'Doel';        P = 'Target';     W = '*'; Min = 260 }
        @{ H = 'Filter';      P = 'Filter';     W = 200 }
        @{ H = 'Filtermodus'; P = 'FilterMode'; W = 100 }
        @{ H = 'Intent';      P = 'Intent';     W = 160 }
    )
    states = @(
        @{ H = 'Soort';    P = 'Kind';     W = 110 }
        @{ H = 'Policy';   P = 'Policy';   W = '*'; Min = 300 }
        @{ H = 'Platform'; P = 'Platform'; W = 120 }
        @{ H = 'Status';   P = 'State';    W = 130; Style = 'StateCell' }
    )
}

# Zet de kolommen van een DataGrid.
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

# Zoekt in alle tekst/getal-eigenschappen van een rij (hoofdletterongevoelig).
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
# 3. WEERGAVEN (de onderdelen in de linker navigatie)
# =============================================================================
# Kind    : dashboard | search | items | rows
# Filter  : (items) scriptblock dat bepaalt welke items erbij horen
# Rows    : (rows) scriptblock dat de rijen levert
# Cols    : sleutel in $script:ColSets
# Hide    : kolommen (eigenschapsnamen) die niet getoond worden
# Open    : scriptblock dat bij dubbelklik op een rij wordt uitgevoerd
# Action  : (optioneel) extra knop; ActionLabel is de tekst
# Ensure  : (optioneel) scriptblock dat eerst de gegevens ophaalt

$script:Views = @(
    @{ Key = 'dashboard'; Group = 'Overzicht'; Label = 'Dashboard'; Kind = 'dashboard' }
    @{ Key = 'attention'; Group = 'Overzicht'; Label = 'Aandachtspunten'; Kind = 'rows'; Cols = 'attention'
       Title = 'Aandachtspunten'
       Hint  = 'Dingen die meestal een controle of opruimactie waard zijn. Dubbelklik op een regel voor details. Met "Ledenaantal ophalen" worden ook lege groepen herkend.'
       Rows  = { Get-AttentionRows }
       Open  = { param($r) Open-ItemById $r.ItemId }
       Action = { Invoke-MemberCountAction }; ActionLabel = 'Ledenaantal ophalen' }
    @{ Key = 'all'; Group = 'Overzicht'; Label = 'Alle assignments'; Kind = 'items'; Cols = 'items'
       Title = 'Alle assignments'
       Hint  = 'Alle policies, scripts, apps en profielen met hun include- en exclude-assignments. Dubbelklik op een regel voor details en instellingen.'
       Filter = { $true }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'search'; Group = 'Overzicht'; Label = 'Zoeken op groep/gebruiker'; Kind = 'search' }

    @{ Key = 'config'; Group = 'Apparaten'; Label = 'Configuration profiles'; Kind = 'items'; Cols = 'items'; Hide = @()
       Title = 'Configuration profiles'
       Hint  = 'Device configuration profiles, administrative templates en Settings catalog.'
       Filter = { $_.Type -eq 'Device configuration' -or $_.Type -eq 'Settings catalog' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'compliance'; Group = 'Apparaten'; Label = 'Compliance policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Subtype')
       Title = 'Compliance policies'; Hint = 'Compliance policies voor apparaten.'
       Filter = { $_.Type -eq 'Compliance' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'scripts'; Group = 'Apparaten'; Label = 'Platform scripts'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Platform scripts'; Hint = 'PowerShell- en shell-scripts. Dubbelklik om de scriptinhoud te lezen.'
       Filter = { $_.Type -eq 'Platform script' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'remediations'; Group = 'Apparaten'; Label = 'Remediations'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Remediations'; Hint = 'Proactive remediations. Dubbelklik om detection- en remediation-script te lezen.'
       Filter = { $_.Type -eq 'Remediation' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'winupdate'; Group = 'Apparaten'; Label = 'Windows updates'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Windows updates'; Hint = 'Update rings, feature updates, quality updates en driver updates.'
       Filter = { $_.Type -eq 'Windows update' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'autopilot'; Group = 'Apparaten'; Label = 'Autopilot profielen'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Platform')
       Title = 'Autopilot deployment profiles'; Hint = 'Windows Autopilot deployment profiles en hun assignments.'
       Filter = { $_.Type -eq 'Autopilot' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'enrollment'; Group = 'Apparaten'; Label = 'Enrollment (ESP, restricties)'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Enrollment'; Hint = 'Enrollment Status Page, platform restrictions, device limits en Windows Hello for Business.'
       Filter = { $_.Type -eq 'Enrollment' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'devices'; Group = 'Apparaten'; Label = 'Apparaten'; Kind = 'rows'; Cols = 'devices'
       Title = 'Beheerde apparaten'
       Hint  = 'Alle beheerde apparaten (wordt geladen zodra je deze pagina opent). Dubbelklik op een apparaat om de status van zijn policies te zien.'
       Rows  = { $script:Devices }
       Ensure = { Initialize-Devices $false }
       Open  = { param($r) Show-DeviceDetails $r }
       Action = { Initialize-Devices $true }; ActionLabel = 'Vernieuwen' }

    @{ Key = 'es-all'; Group = 'Endpoint security'; Label = 'Alle endpoint security'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Endpoint security'; Hint = 'Alle endpoint security policies (nieuwe en oude stijl) en security baselines.'
       Filter = { $_.Type -eq 'Endpoint security' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'asr'; Group = 'Endpoint security'; Label = 'ASR-regels'; Kind = 'rows'; Cols = 'asr'
       Title = 'Attack Surface Reduction-regels'
       Hint  = 'Elke ASR-regel per policy met modus en assignments. Overlap of conflicten tussen policies staan rechts. Met de knop scan je ook alle Settings catalog-profielen.'
       Rows  = { $script:AsrRows }
       Open  = { param($r) Open-ItemById $r.PolicyId }
       Action = { Invoke-AsrScanAction }; ActionLabel = 'Scan ook Settings catalog' }
    @{ Key = 'es-asrpol'; Group = 'Endpoint security'; Label = 'Attack surface reduction'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Attack surface reduction policies'; Hint = 'ASR rules, device control, exploit protection, app and browser isolation en web protection.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'ASR' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-av'; Group = 'Endpoint security'; Label = 'Antivirus'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Antivirus'; Hint = 'Microsoft Defender Antivirus policies.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Antivirus' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-fw'; Group = 'Endpoint security'; Label = 'Firewall'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Firewall'; Hint = 'Microsoft Defender Firewall policies en regels.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Firewall' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-de'; Group = 'Endpoint security'; Label = 'Disk encryption'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Disk encryption'; Hint = 'BitLocker en FileVault policies.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Disk encryption' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-edr'; Group = 'Endpoint security'; Label = 'Endpoint detection and response'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Endpoint detection and response'; Hint = 'EDR policies (Defender for Endpoint).'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'EDR' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-acc'; Group = 'Endpoint security'; Label = 'Account protection'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Account protection'; Hint = 'Windows Hello, Credential Guard en lokale beheerders.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Account protection' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'es-base'; Group = 'Endpoint security'; Label = 'Security baselines'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Security baselines'; Hint = 'Security baselines voor Windows, Edge, Microsoft 365 Apps en Defender for Endpoint.'
       Filter = { $_.Type -eq 'Endpoint security' -and $_.Category -eq 'Security baseline' }; Open = { param($r) Show-ItemDetails $r } }

    @{ Key = 'apps'; Group = 'Apps'; Label = 'Apps'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'Apps'; Hint = 'Alle apps met intent (Required / Available / Uninstall) per groep.'
       Filter = { $_.Type -eq 'App' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'appprot'; Group = 'Apps'; Label = 'App protection policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type', 'Subtype')
       Title = 'App protection policies'; Hint = 'MAM-beleid voor iOS/iPadOS, Android en Windows.'
       Filter = { $_.Type -eq 'App protection' }; Open = { param($r) Show-ItemDetails $r } }
    @{ Key = 'appconf'; Group = 'Apps'; Label = 'App configuration policies'; Kind = 'items'; Cols = 'items'; Hide = @('Type')
       Title = 'App configuration policies'; Hint = 'App-configuratie voor beheerde apparaten en voor apps (MAM).'
       Filter = { $_.Type -eq 'App configuration' }; Open = { param($r) Show-ItemDetails $r } }

    @{ Key = 'groups'; Group = 'Groepen en filters'; Label = 'Groepen in gebruik'; Kind = 'rows'; Cols = 'groups'
       Title = 'Groepen in gebruik'
       Hint  = 'Alle groepen die in een assignment voorkomen. Dubbelklik op een groep om te zien waar hij allemaal aan gekoppeld is.'
       Rows  = { Get-GroupRows }
       Open  = { param($r) Open-GroupInSearch $r }
       Action = { Invoke-MemberCountAction }; ActionLabel = 'Ledenaantal ophalen' }
    @{ Key = 'filters'; Group = 'Groepen en filters'; Label = 'Assignment filters'; Kind = 'rows'; Cols = 'filters'
       Title = 'Assignment filters'
       Hint  = 'Alle assignment filters en hoe vaak ze gebruikt worden. Dubbelklik om te zien waar een filter gebruikt wordt.'
       Rows  = { Get-FilterRows }
       Open  = { param($r) Open-FilterUsage $r } }
)

$script:PlatformChoices = @('Alle platformen', 'Windows', 'macOS', 'iOS/iPadOS', 'Android', 'Linux')

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

# Past de zoekfilters toe en ververst de tabel.
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
        if ($ChkUnassigned.IsChecked) { $rows = @($rows | Where-Object { $_.Status -ne 'Toegewezen' }) }
    }

    $q = $TxtSearch.Text.Trim()
    if ($q) { $rows = @($rows | Where-Object { Test-RowMatch $_ $q }) }

    $script:VisibleRows = $rows
    $Grid.ItemsSource   = $rows
    $TxtCount.Text      = '{0} regels' -f $rows.Count

    if ($rows.Count -eq 0) {
        if (-not $script:Connected)                                   { $TxtEmpty.Text = 'Meld je aan en klik op "Gegevens laden".' }
        elseif (@($script:AllItems).Count -eq 0 -and $v.Key -ne 'devices') { $TxtEmpty.Text = 'Nog geen gegevens. Klik op "Gegevens laden".' }
        else                                                          { $TxtEmpty.Text = 'Geen resultaten.' }
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

    # Filters terugzetten zonder dat elke wijziging een herberekening triggert
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
            $TxtSearchInfo.Text = 'Meld je aan en laad eerst de gegevens.'
        }
    }
    else {
        $PanelTable.Visibility = 'Visible'
        Show-TableView $v
    }
}

# Zet tellers achter de navigatie-items.
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

# Zoekt de rij waarop dubbelgeklikt is (en negeert klikken op kopjes).
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
    if ($it) { Show-ItemDetails $it } else { Show-Info 'Dit item is niet (meer) gevonden in de geladen gegevens.' }
}

function Open-FilterUsage {
    param($Row)
    $script:NavButtons['all'].IsChecked = $true
    $TxtSearch.Text = 'filter: ' + $Row.Name
}

function Open-GroupInSearch {
    param($Row)
    if ($Row.GroupId -eq 'ALL_DEVICES' -or $Row.GroupId -eq 'ALL_USERS') {
        Show-Info 'Voor "Alle apparaten" en "Alle gebruikers" is geen zoekopdracht per groep mogelijk. Gebruik "Alle assignments" en zoek op de naam.'
        return
    }
    $script:NavButtons['search'].IsChecked = $true
    $cand = [pscustomobject]@{ Kind = 'group'; Id = [string]$Row.GroupId; Name = [string]$Row.Name; Display = ('[Groep] ' + [string]$Row.Name) }
    $LstCandidates.ItemsSource = @($cand)
    $LstCandidates.SelectedIndex = 0
}


# =============================================================================
# 4. ACTIES (extra knoppen)
# =============================================================================

function Invoke-ViewAction {
    $v = $script:CurrentView
    if ($v -and $v.Action) { & $v.Action }
}

function Invoke-MemberCountAction {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Laad eerst de gegevens.'; return }
    try {
        Set-Busy $true
        $r = Invoke-MemberCountCheck
        Set-Status ('Ledenaantal opgehaald voor {0} groepen ({1} mislukt).' -f ($r.Total - $r.Failed), $r.Failed)
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { Set-Busy $false }
    if ($r -and $r.Failed -gt 0) {
        Show-Info ('Voor {0} groepen kon het ledenaantal niet worden opgehaald (waarschijnlijk ontbrekende rechten: GroupMember.Read.All).' -f $r.Failed)
    }
    Update-NavCounts
    Update-Table
}

function Invoke-AsrScanAction {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Laad eerst de gegevens.'; return }
    try {
        Set-Busy $true
        $script:LoadErrors.Clear()
        $script:AsrRows = @(Get-AsrRows -ScanAllCatalog $true)
        Set-Status ('{0} ASR-regels gevonden (inclusief Settings catalog).' -f $script:AsrRows.Count)
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
        Set-Status 'Apparaten ophalen...'
        $script:Devices = @(Get-DeviceRows)
        Set-Status ('{0} apparaten geladen.' -f $script:Devices.Count)
    }
    catch {
        $script:Devices = @()
        Show-Error ('Apparaten ophalen mislukt (recht DeviceManagementManagedDevices.Read.All nodig): ' + (Get-ErrorText $_))
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

# Klik op een kaart: ga naar de bijbehorende weergave (eventueel met filter).
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
        $TxtDashHint.Text = 'Nog geen gegevens. Meld je aan en klik op "Gegevens laden".'
        $DashTypeGrid.ItemsSource = @()
        $DashAttentionGrid.ItemsSource = @()
        return
    }

    $attention = @(Get-AttentionRows)
    $typeRows  = @(Get-TypeSummaryRows)

    $assigned   = @($items | Where-Object { $_.Status -eq 'Toegewezen' }).Count
    $unassigned = @($items | Where-Object { $_.Status -eq 'Niet toegewezen' }).Count
    $exclOnly   = @($items | Where-Object { $_.Status -eq 'Alleen uitsluiting' }).Count
    $allDev     = ($typeRows | Measure-Object -Property AllDevices -Sum).Sum
    $allUsr     = ($typeRows | Measure-Object -Property AllUsers -Sum).Sum
    $groupIds   = @($items | ForEach-Object { $_.Assignments } | Where-Object { $_.Kind -eq 'group' } | Select-Object -ExpandProperty GroupId -Unique)
    $errors     = @($attention | Where-Object { $_.Severity -eq 'Fout' }).Count
    $asrRules   = @($script:AsrRows | Select-Object -ExpandProperty Rule -Unique).Count
    $asrConf    = @($script:AsrRows | Where-Object { $_.Conflict -like 'CONFLICT*' } | Select-Object -ExpandProperty DefinitionId -Unique).Count

    $TxtDashHint.Text = 'Samenvatting van alle geladen items. Klik op een kaart om naar de details te gaan.'

    [void]$CardPanel.Children.Add((New-Card 'Items totaal' $items.Count '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Toegewezen' $assigned '#107C10' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Niet toegewezen' $unassigned '#C50F1F' 'all|unassigned'))
    [void]$CardPanel.Children.Add((New-Card 'Alleen uitsluitingen' $exclOnly '#B45F06' 'all|unassigned'))
    [void]$CardPanel.Children.Add((New-Card 'Toegewezen aan Alle apparaten' $allDev '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Toegewezen aan Alle gebruikers' $allUsr '#0F6CBD' 'all'))
    [void]$CardPanel.Children.Add((New-Card 'Groepen in gebruik' $groupIds.Count '#0F6CBD' 'groups'))
    [void]$CardPanel.Children.Add((New-Card 'Aandachtspunten' $attention.Count '#B45F06' 'attention'))
    [void]$CardPanel.Children.Add((New-Card 'Fouten' $errors '#C50F1F' 'attention'))
    [void]$CardPanel.Children.Add((New-Card 'ASR-regels in gebruik' $asrRules '#0F6CBD' 'asr'))
    [void]$CardPanel.Children.Add((New-Card 'ASR-conflicten' $asrConf '#C50F1F' 'asr'))

    Set-GridColumns $DashTypeGrid $script:ColSets.typesummary
    $DashTypeGrid.ItemsSource = $typeRows

    Set-GridColumns $DashAttentionGrid $script:ColSets.attention
    $DashAttentionGrid.ItemsSource = @($attention | Select-Object -First 12)
}


# =============================================================================
# 6. ZOEKEN OP GROEP OF GEBRUIKER
# =============================================================================

function Invoke-DirectorySearch {
    if (-not $script:Connected) { $TxtSearchInfo.Text = 'Meld je eerst aan.'; return }
    $text = $TxtDirQuery.Text.Trim()
    if ($text.Length -lt 2) { $TxtSearchInfo.Text = 'Typ minimaal 2 tekens.'; return }

    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $TxtSearchInfo.Text = 'Zoeken...'
        Invoke-UiRefresh
        $res = Search-Directory -Text $text
        $LstCandidates.ItemsSource = @($res.Results)
        $SearchGrid.ItemsSource = @()
        if (@($res.Results).Count -eq 0) {
            $msg = 'Niets gevonden voor "' + $text + '".'
            if ($res.Error) { $msg += ' ' + $res.Error }
            $TxtSearchInfo.Text = $msg
        } else {
            $TxtSearchInfo.Text = '{0} resultaten. Kies er een in de lijst.' -f @($res.Results).Count
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
        $TxtSearchInfo.Text = 'Laad eerst de gegevens (knop "Gegevens laden").'
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
                    if ($Cand.Kind -eq 'user') { $map[$gid] = 'Lid van: ' + [string]$g.displayName }
                    else                       { $map[$gid] = 'Bovenliggende groep: ' + [string]$g.displayName }
                }
            }
            catch { $note = ' Lidmaatschappen konden niet worden opgehaald: ' + (Get-ErrorText $_) }
        }

        $rows = @(Find-AssignmentRows -GroupMap $map -IncludeAll ([bool]$ChkAll.IsChecked))
        $SearchGrid.ItemsSource = $rows
        Set-GridColumns $SearchGrid $script:ColSets.search

        $effective = @($rows | Where-Object { $_.Effect -eq 'Van toepassing' }).Count
        $TxtSearchInfo.Text = ('{0}: {1} assignments gevonden via {2} groep(en); {3} daarvan van toepassing. Dubbelklik voor details.{4}' -f $Cand.Display, $rows.Count, $map.Count, $effective, $note)
    }
    catch { $TxtSearchInfo.Text = Get-ErrorText $_ }
    finally { $window.Cursor = $null }
}


# =============================================================================
# 7. DETAILVENSTERS
# =============================================================================

# Maakt een leeg detailvenster met een TabControl.
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

# Tab met een tabel. Met -Searchable komt er een zoekveld boven.
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

# Tab met platte tekst (scripts).
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
            if ($a.Mode -eq 'Exclude') { $modeText = 'Uitsluiten' } else { $modeText = 'Opnemen' }
            [pscustomobject]@{ ModeText = $modeText; Target = $a.Name; Filter = $a.FilterName; FilterMode = $a.FilterMode; Intent = $a.Intent }
        })
        if ($arows.Count -eq 0) { $arows = @([pscustomobject]@{ ModeText = ''; Target = '(niet toegewezen)'; Filter = ''; FilterMode = ''; Intent = '' }) }
        Add-GridTab $tabs ('Assignments (' + @($Item.Assignments).Count + ')') $arows $script:ColSets.assignments

        # 2. Algemene eigenschappen
        $props = @(
            [pscustomobject]@{ Eigenschap = 'Naam';           Waarde = $Item.Name }
            [pscustomobject]@{ Eigenschap = 'Type';           Waarde = $Item.Type }
            [pscustomobject]@{ Eigenschap = 'Subtype';        Waarde = $Item.Subtype }
            [pscustomobject]@{ Eigenschap = 'Platform';       Waarde = $Item.Platform }
            [pscustomobject]@{ Eigenschap = 'Beschrijving';   Waarde = $Item.Description }
            [pscustomobject]@{ Eigenschap = 'Laatst gewijzigd'; Waarde = $Item.Modified }
            [pscustomobject]@{ Eigenschap = 'Id';             Waarde = $Item.Id }
            [pscustomobject]@{ Eigenschap = 'Graph-bron';     Waarde = ($Item.Root + '/' + $Item.Path) }
        )
        Add-GridTab $tabs 'Eigenschappen' $props $script:ColSets.kv

        # 3. Instellingen of scriptinhoud (apart afgevangen, zodat een fout de rest niet blokkeert)
        try {
            switch ($Item.SrcKey) {
                'sc' {
                    $rows = @(Get-ConfigPolicySettings -Id $Item.Id)
                    Add-GridTab $tabs ('Instellingen (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                'intents' {
                    $rows = @(Get-IntentSettings -Id $Item.Id)
                    Add-GridTab $tabs ('Instellingen (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                'gpo' {
                    $rows = @(Get-GpoSettings -Id $Item.Id)
                    Add-GridTab $tabs ('Instellingen (' + $rows.Count + ')') $rows $script:ColSets.settings $true
                }
                { $_ -eq 'scripts' -or $_ -eq 'shell' -or $_ -eq 'health' } {
                    $texts = Get-ScriptTexts -Item $Item
                    foreach ($k in $texts.Keys) { Add-TextTab $tabs $k ([string]$texts[$k]) }
                }
                default {
                    $rows = @(Get-ItemProperties -Item $Item)
                    Add-GridTab $tabs ('Instellingen (' + $rows.Count + ')') $rows $script:ColSets.kv $true
                }
            }
        }
        catch {
            Add-TextTab $tabs 'Fout bij ophalen' ('De instellingen konden niet worden opgehaald:' + [Environment]::NewLine + (Get-ErrorText $_))
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
        $dw  = New-DetailWindow -Title ('Apparaat - ' + $Device.Name) -Heading $Device.Name -Sub $sub

        try {
            $states = @(Get-DeviceStateRows -DeviceId $Device.Id)
            Add-GridTab $dw.Tabs ('Policy-status (' + $states.Count + ')') $states $script:ColSets.states $true
        }
        catch { Add-TextTab $dw.Tabs 'Policy-status' ('Status kon niet worden opgehaald:' + [Environment]::NewLine + (Get-ErrorText $_)) }

        try {
            $obj = Invoke-Graph -Uri ($script:GraphBeta + '/deviceManagement/managedDevices/' + $Device.Id)
            $out = New-Object System.Collections.Generic.List[object]
            ConvertTo-FlatProperties -Obj $obj -Prefix '' -Out $out
            Add-GridTab $dw.Tabs 'Eigenschappen' @($out.ToArray()) $script:ColSets.kv $true
        }
        catch { Add-TextTab $dw.Tabs 'Eigenschappen' ('Eigenschappen konden niet worden opgehaald:' + [Environment]::NewLine + (Get-ErrorText $_)) }

        $dw.Window.Show()
    }
    catch { Show-Error (Get-ErrorText $_) }
    finally { $window.Cursor = $null }
}


# =============================================================================
# 8. EXPORT
# =============================================================================

# Vraagt waar het bestand heen moet. Geeft $null bij annuleren.
function Select-ExportPath {
    param([string]$Name)

    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Excel-werkmap (*.xlsx)|*.xlsx|CSV, puntkomma (*.csv)|*.csv|HTML-rapport (*.html)|*.html'
    $safe = [regex]::Replace($Name, '[^A-Za-z0-9]+', '_')
    $dlg.FileName = '{0}_{1}' -f $safe, (Get-Date -Format 'yyyyMMdd_HHmm')
    $dlg.InitialDirectory = $script:Root
    if (-not $dlg.ShowDialog($window)) { return $null }
    return $dlg.FileName
}

function Open-ExportedFile {
    param([string]$Path)
    $ans = [System.Windows.MessageBox]::Show(('Opgeslagen:' + [Environment]::NewLine + $Path + [Environment]::NewLine + [Environment]::NewLine + 'Nu openen?'), 'Export gereed', 'YesNo', 'Question')
    if ($ans -eq 'Yes') { Start-Process -FilePath $Path }
}

# Exporteert 1 tabel (de zichtbare regels).
function Invoke-ExportDialog {
    param($Cols, $Rows, [string]$Title)

    $data = @($Rows)
    if ($data.Count -eq 0) { Show-Info 'Er zijn geen regels om te exporteren.'; return }
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
    catch { Show-Error ('Export mislukt: ' + (Get-ErrorText $_)); return }
    finally { $window.Cursor = $null }

    Set-Status ('Geexporteerd: {0} regels naar {1}' -f $data.Count, $path)
    Open-ExportedFile $path
}

function Export-CurrentTable {
    $v = $script:CurrentView
    if ($null -eq $v) { return }
    Invoke-ExportDialog -Cols $script:CurrentCols -Rows $script:VisibleRows -Title ([string]$v.Title)
}

# Volledig rapport vanaf het dashboard: Excel (meerdere bladen), HTML of CSV (plat).
function Export-FullReport {
    if (@($script:AllItems).Count -eq 0) { Show-Info 'Laad eerst de gegevens.'; return }
    $path = Select-ExportPath 'Intune_rapport'
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
                    @{ Label = 'Items totaal';     Value = @($script:AllItems).Count }
                    @{ Label = 'Niet toegewezen';  Value = @($script:AllItems | Where-Object { $_.Status -eq 'Niet toegewezen' }).Count }
                    @{ Label = 'Aandachtspunten';  Value = $attention.Count }
                    @{ Label = 'ASR-regels';       Value = @($script:AsrRows).Count }
                )
                $sections = @(
                    @{ Title = 'Per type';          Cols = $script:ColSets.typesummary; Rows = $typeRows }
                    @{ Title = 'Aandachtspunten';   Cols = $script:ColSets.attention;   Rows = $attention }
                    @{ Title = 'Alle assignments';  Cols = $itemCols;                   Rows = @($script:AllItems) }
                )
                if (@($script:AsrRows).Count -gt 0) { $sections += @{ Title = 'ASR-regels'; Cols = $script:ColSets.asr; Rows = @($script:AsrRows) } }
                Export-HtmlReport -Path $path -Title 'Intune assignments - rapport' -Cards $cards -Sections $sections -Subtitle $TxtAccount.Text
            }
            default {
                $sheets = @(
                    @{ Name = 'Per type';             Cols = $script:ColSets.typesummary; Rows = $typeRows }
                    @{ Name = 'Aandachtspunten';      Cols = $script:ColSets.attention;   Rows = $attention }
                    @{ Name = 'Alle items';           Cols = $itemCols;                   Rows = @($script:AllItems) }
                    @{ Name = 'Assignments (plat)';   Cols = $script:ColSets.flat;        Rows = $flat }
                )
                if (@($script:AsrRows).Count -gt 0) { $sheets += @{ Name = 'ASR-regels'; Cols = $script:ColSets.asr; Rows = @($script:AsrRows) } }
                Export-Xlsx -Path $path -Sheets $sheets
            }
        }
    }
    catch { Show-Error ('Export mislukt: ' + (Get-ErrorText $_)); return }
    finally { $window.Cursor = $null }

    Set-Status ('Rapport opgeslagen: ' + $path)
    Open-ExportedFile $path
}


# =============================================================================
# 9. AANMELDEN EN LADEN
# =============================================================================

function Invoke-Login {
    try {
        Set-Busy $true $false
        Set-Status 'Aanmelden... er opent een inlogvenster.'
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
        $BtnLogin.Content  = 'Account wisselen'

        $missing = @($script:Scopes | Where-Object { @($ctx.Scopes) -notcontains $_ })
        if ($missing.Count -gt 0) {
            Show-Info ('Let op: niet alle gevraagde rechten zijn verleend. Onderdelen die deze rechten nodig hebben kunnen leeg blijven.' + [Environment]::NewLine + [Environment]::NewLine + ($missing -join [Environment]::NewLine))
        }
    }
    catch {
        $script:Connected = $false
        Show-Error ('Aanmelden mislukt: ' + (Get-ErrorText $_))
        Set-Status 'Aanmelden mislukt.'
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
    catch { Show-Error ('Laden mislukt: ' + (Get-ErrorText $_)) }
    finally { Set-Busy $false }

    Update-NavCounts
    if ($script:CurrentView) { Switch-View ([string]$script:CurrentView.Key) }

    if ($script:Cancel) { Set-Status 'Laden geannuleerd; de gegevens zijn mogelijk onvolledig.' }
    if ($script:LoadErrors.Count -gt 0) {
        $first = @($script:LoadErrors | Select-Object -First 8)
        $msg = 'Sommige onderdelen konden niet worden opgehaald (meestal ontbrekende rechten of een niet-gelicentieerde functie):' + [Environment]::NewLine + [Environment]::NewLine + ($first -join [Environment]::NewLine)
        if ($script:LoadErrors.Count -gt 8) { $msg += [Environment]::NewLine + '... en ' + ($script:LoadErrors.Count - 8) + ' meer.' }
        Show-Info $msg 'Waarschuwing'
    }
}
