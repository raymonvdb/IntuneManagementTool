# =============================================================================
# Graph.ps1
# Alles rond Microsoft Graph: aanroepen, bronnen, groepen/filters en het
# ophalen van alle Intune-items met hun assignments.
#
# Deze file wordt met dot-sourcing geladen door IntuneTool.ps1, dus de
# $script:-variabelen hieronder zijn voor het hele programma beschikbaar.
# =============================================================================

$script:GraphBeta = 'https://graph.microsoft.com/beta'
$script:GraphV1   = 'https://graph.microsoft.com/v1.0'

# Caches, zodat we elke groep/filter/template maar 1x ophalen
$script:GroupCache   = @{}   # groupId  -> pscustomobject (Id, Name, Dynamic, Exists)
$script:FilterCache  = @{}   # filterId -> filterobject uit Graph
$script:MemberCounts = @{}   # groupId  -> aantal leden (alleen na "Ledenaantal ophalen")
$script:TemplateMap  = @{}   # templateId -> template (voor endpoint security intents)

# Resultaten
$script:AllItems  = @()      # alle policies/scripts/apps, 1 object per item
$script:AsrRows   = @()      # ASR-regels per policy
$script:Devices   = $null    # beheerde apparaten (wordt pas geladen als je die pagina opent)

$script:LoadErrors = New-Object System.Collections.Generic.List[string]
$script:Cancel     = $false
$script:Connected  = $false


# =============================================================================
# 1. BASISFUNCTIES
# =============================================================================

function Get-ErrorText {
    param($ErrorRecord)
    $m = [string]$ErrorRecord.Exception.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $m += ' ' + [string]$ErrorRecord.ErrorDetails.Message
    }
    return $m
}

# Eenmalige Graph-aanroep met automatische retry bij throttling (429) en
# tijdelijke serverfouten (503/504).
function Invoke-Graph {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Method = 'GET',
        [string]$Body = '',
        [hashtable]$Headers = $null
    )

    $attempt = 0
    while ($true) {
        try {
            $callArgs = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject' }
            if ($Body)    { $callArgs.Body = $Body; $callArgs.ContentType = 'application/json' }
            if ($Headers) { $callArgs.Headers = $Headers }
            return (Invoke-MgGraphRequest @callArgs)
        }
        catch {
            $attempt++
            $text = Get-ErrorText $_
            if ($attempt -le 4 -and $text -match '429|TooManyRequests|503|504|ServiceUnavailable|GatewayTimeout') {
                Start-Sleep -Seconds ([int][math]::Pow(2, $attempt))
                continue
            }
            throw
        }
    }
}

# Haalt alle pagina's van een Graph-lijst op ('@odata.nextLink' volgen).
function Invoke-GraphAll {
    param([Parameter(Mandatory = $true)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next  = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        if ($resp.value) {
            foreach ($v in $resp.value) { $items.Add($v) }
        }
        $next = $resp.'@odata.nextLink'
    }
    return $items.ToArray()
}

function Format-GraphDate {
    param($Value)
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') } catch { return [string]$Value }
}

function Split-Camel {
    param([string]$Text)
    return [regex]::Replace($Text, '(?<=[a-z0-9])(?=[A-Z])', ' ')
}

function Format-FirstUpper {
    param([string]$Text)
    if (-not $Text) { return '' }
    return $Text.Substring(0, 1).ToUpperInvariant() + $Text.Substring(1)
}


# =============================================================================
# 2. LEESBARE NAMEN: platform, categorie, type
# =============================================================================

# Platform bepalen uit een @odata.type, bijvoorbeeld '#microsoft.graph.iosVppApp'
function Get-PlatformFromOData {
    param([string]$ODataType)
    $x = $ODataType.ToLowerInvariant()
    if ($x -match 'macos')                                                   { return 'macOS' }
    if ($x -match 'ios|ipados')                                              { return 'iOS/iPadOS' }
    if ($x -match 'android|aosp')                                            { return 'Android' }
    if ($x -match 'windows|win32|winget|officesuite|microsoftstore|microsoftedge|sharedpc|editionupgrade') { return 'Windows' }
    if ($x -match 'webapp|weblink')                                          { return 'Web' }
    return ''
}

# Platform uit een tekstwaarde zoals 'windows10', 'macOS', 'androidEnterprise'
function Get-PlatformLabel {
    param([string]$Value)
    if (-not $Value) { return '' }
    $x = $Value.ToLowerInvariant()
    if ($x -match 'windows')     { return 'Windows' }
    if ($x -match 'macos')       { return 'macOS' }
    if ($x -match 'ios|ipados')  { return 'iOS/iPadOS' }
    if ($x -match 'android|aosp'){ return 'Android' }
    if ($x -match 'linux')       { return 'Linux' }
    return $Value
}

# Categorie voor endpoint security, op basis van templateFamily / templateType
function Get-CategoryFromText {
    param([string]$Text)
    $x = $Text.ToLowerInvariant()
    if ($x -match 'attacksurface')        { return 'ASR' }
    if ($x -match 'antivirus')            { return 'Antivirus' }
    if ($x -match 'diskencryption')       { return 'Disk encryption' }
    if ($x -match 'firewall')             { return 'Firewall' }
    if ($x -match 'endpointdetection')    { return 'EDR' }
    if ($x -match 'accountprotection')    { return 'Account protection' }
    if ($x -match 'applicationcontrol')   { return 'Application control' }
    if ($x -match 'privilegemanagement')  { return 'Endpoint privilege management' }
    if ($x -match 'baseline')             { return 'Security baseline' }
    return ''
}

function Get-FriendlyOData {
    param([string]$ODataType)

    $n = $ODataType -replace '^#?microsoft\.graph\.', ''
    $map = @{
        windowsUpdateForBusinessConfiguration           = 'Update ring'
        windows10CustomConfiguration                    = 'Custom (OMA-URI)'
        windows10GeneralConfiguration                   = 'Device restrictions'
        windows10EndpointProtectionConfiguration        = 'Endpoint protection'
        windows10DeviceFirmwareConfigurationInterface   = 'Device firmware (DFCI)'
        windowsDeliveryOptimizationConfiguration        = 'Delivery optimization'
        windowsHealthMonitoringConfiguration            = 'Health monitoring'
        windowsIdentityProtectionConfiguration          = 'Identity protection'
        windows10VpnConfiguration                       = 'VPN'
        windowsWifiConfiguration                        = 'Wi-Fi'
        windowsWifiEnterpriseEAPConfiguration           = 'Wi-Fi (enterprise)'
        windows10SecureAssessmentConfiguration          = 'Secure assessment (onderwijs)'
        windows10TeamGeneralConfiguration               = 'Surface Hub'
        windows10NetworkBoundaryConfiguration           = 'Network boundary'
        windows10PkcsCertificateProfile                 = 'PKCS-certificaat'
        windows10ImportedPFXCertificateProfile          = 'PFX-certificaat (geimporteerd)'
        windowsDomainJoinConfiguration                  = 'Domain join'
        windowsKioskConfiguration                       = 'Kiosk'
        sharedPCConfiguration                           = 'Shared multi-user device'
        editionUpgradeConfiguration                     = 'Edition upgrade'
        windowsDefenderAdvancedThreatProtectionConfiguration = 'Defender for Endpoint onboarding'
        macOSCustomConfiguration                        = 'Custom (macOS)'
        macOSGeneralDeviceConfiguration                 = 'Device restrictions'
        macOSEndpointProtectionConfiguration            = 'Endpoint protection'
        iosGeneralDeviceConfiguration                   = 'Device restrictions'
        iosCustomConfiguration                          = 'Custom (iOS)'
        androidGeneralDeviceConfiguration               = 'Device restrictions'
        androidDeviceOwnerGeneralDeviceConfiguration    = 'Device restrictions (fully managed)'
        androidWorkProfileGeneralDeviceConfiguration    = 'Device restrictions (work profile)'
        androidCustomConfiguration                      = 'Custom (Android)'
        azureADWindowsAutopilotDeploymentProfile        = 'Entra join'
        activeDirectoryWindowsAutopilotDeploymentProfile = 'Hybrid Entra join'
        windows10EnrollmentCompletionPageConfiguration  = 'Enrollment Status Page (ESP)'
        deviceEnrollmentLimitConfiguration              = 'Device limit'
        deviceEnrollmentPlatformRestrictionsConfiguration = 'Platform restrictions (alle platformen)'
        deviceEnrollmentPlatformRestrictionConfiguration  = 'Platform restriction'
        deviceEnrollmentWindowsHelloForBusinessConfiguration = 'Windows Hello for Business'
        deviceEnrollmentNotificationConfiguration       = 'Enrollment notifications'
    }
    if ($map.ContainsKey($n)) { return $map[$n] }

    $t = $n -replace '(DeviceConfiguration|Configuration|CompliancePolicy|Policy|Profile)$', ''
    return (Format-FirstUpper (Split-Camel $t))
}

function Get-FriendlyApp {
    param([string]$ODataType)

    $n = $ODataType -replace '^#?microsoft\.graph\.', ''
    $map = @{
        win32LobApp                 = 'Windows app (Win32)'
        windowsMobileMSI            = 'Windows MSI (LOB)'
        microsoftStoreForBusinessApp= 'Microsoft Store (zakelijk)'
        winGetApp                   = 'Microsoft Store (winget)'
        officeSuiteApp              = 'Microsoft 365 Apps'
        windowsUniversalAppX        = 'Windows universal app'
        windowsAppX                 = 'Windows AppX'
        windowsWebApp               = 'Windows webkoppeling'
        webApp                      = 'Webkoppeling'
        windowsMicrosoftEdgeApp     = 'Microsoft Edge'
        iosVppApp                   = 'iOS VPP-app'
        iosStoreApp                 = 'iOS store-app'
        iosLobApp                   = 'iOS LOB-app'
        managedIOSStoreApp          = 'iOS store-app (beheerd)'
        managedIOSLobApp            = 'iOS LOB-app (beheerd)'
        iosiPadOSWebClip            = 'iOS web clip'
        managedAndroidStoreApp      = 'Android store-app (beheerd)'
        managedAndroidLobApp        = 'Android LOB-app'
        androidManagedStoreApp      = 'Managed Google Play-app'
        androidStoreApp             = 'Android store-app'
        androidForWorkApp           = 'Android Enterprise-app'
        macOSDmgApp                 = 'macOS DMG'
        macOSPkgApp                 = 'macOS PKG'
        macOSLobApp                 = 'macOS LOB-app'
        macOSMicrosoftEdgeApp       = 'Microsoft Edge (macOS)'
        macOSOfficeSuiteApp         = 'Microsoft 365 Apps (macOS)'
        macOSMicrosoftDefenderApp   = 'Microsoft Defender (macOS)'
    }
    if ($map.ContainsKey($n)) { return $map[$n] }
    return (Get-FriendlyOData $ODataType)
}

function Get-IntentText {
    param([string]$Intent)
    switch ($Intent.ToLowerInvariant()) {
        'required'                  { return 'Required' }
        'available'                 { return 'Available' }
        'uninstall'                 { return 'Uninstall' }
        'availablewithoutenrollment'{ return 'Available (zonder enrollment)' }
        default                     { return $Intent }
    }
}


# =============================================================================
# 3. BRONNEN: welke Graph-endpoints lezen we uit?
# =============================================================================
# Per bron:
#   Key  : korte sleutel, wordt gebruikt in de rest van de code
#   Root : 'deviceManagement' of 'deviceAppManagement'
#   Path : collectie onder /beta/<Root>/
#   Type : weergavenaam (tekst of scriptblock dat $p ontvangt)
#   Sub  : subtype (tekst of scriptblock)
#   Plat : platform (tekst of scriptblock)
#   Cat  : categorie voor endpoint security (tekst of scriptblock)
# Een nieuwe bron toevoegen = hier 1 blok bijplakken.

$script:Sources = @(
    [pscustomobject]@{
        Key = 'conf'; Root = 'deviceManagement'; Path = 'deviceConfigurations'
        Type = { param($p) if (([string]$p.'@odata.type') -match 'windowsUpdateForBusiness') { 'Windows update' } else { 'Device configuration' } }
        Sub  = { param($p) Get-FriendlyOData ([string]$p.'@odata.type') }
        Plat = { param($p) Get-PlatformFromOData ([string]$p.'@odata.type') }
        Cat  = ''
    }
    [pscustomobject]@{
        Key = 'gpo'; Root = 'deviceManagement'; Path = 'groupPolicyConfigurations'
        Type = 'Device configuration'; Sub = 'Administrative templates'; Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'comp'; Root = 'deviceManagement'; Path = 'deviceCompliancePolicies'
        Type = 'Compliance'; Sub = ''
        Plat = { param($p) Get-PlatformFromOData ([string]$p.'@odata.type') }
        Cat  = ''
    }
    [pscustomobject]@{
        Key = 'sc'; Root = 'deviceManagement'; Path = 'configurationPolicies'
        Type = {
            param($p)
            $fam = [string]$p.templateReference.templateFamily
            if ($fam -match 'endpointSecurity|baseline') { 'Endpoint security' } else { 'Settings catalog' }
        }
        Sub  = { param($p) [string]$p.templateReference.templateDisplayName }
        Plat = { param($p) Get-PlatformLabel ([string]$p.platforms) }
        Cat  = { param($p) Get-CategoryFromText ([string]$p.templateReference.templateFamily) }
    }
    [pscustomobject]@{
        Key = 'intents'; Root = 'deviceManagement'; Path = 'intents'
        Type = 'Endpoint security'
        Sub  = { param($p) $t = $script:TemplateMap[[string]$p.templateId]; if ($t) { [string]$t.displayName } else { '' } }
        Plat = { param($p) $t = $script:TemplateMap[[string]$p.templateId]; if ($t) { Get-PlatformLabel ([string]$t.platformType) } else { '' } }
        Cat  = { param($p) $t = $script:TemplateMap[[string]$p.templateId]; if ($t) { Get-CategoryFromText (([string]$t.templateType) + ' ' + ([string]$t.templateSubtype)) } else { '' } }
    }
    [pscustomobject]@{
        Key = 'scripts'; Root = 'deviceManagement'; Path = 'deviceManagementScripts'
        Type = 'Platform script'; Sub = 'PowerShell (Windows)'; Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'shell'; Root = 'deviceManagement'; Path = 'deviceShellScripts'
        Type = 'Platform script'; Sub = 'Shell (macOS)'; Plat = 'macOS'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'health'; Root = 'deviceManagement'; Path = 'deviceHealthScripts'
        Type = 'Remediation'
        Sub  = { param($p) if ($p.isGlobalScript) { 'Microsoft (ingebouwd)' } else { 'Custom' } }
        Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'apps'; Root = 'deviceAppManagement'; Path = 'mobileApps'
        Type = 'App'
        Sub  = { param($p) Get-FriendlyApp ([string]$p.'@odata.type') }
        Plat = { param($p) Get-PlatformFromOData ([string]$p.'@odata.type') }
        Cat  = ''
    }
    [pscustomobject]@{
        Key = 'appios'; Root = 'deviceAppManagement'; Path = 'iosManagedAppProtections'
        Type = 'App protection'; Sub = 'iOS/iPadOS'; Plat = 'iOS/iPadOS'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'appand'; Root = 'deviceAppManagement'; Path = 'androidManagedAppProtections'
        Type = 'App protection'; Sub = 'Android'; Plat = 'Android'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'appwin'; Root = 'deviceAppManagement'; Path = 'windowsManagedAppProtections'
        Type = 'App protection'; Sub = 'Windows'; Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'appcfgmam'; Root = 'deviceAppManagement'; Path = 'targetedManagedAppConfigurations'
        Type = 'App configuration'; Sub = 'Apps (MAM)'; Plat = ''; Cat = ''
    }
    [pscustomobject]@{
        Key = 'appcfgmdm'; Root = 'deviceAppManagement'; Path = 'mobileAppConfigurations'
        Type = 'App configuration'; Sub = 'Beheerde apparaten (MDM)'
        Plat = { param($p) Get-PlatformFromOData ([string]$p.'@odata.type') }
        Cat  = ''
    }
    [pscustomobject]@{
        Key = 'autopilot'; Root = 'deviceManagement'; Path = 'windowsAutopilotDeploymentProfiles'
        Type = 'Autopilot'
        Sub  = { param($p) Get-FriendlyOData ([string]$p.'@odata.type') }
        Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'enroll'; Root = 'deviceManagement'; Path = 'deviceEnrollmentConfigurations'
        Type = 'Enrollment'
        Sub  = { param($p) Get-FriendlyOData ([string]$p.'@odata.type') }
        Plat = { param($p) Get-PlatformFromOData ([string]$p.'@odata.type') }
        Cat  = ''
    }
    [pscustomobject]@{
        Key = 'feat'; Root = 'deviceManagement'; Path = 'windowsFeatureUpdateProfiles'
        Type = 'Windows update'; Sub = 'Feature update'; Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'qual'; Root = 'deviceManagement'; Path = 'windowsQualityUpdateProfiles'
        Type = 'Windows update'; Sub = 'Quality update'; Plat = 'Windows'; Cat = ''
    }
    [pscustomobject]@{
        Key = 'drv'; Root = 'deviceManagement'; Path = 'windowsDriverUpdateProfiles'
        Type = 'Windows update'; Sub = 'Driver update'; Plat = 'Windows'; Cat = ''
    }
)

# Waarde van een bronveld: tekst of scriptblock dat het item ontvangt.
function Get-Field {
    param($Field, $P)
    if ($Field -is [scriptblock]) {
        $r = & $Field $P
        return [string](@($r)[0])
    }
    return [string]$Field
}


# =============================================================================
# 4. GROEPEN EN FILTERS
# =============================================================================

function Add-GroupToCache {
    param($G)
    $script:GroupCache[[string]$G.id] = [pscustomobject]@{
        Id      = [string]$G.id
        Name    = [string]$G.displayName
        Dynamic = (@($G.groupTypes) -contains 'DynamicMembership')
        Exists  = $true
    }
}

# Zoekt groepsnamen in bulk op (max 100 per aanroep) en valt per groep terug op
# een losse aanroep. Groepen die niet bestaan krijgen Exists = $false.
function Resolve-GroupIds {
    param([string[]]$Ids)

    $todo = @($Ids | Where-Object { $_ -and -not $script:GroupCache.ContainsKey($_) } | Select-Object -Unique)
    for ($i = 0; $i -lt $todo.Count; $i += 100) {
        $last  = [math]::Min($i + 99, $todo.Count - 1)
        $chunk = @($todo[$i..$last])
        try {
            $body = @{ ids = $chunk; types = @('group') } | ConvertTo-Json -Depth 3 -Compress
            $r = Invoke-Graph -Uri ($script:GraphV1 + '/directoryObjects/getByIds') -Method 'POST' -Body $body
            foreach ($g in @($r.value)) { Add-GroupToCache $g }
        } catch { }
    }

    foreach ($id in $todo) {
        if ($script:GroupCache.ContainsKey($id)) { continue }
        try {
            $g = Invoke-Graph -Uri ($script:GraphV1 + '/groups/' + $id + '?$select=id,displayName,groupTypes')
            Add-GroupToCache $g
        } catch {
            $text = Get-ErrorText $_
            if ($text -match '404|NotFound|ResourceNotFound') {
                $script:GroupCache[$id] = [pscustomobject]@{ Id = $id; Name = '(verwijderde groep ' + $id + ')'; Dynamic = $false; Exists = $false }
            } else {
                $script:GroupCache[$id] = [pscustomobject]@{ Id = $id; Name = '(onbekende groep ' + $id + ')'; Dynamic = $false; Exists = $true }
            }
        }
    }
}

function Get-GroupDisplay {
    param([string]$Id)
    if (-not $script:GroupCache.ContainsKey($Id)) { Resolve-GroupIds @($Id) }
    return [string]$script:GroupCache[$Id].Name
}

function Initialize-Filters {
    $script:FilterCache = @{}
    foreach ($f in @(Invoke-GraphAll -Uri ($script:GraphBeta + '/deviceManagement/assignmentFilters'))) {
        $script:FilterCache[[string]$f.id] = $f
    }
}

function Get-FilterDisplay {
    param([string]$Id)
    if ($script:FilterCache.ContainsKey($Id)) { return [string]$script:FilterCache[$Id].displayName }
    return '(onbekend filter ' + $Id + ')'
}

# Templates van endpoint security intents (geeft naam, categorie en platform)
function Initialize-Templates {
    $script:TemplateMap = @{}
    foreach ($t in @(Invoke-GraphAll -Uri ($script:GraphBeta + '/deviceManagement/templates'))) {
        $script:TemplateMap[[string]$t.id] = $t
    }
}

function Get-GroupMemberCount {
    param([string]$Id)
    $uri = $script:GraphV1 + '/groups/' + $Id + '/transitiveMembers/$count'
    $r = Invoke-Graph -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' }
    return [int](([string]$r).Trim())
}


# =============================================================================
# 5. ASSIGNMENTS OMZETTEN
# =============================================================================

# Zet 1 ruwe assignment om naar een bruikbaar object.
# Kind: group | allDevices | allUsers | collection | other
# Mode: Include | Exclude
function ConvertTo-AssignmentInfo {
    param($Assignment)

    $t     = $Assignment.target
    $odata = [string]$t.'@odata.type'
    $kind  = 'other'
    $mode  = 'Include'
    $gid   = ''
    $name  = $odata

    # Let op de volgorde: 'exclusionGroupAssignmentTarget' eindigt ook op 'groupAssignmentTarget'
    if     ($odata -like '*exclusionGroupAssignmentTarget')  { $kind = 'group'; $mode = 'Exclude'; $gid = [string]$t.groupId }
    elseif ($odata -like '*groupAssignmentTarget')           { $kind = 'group'; $gid = [string]$t.groupId }
    elseif ($odata -like '*allDevicesAssignmentTarget')      { $kind = 'allDevices'; $name = 'Alle apparaten' }
    elseif ($odata -like '*allLicensedUsersAssignmentTarget'){ $kind = 'allUsers';   $name = 'Alle gebruikers' }
    elseif ($odata -like '*collectionAssignmentTarget')      { $kind = 'collection'; $name = 'Configuration Manager-collectie ' + [string]$t.collectionId }

    $groupExists = $true
    if ($kind -eq 'group') {
        $name = Get-GroupDisplay $gid
        $gi = $script:GroupCache[$gid]
        if ($gi -and $gi.Exists -eq $false) { $groupExists = $false }
    }

    # Assignment filter (alleen tonen als er echt een filter actief is)
    $fid   = [string]$t.deviceAndAppManagementAssignmentFilterId
    $fmode = [string]$t.deviceAndAppManagementAssignmentFilterType
    $fname = ''
    if ($fid -and $fid -ne '00000000-0000-0000-0000-000000000000' -and $fmode -and $fmode -ne 'none') {
        $fname = Get-FilterDisplay $fid
    } else {
        $fid = ''; $fmode = ''
    }

    # Intent (alleen apps): required / available / uninstall
    $intent = ''
    if ($Assignment.intent) { $intent = Get-IntentText ([string]$Assignment.intent) }

    $extra = ''
    if ($fname)  { $extra += ' [filter: {0} ({1})]' -f $fname, $fmode }
    if ($intent) { $extra += ' - ' + $intent }

    return [pscustomobject]@{
        Kind        = $kind
        Mode        = $mode
        GroupId     = $gid
        Name        = $name
        FilterId    = $fid
        FilterName  = $fname
        FilterMode  = $fmode
        Intent      = $intent
        Label       = $name + $extra
        GroupExists = $groupExists
    }
}

# Bouwt het uiteindelijke item-object (1 regel in de tabel).
function New-ItemObject {
    param($Src, $P, $RawAssignments)

    $infos = @(foreach ($a in $RawAssignments) { ConvertTo-AssignmentInfo -Assignment $a })

    $sortKey = @{ Expression = { if ($_.Kind -eq 'group') { 1 } else { 0 } } }
    $incInfos = @($infos | Where-Object { $_.Mode -eq 'Include' } | Sort-Object $sortKey, Label)
    $excInfos = @($infos | Where-Object { $_.Mode -eq 'Exclude' } | Sort-Object $sortKey, Label)
    $inc = @($incInfos | ForEach-Object { $_.Label })
    $exc = @($excInfos | ForEach-Object { $_.Label })

    if ($infos.Count -eq 0)      { $status = 'Niet toegewezen' }
    elseif ($incInfos.Count -eq 0) { $status = 'Alleen uitsluiting' }
    else                         { $status = 'Toegewezen' }

    $name = [string]$P.displayName
    if (-not $name) { $name = [string]$P.name }

    return [pscustomobject]@{
        Name        = $name
        Type        = Get-Field $Src.Type $P
        Subtype     = Get-Field $Src.Sub  $P
        Platform    = Get-Field $Src.Plat $P
        Category    = Get-Field $Src.Cat  $P
        Include     = ($inc -join "`n")
        Exclude     = ($exc -join "`n")
        Status      = $status
        Modified    = Format-GraphDate $P.lastModifiedDateTime
        Description = [string]$P.description
        Id          = [string]$P.id
        SrcKey      = $Src.Key
        Root        = $Src.Root
        Path        = $Src.Path
        Assignments = $infos
        Raw         = $P
    }
}


# =============================================================================
# 6. ALLES OPHALEN
# =============================================================================

# Haalt 1 bron op. Probeert eerst alles in 1 keer met $expand=assignments
# (snel). Lukt dat niet, dan wordt per item een aparte assignments-call gedaan.
function Get-SourceRaw {
    param($Src, $Into)

    $base     = '{0}/{1}/{2}' -f $script:GraphBeta, $Src.Root, $Src.Path
    $expanded = $true
    try {
        $policies = @(Invoke-GraphAll -Uri ($base + '?$expand=assignments'))
    } catch {
        $expanded = $false
        $policies = @(Invoke-GraphAll -Uri $base)
    }

    $i = 0
    foreach ($p in $policies) {
        if ($script:Cancel) { break }
        $i++
        if (($i % 15) -eq 0) { Set-Status ('{0}: {1}/{2}' -f $Src.Path, $i, $policies.Count) }

        $names = @($p.PSObject.Properties.Name)
        if ($expanded -and ($names -contains 'assignments')) {
            $assignments = @($p.assignments | Where-Object { $null -ne $_ })
        }
        elseif (($names -contains 'isAssigned') -and ($p.isAssigned -eq $false)) {
            $assignments = @()
        }
        else {
            try {
                $assignments = @(Invoke-GraphAll -Uri ('{0}/{1}/assignments' -f $base, $p.id))
            } catch {
                $assignments = @()
                $script:LoadErrors.Add(('{0}/{1}: assignments niet op te halen' -f $Src.Path, $p.id))
            }
        }
        $Into.Add([pscustomobject]@{ Src = $Src; P = $p; Assignments = $assignments })
    }
}

# Hoofdfunctie: haalt alle bronnen op en bouwt $script:AllItems.
function Get-AllItems {
    $script:LoadErrors.Clear()
    $script:Cancel   = $false
    $script:AsrRows  = @()
    $script:Devices  = $null
    $script:GroupCache   = @{}
    $script:MemberCounts = @{}

    $raw = New-Object System.Collections.Generic.List[object]
    $n   = $script:Sources.Count + 3

    Set-Progress 0 $n
    Set-Status 'Templates en assignment filters ophalen...'
    try { Initialize-Templates } catch { $script:LoadErrors.Add('templates: ' + (Get-ErrorText $_)) }
    try { Initialize-Filters }   catch { $script:LoadErrors.Add('assignmentFilters: ' + (Get-ErrorText $_)) }

    $i = 1
    foreach ($src in $script:Sources) {
        if ($script:Cancel) { break }
        Set-Progress $i $n
        Set-Status ('Ophalen ({0}/{1}): {2}' -f ($i - 1), $script:Sources.Count, $src.Path)
        try { Get-SourceRaw -Src $src -Into $raw }
        catch { $script:LoadErrors.Add(('{0}: {1}' -f $src.Path, (Get-ErrorText $_))) }
        $i++
    }

    # Groepsnamen in bulk opzoeken
    Set-Progress ($n - 2) $n
    Set-Status 'Groepsnamen opzoeken...'
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($r in $raw) {
        foreach ($a in $r.Assignments) {
            $gid = [string]$a.target.groupId
            if ($gid) { $ids.Add($gid) }
        }
    }
    Resolve-GroupIds @($ids.ToArray() | Select-Object -Unique)

    # Items bouwen
    Set-Progress ($n - 1) $n
    Set-Status 'Overzicht samenstellen...'
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($r in $raw) { $items.Add((New-ItemObject -Src $r.Src -P $r.P -RawAssignments $r.Assignments)) }
    $script:AllItems = @($items.ToArray() | Sort-Object Type, Name)

    # ASR-regels (klein aantal extra calls)
    if (-not $script:Cancel) {
        Set-Progress $n $n
        try { $script:AsrRows = @(Get-AsrRows -ScanAllCatalog $false) }
        catch { $script:LoadErrors.Add('ASR: ' + (Get-ErrorText $_)) }
    }

    Set-Status ('Klaar: {0} items geladen.' -f $script:AllItems.Count)
}


# =============================================================================
# 7. APPARATEN
# =============================================================================

function Get-DeviceRows {
    $sel = 'id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,model,manufacturer,serialNumber,lastSyncDateTime,isEncrypted'
    $uri = $script:GraphBeta + '/deviceManagement/managedDevices?$select=' + $sel
    $list = @(Invoke-GraphAll -Uri $uri)
    return @($list | ForEach-Object {
        [pscustomobject]@{
            Name       = [string]$_.deviceName
            User       = [string]$_.userPrincipalName
            OS         = [string]$_.operatingSystem
            Version    = [string]$_.osVersion
            Compliance = [string]$_.complianceState
            Owner      = [string]$_.managedDeviceOwnerType
            Model      = (([string]$_.manufacturer + ' ' + [string]$_.model).Trim())
            Serial     = [string]$_.serialNumber
            Encrypted  = $(if ($_.isEncrypted) { 'Ja' } else { 'Nee' })
            LastSync   = Format-GraphDate $_.lastSyncDateTime
            Id         = [string]$_.id
        }
    })
}

# Status van configuratie- en compliance-policies op 1 apparaat
function Get-DeviceStateRows {
    param([string]$DeviceId)

    $rows = New-Object System.Collections.Generic.List[object]
    $base = $script:GraphBeta + '/deviceManagement/managedDevices/' + $DeviceId

    foreach ($s in @(Invoke-GraphAll -Uri ($base + '/deviceConfigurationStates'))) {
        $rows.Add([pscustomobject]@{ Kind = 'Configuration'; Policy = [string]$s.displayName; Platform = [string]$s.platformType; State = [string]$s.state })
    }
    foreach ($s in @(Invoke-GraphAll -Uri ($base + '/deviceCompliancePolicyStates'))) {
        $rows.Add([pscustomobject]@{ Kind = 'Compliance'; Policy = [string]$s.displayName; Platform = [string]$s.platformType; State = [string]$s.state })
    }
    return @($rows.ToArray() | Sort-Object Kind, Policy)
}


# =============================================================================
# 8. ZOEKEN IN DE DIRECTORY (groepen en gebruikers)
# =============================================================================

function Search-Directory {
    param([string]$Text)

    $q   = $Text.Replace("'", "''")
    $res = New-Object System.Collections.Generic.List[object]
    $err = ''

    try {
        $filter = "startswith(displayName,'" + $q + "')"
        $uri = $script:GraphV1 + '/groups?$filter=' + [uri]::EscapeDataString($filter) + '&$select=id,displayName,groupTypes&$top=25'
        foreach ($g in @((Invoke-Graph -Uri $uri).value)) {
            $res.Add([pscustomobject]@{ Kind = 'group'; Id = [string]$g.id; Name = [string]$g.displayName; Display = ('[Groep] ' + [string]$g.displayName) })
        }
    } catch { $err += 'Groepen: ' + (Get-ErrorText $_) + ' ' }

    try {
        $filter = "startswith(displayName,'" + $q + "') or startswith(userPrincipalName,'" + $q + "') or startswith(mail,'" + $q + "')"
        $uri = $script:GraphV1 + '/users?$filter=' + [uri]::EscapeDataString($filter) + '&$select=id,displayName,userPrincipalName&$top=25'
        foreach ($u in @((Invoke-Graph -Uri $uri).value)) {
            $res.Add([pscustomobject]@{ Kind = 'user'; Id = [string]$u.id; Name = [string]$u.displayName; Display = ('[Gebruiker] ' + [string]$u.displayName + ' (' + [string]$u.userPrincipalName + ')') })
        }
    } catch { $err += 'Gebruikers: ' + (Get-ErrorText $_) }

    return [pscustomobject]@{ Results = $res.ToArray(); Error = $err }
}

# Alle groepen waar een gebruiker of groep (geneste) lid van is.
function Get-ParentGroups {
    param([string]$Kind, [string]$Id)

    if ($Kind -eq 'user') { $seg = 'users' } else { $seg = 'groups' }
    $uri = $script:GraphV1 + '/' + $seg + '/' + $Id + '/transitiveMemberOf/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999'
    $list = @(Invoke-GraphAll -Uri $uri)
    foreach ($g in $list) { Add-GroupToCache $g }
    return $list
}
