# =============================================================================
# Settings.ps1
# Making policy settings readable: Settings catalog, endpoint security
# intents, administrative templates, scripts and ASR rules.
# =============================================================================

# -----------------------------------------------------------------------------
# Making names and values friendly
# -----------------------------------------------------------------------------

# Falls back to a cleaned-up definition ID when Graph gives no displayName.
function Get-PrettyId {
    param([string]$Id)
    $s = $Id -replace '^(device|user)_vendor_msft_(policy_config_)?', ''
    $s = $s -replace '_', ' '
    return $s
}

function Get-SettingName {
    param([string]$Id, $Defs)
    if ($Defs.ContainsKey($Id) -and $Defs[$Id].displayName) { return [string]$Defs[$Id].displayName }
    return (Get-PrettyId $Id)
}

# Translates the raw choice value (e.g. '..._block') into the display name (e.g. 'Block').
function Get-ChoiceText {
    param([string]$DefId, [string]$Value, $Defs)

    if ($Defs.ContainsKey($DefId)) {
        $opt = @($Defs[$DefId].options) | Where-Object { $_.itemId -eq $Value } | Select-Object -First 1
        if ($opt -and $opt.displayName) { return [string]$opt.displayName }
    }
    if ($Value.StartsWith($DefId + '_', [System.StringComparison]::OrdinalIgnoreCase)) {
        $tail = $Value.Substring($DefId.Length + 1)
        return (Format-FirstUpper $tail)
    }
    return $Value
}


# -----------------------------------------------------------------------------
# Settings catalog (configurationPolicies)
# -----------------------------------------------------------------------------

# Recursively walks a settingInstance and adds one row per setting.
# Level determines the indentation of nested settings.
function Add-SettingRows {
    param($Instance, $Defs, $Out, [int]$Level = 0)

    if ($null -eq $Instance) { return }

    $type = [string]$Instance.'@odata.type'
    $id   = [string]$Instance.settingDefinitionId
    $name = Get-SettingName $id $Defs
    $pad  = ('    ' * $Level)

    if ($type -like '*ChoiceSettingCollectionInstance') {
        $vals  = @($Instance.choiceSettingCollectionValue | Where-Object { $null -ne $_ })
        $texts = @($vals | ForEach-Object { Get-ChoiceText $id ([string]$_.value) $Defs })
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = ($texts -join '; ') })
        foreach ($v in $vals) {
            foreach ($c in @($v.children | Where-Object { $null -ne $_ })) { Add-SettingRows -Instance $c -Defs $Defs -Out $Out -Level ($Level + 1) }
        }
    }
    elseif ($type -like '*ChoiceSettingInstance') {
        $v = $Instance.choiceSettingValue
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = (Get-ChoiceText $id ([string]$v.value) $Defs) })
        foreach ($c in @($v.children | Where-Object { $null -ne $_ })) { Add-SettingRows -Instance $c -Defs $Defs -Out $Out -Level ($Level + 1) }
    }
    elseif ($type -like '*SimpleSettingCollectionInstance') {
        $vals = @($Instance.simpleSettingCollectionValue | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_.value })
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = ($vals -join '; ') })
    }
    elseif ($type -like '*SimpleSettingInstance') {
        $v = $Instance.simpleSettingValue
        if (([string]$v.'@odata.type') -like '*Secret*') { $text = '(encrypted)' } else { $text = [string]$v.value }
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = $text })
    }
    elseif ($type -like '*GroupSettingCollectionInstance') {
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = '' })
        foreach ($item in @($Instance.groupSettingCollectionValue | Where-Object { $null -ne $_ })) {
            foreach ($c in @($item.children | Where-Object { $null -ne $_ })) { Add-SettingRows -Instance $c -Defs $Defs -Out $Out -Level ($Level + 1) }
        }
    }
    elseif ($type -like '*GroupSettingInstance') {
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = '' })
        foreach ($c in @($Instance.groupSettingValue.children | Where-Object { $null -ne $_ })) {
            Add-SettingRows -Instance $c -Defs $Defs -Out $Out -Level ($Level + 1)
        }
    }
    else {
        $Out.Add([pscustomobject]@{ DefinitionId = $id; Name = $name; Level = $Level; Setting = $pad + $name; Value = ('(' + $type + ')') })
    }
}

# Retrieves all settings of a Settings catalog policy. With
# $expand=settingDefinitions we get the real names from Intune; if that fails
# we show cleaned-up IDs instead.
function Get-ConfigPolicySettings {
    param([string]$Id)

    $base = $script:GraphBeta + '/deviceManagement/configurationPolicies/' + $Id + '/settings'
    try   { $settings = @(Invoke-GraphAll -Uri ($base + '?$expand=settingDefinitions')) }
    catch { $settings = @(Invoke-GraphAll -Uri $base) }

    $defs = @{}
    foreach ($s in $settings) {
        foreach ($d in @($s.settingDefinitions | Where-Object { $null -ne $_ })) {
            if ($d.id) { $defs[[string]$d.id] = $d }
        }
    }

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($s in $settings) { Add-SettingRows -Instance $s.settingInstance -Defs $defs -Out $out }
    return $out.ToArray()
}


# -----------------------------------------------------------------------------
# Endpoint security intents (old style) and administrative templates
# -----------------------------------------------------------------------------

function Get-IntentSettings {
    param([string]$Id)

    $list = @(Invoke-GraphAll -Uri ($script:GraphBeta + '/deviceManagement/intents/' + $Id + '/settings'))
    $out  = New-Object System.Collections.Generic.List[object]
    foreach ($s in $list) {
        $defId = [string]$s.definitionId
        $short = $defId
        if ($short.Contains('_')) { $short = $short.Substring($short.LastIndexOf('_') + 1) }
        $name = Format-FirstUpper (Split-Camel $short)

        $names = @($s.PSObject.Properties.Name)
        if (($names -contains 'value') -and ($null -ne $s.value)) { $val = [string]$s.value } else { $val = [string]$s.valueJson }

        $out.Add([pscustomobject]@{ DefinitionId = $defId; Name = $name; Level = 0; Setting = $name; Value = $val })
    }
    return $out.ToArray()
}

function Get-GpoSettings {
    param([string]$Id)

    $base = $script:GraphBeta + '/deviceManagement/groupPolicyConfigurations/' + $Id + '/definitionValues'
    $dvs  = @(Invoke-GraphAll -Uri ($base + '?$expand=definition'))
    $out  = New-Object System.Collections.Generic.List[object]

    foreach ($dv in $dvs) {
        if ($dv.enabled) { $state = 'Enabled' } else { $state = 'Disabled' }
        $vals = ''
        try {
            $pvs = @(Invoke-GraphAll -Uri ($base + '/' + $dv.id + '/presentationValues'))
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($pv in $pvs) {
                if ($null -ne $pv.value) { $parts.Add([string]$pv.value) }
                elseif ($pv.values) {
                    $items = @($pv.values | ForEach-Object { if ($_.value) { [string]$_.value } else { [string]$_.name } })
                    $parts.Add(($items -join ', '))
                }
            }
            $vals = ($parts.ToArray() -join '; ')
        } catch { }

        $name = [string]$dv.definition.displayName
        if ($dv.definition.categoryPath) { $name += '  (' + [string]$dv.definition.categoryPath + ')' }
        $text = $state
        if ($vals) { $text += ' - ' + $vals }
        $out.Add([pscustomobject]@{ DefinitionId = [string]$dv.id; Name = $name; Level = 0; Setting = $name; Value = $text })
    }
    return $out.ToArray()
}


# -----------------------------------------------------------------------------
# Generic properties (for device configuration, compliance, apps, ...)
# -----------------------------------------------------------------------------

# Turns any object into a flat list of "Property = Value".
function ConvertTo-FlatProperties {
    param($Obj, [string]$Prefix, $Out, [int]$Depth = 0)

    if ($Depth -gt 6 -or $null -eq $Obj) { return }

    foreach ($prop in $Obj.PSObject.Properties) {
        $pname = [string]$prop.Name
        if ($pname -like '@odata*' -or $pname -eq 'assignments') { continue }
        $val = $prop.Value
        if ($null -eq $val) { continue }

        if ($Prefix) { $key = $Prefix + '.' + $pname } else { $key = $pname }

        if ($val -is [string] -or $val -is [ValueType]) {
            if ([string]$val -ne '') { $Out.Add([pscustomobject]@{ Property = $key; Value = [string]$val }) }
        }
        elseif ($val -is [System.Collections.IEnumerable]) {
            $arr = @($val)
            $i = 0
            foreach ($e in $arr) {
                $ekey = '{0}[{1}]' -f $key, $i
                if ($e -is [string] -or $e -is [ValueType]) {
                    $Out.Add([pscustomobject]@{ Property = $ekey; Value = [string]$e })
                } else {
                    ConvertTo-FlatProperties -Obj $e -Prefix $ekey -Out $Out -Depth ($Depth + 1)
                }
                $i++
            }
        }
        else {
            ConvertTo-FlatProperties -Obj $val -Prefix $key -Out $Out -Depth ($Depth + 1)
        }
    }
}

# Retrieves one item in full and returns its flat properties.
function Get-ItemProperties {
    param($Item)
    $obj = Invoke-Graph -Uri ('{0}/{1}/{2}/{3}' -f $script:GraphBeta, $Item.Root, $Item.Path, $Item.Id)
    $out = New-Object System.Collections.Generic.List[object]
    ConvertTo-FlatProperties -Obj $obj -Prefix '' -Out $out
    return $out.ToArray()
}


# -----------------------------------------------------------------------------
# Script content (platform scripts, shell scripts, remediations)
# -----------------------------------------------------------------------------

function ConvertFrom-Base64Text {
    param([string]$Base64)
    if (-not $Base64) { return '' }
    $bytes = [System.Convert]::FromBase64String($Base64)
    $text  = [System.Text.Encoding]::UTF8.GetString($bytes)
    return $text.TrimStart([char]0xFEFF)
}

# Returns an ordered hashtable: tab name -> script text
function Get-ScriptTexts {
    param($Item)

    $obj = Invoke-Graph -Uri ('{0}/{1}/{2}/{3}' -f $script:GraphBeta, $Item.Root, $Item.Path, $Item.Id)
    $texts = [ordered]@{}
    switch ($Item.SrcKey) {
        'scripts' { $texts['Script'] = ConvertFrom-Base64Text ([string]$obj.scriptContent) }
        'shell'   { $texts['Script'] = ConvertFrom-Base64Text ([string]$obj.scriptContent) }
        'health'  {
            $texts['Detection script']   = ConvertFrom-Base64Text ([string]$obj.detectionScriptContent)
            $texts['Remediation script'] = ConvertFrom-Base64Text ([string]$obj.remediationScriptContent)
        }
    }
    return $texts
}


# -----------------------------------------------------------------------------
# ASR rules (Attack Surface Reduction)
# -----------------------------------------------------------------------------

# Collects all rules with their mode (Block/Audit/Warn/...) per ASR policy.
# By default only the endpoint security ASR policies. With -ScanAllCatalog
# all Settings catalog profiles are also searched for ASR rules.
function Get-AsrRows {
    param([bool]$ScanAllCatalog = $false)

    $rows  = New-Object System.Collections.Generic.List[object]
    $cands = @($script:AllItems | Where-Object {
        $_.Category -eq 'ASR' -or ($ScanAllCatalog -and $_.SrcKey -eq 'sc' -and $_.Type -eq 'Settings catalog')
    })

    $n = 0
    foreach ($it in $cands) {
        if ($script:Cancel) { break }
        $n++
        Set-Progress $n $cands.Count
        Set-Status ('Reading ASR rules {0}/{1}: {2}' -f $n, $cands.Count, $it.Name)

        try {
            if ($it.SrcKey -eq 'sc') {
                $settings = @(Get-ConfigPolicySettings -Id $it.Id | Where-Object { $_.DefinitionId -match 'attacksurfacereduction' })
            }
            elseif ($it.SrcKey -eq 'intents') {
                $settings = @(Get-IntentSettings -Id $it.Id | Where-Object { $_.DefinitionId -match 'attacksurface' })
            }
            else { $settings = @() }

            foreach ($s in $settings) {
                # Skip container rows without a value (groups)
                if ($s.Value -eq '' -and $s.Level -eq 0 -and $it.SrcKey -eq 'sc') { continue }
                $rows.Add([pscustomobject]@{
                    Policy       = $it.Name
                    Rule         = $s.Name
                    Mode         = $s.Value
                    Include      = $it.Include
                    Exclude      = $it.Exclude
                    Status       = $it.Status
                    Source       = $it.Type
                    Conflict     = ''
                    PolicyId     = $it.Id
                    DefinitionId = $s.DefinitionId
                })
            }
        }
        catch {
            $script:LoadErrors.Add(('ASR {0}: {1}' -f $it.Name, (Get-ErrorText $_)))
        }
    }

    # Same rule in multiple policies? Mark overlaps and conflicts.
    foreach ($g in @($rows.ToArray() | Group-Object DefinitionId)) {
        if ($g.Name -match 'exclusion') { continue }
        $policies = @($g.Group | Select-Object -ExpandProperty PolicyId -Unique)
        if ($policies.Count -le 1) { continue }
        $modes   = @($g.Group | Select-Object -ExpandProperty Mode -Unique)
        $details = (@($g.Group | ForEach-Object { '{0}: {1}' -f $_.Policy, $_.Mode }) -join ' | ')
        if ($modes.Count -gt 1) { $text = 'CONFLICT - different modes: ' + $details }
        else                    { $text = 'Overlap - same mode in ' + $policies.Count + ' policies: ' + $details }
        foreach ($r in $g.Group) { $r.Conflict = $text }
    }

    return $rows.ToArray()
}
