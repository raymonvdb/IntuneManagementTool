# =============================================================================
# Analysis.ps1
# Berekeningen op de geladen gegevens: dashboard, aandachtspunten,
# groepen- en filteroverzicht, plat assignment-overzicht en omgekeerd zoeken.
# Alles hier werkt op $script:AllItems; er worden geen API-calls gedaan
# (behalve het optionele tellen van groepsleden).
# =============================================================================

# -----------------------------------------------------------------------------
# Dashboard
# -----------------------------------------------------------------------------

function Get-TypeSummaryRows {
    $items = @($script:AllItems)
    return @($items | Group-Object Type | Sort-Object Name | ForEach-Object {
        $g = @($_.Group)
        [pscustomobject]@{
            Type        = [string]$_.Name
            Total       = $g.Count
            Assigned    = @($g | Where-Object { $_.Status -eq 'Toegewezen' }).Count
            Unassigned  = @($g | Where-Object { $_.Status -eq 'Niet toegewezen' }).Count
            ExcludeOnly = @($g | Where-Object { $_.Status -eq 'Alleen uitsluiting' }).Count
            AllDevices  = @($g | Where-Object { @($_.Assignments | Where-Object { $_.Kind -eq 'allDevices' -and $_.Mode -eq 'Include' }).Count -gt 0 }).Count
            AllUsers    = @($g | Where-Object { @($_.Assignments | Where-Object { $_.Kind -eq 'allUsers' -and $_.Mode -eq 'Include' }).Count -gt 0 }).Count
        }
    })
}


# -----------------------------------------------------------------------------
# Aandachtspunten
# -----------------------------------------------------------------------------

# Zoekt zaken die meestal een opruimactie of controle waard zijn.
function Get-AttentionRows {
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($it in $script:AllItems) {
        $builtIn = ($it.Subtype -eq 'Microsoft (ingebouwd)')

        if (-not $builtIn) {
            if ($it.Status -eq 'Niet toegewezen') {
                $rows.Add([pscustomobject]@{ Severity = 'Waarschuwing'; Type = $it.Type; Item = $it.Name; Issue = 'Niet toegewezen aan een groep of "Alle apparaten/gebruikers".'; ItemId = $it.Id })
            }
            elseif ($it.Status -eq 'Alleen uitsluiting') {
                $rows.Add([pscustomobject]@{ Severity = 'Waarschuwing'; Type = $it.Type; Item = $it.Name; Issue = 'Heeft alleen uitsluitingen en geen include-assignment: geldt voor niemand.'; ItemId = $it.Id })
            }
        }

        $seen = @{}
        $includedGroups = @($it.Assignments | Where-Object { $_.Kind -eq 'group' -and $_.Mode -eq 'Include' } | ForEach-Object { $_.GroupId })

        foreach ($a in $it.Assignments) {
            if ($a.Kind -ne 'group') { continue }
            $key = $a.Mode + '|' + $a.GroupId
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true

            if ($a.GroupExists -eq $false) {
                $rows.Add([pscustomobject]@{ Severity = 'Fout'; Type = $it.Type; Item = $it.Name; Issue = ('Verwijst naar een groep die niet meer bestaat: ' + $a.GroupId); ItemId = $it.Id })
            }
            elseif ($script:MemberCounts.ContainsKey($a.GroupId) -and $script:MemberCounts[$a.GroupId] -eq 0) {
                $rows.Add([pscustomobject]@{ Severity = 'Waarschuwing'; Type = $it.Type; Item = $it.Name; Issue = ('Groep "' + $a.Name + '" heeft 0 leden (' + $a.Mode + ').'); ItemId = $it.Id })
            }

            if ($a.Mode -eq 'Exclude' -and ($includedGroups -contains $a.GroupId)) {
                $rows.Add([pscustomobject]@{ Severity = 'Fout'; Type = $it.Type; Item = $it.Name; Issue = ('Groep "' + $a.Name + '" staat zowel bij include als exclude.'); ItemId = $it.Id })
            }
        }

        $hasAllDev   = @($it.Assignments | Where-Object { $_.Kind -eq 'allDevices' -and $_.Mode -eq 'Include' }).Count -gt 0
        $hasAllUsers = @($it.Assignments | Where-Object { $_.Kind -eq 'allUsers'   -and $_.Mode -eq 'Include' }).Count -gt 0
        if ($hasAllDev -and $hasAllUsers) {
            $rows.Add([pscustomobject]@{ Severity = 'Info'; Type = $it.Type; Item = $it.Name; Issue = 'Toegewezen aan zowel "Alle apparaten" als "Alle gebruikers".'; ItemId = $it.Id })
        }
    }

    # ASR-regels die in meerdere policies voorkomen met verschillende modi
    foreach ($g in @($script:AsrRows | Where-Object { $_.Conflict -like 'CONFLICT*' } | Group-Object DefinitionId)) {
        $first = $g.Group[0]
        $rows.Add([pscustomobject]@{ Severity = 'Waarschuwing'; Type = 'ASR'; Item = $first.Rule; Issue = $first.Conflict; ItemId = $first.PolicyId })
    }

    return @($rows.ToArray() | Sort-Object @{ Expression = { switch ($_.Severity) { 'Fout' { 0 } 'Waarschuwing' { 1 } default { 2 } } } }, Type, Item)
}


# -----------------------------------------------------------------------------
# Groepen in gebruik
# -----------------------------------------------------------------------------

function Get-GroupRows {
    $map = @{}

    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            if     ($a.Kind -eq 'group')      { $key = $a.GroupId;      $name = $a.Name }
            elseif ($a.Kind -eq 'allDevices') { $key = 'ALL_DEVICES';   $name = 'Alle apparaten' }
            elseif ($a.Kind -eq 'allUsers')   { $key = 'ALL_USERS';     $name = 'Alle gebruikers' }
            else { continue }

            if (-not $map.ContainsKey($key)) {
                $map[$key] = [pscustomobject]@{ GroupId = $key; Name = $name; IncCount = 0; ExcCount = 0; Types = @{} }
            }
            $g = $map[$key]
            if ($a.Mode -eq 'Exclude') { $g.ExcCount++ } else { $g.IncCount++ }
            $g.Types[[string]$it.Type] = $true
        }
    }

    $rows = foreach ($g in $map.Values) {
        $dyn = ''
        $members = $null
        if ($script:GroupCache.ContainsKey($g.GroupId)) {
            if ($script:GroupCache[$g.GroupId].Dynamic) { $dyn = 'Ja' } else { $dyn = 'Nee' }
        }
        if ($script:MemberCounts.ContainsKey($g.GroupId)) { $members = $script:MemberCounts[$g.GroupId] }
        [pscustomobject]@{
            Name     = $g.Name
            Dynamic  = $dyn
            Members  = $members
            IncCount = $g.IncCount
            ExcCount = $g.ExcCount
            Types    = (@($g.Types.Keys | Sort-Object) -join ', ')
            GroupId  = $g.GroupId
        }
    }
    return @($rows | Sort-Object Name)
}

# Telt de leden van alle groepen die in assignments voorkomen.
function Invoke-MemberCountCheck {
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            if ($a.Kind -eq 'group' -and $a.GroupExists -and -not $script:MemberCounts.ContainsKey($a.GroupId)) { $ids.Add($a.GroupId) }
        }
    }
    $todo = @($ids.ToArray() | Select-Object -Unique)
    $failed = 0
    $i = 0
    foreach ($id in $todo) {
        if ($script:Cancel) { break }
        $i++
        Set-Progress $i $todo.Count
        if (($i % 5) -eq 1) { Set-Status ('Ledenaantal ophalen {0}/{1}...' -f $i, $todo.Count) }
        try { $script:MemberCounts[$id] = Get-GroupMemberCount -Id $id }
        catch { $failed++ }
    }
    return [pscustomobject]@{ Total = $todo.Count; Failed = $failed }
}


# -----------------------------------------------------------------------------
# Assignment filters
# -----------------------------------------------------------------------------

function Get-FilterRows {
    $usage = @{}
    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            if ($a.FilterId) {
                if ($usage.ContainsKey($a.FilterId)) { $usage[$a.FilterId]++ } else { $usage[$a.FilterId] = 1 }
            }
        }
    }
    $rows = foreach ($f in $script:FilterCache.Values) {
        $used = 0
        if ($usage.ContainsKey([string]$f.id)) { $used = $usage[[string]$f.id] }
        [pscustomobject]@{
            Name           = [string]$f.displayName
            Platform       = Get-PlatformLabel ([string]$f.platform)
            ManagementType = [string]$f.assignmentFilterManagementType
            Used           = $used
            Rule           = [string]$f.rule
            Description    = [string]$f.description
            FilterId       = [string]$f.id
        }
    }
    return @($rows | Sort-Object Name)
}


# -----------------------------------------------------------------------------
# Plat overzicht: 1 regel per assignment (handig voor draaitabellen in Excel)
# -----------------------------------------------------------------------------

function Get-FlatAssignmentRows {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:AllItems) {
        if (@($it.Assignments).Count -eq 0) {
            $rows.Add([pscustomobject]@{ Type = $it.Type; Subtype = $it.Subtype; Platform = $it.Platform; Name = $it.Name; Mode = ''; Target = '(niet toegewezen)'; Filter = ''; FilterMode = ''; Intent = '' })
            continue
        }
        foreach ($a in $it.Assignments) {
            if ($a.Mode -eq 'Exclude') { $modeText = 'Uitsluiten' } else { $modeText = 'Opnemen' }
            $rows.Add([pscustomobject]@{ Type = $it.Type; Subtype = $it.Subtype; Platform = $it.Platform; Name = $it.Name; Mode = $modeText; Target = $a.Name; Filter = $a.FilterName; FilterMode = $a.FilterMode; Intent = $a.Intent })
        }
    }
    return $rows.ToArray()
}


# -----------------------------------------------------------------------------
# Omgekeerd zoeken: wat is er aan deze groep/gebruiker toegewezen?
# -----------------------------------------------------------------------------

# GroupMap: groupId -> tekst waarmee de relatie wordt omschreven ("Via")
function Find-AssignmentRows {
    param([hashtable]$GroupMap, [bool]$IncludeAll)

    $tmp = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            $via = $null
            if     ($a.Kind -eq 'group' -and $GroupMap.ContainsKey($a.GroupId)) { $via = [string]$GroupMap[$a.GroupId] }
            elseif ($IncludeAll -and $a.Kind -eq 'allUsers')                    { $via = 'Alle gebruikers' }
            elseif ($IncludeAll -and $a.Kind -eq 'allDevices')                  { $via = 'Alle apparaten' }
            if ($null -ne $via) { $tmp.Add([pscustomobject]@{ Item = $it; A = $a; Via = $via }) }
        }
    }

    # Uitsluiting wint van opname binnen hetzelfde item
    $excluded = @{}
    foreach ($t in $tmp) { if ($t.A.Mode -eq 'Exclude') { $excluded[$t.Item.SrcKey + '|' + $t.Item.Id] = $true } }

    $rows = foreach ($t in $tmp) {
        $isExcluded = $excluded.ContainsKey($t.Item.SrcKey + '|' + $t.Item.Id)
        if ($t.A.Mode -eq 'Exclude')  { $mode = 'Uitsluiten'; $effect = 'Uitgesloten' }
        elseif ($isExcluded)          { $mode = 'Opnemen';    $effect = 'Uitgesloten (exclude wint)' }
        else                          { $mode = 'Opnemen';    $effect = 'Van toepassing' }

        $detail = ''
        if ($t.A.FilterName) { $detail += 'filter: {0} ({1})  ' -f $t.A.FilterName, $t.A.FilterMode }
        if ($t.A.Intent)     { $detail += $t.A.Intent }

        [pscustomobject]@{
            Name    = $t.Item.Name
            Type    = $t.Item.Type
            Subtype = $t.Item.Subtype
            Via     = $t.Via
            Mode    = $mode
            Effect  = $effect
            Detail  = $detail.Trim()
            Item    = $t.Item
        }
    }
    return @($rows | Sort-Object Type, Name)
}
