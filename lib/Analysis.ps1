# =============================================================================
# Analysis.ps1
# Calculations on the loaded data: dashboard, attention points,
# group and filter overview, flat assignment overview and reverse search.
# Everything here works on $script:AllItems; no API calls are made
# (except the optional counting of group members).
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
            Assigned    = @($g | Where-Object { $_.Status -eq 'Assigned' }).Count
            Unassigned  = @($g | Where-Object { $_.Status -eq 'Not assigned' }).Count
            ExcludeOnly = @($g | Where-Object { $_.Status -eq 'Exclusion only' }).Count
            AllDevices  = @($g | Where-Object { @($_.Assignments | Where-Object { $_.Kind -eq 'allDevices' -and $_.Mode -eq 'Include' }).Count -gt 0 }).Count
            AllUsers    = @($g | Where-Object { @($_.Assignments | Where-Object { $_.Kind -eq 'allUsers' -and $_.Mode -eq 'Include' }).Count -gt 0 }).Count
        }
    })
}


# -----------------------------------------------------------------------------
# Attention points
# -----------------------------------------------------------------------------

# Looks for things that usually deserve a cleanup or a check.
function Get-AttentionRows {
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($it in $script:AllItems) {
        $builtIn = ($it.Subtype -eq 'Microsoft (built-in)')

        if (-not $builtIn) {
            if ($it.Status -eq 'Not assigned') {
                $rows.Add([pscustomobject]@{ Severity = 'Warning'; Type = $it.Type; Item = $it.Name; Issue = 'Not assigned to any group or to "All devices/users".'; ItemId = $it.Id })
            }
            elseif ($it.Status -eq 'Exclusion only') {
                $rows.Add([pscustomobject]@{ Severity = 'Warning'; Type = $it.Type; Item = $it.Name; Issue = 'Has only exclusions and no include assignment: applies to nobody.'; ItemId = $it.Id })
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
                $rows.Add([pscustomobject]@{ Severity = 'Error'; Type = $it.Type; Item = $it.Name; Issue = ('References a group that no longer exists: ' + $a.GroupId); ItemId = $it.Id })
            }
            elseif ($script:MemberCounts.ContainsKey($a.GroupId) -and $script:MemberCounts[$a.GroupId] -eq 0) {
                $rows.Add([pscustomobject]@{ Severity = 'Warning'; Type = $it.Type; Item = $it.Name; Issue = ('Group "' + $a.Name + '" has 0 members (' + $a.Mode + ').'); ItemId = $it.Id })
            }

            if ($a.Mode -eq 'Exclude' -and ($includedGroups -contains $a.GroupId)) {
                $rows.Add([pscustomobject]@{ Severity = 'Error'; Type = $it.Type; Item = $it.Name; Issue = ('Group "' + $a.Name + '" is both included and excluded.'); ItemId = $it.Id })
            }
        }

        $hasAllDev   = @($it.Assignments | Where-Object { $_.Kind -eq 'allDevices' -and $_.Mode -eq 'Include' }).Count -gt 0
        $hasAllUsers = @($it.Assignments | Where-Object { $_.Kind -eq 'allUsers'   -and $_.Mode -eq 'Include' }).Count -gt 0
        if ($hasAllDev -and $hasAllUsers) {
            $rows.Add([pscustomobject]@{ Severity = 'Info'; Type = $it.Type; Item = $it.Name; Issue = 'Assigned to both "All devices" and "All users".'; ItemId = $it.Id })
        }
    }

    # ASR rules that occur in multiple policies with different modes
    foreach ($g in @($script:AsrRows | Where-Object { $_.Conflict -like 'CONFLICT*' } | Group-Object DefinitionId)) {
        $first = $g.Group[0]
        $rows.Add([pscustomobject]@{ Severity = 'Warning'; Type = 'ASR'; Item = $first.Rule; Issue = $first.Conflict; ItemId = $first.PolicyId })
    }

    return @($rows.ToArray() | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } }, Type, Item)
}


# -----------------------------------------------------------------------------
# Groups in use
# -----------------------------------------------------------------------------

function Get-GroupRows {
    $map = @{}

    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            if     ($a.Kind -eq 'group')      { $key = $a.GroupId;      $name = $a.Name }
            elseif ($a.Kind -eq 'allDevices') { $key = 'ALL_DEVICES';   $name = 'All devices' }
            elseif ($a.Kind -eq 'allUsers')   { $key = 'ALL_USERS';     $name = 'All users' }
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
            if ($script:GroupCache[$g.GroupId].Dynamic) { $dyn = 'Yes' } else { $dyn = 'No' }
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

# Counts the members of all groups that occur in assignments.
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
        if (($i % 5) -eq 1) { Set-Status ('Getting member counts {0}/{1}...' -f $i, $todo.Count) }
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
# Flat overview: one row per assignment (handy for pivot tables in Excel)
# -----------------------------------------------------------------------------

function Get-FlatAssignmentRows {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:AllItems) {
        if (@($it.Assignments).Count -eq 0) {
            $rows.Add([pscustomobject]@{ Type = $it.Type; Subtype = $it.Subtype; Platform = $it.Platform; Name = $it.Name; Mode = ''; Target = '(not assigned)'; Filter = ''; FilterMode = ''; Intent = '' })
            continue
        }
        foreach ($a in $it.Assignments) {
            $rows.Add([pscustomobject]@{ Type = $it.Type; Subtype = $it.Subtype; Platform = $it.Platform; Name = $it.Name; Mode = $a.Mode; Target = $a.Name; Filter = $a.FilterName; FilterMode = $a.FilterMode; Intent = $a.Intent })
        }
    }
    return $rows.ToArray()
}


# -----------------------------------------------------------------------------
# Reverse search: what is assigned to this group/user?
# -----------------------------------------------------------------------------

# GroupMap: groupId -> text that describes the relationship ("Via")
function Find-AssignmentRows {
    param([hashtable]$GroupMap, [bool]$IncludeAll)

    $tmp = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:AllItems) {
        foreach ($a in $it.Assignments) {
            $via = $null
            if     ($a.Kind -eq 'group' -and $GroupMap.ContainsKey($a.GroupId)) { $via = [string]$GroupMap[$a.GroupId] }
            elseif ($IncludeAll -and $a.Kind -eq 'allUsers')                    { $via = 'All users' }
            elseif ($IncludeAll -and $a.Kind -eq 'allDevices')                  { $via = 'All devices' }
            if ($null -ne $via) { $tmp.Add([pscustomobject]@{ Item = $it; A = $a; Via = $via }) }
        }
    }

    # Exclusion wins over inclusion within the same item
    $excluded = @{}
    foreach ($t in $tmp) { if ($t.A.Mode -eq 'Exclude') { $excluded[$t.Item.SrcKey + '|' + $t.Item.Id] = $true } }

    $rows = foreach ($t in $tmp) {
        $isExcluded = $excluded.ContainsKey($t.Item.SrcKey + '|' + $t.Item.Id)
        if ($t.A.Mode -eq 'Exclude') { $effect = 'Excluded' }
        elseif ($isExcluded)         { $effect = 'Excluded (exclude wins)' }
        else                         { $effect = 'Applies' }

        $detail = ''
        if ($t.A.FilterName) { $detail += 'filter: {0} ({1})  ' -f $t.A.FilterName, $t.A.FilterMode }
        if ($t.A.Intent)     { $detail += $t.A.Intent }

        [pscustomobject]@{
            Name    = $t.Item.Name
            Type    = $t.Item.Type
            Subtype = $t.Item.Subtype
            Via     = $t.Via
            Mode    = $t.A.Mode
            Effect  = $effect
            Detail  = $detail.Trim()
            Item    = $t.Item
        }
    }
    return @($rows | Sort-Object Type, Name)
}
