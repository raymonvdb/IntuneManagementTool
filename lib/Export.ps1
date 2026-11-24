# =============================================================================
# Export.ps1
# Exporteren naar Excel (.xlsx, zonder dat Excel geinstalleerd hoeft te zijn),
# CSV (puntkomma-gescheiden) en een zelfstandig HTML-rapport.
#
# Kolomdefinitie ($Cols): lijst van @{ H = 'Koptekst'; P = 'Eigenschapsnaam' }
# -----------------------------------------------------------------------------

function Get-CellText {
    param($Row, [string]$Prop)
    $v = $Row.$Prop
    if ($null -eq $v) { return '' }
    return [string]$v
}

# -----------------------------------------------------------------------------
# CSV
# -----------------------------------------------------------------------------

function Export-RowsCsv {
    param([string]$Path, $Cols, $Rows)

    $out = foreach ($r in $Rows) {
        $o = [ordered]@{}
        foreach ($c in $Cols) { $o[[string]$c.H] = (Get-CellText $r $c.P) -replace "`r?`n", '; ' }
        [pscustomobject]$o
    }
    @($out) | Export-Csv -Path $Path -NoTypeInformation -Delimiter ';' -Encoding UTF8
}


# -----------------------------------------------------------------------------
# XLSX
# -----------------------------------------------------------------------------

function ConvertTo-XmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $t = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    if ($t.Length -gt 32000) { $t = $t.Substring(0, 32000) }
    return [System.Security.SecurityElement]::Escape($t)
}

function Get-ColumnLetter {
    param([int]$Index)
    $s = ''
    $n = $Index + 1
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = ([string][char](65 + $r)) + $s
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $s
}

# Bouwt de XML van 1 werkblad. Header: string[]; Matrix: lijst van string[]
function New-SheetXml {
    param($Header, $Matrix)

    $colCount = @($Header).Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')

    # Kolombreedtes op basis van de inhoud (max 60 tekens)
    [void]$sb.Append('<cols>')
    for ($c = 0; $c -lt $colCount; $c++) {
        $max = ([string]$Header[$c]).Length
        $sample = 0
        foreach ($row in $Matrix) {
            if ($sample -ge 200) { break }
            $sample++
            foreach ($line in (([string]$row[$c]) -split "`n")) { if ($line.Length -gt $max) { $max = $line.Length } }
        }
        $w = [math]::Min(60, [math]::Max(10, $max + 2))
        [void]$sb.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($c + 1), $w))
    }
    [void]$sb.Append('</cols>')

    [void]$sb.Append('<sheetData>')
    [void]$sb.Append('<row r="1">')
    for ($c = 0; $c -lt $colCount; $c++) {
        [void]$sb.Append(('<c r="{0}1" s="1" t="inlineStr"><is><t xml:space="preserve">{1}</t></is></c>' -f (Get-ColumnLetter $c), (ConvertTo-XmlText ([string]$Header[$c]))))
    }
    [void]$sb.Append('</row>')

    $r = 1
    foreach ($row in $Matrix) {
        $r++
        [void]$sb.Append(('<row r="{0}">' -f $r))
        for ($c = 0; $c -lt $colCount; $c++) {
            $text = [string]$row[$c]
            if ($text -eq '') { continue }
            [void]$sb.Append(('<c r="{0}{1}" s="2" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f (Get-ColumnLetter $c), $r, (ConvertTo-XmlText $text)))
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')

    $lastCol = Get-ColumnLetter ($colCount - 1)
    [void]$sb.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $lastCol, [math]::Max(1, $r)))
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Add-ZipText {
    param($Zip, [string]$EntryName, [string]$Text)
    $entry  = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Text)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
}

# Sheets: lijst van @{ Name = 'Blad'; Cols = ...; Rows = ... }
function Export-Xlsx {
    param([string]$Path, $Sheets)

    Add-Type -AssemblyName System.IO.Compression

    # Werkbladen voorbereiden (naam opschonen, matrix bouwen)
    $prepared = New-Object System.Collections.Generic.List[object]
    $usedNames = @{}
    foreach ($s in $Sheets) {
        $name = [regex]::Replace([string]$s.Name, '[\[\]\:\*\?\/\\]', '-')
        if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
        if (-not $name) { $name = 'Blad' }
        $base = $name; $k = 2
        while ($usedNames.ContainsKey($name.ToLowerInvariant())) { $name = $base.Substring(0, [math]::Min($base.Length, 28)) + ' ' + $k; $k++ }
        $usedNames[$name.ToLowerInvariant()] = $true

        $header = @($s.Cols | ForEach-Object { [string]$_.H })
        $matrix = New-Object System.Collections.Generic.List[object]
        foreach ($r in @($s.Rows)) {
            $line = New-Object 'string[]' $header.Count
            for ($c = 0; $c -lt $header.Count; $c++) { $line[$c] = Get-CellText $r $s.Cols[$c].P }
            $matrix.Add($line)
        }
        $prepared.Add([pscustomobject]@{ Name = $name; Header = $header; Matrix = $matrix })
    }

    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $ct  = New-Object System.Text.StringBuilder
    [void]$ct.Append($xml)
    [void]$ct.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$ct.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    [void]$ct.Append('<Default Extension="xml" ContentType="application/xml"/>')
    [void]$ct.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    [void]$ct.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')

    $wb  = New-Object System.Text.StringBuilder
    [void]$wb.Append($xml)
    [void]$wb.Append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')

    $rel = New-Object System.Text.StringBuilder
    [void]$rel.Append($xml)
    [void]$rel.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

    $n = 0
    foreach ($p in $prepared) {
        $n++
        [void]$ct.Append(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $n))
        [void]$wb.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (ConvertTo-XmlText $p.Name), $n))
        [void]$rel.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $n))
    }
    [void]$ct.Append('</Types>')
    [void]$wb.Append('</sheets></workbook>')
    [void]$rel.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' -f ($n + 1)))
    [void]$rel.Append('</Relationships>')

    $rootRels = $xml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

    # Stijlen: 0 = standaard, 1 = kop (wit op blauw), 2 = cel (tekstterugloop, bovenaan uitgelijnd)
    $styles = $xml + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F6CBD"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Add-ZipText $zip '[Content_Types].xml'        $ct.ToString()
            Add-ZipText $zip '_rels/.rels'                $rootRels
            Add-ZipText $zip 'xl/workbook.xml'            $wb.ToString()
            Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rel.ToString()
            Add-ZipText $zip 'xl/styles.xml'              $styles
            $i = 0
            foreach ($p in $prepared) {
                $i++
                Add-ZipText $zip ('xl/worksheets/sheet{0}.xml' -f $i) (New-SheetXml -Header $p.Header -Matrix $p.Matrix)
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}


# -----------------------------------------------------------------------------
# HTML-rapport (1 bestand, met zoeken en sorteren in de browser)
# -----------------------------------------------------------------------------

function ConvertTo-HtmlText {
    param([string]$Text)
    $t = [System.Net.WebUtility]::HtmlEncode($Text)
    return ($t -replace "`r?`n", '<br>')
}

# Cards: lijst van @{ Label; Value }. Sections: lijst van @{ Title; Cols; Rows }
function Export-HtmlReport {
    param([string]$Path, [string]$Title, $Cards, $Sections, [string]$Subtitle = '')

    $head = @'
<!DOCTYPE html>
<html lang="nl"><head><meta charset="utf-8"><title>@@TITLE@@</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f2f1;color:#201f1e}
header{background:#0b3c5d;color:#fff;padding:18px 28px}
header h1{margin:0;font-size:22px} header div{opacity:.8;margin-top:4px;font-size:13px}
main{padding:20px 28px}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:10px}
.card{background:#fff;border-bottom:3px solid #0f6cbd;padding:10px 16px;min-width:150px}
.card b{display:block;font-size:26px} .card span{color:#605e5c;font-size:12px}
h2{font-size:17px;margin:28px 0 8px}
input{padding:6px 8px;margin-bottom:8px;width:320px;border:1px solid #8a8886}
table{border-collapse:collapse;background:#fff;width:100%;font-size:13px}
th{background:#faf9f8;text-align:left;padding:7px 8px;border-bottom:2px solid #e1dfdd;cursor:pointer;position:sticky;top:0}
td{padding:6px 8px;border-bottom:1px solid #edebe9;vertical-align:top}
tr:hover td{background:#f3f9fd}
</style>
<script>
function filt(inp){var q=inp.value.toLowerCase();var t=inp.parentNode.querySelector('table');var rows=t.tBodies[0].rows;for(var i=0;i<rows.length;i++){rows[i].style.display=rows[i].innerText.toLowerCase().indexOf(q)>-1?'':'none';}}
function sortT(th){var t=th;while(t&&t.tagName!=='TABLE'){t=t.parentNode;}var i=th.cellIndex;var asc=th.getAttribute('data-asc')!=='1';th.setAttribute('data-asc',asc?'1':'0');var rows=Array.prototype.slice.call(t.tBodies[0].rows);rows.sort(function(a,b){var x=a.cells[i].innerText.toLowerCase(),y=b.cells[i].innerText.toLowerCase();var nx=parseFloat(x),ny=parseFloat(y);if(!isNaN(nx)&&!isNaN(ny)&&String(nx)===x&&String(ny)===y){return asc?nx-ny:ny-nx;}return asc?x.localeCompare(y):y.localeCompare(x);});for(var k=0;k<rows.length;k++){t.tBodies[0].appendChild(rows[k]);}}
</script></head><body>
'@
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($head.Replace('@@TITLE@@', (ConvertTo-HtmlText $Title)))
    [void]$sb.Append('<header><h1>' + (ConvertTo-HtmlText $Title) + '</h1><div>' + (ConvertTo-HtmlText ('Gegenereerd op ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + ' ' + $Subtitle)) + '</div></header><main>')

    if ($Cards) {
        [void]$sb.Append('<div class="cards">')
        foreach ($c in $Cards) {
            [void]$sb.Append('<div class="card"><b>' + (ConvertTo-HtmlText ([string]$c.Value)) + '</b><span>' + (ConvertTo-HtmlText ([string]$c.Label)) + '</span></div>')
        }
        [void]$sb.Append('</div>')
    }

    foreach ($s in $Sections) {
        $rows = @($s.Rows)
        [void]$sb.Append('<h2>' + (ConvertTo-HtmlText ([string]$s.Title)) + ' (' + $rows.Count + ')</h2><div>')
        [void]$sb.Append('<input placeholder="Zoeken..." oninput="filt(this)"><table><thead><tr>')
        foreach ($c in $s.Cols) { [void]$sb.Append('<th onclick="sortT(this)">' + (ConvertTo-HtmlText ([string]$c.H)) + '</th>') }
        [void]$sb.Append('</tr></thead><tbody>')
        foreach ($r in $rows) {
            [void]$sb.Append('<tr>')
            foreach ($c in $s.Cols) { [void]$sb.Append('<td>' + (ConvertTo-HtmlText (Get-CellText $r $c.P)) + '</td>') }
            [void]$sb.Append('</tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
    }
    [void]$sb.Append('</main></body></html>')

    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
}
