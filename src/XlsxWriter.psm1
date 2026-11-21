<#
XlsxWriter.psm1 - writes a small Excel workbook (.xlsx) with .NET only: no Excel, no module.

  Purpose:     The results file of Ion: one workbook with tabs (the PAM, 2026-11-20: a CSV cannot have tabs).
               Each tab has a bold header row; a list tab also has the header frozen and filter buttons on, and its
               link column holds clickable links (the PAM, 2026-11-20); chosen columns wrap their text, and columns the
               person fills in can have a dropdown list of fixed words (the PAM, 2026-11-21).
  Inputs:      the tabs to write (New-XlsxSheet), each with its columns and rows of cells (text or whole numbers).
  Outputs:     one .xlsx file (Office Open XML, ECMA-376: a zip of XML parts).
  Depends on:  .NET System.IO.Compression (loaded here).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the one .xlsx file the caller names; it never overwrites a file that exists.
  Credentials: none.

  Safety:      every text cell is an inline string, never a formula, so text from a scanned file cannot run in Excel.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# Excel limits: 31 characters in a tab name, 32,767 in a cell (ECMA-376 and Excel's specifications and limits).
$script:MaxSheetName = 31
$script:MaxCellText = 32767

# Excel opens a link of at most 2,079 characters (a longer one stays plain text in its cell).
$script:MaxLinkChars = 2079

# XML namespaces of the parts.
$script:MainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Four cell styles: 0 = normal, 1 = bold (the header row and the Summary labels), 2 = a link (blue, underlined),
# 3 = wrapped text, top-aligned (long prose, for example the Method tab; the PAM, 2026-11-21).
$script:StylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>'

function New-XlsxSheet {
    <#
    .SYNOPSIS
      One tab of the workbook.
    .PARAMETER Name
      The tab name (cut to 31 characters; characters Excel refuses are replaced).
    .PARAMETER Columns
      The header row.
    .PARAMETER Rows
      The rows: each an array of cells, a cell being text or a whole number.
    .PARAMETER Widths
      Column widths in characters, one per column.
    .PARAMETER List
      Freeze the header row and turn the filter buttons on (a list of hits); off for the Summary tab.
    .PARAMETER LinkColumn
      The column (0-based) whose https:// cells become clickable links; -1 for none.
    .PARAMETER WrapColumns
      Columns (0-based) whose text wraps in the cell (style 3).
    .PARAMETER Choices
      Column (0-based) -> the values its cells may hold: a dropdown list in Excel (data validation, a fixed list of words;
      never a formula from scanned text). For columns the person fills in.
    .OUTPUTS
      [pscustomobject]
    #>
    param([string]$Name, [string[]]$Columns, [object[]]$Rows, [int[]]$Widths, [switch]$List, [int]$LinkColumn = -1,
          [int[]]$WrapColumns = @(), [hashtable]$Choices = @{})
    [pscustomobject]@{ Name = $Name; Columns = $Columns; Rows = @($Rows); Widths = $Widths; List = [bool]$List; LinkColumn = $LinkColumn
        WrapColumns = $WrapColumns; Choices = $Choices }
}

function ConvertTo-XmlText {
    # Text for an XML part: characters XML does not allow are dropped, the rest is escaped, and it is cut to Excel's cell limit.
    param([AllowEmptyString()][string]$Text)
    $clean = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', ''
    if ($clean.Length -gt $script:MaxCellText) { $clean = $clean.Substring(0, $script:MaxCellText) }
    return [System.Security.SecurityElement]::Escape($clean)
}

function Get-ColumnLetter {
    # The column letters of a 1-based column number: 1 = A, 27 = AA.
    param([int]$Number)
    $letters = ''
    while ($Number -gt 0) { $Number--; $letters = [char](65 + $Number % 26) + $letters; $Number = [Math]::Floor($Number / 26) }
    return $letters
}

function ConvertTo-CellXml {
    # One cell: a whole number as a number, anything else as an inline string (never a formula).
    param([string]$Ref, $Value, [int]$Style)
    $styleAttr = ''
    if ($Style -gt 0) { $styleAttr = ' s="' + $Style + '"' }
    if ($Value -is [int] -or $Value -is [long]) { return '<c r="{0}"{1}><v>{2}</v></c>' -f $Ref, $styleAttr, $Value }
    return '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Ref, $styleAttr, (ConvertTo-XmlText ([string]$Value))
}

function Get-SheetLinks {
    <#
    .SYNOPSIS
      The clickable links of one tab: each https:// cell in its link column, with its cell and relationship ID.
      A link is a relationship of the sheet (ECMA-376 hyperlink), never a HYPERLINK formula (the safety rule above).
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .OUTPUTS
      [pscustomobject] Ref (for example G2), Id (rId1, rId2 ...), Target.
    #>
    param($Sheet)
    if ($Sheet.LinkColumn -lt 0) { return }
    $letter = Get-ColumnLetter ($Sheet.LinkColumn + 1)
    $rowNumber = 1
    $n = 0
    foreach ($row in $Sheet.Rows) {
        $rowNumber++
        $cells = @($row)
        if ($Sheet.LinkColumn -ge $cells.Count) { continue }
        $target = [string]$cells[$Sheet.LinkColumn]
        if ($target -notmatch '^https://\S+$' -or $target.Length -gt $script:MaxLinkChars) { continue }
        $n++
        [pscustomobject]@{ Ref = $letter + $rowNumber; Id = 'rId' + $n; Target = $target }
    }
}

function ConvertTo-SheetRelsXml {
    # The relationships part of one tab: one external target per link.
    param([object[]]$Links)
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    foreach ($link in $Links) {
        $xml += '<Relationship Id="{0}" Type="{1}/hyperlink" Target="{2}" TargetMode="External"/>' -f $link.Id, $script:RelNs, (ConvertTo-XmlText $link.Target)
    }
    return $xml + '</Relationships>'
}

function ConvertTo-SheetXml {
    <#
    .SYNOPSIS
      The worksheet XML of one tab.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .PARAMETER Links
      Output of Get-SheetLinks for this tab.
    .OUTPUTS
      [string]
    #>
    param($Sheet, [object[]]$Links = @())
    $linked = @{}
    foreach ($link in $Links) { $linked[$link.Ref] = $true }
    $xml = New-Object System.Text.StringBuilder
    [void]$xml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + $script:MainNs + '" xmlns:r="' + $script:RelNs + '">')
    if ($Sheet.List) {
        [void]$xml.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    }
    if ($Sheet.Widths) {
        [void]$xml.Append('<cols>')
        for ($i = 0; $i -lt $Sheet.Widths.Count; $i++) {
            [void]$xml.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $Sheet.Widths[$i]))
        }
        [void]$xml.Append('</cols>')
    }
    [void]$xml.Append('<sheetData><row r="1">')
    for ($i = 0; $i -lt $Sheet.Columns.Count; $i++) { [void]$xml.Append((ConvertTo-CellXml ((Get-ColumnLetter ($i + 1)) + '1') $Sheet.Columns[$i] 1)) }
    [void]$xml.Append('</row>')
    $rowNumber = 1
    foreach ($row in $Sheet.Rows) {
        $rowNumber++
        [void]$xml.Append('<row r="' + $rowNumber + '">')
        $cells = @($row)
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $style = 0
            $ref = (Get-ColumnLetter ($i + 1)) + $rowNumber
            if (-not $Sheet.List -and $i -eq 0) { $style = 1 }   # Summary tab: the label column is bold
            if ($Sheet.WrapColumns -contains $i) { $style = 3 }
            if ($linked.ContainsKey($ref)) { $style = 2 }
            [void]$xml.Append((ConvertTo-CellXml $ref $cells[$i] $style))
        }
        [void]$xml.Append('</row>')
    }
    [void]$xml.Append('</sheetData>')
    if ($Sheet.List -and $Sheet.Columns.Count -gt 0) {
        [void]$xml.Append(('<autoFilter ref="A1:{0}{1}"/>' -f (Get-ColumnLetter $Sheet.Columns.Count), [Math]::Max(2, $rowNumber)))
    }
    [void]$xml.Append((ConvertTo-ChoicesXml $Sheet ([Math]::Max(2, $rowNumber))))
    if ($Links.Count -gt 0) {
        [void]$xml.Append('<hyperlinks>')
        foreach ($link in $Links) { [void]$xml.Append(('<hyperlink ref="{0}" r:id="{1}"/>' -f $link.Ref, $link.Id)) }
        [void]$xml.Append('</hyperlinks>')
    }
    [void]$xml.Append('</worksheet>')
    return $xml.ToString()
}

function ConvertTo-ChoicesXml {
    <#
    .SYNOPSIS
      The dropdown lists of one tab (ECMA-376 dataValidations, after autoFilter and before hyperlinks): one list per column
      in Choices, for rows 2 to the last row; blank cells are allowed.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .PARAMETER LastRow
      The last row number.
    .OUTPUTS
      [string] empty when the tab has no choices.
    #>
    param($Sheet, [int]$LastRow)
    if (-not $Sheet.Choices -or $Sheet.Choices.Count -eq 0) { return '' }
    $xml = '<dataValidations count="' + $Sheet.Choices.Count + '">'
    foreach ($column in @($Sheet.Choices.Keys | Sort-Object)) {
        $letter = Get-ColumnLetter ([int]$column + 1)
        $list = (@($Sheet.Choices[$column]) | ForEach-Object { ([string]$_) -replace '[",]', '' }) -join ','
        $xml += '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{0}2:{0}{1}"><formula1>"{2}"</formula1></dataValidation>' -f
            $letter, $LastRow, (ConvertTo-XmlText $list)
    }
    return $xml + '</dataValidations>'
}

function Get-SafeSheetNames {
    <#
    .SYNOPSIS
      Tab names Excel accepts: no : \ / ? * [ ], at most 31 characters, none twice (a number is added).
    .PARAMETER Names
      The wanted names, in order.
    .OUTPUTS
      [string[]] the names to use, in the same order.
    #>
    param([string[]]$Names)
    $used = @{}
    foreach ($name in $Names) {
        $clean = ($name -replace '[:\\/?*\[\]]', '-').Trim()
        if (-not $clean) { $clean = 'Sheet' }
        if ($clean.Length -gt $script:MaxSheetName) { $clean = $clean.Substring(0, $script:MaxSheetName) }
        $candidate = $clean
        $n = 2
        while ($used.ContainsKey($candidate.ToLower())) {
            $suffix = ' (' + $n + ')'
            $candidate = $clean.Substring(0, [Math]::Min($clean.Length, $script:MaxSheetName - $suffix.Length)) + $suffix
            $n++
        }
        $used[$candidate.ToLower()] = $true
        $candidate
    }
}

function Add-ZipText {
    # Adds one UTF-8 text part (no byte order mark) to the zip.
    param($Zip, [string]$Name, [string]$Text)
    $writer = New-Object System.IO.StreamWriter($Zip.CreateEntry($Name).Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Export-Xlsx {
    <#
    .SYNOPSIS
      Writes the tabs to a new .xlsx file. Throws if the file exists (the caller picks a new name).
    .PARAMETER Path
      The new file.
    .PARAMETER Sheets
      The tabs, in order (New-XlsxSheet).
    #>
    param([string]$Path, [object[]]$Sheets)
    $names = @(Get-SafeSheetNames @($Sheets | ForEach-Object { $_.Name }))
    $types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $book = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + $script:MainNs + '" xmlns:r="' + $script:RelNs + '"><sheets>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $types += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        $book += '<sheet name="' + (ConvertTo-XmlText $names[$i - 1]) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'
        $rels += '<Relationship Id="rId' + $i + '" Type="' + $script:RelNs + '/worksheet" Target="worksheets/sheet' + $i + '.xml"/>'
    }
    $rels += '<Relationship Id="rId' + ($Sheets.Count + 1) + '" Type="' + $script:RelNs + '/styles" Target="styles.xml"/></Relationships>'
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::CreateNew)   # CreateNew: never overwrite
    $zip = $null
    $written = $false
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        Add-ZipText $zip '[Content_Types].xml' ($types + '</Types>')
        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">' +
            '<Relationship Id="rId1" Type="' + $script:RelNs + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        Add-ZipText $zip 'xl/workbook.xml' ($book + '</sheets></workbook>')
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rels
        Add-ZipText $zip 'xl/styles.xml' $script:StylesXml
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $links = @(Get-SheetLinks $Sheets[$i - 1])
            Add-ZipText $zip ('xl/worksheets/sheet' + $i + '.xml') (ConvertTo-SheetXml $Sheets[$i - 1] $links)
            if ($links.Count -gt 0) { Add-ZipText $zip ('xl/worksheets/_rels/sheet' + $i + '.xml.rels') (ConvertTo-SheetRelsXml $links) }
        }
        $written = $true
    }
    finally {
        if ($zip) { $zip.Dispose() }
        $stream.Dispose()
        # A half-written workbook would not open: remove it, so a rerun can use the name.
        if (-not $written) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
    }
}

Export-ModuleMember -Function New-XlsxSheet, Export-Xlsx
