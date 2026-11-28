<#
ReportExport.psm1 - screen summary formatting, workbook sheets/tabs, breach advice, and ECMA-376 .xlsx writer.

  Purpose:     Builds report rows, severity counts, not-scanned groups, the Summary/Hits/Files/Not scanned/
               By account/Method tabs, and writes the CUI-marked Excel workbook (.xlsx) with .NET only.
  Inputs:      scan rows, summary metrics, output folder, site URL.
  Outputs:     screen summaries and saved <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx workbook.
  Depends on:  PiiEngine.psm1 (Get-ConfidenceRank, Protect-PiiName), imported first; ScanQueue / Accountability at run time.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the one .xlsx workbook and its parent folders under OutputFolder; never overwrites an existing file.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# ==== XlsxWriter.psm1 ====
# Excel limits: 31 characters in a tab name, 32,767 in a cell (ECMA-376 and Excel's specifications and limits).
$script:MaxSheetName = 31
$script:MaxCellText = 32767

# Excel opens a link of at most 2,079 characters (a longer one stays plain text in its cell).
$script:MaxLinkChars = 2079

# XML namespaces of the parts.
$script:MainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Cell styles: 0 = normal, 1 = bold (the header row and the Summary labels), 2 = a link (blue, underlined),
# 3 = wrapped text, top-aligned (long prose, for example the Method tab; the Administrator, 2026-11-21), 4 = a date and time,
# 5, 6, 7 = the High, Medium, Low fill of a confidence cell (the Administrator, 2026-11-24, the approved preview).
$script:StylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts>' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts>' +
    '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>'
$script:FillStyle = @{ 'High' = 5; 'Medium' = 6; 'Low' = 7 }
# The date and time texts Ion writes ("Last saved" and the like), read as Excel date cells; any other text stays text.
$script:DateFormats = [string[]]@('yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd')
# Every tab prints landscape, one page wide, "CUI" centred at the top and bottom (the Administrator, 2026-11-24).
$script:PrintXml = '<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
    '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
    '<headerFooter><oddHeader>&amp;CCUI</oddHeader><oddFooter>&amp;CCUI</oddFooter></headerFooter>'

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
    .PARAMETER Table
      With List: the Excel Table name (letters and digits, unique in the workbook); the tab is then a real table.
    .PARAMETER LinkTarget
      The column (0-based) that holds the https:// address of the link in LinkColumn; -1: LinkColumn itself.
    .PARAMETER DateColumns
      Columns (0-based) whose "yyyy-MM-dd HH:mm" texts become date cells (style 4).
    .PARAMETER FreezeColumns
      With List: how many columns on the left stay in view when the tab scrolls right.
    .PARAMETER FillColumn
      The column (0-based) of High, Medium, Low cells that get a fill (styles 5-7); -1 for none.
    .PARAMETER DateChecks
      Columns (0-based) the person fills in where Excel accepts only a date (2000-01-01 or later), with a short message.
    .OUTPUTS
      [pscustomobject]
    #>
    param([string]$Name, [string[]]$Columns, [object[]]$Rows, [int[]]$Widths, [switch]$List, [int]$LinkColumn = -1,
          [int[]]$WrapColumns = @(), [hashtable]$Choices = @{}, [string]$Table = '', [int]$LinkTarget = -1,
          [int[]]$DateColumns = @(), [int]$FreezeColumns = 0, [int]$FillColumn = -1, [int[]]$DateChecks = @())
    [pscustomobject]@{ Name = $Name; Columns = $Columns; Rows = @($Rows); Widths = $Widths; List = [bool]$List; LinkColumn = $LinkColumn
        WrapColumns = $WrapColumns; Choices = $Choices; Table = $Table; LinkTarget = $LinkTarget; DateColumns = $DateColumns
        FreezeColumns = $FreezeColumns; FillColumn = $FillColumn; DateChecks = $DateChecks }
}

function New-XlsxFormula {
    <#
    .SYNOPSIS
      A formula cell: the formula text (without "=") and the value Ion worked out, shown until Excel recalculates.
      Only Ion's code builds the text (cell references, table names): never pass text from a scanned file.
    .PARAMETER Formula
      For example COUNTIFS(tblHits[Type],A20,tblHits[Confidence],"High").
    .PARAMETER Value
      The value now (a whole number).
    .OUTPUTS
      [pscustomobject]
    #>
    param([string]$Formula, [long]$Value = 0)
    [pscustomobject]@{ XlsxFormula = $Formula; Value = $Value }
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
    <#
    .SYNOPSIS
      One cell: a whole number or a date as a number, a New-XlsxFormula as a formula with its value, anything else as an
      inline string (never a formula). An empty cell is not written, so Excel counts it as blank.
    .PARAMETER Ref
      The cell, for example B7.
    .PARAMETER Value
      The cell's value.
    .PARAMETER Style
      The style number (see $StylesXml); 0 for none.
    .OUTPUTS
      [string] the cell XML, or empty.
    #>
    param([string]$Ref, $Value, [int]$Style)
    $styleAttr = ''
    if ($Style -gt 0) { $styleAttr = ' s="' + $Style + '"' }
    if ($null -eq $Value -or ($Value -is [string] -and $Value -eq '')) { return '' }
    if ($Value -is [int] -or $Value -is [long]) { return '<c r="{0}"{1}><v>{2}</v></c>' -f $Ref, $styleAttr, $Value }
    if ($Value -is [double]) { return '<c r="{0}"{1}><v>{2}</v></c>' -f $Ref, $styleAttr, $Value.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value.PSObject.Properties['XlsxFormula']) {
        return '<c r="{0}"{1}><f>{2}</f><v>{3}</v></c>' -f $Ref, $styleAttr, (ConvertTo-XmlText $Value.XlsxFormula), $Value.Value
    }
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
        $from = if ($Sheet.LinkTarget -ge 0) { $Sheet.LinkTarget } else { $Sheet.LinkColumn }
        if ($Sheet.LinkColumn -ge $cells.Count -or $from -ge $cells.Count) { continue }
        $target = [string]$cells[$from]
        if ($target -notmatch '^https://\S+$' -or $target.Length -gt $script:MaxLinkChars) { continue }
        $n++
        [pscustomobject]@{ Ref = $letter + $rowNumber; Id = 'rId' + $n; Target = $target }
    }
}

function ConvertTo-SheetRelsXml {
    # The relationships part of one tab: one external target per link, and its table part (TableNumber > 0).
    param([object[]]$Links, [int]$TableNumber = 0)
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    foreach ($link in $Links) {
        $xml += '<Relationship Id="{0}" Type="{1}/hyperlink" Target="{2}" TargetMode="External"/>' -f $link.Id, $script:RelNs, (ConvertTo-XmlText $link.Target)
    }
    if ($TableNumber -gt 0) { $xml += '<Relationship Id="rIdTable" Type="{0}/table" Target="../tables/table{1}.xml"/>' -f $script:RelNs, $TableNumber }
    return $xml + '</Relationships>'
}

function ConvertTo-TableXml {
    <#
    .SYNOPSIS
      The table part of a tab (ECMA-376 table): its range, its filter buttons, one column per header cell, banded rows.
    .PARAMETER Sheet
      Output of New-XlsxSheet, with Table set.
    .PARAMETER Number
      The table's number in the workbook (1, 2 ...).
    .OUTPUTS
      [string]
    #>
    param($Sheet, [int]$Number)
    $range = 'A1:{0}{1}' -f (Get-ColumnLetter $Sheet.Columns.Count), [Math]::Max(2, $Sheet.Rows.Count + 1)
    $xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{0}" id="{1}" name="{2}" displayName="{2}" ref="{3}" ' +
        'totalsRowShown="0"><autoFilter ref="{3}"/><tableColumns count="{4}">') -f $script:MainNs, $Number, $Sheet.Table, $range, $Sheet.Columns.Count
    for ($i = 0; $i -lt $Sheet.Columns.Count; $i++) { $xml += '<tableColumn id="{0}" name="{1}"/>' -f ($i + 1), (ConvertTo-XmlText $Sheet.Columns[$i]) }
    return $xml + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
}

function Get-SheetTopXml {
    <#
    .SYNOPSIS
      The start of a tab's XML, before its cells: fit to the page width, the frozen header row (and columns), the column widths.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .OUTPUTS
      [string]
    #>
    param($Sheet)
    $top = New-Object System.Text.StringBuilder
    [void]$top.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    if ($Sheet.List -and $Sheet.FreezeColumns -gt 0) {
        $pane = ('<sheetViews><sheetView workbookViewId="0"><pane xSplit="{0}" ySplit="1" topLeftCell="{1}2" activePane="bottomRight" ' +
            'state="frozen"/></sheetView></sheetViews>') -f $Sheet.FreezeColumns, (Get-ColumnLetter ($Sheet.FreezeColumns + 1))
        [void]$top.Append($pane)
    }
    elseif ($Sheet.List) {
        [void]$top.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    }
    if ($Sheet.Widths) {
        [void]$top.Append('<cols>')
        for ($i = 0; $i -lt $Sheet.Widths.Count; $i++) {
            [void]$top.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $Sheet.Widths[$i]))
        }
        [void]$top.Append('</cols>')
    }
    return $top.ToString()
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
    [void]$xml.Append((Get-SheetTopXml $Sheet))
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
            $value = $cells[$i]
            $when = [datetime]::MinValue
            if ($Sheet.DateColumns -contains $i -and $value -is [string] -and [datetime]::TryParseExact($value, $script:DateFormats,
                    [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$when)) { $value = $when.ToOADate(); $style = 4 }
            if ($i -eq $Sheet.FillColumn -and $script:FillStyle.ContainsKey([string]$value)) { $style = $script:FillStyle[[string]$value] }
            [void]$xml.Append((ConvertTo-CellXml $ref $value $style))
        }
        [void]$xml.Append('</row>')
    }
    [void]$xml.Append('</sheetData>')
    if ($Sheet.List -and $Sheet.Columns.Count -gt 0 -and -not $Sheet.Table) {   # a table has its own filter buttons
        [void]$xml.Append(('<autoFilter ref="A1:{0}{1}"/>' -f (Get-ColumnLetter $Sheet.Columns.Count), [Math]::Max(2, $rowNumber)))
    }
    [void]$xml.Append((ConvertTo-ChoicesXml $Sheet ([Math]::Max(2, $rowNumber))))
    if ($Links.Count -gt 0) {
        [void]$xml.Append('<hyperlinks>')
        foreach ($link in $Links) { [void]$xml.Append(('<hyperlink ref="{0}" r:id="{1}"/>' -f $link.Ref, $link.Id)) }
        [void]$xml.Append('</hyperlinks>')
    }
    [void]$xml.Append($script:PrintXml)
    if ($Sheet.List -and $Sheet.Table) { [void]$xml.Append('<tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>') }
    [void]$xml.Append('</worksheet>')
    return $xml.ToString()
}

function ConvertTo-ChoicesXml {
    <#
    .SYNOPSIS
      The checks of one tab (ECMA-376 dataValidations, after autoFilter and before hyperlinks): a date check per column in
      DateChecks and a dropdown list per column in Choices, for rows 2 to the last row; blank cells are allowed.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .PARAMETER LastRow
      The last row number.
    .OUTPUTS
      [string] empty when the tab has no choices.
    #>
    param($Sheet, [int]$LastRow)
    $checks = @($Sheet.DateChecks)
    if ((-not $Sheet.Choices -or $Sheet.Choices.Count -eq 0) -and $checks.Count -eq 0) { return '' }
    $xml = '<dataValidations count="' + ($Sheet.Choices.Count + $checks.Count) + '">'
    foreach ($column in $checks) {   # 36525 = 1999-12-31 as an Excel date number: 2000-01-01 and later
        $xml += ('<dataValidation type="date" operator="greaterThan" allowBlank="1" showErrorMessage="1" errorTitle="Review date" ' +
            'error="Type a date of 2000-01-01 or later, for example 2026-11-24." sqref="{0}2:{0}{1}"><formula1>36525</formula1></dataValidation>') -f (Get-ColumnLetter ($column + 1)), $LastRow
    }
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
    $tables = @{}   # tab number -> table number
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        if (-not ($Sheets[$i - 1].List -and $Sheets[$i - 1].Table)) { continue }
        $tables[$i] = $tables.Count + 1
        $types += '<Override PartName="/xl/tables/table' + $tables[$i] + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
    }
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::CreateNew)   # CreateNew: never overwrite
    $zip = $null
    $written = $false
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        Add-ZipText $zip '[Content_Types].xml' ($types + '</Types>')
        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">' +
            '<Relationship Id="rId1" Type="' + $script:RelNs + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        Add-ZipText $zip 'xl/workbook.xml' ($book + '</sheets><calcPr fullCalcOnLoad="1"/></workbook>')   # formulas recalculate on open
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rels
        Add-ZipText $zip 'xl/styles.xml' $script:StylesXml
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $links = @(Get-SheetLinks $Sheets[$i - 1])
            Add-ZipText $zip ('xl/worksheets/sheet' + $i + '.xml') (ConvertTo-SheetXml $Sheets[$i - 1] $links)
            $table = if ($tables.ContainsKey($i)) { $tables[$i] } else { 0 }
            if ($links.Count -gt 0 -or $table) { Add-ZipText $zip ('xl/worksheets/_rels/sheet' + $i + '.xml.rels') (ConvertTo-SheetRelsXml $links $table) }
            if ($table) { Add-ZipText $zip ('xl/tables/table' + $table + '.xml') (ConvertTo-TableXml $Sheets[$i - 1] $table) }
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

# ==== Report.psm1 ====
# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# PII types from most to least damaging, for sorting and for the order of the tabs. AFI 33-332 gives no ranking
# (A3.1 lists only the factors for judging harm), so this is the Administrator's order (2026-11-20): SSN and date of birth
# first, then identity documents and accounts, then contact details.
$script:SeverityOrder = @('SSN', 'Date of birth', 'Passport', 'Payment card', 'DoD ID', 'Home address', 'Name', 'Phone', 'Personal e-mail')

$script:TopItems = 10   # how many files the screen summary lists (the workbook lists all)
# Under the table of types: what High, Medium and Low mean (the Administrator, 2026-11-21: plain words for a unit member).
$script:ConfidenceKey = @('High = labelled ("SSN:") or a sure pattern. Medium = right form, no label.', 'Low = weak match, often not PII.')

# Screen colours (the Administrator, 2026-11-20: less grey, more format), of the 16 console colours: Cyan = headings and labels,
# White = values, Green = clean or done, Yellow = check by hand and Medium, Red = possible PII, High and errors,
# Gray = hints and folder-row labels, DarkGray = rules and table borders, DarkCyan (teal) = folder names (the Administrator, 2026-11-20).
$script:ScreenColors = @{ Title = 'Cyan'; Label = 'Cyan'; Text = 'White'; Good = 'Green'; Accent = 'Yellow'; Alert = 'Red'
    Muted = 'Gray'; Rule = 'DarkGray'; Folder = 'DarkCyan' }

# Not-scanned reasons in groups for the counts on screen and on the Summary tab: a pattern on the reason, its group.
# A file type that is not read is counted per extension (Get-NotScannedGroup).
$script:NotScannedGroups = [ordered]@{
    '^Sensitivity label' = 'Sensitivity label with encryption'
    '^Encrypted' = 'Encrypted (password or author limits)'
    '^An old Office file' = 'Old Office format with a new name'
    '^Larger than' = 'Too large'
    '^Empty file' = 'Empty file'
    '^No text in this PDF' = 'PDF without text (scanned image?)'
    'web page' = 'Download gave a web page'
    '^Download failed' = 'Download failed'
    '^Reading took longer than' = 'File time limit reached (file may be fine)'
    'longer than the time limit|longer than its time limit' = 'PDF time limit reached (file may be fine)'
    '^Part of this PDF is not well formed|^Part of this PDF could not be read' = 'PDF parts not read (file may open fine)'
    '^Part of this PDF' = 'PDF parts not unpacked'
    '^Not a valid (Office file|PDF)' = 'Not a valid file (damaged or wrong name)'
    '^Could not read' = 'Could not read (in use, damaged or protected)'
    'could not be fully read|^A font in this PDF|^Some characters in this PDF' = 'PDF fonts not fully readable'
    '^Files attached inside this PDF' = 'PDF attachments'
    '^Image: Windows OCR|^Scanned PDF: Image: Windows OCR' = 'Images: OCR not available'
    '^Image:|^Scanned PDF' = 'Images and scans not fully read'
    '^Old Excel file' = 'Old Excel: numbers not read'
    'old Office or Outlook' = 'Old Office or Outlook not readable'
    '^An archive or e-mail inside another|in this archive|Too large to unpack|Could not unpack' = 'Archives: files not unpacked'
}

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the folder (masked by the caller).
    .PARAMETER Link
      A link to the item, or empty.
    .PARAMETER Modified
      When the item last changed, or empty.
    .PARAMETER Detector
      The PII type, or the note word.
    .PARAMETER Confidence
      High, Medium, Low, or empty for a note.
    .PARAMETER MaskedValue
      The masked value, or empty.
    .PARAMETER Context
      The masked snippet, or the reason for a note.
    .OUTPUTS
      [pscustomobject] Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link, and the
      accountability fields (empty until Set-RowAccountability fills them): ModifiedBy, Created, CreatedBy, AddedOn, AddedBy,
      HistoryNote, OwnPermissions, AccessEntries (plan step 5: a web file with a hit).
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source = $Source; Channel = $Channel; Item = $Item; Detector = $Detector; Confidence = $Confidence
        MaskedValue = $MaskedValue; Context = $Context; Modified = $Modified; Link = $Link
        ModifiedBy = ''; Created = ''; CreatedBy = ''; AddedOn = ''; AddedBy = ''; HistoryNote = ''
        OwnPermissions = ''; AccessEntries = ''
    }
}

function Add-IncompleteRow {
    <#
    .SYNOPSIS
      Adds one INCOMPLETE row when the scan did not reach its end and no INCOMPLETE row is there yet.
      Called from the script's 'finally' block, which also runs after Ctrl+C (when 'catch' does not).
    .PARAMETER Rows
      The list of report rows (changed in place).
    .PARAMETER Finished
      True when every step ran.
    .OUTPUTS
      None.
    #>
    param($Rows, [bool]$Finished)
    if ($Finished) { return }
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) { return }
    $Rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
}

function Write-Ui {
    <#
    .SYNOPSIS
      Writes one piece of screen text in a palette colour ($ScreenColors). Every screen line of Ion goes through here.
    .PARAMETER Text
      The text (already masked by the caller).
    .PARAMETER Kind
      Title, Label, Text, Good, Accent, Alert, Muted, Rule or Folder.
    .PARAMETER NoNewline
      Keep writing on the same line (for a line with several colours).
    #>
    param([AllowEmptyString()][string]$Text, [ValidateSet('Title', 'Label', 'Text', 'Good', 'Accent', 'Alert', 'Muted', 'Rule', 'Folder')][string]$Kind = 'Text',
          [switch]$NoNewline)
    Write-Host $Text -ForegroundColor $script:ScreenColors[$Kind] -NoNewline:$NoNewline
}

function Get-IonMark {
    # The app's mark as text: the Greek letter gamma (Logo Studies, the Administrator's choice of 2026-11-20), written as a code so the script stays ASCII.
    return [string][char]0x3B3
}

function Write-Banner {
    # A heading line "==== <text> ====" in the title colour. With -Mark (the start line next to "Ion" only; the Administrator,
    # 2026-11-21): "==== <mark> <text> ====", the mark in the mark's eye colour (yellow).
    param([string]$Text, [switch]$Mark)
    if (-not $Mark) { Write-Ui ('==== ' + $Text + ' ====') 'Title'; return }
    Write-Ui '==== ' 'Title' -NoNewline; Write-Ui (Get-IonMark) 'Accent' -NoNewline; Write-Ui (' ' + $Text + ' ====') 'Title'
}

function Get-SeverityIndex {
    # Where a PII type stands in $SeverityOrder (0 = worst); a type not in the list comes last.
    param([string]$Type)
    $index = [Array]::IndexOf($script:SeverityOrder, $Type)
    if ($index -lt 0) { return $script:SeverityOrder.Count }
    return $index
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: its worst type, highest confidence, number of hits and types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Worst, Highest, Hits, Channel, Source, Item, Detectors; worst type first, then highest confidence.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Channel, Item)) {
        $first = $group.Group[0]
        $worst = ($group.Group | Sort-Object { Get-SeverityIndex $_.Detector }, { -$script:Rank[$_.Confidence] } | Select-Object -First 1)
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Worst = $worst.Detector; Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = ((@($group.Group | Select-Object -ExpandProperty Detector -Unique) | Sort-Object { Get-SeverityIndex $_ }) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { Get-SeverityIndex $_.Worst } }, @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true },
        @{ Expression = 'Hits'; Descending = $true })
}

function Get-TypeCounts {
    <#
    .SYNOPSIS
      The hits per PII type, worst type first, split by confidence.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, High, Medium, Low, Total.
    #>
    param([object[]]$Hits)
    foreach ($group in ($Hits | Group-Object Detector | Sort-Object { Get-SeverityIndex $_.Name })) {
        $count = @{ High = 0; Medium = 0; Low = 0 }
        foreach ($hit in $group.Group) { $count[$hit.Confidence]++ }
        [pscustomobject]@{ Type = $group.Name; High = $count.High; Medium = $count.Medium; Low = $count.Low; Total = $group.Count }
    }
}

function Get-NotScannedGroup {
    <#
    .SYNOPSIS
      The group of one not-scanned row: a type not read by its extension, else by its reason ($NotScannedGroups).
    .PARAMETER Row
      A report row that is not a hit (NOT SCANNED, BLOCKED or INCOMPLETE).
    .OUTPUTS
      [string] the group name.
    #>
    param($Row)
    if ($Row.Detector -ne 'NOT SCANNED') { return $Row.Detector }   # BLOCKED, INCOMPLETE
    if ($Row.Context -match '^File type') {
        $extension = [System.IO.Path]::GetExtension($Row.Item).ToLower()
        if (-not $extension) { $extension = '(no extension)' }
        return 'Type not read: ' + $extension
    }
    foreach ($pattern in $script:NotScannedGroups.Keys) { if ($Row.Context -match $pattern) { return $script:NotScannedGroups[$pattern] } }
    return 'Other'
}

function Get-NotScannedCounts {
    <#
    .SYNOPSIS
      How many files were not scanned, per reason, most first (the Administrator, 2026-11-20: see at once what to check by hand).
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Reason, Count.
    #>
    param([object[]]$Rows)
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    foreach ($group in ($notes | Group-Object { Get-NotScannedGroup $_ } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)) {
        [pscustomobject]@{ Reason = $group.Name; Count = $group.Count }
    }
}

function Get-ScanMetrics {
    <#
    .SYNOPSIS
      The scan's totals, as label and value pairs (the screen summary and the Summary tab show the same).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([object[]]$Rows, [int]$ItemsScanned)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $metrics = [ordered]@{}
    $metrics['Files read'] = $ItemsScanned
    $metrics['Files with possible PII'] = @(Get-ReportSummary $Rows).Count
    $metrics['Possible PII hits'] = $hits.Count
    $metrics['High / Medium / Low'] = '{0} / {1} / {2}' -f @($hits | Where-Object { $_.Confidence -eq 'High' }).Count,
        @($hits | Where-Object { $_.Confidence -eq 'Medium' }).Count, @($hits | Where-Object { $_.Confidence -eq 'Low' }).Count
    $metrics['Not scanned / incomplete'] = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) }).Count
    return $metrics
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The summary for the screen, line by line with its palette colour: the scan's details and totals, a table of
      possible PII per type (worst first; red with High hits, yellow with Medium, white otherwise), a table of the
      files with the most serious hits, and the not-scanned files per reason. Every hit is listed in the workbook.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [pscustomobject] Text, Kind (a Write-Ui colour), LabelLength (characters at the start in the label colour), one per line.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $out = New-Object System.Collections.Generic.List[object]
    # Adds one screen line with its colour; the first LabelLength characters are written in the label colour.
    $add = { param($text, $kind, $labelLength = 0) $out.Add([pscustomobject]@{ Text = $text; Kind = $kind; LabelLength = $labelLength }) }
    & $add '' 'Text'; & $add '==== Results ====' 'Title'
    # A scan that stopped early says so first, in red: its results cover only the files read so far.
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) {
        & $add ('INCOMPLETE: the scan stopped early. Only the {0} {1} read are in these results; scan again.' -f $ItemsScanned, @('files', 'file')[[int]($ItemsScanned -eq 1)]) 'Alert'
    }
    # The values start in one column, after the longest label (labels differ: site, drive, timing lines).
    $pairs = @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())
    $labelWidth = ($pairs | ForEach-Object { ([string]$_.Key).Length + 1 } | Measure-Object -Maximum).Maximum
    foreach ($pair in $pairs) {
        & $add (($pair.Key + ':').PadRight($labelWidth) + ' ' + $pair.Value) 'Text' ($labelWidth + 1)
    }
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    if ($hits.Count -eq 0) { & $add '' 'Text'; & $add 'No possible PII found in the files that were read.' 'Good' }
    else {
        $types = @(Get-TypeCounts $hits)
        & $add '' 'Text'; & $add 'POSSIBLE PII BY TYPE (worst first)' 'Title'
        $kinds = @($types | ForEach-Object { if ($_.High -gt 0) { 'Alert' } elseif ($_.Medium -gt 0) { 'Accent' } else { 'Text' } })
        Get-TableLines @('Type', 'High', 'Medium', 'Low', 'Total') @($types | ForEach-Object { , @($_.Type, $_.High, $_.Medium, $_.Low, $_.Total) }) $kinds |
            ForEach-Object { $out.Add($_) }
        foreach ($keyLine in $script:ConfidenceKey) { & $add $keyLine 'Muted' }
        $top = @(Get-ReportSummary $Rows | Select-Object -First $script:TopItems)
        & $add '' 'Text'; & $add ('FILES WITH THE MOST SERIOUS HITS (first {0}; all are in the workbook)' -f $script:TopItems) 'Title'
        $kinds = @($top | ForEach-Object { if ($_.Highest -eq 'High') { 'Alert' } else { 'Accent' } })
        Get-TableLines @('Worst', 'Highest', 'Hits', $FolderLabel, 'File') @($top | ForEach-Object { , @($_.Worst, $_.Highest, $_.Hits, $_.Channel, $_.Item) }) $kinds |
            ForEach-Object { $out.Add($_) }
    }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        & $add '' 'Text'; & $add 'NOT SCANNED: CHECK THESE BY HAND (the workbook lists each file)' 'Title'
        Get-TableLines @('Reason', 'Files') @($skipped | ForEach-Object { , @($_.Reason, $_.Count) }) @($skipped | ForEach-Object { 'Accent' }) |
            ForEach-Object { $out.Add($_) }
    }
    return $out.ToArray()
}

function Set-RowAccountability {
    <#
    .SYNOPSIS
      Fills the accountability fields of a file's rows: who last changed it, when it was made and by whom, and who
      put the PII in and when (where the file's history shows it; else the file's maker or "not recorded").
    .PARAMETER Rows
      The file's rows (New-ReportRow).
    .PARAMETER Who
      @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; HistoryNote; OwnPermissions; AccessEntries; ByHit } (any may be missing); ByHit:
      "type|masked value" -> @{ AddedOn; AddedBy; Note } for each hit dated from the file's history.
    #>
    param([object[]]$Rows, [hashtable]$Who)
    if (-not $Who) { return }
    foreach ($row in $Rows) {
        foreach ($key in @('ModifiedBy', 'Created', 'CreatedBy', 'AddedOn', 'AddedBy', 'HistoryNote', 'OwnPermissions', 'AccessEntries')) {
            if ($Who.ContainsKey($key) -and $Who[$key]) { $row.$key = [string]$Who[$key] }
        }
        $hitKey = $row.Detector + '|' + $row.MaskedValue
        if ($Who.ContainsKey('ByHit') -and $Who.ByHit -and $Who.ByHit.ContainsKey($hitKey)) {
            $entry = $Who.ByHit[$hitKey]
            $row.AddedOn = [string]$entry.AddedOn; $row.AddedBy = [string]$entry.AddedBy
            if ($entry.ContainsKey('Note')) { $row.HistoryNote = [string]$entry.Note }
        }
    }
}

function Add-PartRows {
    <#
    .SYNOPSIS
      The rows of one file inside a zip or an e-mail, named "<file>/<inner path>" (PII in the name masked): its
      hits, a hit for PII in its name, and a NOT SCANNED row when it was not read or read only in part.
    .PARAMETER Part
      One result of Get-TextFromBytes Parts (with its Name).
    .PARAMETER Channel
      The channel name as the report shows it.
    .PARAMETER Item
      The container file as the report shows it.
    .PARAMETER Link
      The container file's link (an inner file has no link of its own).
    .PARAMETER Modified
      The container file's change time.
    .PARAMETER Rows
      The list that receives the rows.
    #>
    param($Part, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified, $Rows)
    $name = $Item + '/' + (Protect-PiiName $Part.Name)
    foreach ($hit in @(Find-PiiInText $Part.Name)) {
        $Rows.Add((New-ReportRow 'File name' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the name of a file inside it'))
    }
    if ($Part.Status -ne 'OK') { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Reason))); return }
    if ($Part.Note) { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Note))) }
    foreach ($hit in @(Find-PiiInText $Part.Text)) {
        $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
}

# ==== ReportBook.psm1 ====
# =====================================================================================
# Settings
# =====================================================================================

# The tool's name on screen and in the workbook (the Administrator, 2026-11-20).
$script:ToolTitle = 'Ion - DoD365 PII Reporter'

$script:MaxNameInPath = 40   # the site name is in the path twice: at most 40 characters keeps it under 260 (review of PR 41)
# C6, the review columns the person fills in (the Administrator, 2026-11-21), each with its dropdown (empty = free text); Ion never fills them.
# Order: the answers first, then who and when (the Administrator, 2026-11-24, the approved preview).
$script:ReviewColumns = [ordered]@{ 'Is it PII?' = @('Yes', 'No', 'Not sure')
    'Action taken' = @('Removed', 'Moved', 'Access limited', 'Left (not PII)', 'Other')
    'Reported to Privacy Manager' = @('Yes', 'No', 'Not needed'); 'Reviewed by' = @(); 'Review date' = @() }
# What to do with PII found where it should not be (rule source: AFI 33-332 para 3.1.2 as changed by AFI33-332_DAFGM2026-01,
# 7 July 2026; edition: docs/DECISIONS.md). The wording is the README's "If it finds PII" section; scripts/build_site.py --check fails when they differ (the Administrator, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01 (7 July 2026), ' +
                       'page 1, second paragraph).'
# "Start here" on the Summary tab: 4 steps (the approved workbook list of 2026-11-21, item 4 of
# docs/reviews/2026-11-24-ion-detector-handoff.md). STE: each step is one screen, its actions numbered, then what the
# person sees; checked by hand, as scripts/check_instructions.py cannot read text in code. Reporting a breach is the
# "If you find PII" row (one home).
$script:StartHere = @(
    @('Start here: 1', ('On the tab Hits, do these actions: (1) Click the filter button of the column Confidence. (2) Clear ' +
        '(Select All). (3) Select High and Medium. (4) Click OK. You see the High and Medium rows.')),
    @('Start here: 2', 'In a row, click the name in the column File. The file opens.'),
    @('Start here: 3', 'In the file, find the words that the cell Context shows next to the masked value. You see the full value.'),
    @('Start here: 4', ('On the tab Hits, in the same row, do these actions: (1) In Is it PII?, select Yes, No or Not sure. ' +
        '(2) In Action taken, select the action that you did. (3) In Reported to Privacy Manager, select Yes, No or Not needed. ' +
        '(4) In Reviewed by, type your name. (5) In Review date, type the date of your check. The row is complete.')))
# Shown as two lines, the advice and its rule source (the Administrator, 2026-11-21); one line if no "(AFI ...)." part is found.
$script:BreachLines = @($script:BreachAdvice)
if ($script:BreachAdvice -match '^(.*) \((AFI .*)\)\.$') { $script:BreachLines = @(($Matches[1] + '.'), ('Source: ' + $Matches[2] + '.')) }

# =====================================================================================
# The workbook
# =====================================================================================

function Get-ResultPath {
    <#
    .SYNOPSIS
      A new workbook path: <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx (the Administrator, 2026-11-21), with (2), (3) ...
      if the name exists, so a result is never overwritten. Creates the folder. The site name is masked like any name the
      tool shows (found PII only); characters Windows refuses in a folder name are replaced.
    .PARAMETER OutputFolder
      The base folder.
    .PARAMETER SiteUrl
      The site address; its last part (teams/<name>) names the site folder.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$SiteUrl)
    $siteName = [uri]::UnescapeDataString(($SiteUrl.TrimEnd('/') -split '/')[-1])
    $siteName = ((Protect-PiiName $siteName) -replace '[\\/:*?"<>|]', '-').Trim(' .')
    if (-not $siteName) { $siteName = 'site' }
    if ($siteName.Length -gt $script:MaxNameInPath) { $siteName = $siteName.Substring(0, $script:MaxNameInPath).Trim(' .') }   # MAX_PATH 260
    if ($siteName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $siteName = '_' + $siteName }   # names Windows keeps for devices
    $now = Get-Date
    $folder = Join-Path $OutputFolder $siteName
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
    $base = 'Ion {0} {1}' -f $siteName, $now.ToString('yyyy-MM-dd HHmm')
    $path = Join-Path $folder ($base + '.xlsx')
    $copy = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $folder ('{0} ({1}).xlsx' -f $base, $copy); $copy++ }
    return $path
}

function Get-HitCells {
    # One row of the Hits tab, in the order of its columns: ID, what was found, the review columns, then who and when.
    param($Hit, [int]$Number)
    @(('H{0:000}' -f $Number), $Hit.Detector, $Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context) +
        @('') * $script:ReviewColumns.Count +
        @($Hit.Modified, $Hit.ModifiedBy, $Hit.AddedOn, $Hit.AddedBy, $Hit.Created, $Hit.CreatedBy, $Hit.HistoryNote, $Hit.OwnPermissions,
            $Hit.AccessEntries, $Hit.Source, $Hit.Link)
}

function Get-FileRows {
    <#
    .SYNOPSIS
      The rows of the Files tab: one per file with hits, worst first: its counts, its types, how many hits a reviewer
      has answered "Is it PII?" for (a formula on the Hits table), who saved it, its history notes and who has access.
    .PARAMETER Hits
      The hit rows, in the Hits tab's order.
    .PARAMETER FolderLabel
      The name of the folder column (Channel or Folder), as in the Hits table.
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Hits, [string]$FolderLabel)
    $groups = @($Hits | Group-Object { $_.Channel + "`n" + $_.Item } | ForEach-Object {
        $g = @($_.Group)
        [pscustomobject]@{ First = $g[0]; Count = $g.Count; High = @($g | Where-Object { $_.Confidence -eq 'High' }).Count
            Medium = @($g | Where-Object { $_.Confidence -eq 'Medium' }).Count; Low = @($g | Where-Object { $_.Confidence -eq 'Low' }).Count
            Types = (@($g | ForEach-Object { $_.Detector } | Select-Object -Unique) -join ', ')
            AddedBy = (@($g | ForEach-Object { $_.AddedBy } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
            Notes = (@($g | ForEach-Object { $_.HistoryNote } | Where-Object { $_ } | Select-Object -Unique) -join '; ') }
    } | Sort-Object @{ Expression = 'High'; Descending = $true }, @{ Expression = 'Medium'; Descending = $true },
        @{ Expression = 'Low'; Descending = $true }, @{ Expression = { $_.First.Channel } }, @{ Expression = { $_.First.Item } })
    $row = 1
    foreach ($f in $groups) {
        $row++
        $reviewed = New-XlsxFormula ('SUMPRODUCT((tblHits[{0}]=A{1})*(tblHits[File]=B{1})*(tblHits[Is it PII?]<>""))' -f $FolderLabel, $row) 0
        , @($f.First.Channel, $f.First.Item, [long]$f.Count, [long]$f.High, [long]$f.Medium, [long]$f.Low, $f.Types, $reviewed,
            $f.First.Modified, $f.First.ModifiedBy, $f.AddedBy, $f.Notes, $f.First.OwnPermissions, $f.First.AccessEntries, $f.First.Link)
    }
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first; Hits (every hit, one Excel Table: filter, sort, pivot) and Files (one row
      per file) when there are hits; then Not scanned, By account and Method (the Administrator, 2026-11-24: as previewed).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the Summary tab.
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site): the name of the folder column.
    .OUTPUTS
      [object[]] New-XlsxSheet objects.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $hits = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -gt 0 } |
        Sort-Object @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true }, @{ Expression = { Get-SeverityIndex $_.Detector } }, Channel, Item)
    # The review columns right after the hit, then who and when (accountability names: the Administrator, 2026-11-21; the Method tab explains them).
    $columns = @('ID', 'Type', 'Confidence', $FolderLabel, 'File', 'Masked value', 'Context') + @($script:ReviewColumns.Keys) +
        @('Last saved', 'Last saved by', 'PII first saved', 'PII first saved by', 'Created', 'Created by', 'History note', 'Own permissions',
            'Access entries', 'Found in', 'Link')
    $widths = @(7, 15, 11, 16, 40, 16, 50) + @(16) * $script:ReviewColumns.Count + @(16, 22, 16, 22, 16, 22, 40, 17, 34, 10, 50)
    $choices = @{}; $offset = 7
    foreach ($name in @($script:ReviewColumns.Keys)) { if ($script:ReviewColumns[$name]) { $choices[$offset] = $script:ReviewColumns[$name] }; $offset++ }
    $after = 7 + $script:ReviewColumns.Count   # the first column after the review columns
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header $hits) @(34, 70, 9, 9, 9, 10)))
    if ($hits.Count -gt 0) {
        $n = 0
        $sheets.Add((New-XlsxSheet 'Hits' $columns @($hits | ForEach-Object { $n++; , (Get-HitCells $_ $n) }) $widths -List -Table 'tblHits' `
            -LinkColumn 4 -LinkTarget ($columns.Count - 1) -Choices $choices -WrapColumns @(6) -DateColumns @($after, ($after + 2), ($after + 4)) `
            -DateChecks @(7 + @($script:ReviewColumns.Keys).IndexOf('Review date')) `
            -FreezeColumns 5 -FillColumn 2))
        $sheets.Add((New-XlsxSheet 'Files' @($FolderLabel, 'File', 'Hits', 'High', 'Medium', 'Low', 'Types', 'Reviewed', 'Last saved',
            'Last saved by', 'PII first saved by', 'History note', 'Own permissions', 'Access entries', 'Link') @(Get-FileRows $hits $FolderLabel) `
            @(16, 40, 7, 7, 9, 7, 40, 10, 16, 22, 30, 40, 17, 34, 50) -List -Table 'tblFiles' -LinkColumn 1 -LinkTarget 14 -DateColumns @(8) -FreezeColumns 2))
    }
    $notes = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -eq 0 })
    if ($notes.Count -gt 0) {
        # A Group column, so the tab filters by the groups the Summary counts (the Administrator, 2026-11-21).
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'File', 'Reason', 'Group', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, (Get-NotScannedGroup $_), $_.Link) }) `
            @(14, 22, 45, 70, 30, 50) -List -Table 'tblNotScanned' -LinkColumn 2 -LinkTarget 5))
    }
    $accounts = @(Get-AccountRows $hits)   # C6 (the Administrator, 2026-11-21): whom to contact first
    if ($accounts.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'By account' @('PII first saved by', 'Files', 'High', 'Medium', 'Low') $accounts @(40, 9, 9, 9, 9) -List -Table 'tblAccounts'))
    }
    $method = @(Get-MethodRows | Where-Object { $accounts.Count -gt 0 -or $_[0] -ne 'By account tab' })   # no row for a tab that is not there
    $sheets.Add((New-XlsxSheet 'Method' @('Topic', 'How Ion works it out') $method @(34, 110) -List -Table 'tblMethod' -WrapColumns @(1)))   # the Administrator, 2026-11-21
    return $sheets.ToArray()
}

function Get-TypeTableRows {
    <#
    .SYNOPSIS
      The Summary tab's hits per type (worst first): High, Medium, Low, Total and Reviewed, as formulas on the Hits table
      with Ion's values, and a Total row.
    .PARAMETER Hits
      The hit rows (the Hits tab).
    .PARAMETER StartRow
      The sheet row of the table's header row.
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Hits, [int]$StartRow)
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add(@('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total', 'Reviewed'))
    $first = $StartRow + $list.Count   # the sheet row of the first type (row 1 is the title)
    foreach ($type in @(Get-TypeCounts $Hits)) {
        $r = $StartRow + $list.Count
        $cells = @($type.Type)
        foreach ($level in @(@('High', $type.High), @('Medium', $type.Medium), @('Low', $type.Low))) {
            $cells += New-XlsxFormula ('COUNTIFS(tblHits[Type],$A{0},tblHits[Confidence],"{1}")' -f $r, $level[0]) $level[1]
        }
        $list.Add($cells + @((New-XlsxFormula ('SUM(B{0}:D{0})' -f $r) $type.Total), (New-XlsxFormula ('COUNTIFS(tblHits[Type],$A{0},tblHits[Is it PII?],"<>")' -f $r) 0)))
    }
    if ($Hits.Count -eq 0) { $list.Add(@('None found in the files that were read.')) }
    else {
        $last = $StartRow + $list.Count - 1
        $total = @('Total')
        foreach ($c in @('B', 'C', 'D', 'E', 'F')) {
            $value = switch ($c) { 'B' { @($Hits | Where-Object { $_.Confidence -eq 'High' }).Count } 'C' { @($Hits | Where-Object { $_.Confidence -eq 'Medium' }).Count }
                'D' { @($Hits | Where-Object { $_.Confidence -eq 'Low' }).Count } 'E' { $Hits.Count } default { 0 } }
            $total += New-XlsxFormula ('SUM({0}{1}:{0}{2})' -f $c, $first, $last) $value
        }
        $list.Add($total)
    }
    return $list.ToArray()   # one row per item (each row stays an array)
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the reviewer's progress, the hits per type (worst
      first; counts as formulas on the Hits table, so they follow a reviewer's filter-free edits), the not-scanned files
      per reason, then how to read the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER Hits
      The hit rows (the Hits tab).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [object[]]$Hits)
    $out = New-Object System.Collections.Generic.List[object]
    if ($Hits.Count -gt 0) { foreach ($step in $script:StartHere) { $out.Add(@($step[0], $step[1])) }; $out.Add(@('')) }
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) {
        if ($pair.Key -ne 'High / Medium / Low') { $out.Add(@($pair.Key, $pair.Value)) }   # the type table below has these as numbers
        if ($pair.Key -eq 'Run') { $out.Add(@('Finished', (Get-Date -Format 'yyyy-MM-dd HH:mm'))) }
    }
    if ($Hits.Count -gt 0) {
        $out.Add(@('Hits reviewed ("Is it PII?" filled in)', (New-XlsxFormula 'COUNTIF(tblHits[Is it PII?],"<>")' 0)))
        $out.Add(@('Confirmed PII ("Is it PII?" = Yes)', (New-XlsxFormula 'COUNTIF(tblHits[Is it PII?],"Yes")' 0)))
    }
    $out.Add(@(''))
    foreach ($row in @(Get-TypeTableRows $Hits ($out.Count + 2))) { $out.Add($row) }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        $out.Add(@(''))
        $out.Add(@('Not scanned: check by hand', 'Files'))
        foreach ($group in $skipped) { $out.Add(@($group.Reason, $group.Count)) }
    }
    $out.Add(@(''))
    $out.Add(@('Tabs', ('Hits: one row per hit, one Excel Table (filter, sort, pivot); fill in the review columns there. ' +
        'Files: one row per file. Then Not scanned, By account and Method (how each column is worked out).')))
    $out.Add(@('Confidence', 'A pattern guess, not proof: open each item and check it before you act.'))
    $out.Add(@('If you find PII', $script:BreachLines[0]))
    if ($script:BreachLines.Count -gt 1) { $out.Add(@('Rule source', ($script:BreachLines[1] -replace '^Source: ', ''))) }
    $out.Add(@('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.'))
    return , $out.ToArray()
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the workbook (Summary tab first). Returns its path.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the summary.
    .PARAMETER OutputFolder
      The base folder; the workbook goes in <site> under it.
    .PARAMETER SiteUrl
      The site address (names the site folder).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$OutputFolder, [string]$SiteUrl,
          [string]$FolderLabel = 'Folder')
    foreach ($line in @(Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header -FolderLabel $FolderLabel)) {
        if ($line.LabelLength -gt 0) { Write-Ui $line.Text.Substring(0, $line.LabelLength) 'Label' -NoNewline; Write-Ui $line.Text.Substring($line.LabelLength) $line.Kind }
        else { Write-Ui $line.Text $line.Kind }
    }
    # The advice comes after the saved path (the Administrator, 2026-11-21), and also when the save fails ('finally').
    try {
        $path = Get-ResultPath $OutputFolder $SiteUrl
        Export-Xlsx -Path $path -Sheets (Get-WorkbookSheets $Rows $ItemsScanned $Header $FolderLabel)
        Write-Ui ''; Write-Ui "Saved (values masked): $path" 'Good'; Write-Ui 'Handle this file as CUI: it shows where PII is.' 'Accent'
    }
    finally {
        Write-Ui ''; Write-Ui 'Confidence is a pattern guess, not proof: open each item and check it before you act.' 'Muted'
        Write-Ui $script:BreachLines[0] 'Accent'; if ($script:BreachLines.Count -gt 1) { Write-Ui $script:BreachLines[1] 'Muted' }
    }
    return $path
}

Export-ModuleMember -Function New-XlsxSheet, New-XlsxFormula, Export-Xlsx, Write-Ui, Get-IonMark, Write-Banner, Get-SeverityIndex, New-ReportRow, Add-IncompleteRow, Get-ReportSummary,
    Get-TypeCounts, Get-NotScannedGroup, Get-NotScannedCounts, Get-ScanMetrics, Get-ScanSummaryText, Add-PartRows, Set-RowAccountability, Get-WorkbookSheets, Get-ResultPath, Publish-ScanReport
