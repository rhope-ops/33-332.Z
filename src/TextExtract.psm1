<#
TextExtract.psm1 - turns file bytes into plain text, in memory.

  Purpose:     Plain text from Word, Excel and PowerPoint files (.docx/.xlsx/.pptx and macro versions), PDF,
               OpenDocument, Visio, RTF, text and HTML; images (OCR), old .doc/.xls/.ppt and .msg; and each file
               inside a .zip or an .eml (the PAM, 2026-11-20). Other types (scripts and anything else), encrypted
               files and files that cannot be read are reported as NOT SCANNED, so a person checks them by hand.
  Inputs:      file bytes and the file name (for its extension).
  Outputs:     an object with Status ('OK' or 'NOT SCANNED'), Reason, Text, Note and Parts (Get-TextFromBytes).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  PdfText.psm1 (Get-PdfText), OcrText.psm1 (Get-ImageText), OleText.psm1 (Get-OleText) and
               ContainerText.psm1 (Get-ArchiveParts, Get-EmailParts), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# =====================================================================================
# Settings: which files are read, and which parts of an Office file hold text
# =====================================================================================

# Plain text types, read as they are.
$script:PlainTextExtensions = @('.txt', '.csv', '.md', '.log', '.json', '.xml', '.tsv')

# HTML types, read and stripped of tags.
$script:HtmlExtensions = @('.html', '.htm')

# Office Open XML types (a zip of XML parts, ECMA-376) and the parts that hold their text:
# body, headers, footers, comments, notes; worksheets and cell comments; slides and speaker notes.
# A worksheet is read cell by cell (Convert-SheetXmlToText), with its shared strings put back into their cells.
$script:OfficeParts = @{
    '.docx' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.docm' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.xlsx' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.xlsm' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.pptx' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    '.pptm' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    # OpenDocument (ODF: a zip with content.xml) and Visio (pages in visio/pages); the PAM, 2026-11-20: round 3.
    '.odt' = '^content\.xml$'
    '.ods' = '^content\.xml$'
    '.odp' = '^content\.xml$'
    '.vsdx' = '^visio/pages/page\d+\.xml$'
}

# Images, read with Windows OCR (OcrText.psm1); old binary Office and Outlook files (OleText.psm1); files that
# hold other files (ContainerText.psm1): their parts are read one by one, at most $MaxContainerDepth levels deep.
$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif')
$script:OleExtensions = @('.doc', '.xls', '.ppt', '.msg')
$script:ContainerExtensions = @('.zip', '.eml')
$script:MaxContainerDepth = 2

# First bytes of a zip (every .docx/.xlsx/.pptx) and of an OLE compound file. Office saves a file in OLE form when it
# is encrypted: a password, or a sensitivity label (CUI) with encryption; an old .doc/.xls/.ppt renamed is OLE too.
# The text of an encrypted file cannot be read without the Microsoft Information Protection SDK (not on the gov device).
$script:ZipMagic = '504B'
$script:OleMagic = 'D0CF11E0A1B11AE1'
$script:EncryptedReason = 'Encrypted, password-protected or labelled with encryption (or an old Office format with a new name): check by hand'

# =====================================================================================
# Markup to text
# =====================================================================================

# Between two cells of a worksheet row: a mark that no detector reads across, so values of neighbouring cells
# (123 | 45 | 6789) never join into one SSN, card or phone number (the PAM, 2026-11-20: Excel false positives).
$script:CellBreak = ' | '

function Convert-SheetXmlToText {
    <#
    .SYNOPSIS
      A worksheet as text: one line per row, cells split by $CellBreak. Only what a person sees in a cell is kept:
      text (shared or inline), whole numbers and formula results. Formula text is dropped, and so are numbers with
      a decimal part or an exponent (a time or a calculated value: 45231.416666666664 is no SSN or phone number).
    .PARAMETER Xml
      The worksheet XML.
    .PARAMETER Shared
      The shared strings of the workbook, in order (a cell of type "s" holds an index into them).
    .OUTPUTS
      [string]
    #>
    param([string]$Xml, [string[]]$Shared)
    # Tags may carry a namespace prefix (<x:row>, <x:c>): some tools write them so, and a missed tag would be a silent miss.
    $lines = foreach ($row in [regex]::Matches($Xml, '(?s)<(?:\w+:)?row\b[^>]*>(.*?)</(?:\w+:)?row>')) {
        $cells = foreach ($cell in [regex]::Matches($row.Groups[1].Value, '(?s)<(?:\w+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?c>)')) {
            $type = [regex]::Match($cell.Groups[1].Value, '\bt="(\w+)"').Groups[1].Value
            $inner = $cell.Groups[2].Value -replace '(?s)<(?:\w+:)?f\b[^>]*/>|<(?:\w+:)?f\b[^>]*>.*?</(?:\w+:)?f>', ''   # formula text: not data
            $value = [System.Net.WebUtility]::HtmlDecode([regex]::Match($inner, '(?s)<(?:\w+:)?v>(.*?)</(?:\w+:)?v>').Groups[1].Value)
            switch ($type) {
                's' { $index = 0; if ([int]::TryParse($value, [ref]$index) -and $index -lt $Shared.Count) { $Shared[$index] } }
                'inlineStr' { Convert-XmlToText $inner }
                'str' { $value }
                default { if ($type -notin @('b', 'e') -and $value -match '^-?\d+$') { $value } }   # whole numbers only
            }
        }
        (@($cells | Where-Object { $_ }) -join $script:CellBreak)
    }
    return ($lines -join "`n")
}

function Convert-XmlToText {
    <#
    .SYNOPSIS
      Strips Office XML to plain text. Paragraph, row and cell ends become line breaks or spaces;
      every other tag is removed with no space, so text that Word splits into several runs
      ("123-45-" and "6789") joins again.
    .PARAMETER Xml
      The XML text of one part.
    .OUTPUTS
      [string]
    #>
    param([string]$Xml)
    # Word field codes are not shown on the page: PAGEREF _Toc123456789, HYPERLINK addresses, form IDs. They gave
    # false positives (a 9-digit bookmark number read as an SSN), so their text is dropped (the PAM, 2026-11-20).
    $text = $Xml -replace '(?s)<w:instrText\b[^>]*?(?:/>|>.*?</w:instrText>)', ''
    $text = $text -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
    $text = $text -replace '<text:(?:s|tab|line-break)\b[^>]*/>', ' '
    $text = $text -replace '</w:p>|</a:p>|</row>|</text:p>|</text:h>|</table:table-row>|</Text>', "`n"
    $text = $text -replace '</table:table-cell>', $script:CellBreak   # ODF cells: never joined into one number
    $text = $text -replace '</si>|</c>|</w:tc>|</a:tc>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

function Convert-HtmlToText {
    <#
    .SYNOPSIS
      Strips HTML (.html files) to plain text. Script and style blocks are dropped;
      block ends become line breaks.
    .PARAMETER Html
      The HTML text.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $text = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ' '
    $text = $text -replace '(?i)<br\s*/?>|</p>|</div>|</li>|</tr>', "`n"
    $text = $text -replace '(?i)</td>|</th>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

# =====================================================================================
# Reading bytes
# =====================================================================================

function Read-ZipEntry {
    # The text of one zip entry (UTF-8, as Office writes its XML).
    param($Entry)
    $reader = New-Object System.IO.StreamReader($Entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-SharedStrings {
    # The shared strings of a workbook in order (xl/sharedStrings.xml), each as plain text; none for other files.
    param($Zip)
    $entry = $Zip.GetEntry('xl/sharedStrings.xml')
    if (-not $entry) { return @() }
    foreach ($item in [regex]::Matches((Read-ZipEntry $entry), '(?s)<(?:\w+:)?si>(.*?)</(?:\w+:)?si>')) { Convert-XmlToText $item.Groups[1].Value }
}

function Get-OfficeProps {
    <#
    .SYNOPSIS
      The names an Office file stores about itself (docProps/core.xml: creator, last modified by), as Office saved
      them; empty for other files (the accountability columns, the PAM 2026-11-21).
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [hashtable] Creator, LastModifiedBy.
    #>
    param([byte[]]$Bytes)
    $props = @{ Creator = ''; LastModifiedBy = '' }
    if ($Bytes.Length -lt 4 -or $Bytes[0] -ne 0x50 -or $Bytes[1] -ne 0x4B) { return $props }   # not a zip
    try {
        $memory = New-Object System.IO.MemoryStream(, $Bytes)
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $core = $zip.GetEntry('docProps/core.xml')
            if ($core) {
                $xml = Read-ZipEntry $core
                $props.Creator = [System.Net.WebUtility]::HtmlDecode([regex]::Match($xml, '<dc:creator>([^<]*)</dc:creator>').Groups[1].Value)
                $props.LastModifiedBy = [System.Net.WebUtility]::HtmlDecode([regex]::Match($xml, '<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>').Groups[1].Value)
            }
        }
        finally { $zip.Dispose(); $memory.Dispose() }
    }
    catch { return @{ Creator = ''; LastModifiedBy = '' } }   # a damaged zip: its reader reports it
    return $props
}

function Get-ZipPartText {
    <#
    .SYNOPSIS
      Opens an Office file (a zip) in memory and returns the text of the parts whose names match.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER PartPattern
      Regular expression on the part name (from $OfficeParts).
    .OUTPUTS
      [string] the parts' text, joined with line breaks.
    #>
    param([byte[]]$Bytes, [string]$PartPattern)
    $memory = New-Object System.IO.MemoryStream(, $Bytes)
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $shared = @(Get-SharedStrings $zip)
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match $PartPattern } | Sort-Object FullName)) {
            $xml = Read-ZipEntry $entry
            if ($entry.FullName -match '^xl/worksheets/') { $texts.Add((Convert-SheetXmlToText $xml $shared)) }
            else { $texts.Add((Convert-XmlToText $xml)) }
        }
        return ($texts -join "`n")
    }
    finally { $zip.Dispose(); $memory.Dispose() }
}

function Convert-RtfToText {
    <#
    .SYNOPSIS
      Rich text (.rtf) as plain text: groups that hold no page text (font and colour tables, pictures, \*
      destinations) dropped; \par, \line and \tab as breaks; \'hh and \uN characters decoded; other control
      words and braces removed.
    .PARAMETER Rtf
      The RTF text.
    .OUTPUTS
      [string]
    #>
    param([string]$Rtf)
    # Drop whole groups that are not page text, with the groups nested in them (a balancing group counts the braces).
    $skip = '\{\\(?:\*|fonttbl|colortbl|stylesheet|pict|themedata|colorschememapping|datastore|latentstyles)' +
        '(?>[^{}]+|\{(?<d>)|\}(?<-d>))*(?(d)(?!))\}'
    $Rtf = (New-Object System.Text.RegularExpressions.Regex($skip, 'None', [TimeSpan]::FromSeconds(10))).Replace($Rtf, '')
    $text = $Rtf -replace '\\(?:par|line|row|sect|page)\b ?', "`n" -replace '\\(?:tab|cell)\b ?', ' '
    # \uN (with its fallback character, \'hh or ?, dropped) and \'hh in one pass, so a fallback is never decoded twice.
    $out = New-Object System.Text.StringBuilder
    $at = 0
    foreach ($m in [regex]::Matches($text, '\\u(-?\d+) ?(?:\\''[0-9A-Fa-f]{2}|\?)?|\\''([0-9A-Fa-f]{2})')) {
        [void]$out.Append($text, $at, $m.Index - $at)
        if ($m.Groups[1].Success) { [void]$out.Append([char](([int]$m.Groups[1].Value) -band 0xFFFF)) }
        else { [void]$out.Append([char][Convert]::ToInt32($m.Groups[2].Value, 16)) }
        $at = $m.Index + $m.Length
    }
    [void]$out.Append($text, $at, $text.Length - $at)
    $text = $out.ToString()
    $text = $text -replace '\\[A-Za-z]+-?\d* ?', '' -replace '\\([{}\\])', '$1' -replace '[{}]', ''
    return $text
}

function Get-PlainText {
    <#
    .SYNOPSIS
      Decodes a text file. A UTF-8 or UTF-16 byte order mark is honoured; no mark means UTF-8.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    $reader = New-Object System.IO.StreamReader((New-Object System.IO.MemoryStream(, $Bytes)), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-NotZipReason {
    <#
    .SYNOPSIS
      Why an Office file's bytes are not a zip, from their first bytes; empty when they are one. Without this check
      the zip reader says only "End of central directory not found" (the PAM's DoD365 run, 2026-11-20).
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string] the reason a person can act on, or ''.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return 'Empty file (0 bytes)' }
    $head = ($Bytes[0..([Math]::Min(7, $Bytes.Length - 1))] | ForEach-Object { $_.ToString('X2') }) -join ''
    if ($head.StartsWith($script:ZipMagic)) { return '' }
    if ($head.StartsWith($script:OleMagic)) { return $script:EncryptedReason }
    # Decoded with its byte order mark, if any (UTF-8 or UTF-16), so a web page with a mark is still seen.
    $start = (Get-PlainText ([byte[]]$Bytes[0..([Math]::Min(255, $Bytes.Length - 1))])).TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
    if ($start.StartsWith('<')) { return 'The download gave a web page, not the file: open it in SharePoint and check by hand' }
    return 'Not a valid Office file (damaged, or another type with an Office name): check by hand'
}

function Test-SupportedFile {
    # True when Get-TextFromBytes can read this file type (so other types are not downloaded).
    param([string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($script:PlainTextExtensions -contains $extension) -or ($script:HtmlExtensions -contains $extension) -or
        $script:OfficeParts.ContainsKey($extension) -or $extension -in @('.pdf', '.rtf') -or $script:ImageExtensions -contains $extension -or
        $script:OleExtensions -contains $extension -or $script:ContainerExtensions -contains $extension
}

function Get-TextFromBytes {
    <#
    .SYNOPSIS
      Returns the text of one file, or why it was not scanned. A zip or an e-mail gives Parts: the result of each
      file inside it (read the same way, up to $MaxContainerDepth levels deep).
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER FileName
      The file name; only its extension is used.
    .PARAMETER Depth
      How many containers this file sits in (0 for a file in the library).
    .OUTPUTS
      [pscustomobject] Status ('OK' or 'NOT SCANNED'), Reason, Text, Note (a part of a read file that was not read),
      Parts (the results of the files inside, each with a Name; empty for other files).
    #>
    param([byte[]]$Bytes, [string]$FileName, [int]$Depth = 0)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    $result = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = ''; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) }
    try {
        $read = $null   # a reader's hashtable (Text, Reason, Note) for the types that give one
        if ($script:PlainTextExtensions -contains $extension) { $result.Text = Get-PlainText $Bytes }
        elseif ($script:HtmlExtensions -contains $extension) { $result.Text = Convert-HtmlToText (Get-PlainText $Bytes) }
        elseif ($extension -eq '.rtf') { $result.Text = Convert-RtfToText (Get-PlainText $Bytes) }
        elseif ($extension -eq '.pdf') { $read = Get-PdfText $Bytes }
        elseif ($script:ImageExtensions -contains $extension) { $read = Get-ImageText $Bytes }
        elseif ($script:OleExtensions -contains $extension) { $read = Get-OleText $Bytes $extension }
        elseif ($script:ContainerExtensions -contains $extension) { Add-ContainerParts $result $Bytes $extension $Depth }
        elseif ($script:OfficeParts.ContainsKey($extension)) {
            $notZip = Get-NotZipReason $Bytes
            if ($notZip) { $result.Status = 'NOT SCANNED'; $result.Reason = $notZip }
            else { $result.Text = Get-ZipPartText $Bytes $script:OfficeParts[$extension] }
        }
        else {
            $result.Status = 'NOT SCANNED'
            $result.Reason = "File type '$extension' is not read by this script: check by hand"
        }
        if ($read) {
            if ($read.Reason) { $result.Status = 'NOT SCANNED'; $result.Reason = $read.Reason }
            else { $result.Text = $read.Text; if ($read.ContainsKey('Note')) { $result.Note = $read.Note } }
        }
    }
    catch {
        # A damaged or protected file must not stop the scan: report it and go on.
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file (damaged or protected?): ' + $_.Exception.Message
    }
    return $result
}

function Add-ContainerParts {
    <#
    .SYNOPSIS
      Reads each file in a zip or e-mail into $Result.Parts, named "inner/path" (a part's own parts are flattened
      as "inner.zip/file"). A container deeper than $MaxContainerDepth is not opened: NOT SCANNED with a reason.
    .PARAMETER Result
      The container's result (Get-TextFromBytes).
    .PARAMETER Bytes
      The container bytes.
    .PARAMETER Extension
      The container's extension: zip or eml, with its point.
    .PARAMETER Depth
      How many containers the container sits in.
    #>
    param($Result, [byte[]]$Bytes, [string]$Extension, [int]$Depth)
    if ($Depth -ge $script:MaxContainerDepth) {
        $Result.Status = 'NOT SCANNED'; $Result.Reason = 'An archive or e-mail inside another: open it and check by hand'
        return
    }
    $parts = if ($Extension -eq '.zip') { @(Get-ArchiveParts $Bytes) } else { @(Get-EmailParts $Bytes) }
    foreach ($part in $parts) {
        if ($part.Reason) { $one = [pscustomobject]@{ Status = 'NOT SCANNED'; Reason = $part.Reason; Text = ''; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) } }
        elseif ($null -ne $part.Text) {
            $text = $part.Text; if ($part.Name -like '*.html') { $text = Convert-HtmlToText $text }
            $one = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = $text; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) }
        }
        else { $one = Get-TextFromBytes -Bytes $part.Bytes -FileName $part.Name -Depth ($Depth + 1) }
        $one | Add-Member -NotePropertyName Name -NotePropertyValue ([string]$part.Name) -Force
        $Result.Parts.Add($one)
        $inners = @(); if ($null -ne $one.Parts -and $one.Parts.Count -gt 0) { $inners = $one.Parts.ToArray() }
        foreach ($inner in $inners) {
            $inner.Name = $part.Name + '/' + $inner.Name
            $Result.Parts.Add($inner)
        }
    }
}

Export-ModuleMember -Function Get-TextFromBytes, Convert-HtmlToText, Test-SupportedFile, Get-OfficeProps
