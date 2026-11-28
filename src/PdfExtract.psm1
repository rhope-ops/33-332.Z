<#
PdfExtract.psm1 - pure-.NET PDF reader: budgets, filters, objects, fonts, glyph decoding, layout, content streams, page text.

  Purpose:     Reads PDF objects, ToUnicode character maps, glyph widths, word spacing, content streams, and
               XFA/form text with .NET only, enforcing time and inflation budgets and sending scan-only pages to OCR.
  Inputs:      PDF file bytes (from TextExtract.psm1).
  Outputs:     extracted text and scan/not-read notes (Get-PdfText).
  Depends on:  ContainerText.psm1 (Get-ScanPagesText), imported first.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# ==== PdfCore.psm1 ====
# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)

# A ToUnicode map uses at most this many ranges (real maps use far fewer); ranges are looked up per glyph, not expanded.
# A map with more is read in part, and the PDF gets a "check by hand" note (Add-PdfNote).
$script:MaxCMapRanges = 256

# A range's target is a few UTF-16 code units in hex (<0041>, or a ligature <00660069>); a longer one is not used.
$script:MaxRangeTarget = 66

# Limits, so one damaged or hostile PDF cannot stall the scan (fresh review, 2026-11-20): the time for one PDF, the
# time for one regex search, the size one stream may inflate to, and the inflated size of one PDF in all.
# 120 s: a 76-page table-heavy PDF took 18 s on PowerShell 7 (fresh review, 2026-11-21); Windows PowerShell 5.1 is slower.
$script:PdfSecondsLimit = 120
$script:PdfRegexTimeout = [TimeSpan]::FromSeconds(2)
$script:MaxStreamBytes = 16MB
$script:MaxInflatedBytes = 64MB
$script:PdfBudget = @{ Clock = $null; Inflated = 0L; Notes = $null }

# Streams that hold no page text and are not inflated: images and embedded font programs.
$script:NotTextStream = '/Subtype\s*/(?:Image|Type1C|CIDFontType0C|OpenType)\b|/Length[123]\s|/Type\s*/(?:EmbeddedFile|XRef|Metadata)\b'

function Start-PdfBudget {
    # Starts the clock, the unpack count and the notes of one PDF (Get-PdfText).
    $script:PdfBudget = @{ Clock = [Diagnostics.Stopwatch]::StartNew(); Inflated = 0L; Notes = (New-Object System.Collections.Generic.List[string]) }
}

function Stop-PdfBudget {
    # Ends one PDF's budget and returns its notes (parts not read).
    $notes = $script:PdfBudget.Notes
    $script:PdfBudget = @{ Clock = $null; Inflated = 0L; Notes = $null }
    if ($notes) { return $notes.ToArray() }
}

function ConvertFrom-HexText {
    <#
    .SYNOPSIS
      Hex digits (ASCIIHexDecode data, other characters ignored) as their bytes in Latin-1 characters.
    .PARAMETER Text
      The hex text.
    #>
    param([string]$Text)
    $hex = $Text -replace '[^0-9A-Fa-f]', ''
    if ($hex.Length % 2) { $hex += '0' }
    return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
        if (($i % 200000) -eq 0) { Assert-PdfTime }
        [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
    })
}

function Get-PdfRegex {
    # A regex with a time limit per search ($PdfRegexTimeout); a search over it throws instead of running on.
    param([string]$Pattern)
    return New-Object System.Text.RegularExpressions.Regex($Pattern, [System.Text.RegularExpressions.RegexOptions]::None, $script:PdfRegexTimeout)
}

function Assert-PdfTime {
    # Throws when this PDF has taken longer than $PdfSecondsLimit (Get-PdfText then reports it as too slow to read).
    if ($script:PdfBudget.Clock -and $script:PdfBudget.Clock.Elapsed.TotalSeconds -gt $script:PdfSecondsLimit) {
        throw (New-Object System.TimeoutException 'PDF time limit')
    }
}

function Add-PdfNote {
    # Records a part of this PDF that was not read (a cap, damage, attachments): the report then shows a "check by
    # hand" row next to the hits, so no part is left out without a word (fresh review, 2026-11-21).
    param([string]$Note)
    if (-not $script:PdfBudget.ContainsKey('Notes') -or $null -eq $script:PdfBudget.Notes) { return }   # outside Get-PdfText
    if (-not $script:PdfBudget.Notes.Contains($Note)) { $script:PdfBudget.Notes.Add($Note) }
}

function Expand-PdfStream {
    <#
    .SYNOPSIS
      A FlateDecode stream (zlib: a 2-byte header, then deflate) inflated; what was read before an error is kept.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    #>
    param([string]$Raw)
    $bytes = $script:Latin1.GetBytes($Raw)
    # The zlib header (RFC 1950): method 8 in the low 4 bits, and the two bytes a multiple of 31. Any window size, not
    # only 0x78: a header read as deflate data made good PDFs look damaged (the Administrator's gov run, 2026-11-21).
    $skip = 0
    if ($bytes.Length -ge 2 -and ($bytes[0] -band 0x0F) -eq 8 -and (($bytes[0] * 256 + $bytes[1]) % 31) -eq 0) { $skip = 2 }
    $source = New-Object System.IO.MemoryStream -ArgumentList $bytes, $skip, ($bytes.Length - $skip)
    $inflate = New-Object System.IO.Compression.DeflateStream($source, [System.IO.Compression.CompressionMode]::Decompress)
    $output = New-Object System.IO.MemoryStream
    $buffer = New-Object byte[] 65536
    try {
        # Inflate in steps and stop at the caps: a small stream can unpack to gigabytes.
        while ($true) {
            if ($output.Length -ge $script:MaxStreamBytes -or $script:PdfBudget.Inflated -ge $script:MaxInflatedBytes) {
                # At a cap: a note only when there is more to unpack (a stream of exactly the cap is whole).
                if ($inflate.Read($buffer, 0, 1) -gt 0) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand' }
                break
            }
            $read = $inflate.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $output.Write($buffer, 0, $read)
            $script:PdfBudget.Inflated += $read
        }
    }
    catch { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand' }   # the text read before the damage is kept
    finally { $inflate.Dispose(); $source.Dispose() }
    return $script:Latin1.GetString($output.ToArray())
}

function Test-PdfUnpackCap {
    # True when a stream being decoded by hand (LZW, RunLength) reached the cap of one stream or of the whole PDF.
    param([long]$Length)
    return ($Length -ge $script:MaxStreamBytes -or $script:PdfBudget.Inflated + $Length -ge $script:MaxInflatedBytes)
}

function Close-PdfUnpack {
    <#
    .SYNOPSIS
      Ends a stream decoded by hand: its size counts toward the PDF's unpack total (as Flate output does), and a
      stream stopped at a cap gets a "check by hand" note.
    .PARAMETER Out
      The StringBuilder with the decoded bytes.
    .PARAMETER Whole
      The stream ended by itself (not at a cap).
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param($Out, [switch]$Whole)
    $script:PdfBudget.Inflated += $Out.Length
    if (-not $Whole) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand' }
    return $Out.ToString()
}

function Expand-LzwStream {
    <#
    .SYNOPSIS
      An LZWDecode stream (PDF and TIFF LZW: 9 to 12 bit codes, most significant bit first, 256 = clear table,
      257 = end) decoded, up to $MaxStreamBytes; damaged data ends the stream with a note.
    .PARAMETER Raw
      The stream bytes as Latin-1 text.
    .PARAMETER EarlyChange
      True (the default): the code width grows one code early, as the PDF default /EarlyChange 1 says.
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param([string]$Raw, [bool]$EarlyChange = $true)
    $bytes = $script:Latin1.GetBytes($Raw)
    $out = New-Object System.Text.StringBuilder
    $table = New-Object System.Collections.Generic.List[string]
    # Empties the code table back to its 258 starting entries (bytes 0-255, clear, end).
    $reset = { $table.Clear(); for ($c = 0; $c -lt 256; $c++) { $table.Add([string][char]$c) }; $table.Add(''); $table.Add('') }
    & $reset
    $width = 9; $bits = 0L; $count = 0; $previous = $null; $n = 0
    foreach ($byte in $bytes) {
        $bits = ($bits -shl 8) -bor $byte; $count += 8
        while ($count -ge $width) {
            if ((++$n % 10000) -eq 0) { Assert-PdfTime }
            $count -= $width
            $code = [int](($bits -shr $count) -band ((1 -shl $width) - 1))
            $bits = $bits -band ((1L -shl $count) - 1)
            if ($code -eq 256) { & $reset; $width = 9; $previous = $null; continue }
            if ($code -eq 257) { return (Close-PdfUnpack $out -Whole) }
            if ($code -lt $table.Count) { $entry = $table[$code] }
            elseif ($code -eq $table.Count -and $null -ne $previous) { $entry = $previous + $previous[0] }
            else { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return (Close-PdfUnpack $out -Whole) }
            [void]$out.Append($entry)
            if (Test-PdfUnpackCap $out.Length) { return (Close-PdfUnpack $out) }
            if ($null -ne $previous -and $table.Count -lt 4096) { $table.Add($previous + $entry[0]) }
            $previous = $entry
            $limit = $table.Count; if ($EarlyChange) { $limit++ }
            if ($limit -ge (1 -shl $width) -and $width -lt 12) { $width++ }
        }
    }
    return (Close-PdfUnpack $out -Whole)
}

function Expand-RunLengthStream {
    <#
    .SYNOPSIS
      A RunLengthDecode stream: a length byte 0-127 copies the next 1-128 bytes, 129-255 repeats the next byte
      2-128 times, 128 ends it; up to $MaxStreamBytes.
    .PARAMETER Raw
      The stream bytes as Latin-1 text.
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param([string]$Raw)
    $out = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Raw.Length) {
        if (($out.Length % 65536) -eq 0) { Assert-PdfTime }
        $length = [int]$Raw[$i]
        if ($length -eq 128) { break }
        if ($length -lt 128) { $take = [Math]::Min($length + 1, $Raw.Length - $i - 1); [void]$out.Append($Raw, $i + 1, $take); $i += 1 + $take }
        elseif ($i + 1 -lt $Raw.Length) { [void]$out.Append($Raw[$i + 1], 257 - $length); $i += 2 }
        else { break }
        if (Test-PdfUnpackCap $out.Length) { return (Close-PdfUnpack $out) }
    }
    return (Close-PdfUnpack $out -Whole)
}

function Read-PdfStream {
    <#
    .SYNOPSIS
      A stream's data with its filters undone in order: FlateDecode, ASCII85Decode, ASCIIHexDecode, LZWDecode and
      RunLengthDecode (often chained, such as [/ASCII85Decode /FlateDecode]). Images and fonts are not text: $null.
      Any other filter: $null and a "check by hand" note, never a silent miss.
    .PARAMETER Dict
      The stream's dictionary.
    .PARAMETER Raw
      The stream's bytes as Latin-1 text.
    .OUTPUTS
      [string] or $null.
    #>
    param([string]$Dict, [string]$Raw)
    if ($Dict -match '/Type\s*/EmbeddedFile\b') { Add-PdfNote 'Files attached inside this PDF are not scanned: check by hand' }
    if ($Dict -match $script:NotTextStream) { return $null }
    $filter = [regex]::Match($Dict, '/Filter\s*(\[[^\]]*\]|/\w+)')
    $data = $Raw
    foreach ($name in @([regex]::Matches($filter.Groups[1].Value, '/(\w+)') | ForEach-Object { $_.Groups[1].Value })) {
        switch ($name) {
            'FlateDecode' { $data = Expand-PdfStream $data }
            'ASCII85Decode' { $data = ConvertFrom-Ascii85 $data }
            'ASCIIHexDecode' { $data = ConvertFrom-HexText $data }
            'LZWDecode' { $data = Expand-LzwStream $data ($Dict -notmatch '/EarlyChange\s+0\b') }
            'RunLengthDecode' { $data = Expand-RunLengthStream $data }
            default { Add-PdfNote 'Part of this PDF uses a compression this reader cannot open: check by hand'; return $null }
        }
    }
    return $data
}

function ConvertFrom-Ascii85 {
    <#
    .SYNOPSIS
      ASCII85 text (PDF ASCII85Decode, ending in ~>) to its bytes as Latin-1 text; z = four zero bytes.
    .PARAMETER Text
      The text.
    #>
    param([string]$Text)
    $end = $Text.IndexOf('~>'); if ($end -ge 0) { $Text = $Text.Substring(0, $end) }
    $Text = ($Text -replace '\s', '') -creplace 'z', '!!!!!'   # -creplace: only a small z (a capital Z is data)
    $out = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Text.Length; $i += 5) {
        if (($i % 100000) -eq 0) { Assert-PdfTime }
        $group = $Text.Substring($i, [Math]::Min(5, $Text.Length - $i))
        $count = $group.Length - 1
        if ($count -lt 1) { break }
        $group = $group.PadRight(5, 'u')
        [long]$value = 0
        foreach ($c in $group.ToCharArray()) { $value = $value * 85 + ([int]$c - 33) }
        for ($k = 0; $k -lt $count; $k++) { [void]$out.Append([char](($value -shr (24 - 8 * $k)) -band 0xFF)) }
    }
    return $out.ToString()
}

function ConvertFrom-Utf16Hex {
    # The text of a ToUnicode target: UTF-16BE given as hex.
    param([string]$Hex)
    $bytes = [byte[]]@(for ($i = 0; $i + 1 -lt $Hex.Length; $i += 2) { [Convert]::ToByte($Hex.Substring($i, 2), 16) })
    return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
}

function Get-PdfBlocks {
    <#
    .SYNOPSIS
      The text between each begin word and the end word after it (bfchar or bfrange blocks), found in one linear pass:
      a block without its end word runs to the next begin word, so a damaged map cannot make the search quadratic.
    .PARAMETER Text
      The CMap stream.
    .PARAMETER Begin
      The begin word, for example 'beginbfchar'.
    .PARAMETER End
      The end word, for example 'endbfchar'.
    .OUTPUTS
      [string] one per block.
    #>
    param([string]$Text, [string]$Begin, [string]$End)
    $starts = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($Begin))).Matches($Text)) { $m.Index + $m.Length })
    $stops = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($End))).Matches($Text)) { $m.Index })
    $e = 0
    for ($b = 0; $b -lt $starts.Count; $b++) {
        $stop = $Text.Length; if ($b + 1 -lt $starts.Count) { $stop = $starts[$b + 1] - $Begin.Length }
        while ($e -lt $stops.Count -and $stops[$e] -lt $starts[$b]) { $e++ }
        if ($e -lt $stops.Count -and $stops[$e] -lt $stop) { $stop = $stops[$e] }
        $Text.Substring($starts[$b], $stop - $starts[$b])
    }
}

function Get-PdfCMap {
    <#
    .SYNOPSIS
      A font's ToUnicode map: glyph code (hex, upper case) -> text from its bfchar entries and listed ranges, plus
      its plain ranges (looked up by ConvertFrom-PdfGlyphs, not expanded).
    .PARAMETER Text
      The CMap stream.
    .OUTPUTS
      [hashtable] Bytes (bytes per code), Map, Ranges.
    #>
    param([string]$Text)
    $map = @{}
    $ranges = New-Object System.Collections.Generic.List[object]
    $bytes = 1
    $pairRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>'
    $rangeRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])'
    $listRegex = Get-PdfRegex '<([0-9A-Fa-f]*)>'
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfchar' 'endbfchar')) {
        Assert-PdfTime
        $n = 0
        foreach ($pair in $pairRegex.Matches($block)) {
            if ((++$n % 1000) -eq 0) { Assert-PdfTime }
            $map[$pair.Groups[1].Value.ToUpper()] = ConvertFrom-Utf16Hex $pair.Groups[2].Value
            $bytes = [int]($pair.Groups[1].Value.Length / 2)
        }
    }
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfrange' 'endbfrange')) {
        Assert-PdfTime
        foreach ($range in $rangeRegex.Matches($block)) {
            $width = $range.Groups[1].Value.Length
            $bytes = [int]($width / 2)
            $low = [Convert]::ToInt64($range.Groups[1].Value, 16)
            $high = [Convert]::ToInt64($range.Groups[2].Value, 16)
            $target = $range.Groups[3].Value
            if ($high -lt $low) { continue }
            if ($target.StartsWith('[')) {
                # A listed range names each target: as many entries as the map's own text holds, so no expansion cost.
                $code = $low
                foreach ($item in $listRegex.Matches($target)) {
                    if ($code -gt $high) { break }
                    if ((($code - $low) % 1000) -eq 999) { Assert-PdfTime }
                    $map[$code.ToString('X' + $width)] = ConvertFrom-Utf16Hex $item.Groups[1].Value
                    $code++
                }
            }
            elseif ($ranges.Count -ge $script:MaxCMapRanges) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand' }
            elseif ($target.Length -gt $script:MaxRangeTarget) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand' }
            elseif ($target.Length -ge 6) {
                $hex = $target.Substring(1, $target.Length - 2)
                $ranges.Add(@{ Low = $low; High = $high; Prefix = $hex.Substring(0, $hex.Length - 4); Start = [Convert]::ToInt32($hex.Substring($hex.Length - 4), 16) })
            }
        }
    }
    return @{ Bytes = $bytes; Map = $map; Ranges = $ranges }
}

# ==== PdfObjects.psm1 ====
# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:ObjLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary). A literal holds at most
# 10,000 parts, so an unclosed bracket fails after a bounded scan. A "(" that no string token takes means a string
# was not read: the PDF is then NOT SCANNED, never read in part without a word (Assert-PdfGap).
$script:PdfStringPattern = '\((?>[^()\\]+|\\[\s\S]|\((?<d>)|\)(?<-d>)){0,10000}(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The message of the error a string not read throws (Get-PdfText turns it into "part of this PDF could not be read").
$script:PdfStringError = 'PDF string not read'

# The escapes of a literal string (\n, \053, a line break ...) and how each is undone, built once: building a
# PowerShell delegate per string cost about 450 microseconds (speed, 2026-11-21).
$script:EscapePattern = New-Object System.Text.RegularExpressions.Regex('\\(?:[0-7]{1,3}|\r?\n|.)')
$script:UndoEscape = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    $code = $m.Value.Substring(1)
    if ($code -match '^[0-7]+$') { return [string][char]([Convert]::ToInt32($code, 8) -band 0xFF) }
    switch -CaseSensitive ($code) { 'n' { "`n" } 'r' { "`r" } 't' { "`t" } 'b' { ' ' } 'f' { ' ' } default { $code.Trim("`r", "`n") } }
}

# An object stream (PDF 1.5) is unpacked only within these limits: its header of object numbers and offsets, and its
# object count (real files hold a few hundred objects per stream). More gives a "check by hand" note.
$script:MaxPackedHeader = 1MB

# Earlier versions of objects defined again further on (incremental updates): their text is still read.
$script:PdfOldObjects = New-Object System.Collections.Generic.List[object]
$script:MaxPackedObjects = 100000

# Dictionary keys whose string values are text a person wrote (Get-PdfDictText) -> the label put before the value:
# a form field's value (labelled with the field's name), a comment or note and its rich text, the file's info.
$script:PdfTextKeys = [ordered]@{ 'V' = ''; 'Contents' = 'Comment'; 'RC' = 'Comment'; 'Author' = 'Author'; 'Title' = 'Title'
    'Subject' = 'Subject'; 'Keywords' = 'Keywords' }

function Get-PdfStringPattern {
    # The pattern of one PDF string token (PdfExtract.psm1 builds its token pattern on it).
    return $script:PdfStringPattern
}

function Get-PdfStringError {
    # The message of the error a string not read throws.
    return $script:PdfStringError
}

function Stop-PdfHostile {
    # Ends the read of a PDF with more operands or array entries than any real one has: it is reported as NOT SCANNED
    # ("part could not be read"), so a hostile file cannot hold gigabytes (reviews r14, r18).
    throw (New-Object System.FormatException (Get-PdfStringError))
}

function Reset-PdfOldObjects {
    # Forgets the older object versions of the last PDF (Get-PdfText, before each PDF).
    $script:PdfOldObjects.Clear()
}

function Get-PdfOldObjects {
    # The older versions of objects defined again further on in this PDF (an incremental update).
    return , $script:PdfOldObjects.ToArray()
}

function Get-PdfObjects {
    <#
    .SYNOPSIS
      The objects of a PDF by number: each its dictionary text and its stream (inflated when FlateDecode; $null for
      other filters, such as images). Objects packed in object streams (PDF 1.5) are unpacked too.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary] number -> @{ Dict; Stream; Pos; Image (JPEG data or $null) }, in file order.
    #>
    param([string]$Text)
    $objects = New-Object System.Collections.Specialized.OrderedDictionary
    # Each object runs from its header to the first "endobj" after it, or to the next header when "endobj" is missing.
    # Two linear passes, so a damaged file without "endobj" cannot make the search quadratic.
    $heads = @((Get-PdfRegex '(\d+)\s+\d+\s+obj\b').Matches($Text))
    $ends = @((Get-PdfRegex '\bendobj\b').Matches($Text) | ForEach-Object { $_.Index })
    $e = 0
    for ($h = 0; $h -lt $heads.Count; $h++) {
        Assert-PdfTime
        $start = $heads[$h].Index + $heads[$h].Length
        $stop = $Text.Length; if ($h + 1 -lt $heads.Count) { $stop = $heads[$h + 1].Index }
        while ($e -lt $ends.Count -and $ends[$e] -lt $start) { $e++ }
        if ($e -lt $ends.Count -and $ends[$e] -lt $stop) { $stop = $ends[$e] }
        $body = $Text.Substring($start, $stop - $start)
        $stream = $null
        $parts = Split-PdfStreamBody $body
        $image = $null
        if ($parts) {
            $body = $parts.Dict; $stream = Read-PdfStream $body $parts.Data
            # A JPEG page image (a scanned page): kept as it is, for OCR when the PDF has no text (Get-ScanPagesText).
            if ($body -match '/Subtype\s*/Image\b' -and $body -match '/Filter\s*\[?\s*/DCTDecode\s*\]?') { $image = $parts.Data }
        }
        $number = $heads[$h].Groups[1].Value
        # A file saved with added edits (an incremental update) defines an object again further on: the later one is
        # the one in use; the earlier one's text is still in the file, so it is kept (Get-PdfDictText, Get-PdfLooseText).
        if ($objects.Contains($number)) { $script:PdfOldObjects.Add($objects[$number]) }
        $objects[$number] = @{ Dict = $body; Stream = $stream; Pos = $heads[$h].Index; Image = $image }
    }
    foreach ($packed in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/ObjStm' -and $_.Stream })) { Add-PackedObjects $objects $packed }
    return $objects
}

function Split-PdfStreamBody {
    <#
    .SYNOPSIS
      An object body with a stream, split into its dictionary and its data: from the "stream" keyword (then CR LF, LF
      or CR) to the last "endstream". Plain string searches, so a stream of any size is split in one linear pass.
      $null for an object without a stream; a stream without its end gives a note.
    .PARAMETER Body
      The object's text from its header to endobj.
    .OUTPUTS
      [hashtable] Dict, Data; or $null.
    #>
    param([string]$Body)
    $at = -1
    for ($from = 0; ($at = $Body.IndexOf('stream', $from, [StringComparison]::Ordinal)) -ge 0; $from = $at + 6) {
        $after = $at + 6
        $word = $at -eq 0 -or -not [char]::IsLetter($Body[$at - 1])   # not the end of "endstream"
        if ($word -and $after -lt $Body.Length -and ($Body[$after] -eq "`n" -or $Body[$after] -eq "`r")) { break }
    }
    if ($at -lt 0) { return $null }
    $start = $at + 6
    if ($Body[$start] -eq "`r") { $start++ }
    if ($start -lt $Body.Length -and $Body[$start] -eq "`n") { $start++ }
    $end = $Body.LastIndexOf('endstream', [StringComparison]::Ordinal)
    if ($end -lt $start) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return $null }   # cut short
    $stop = $end
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`n") { $stop-- }
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`r") { $stop-- }
    return @{ Dict = $Body.Substring(0, $at); Data = $Body.Substring($start, $stop - $start) }
}

function Add-PackedObjects {
    <#
    .SYNOPSIS
      Adds the objects of one object stream (/Type /ObjStm: N pairs "number offset", then the objects from /First).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Packed
      One object stream of Get-PdfObjects.
    #>
    param($Objects, $Packed)
    $count = [int][regex]::Match($Packed.Dict, '/N\s+(\d+)').Groups[1].Value
    $first = [int][regex]::Match($Packed.Dict, '/First\s+(\d+)').Groups[1].Value
    if ($first -le 0 -or $first -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return }
    if ($first -gt $script:MaxPackedHeader -or $count -gt $script:MaxPackedObjects) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand'; return }
    $pairs = @($Packed.Stream.Substring(0, $first).Trim() -split '\s+')
    $number = 0
    for ($k = 0; $k -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        if (-not [int]::TryParse($pairs[$k], [ref]$number)) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return }
    }
    for ($k = 0; $k -lt $count -and 2 * $k + 1 -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        $start = $first + [int]$pairs[2 * $k + 1]
        $end = $Packed.Stream.Length
        if (2 * $k + 3 -lt $pairs.Count) { $end = $first + [int]$pairs[2 * $k + 3] }
        if ($start -ge $end -or $end -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; continue }
        $dict = $Packed.Stream.Substring($start, $end - $start)
        $number = $pairs[2 * $k]
        # The later definition in the file wins (an incremental update); the other one's text is kept too.
        if ($Objects.Contains($number) -and $Objects[$number].Pos -gt $Packed.Pos) { $script:PdfOldObjects.Add(@{ Dict = $dict; Stream = $null }); continue }
        if ($Objects.Contains($number)) { $script:PdfOldObjects.Add($Objects[$number]) }
        $Objects[$number] = @{ Dict = $dict; Stream = $null; Pos = $Packed.Pos }
    }
}

function ConvertFrom-PdfString {
    <#
    .SYNOPSIS
      The bytes of one PDF string token, as Latin-1 characters: (literal) with its escapes undone, or <hex>.
    .PARAMETER Token
      One PDF string token.
    #>
    param([string]$Token)
    if ($Token.StartsWith('<')) {
        $hex = $Token.Substring(1, $Token.Length - 2) -replace '\s', ''
        if ($hex.Length % 2) { $hex += '0' }
        return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
            if (($i % 200000) -eq 0) { Assert-PdfTime }
            [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
        })
    }
    $inner = $Token.Substring(1, $Token.Length - 2)
    if ($inner.IndexOf([char]92) -lt 0) { return $inner }   # no escapes (most strings): as it is (speed)
    return $script:EscapePattern.Replace($inner, $script:UndoEscape)
}

function ConvertFrom-PdfTextString {
    # A PDF text string (a form field's name or value): UTF-16 when it starts with the FE FF mark, else Latin-1.
    param([string]$Token)
    $raw = ConvertFrom-PdfString $Token
    if ($raw.Length -ge 2 -and [int]$raw[0] -eq 0xFE -and [int]$raw[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($script:ObjLatin1.GetBytes($raw.Substring(2)))
    }
    return $raw
}

function Assert-PdfGap {
    <#
    .SYNOPSIS
      Throws when the text between two tokens holds a "(": a string there was not read (too long, or damaged), so
      the PDF is reported NOT SCANNED rather than read in part without a word (fresh review, 2026-11-20).
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the gap starts.
    .PARAMETER To
      Where the gap ends.
    #>
    param([string]$Content, [int]$From, [int]$To)
    if ($To -gt $From -and $Content.IndexOf('(', $From, $To - $From) -ge 0) { throw (New-Object System.FormatException $script:PdfStringError) }
}

function Get-PdfFieldName {
    <#
    .SYNOPSIS
      A form field's name (/T) as the label of a value: the last name before the value in the dictionary (a
      dictionary can hold several fields inline), else the first; empty when there is none.
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .PARAMETER Before
      Where the value starts in the dictionary.
    #>
    param([string]$Dict, $NameRegex, [int]$Before)
    $names = @($NameRegex.Matches($Dict))
    if ($names.Count -eq 0) { return '' }
    $pick = $names[0]
    foreach ($name in $names) { if ($name.Index -lt $Before) { $pick = $name } }
    return (ConvertFrom-PdfTextString $pick.Groups[1].Value)
}

function Format-PdfDictLine {
    # One line of dictionary text: "label: text", or the text alone.
    param([string]$Label, [string]$Text)
    if ($Label) { return ($Label + ': ' + $Text) }
    return $Text
}

function Get-PdfDictText {
    <#
    .SYNOPSIS
      Text kept in dictionaries, not on a page: form-field values ("name: value", so labels stay with values),
      comments and sticky notes (/Contents of an annotation, /RC rich text) and the file's info (author, title,
      subject, keywords), in current and older object versions; every value of a key, also several in one
      dictionary. A literal value that is not read throws (Get-PdfText: "part of this PDF could not be read").
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] one line per value.
    #>
    param($Objects)
    $nameRegex = Get-PdfRegex ('/T\s*(' + $script:PdfStringPattern + ')')
    $checks = foreach ($key in $script:PdfTextKeys.Keys) {
        @{ Key = $key; Label = $script:PdfTextKeys[$key]; Value = (Get-PdfRegex ('/' + $key + '\s*(' + $script:PdfStringPattern + ')'))
           Ref = (Get-PdfRegex ('/' + $key + '\s+(\d+)\s+\d+\s+R')); Open = ('/' + $key + '\s*\(') }
    }
    # Older versions too. ToArray(), not @(): @() around an empty generic list fails ("Argument types do not match").
    $all = @($Objects.Values) + $script:PdfOldObjects.ToArray()
    foreach ($obj in $all) {
        Assert-PdfTime
        foreach ($check in $checks) {
            if ($obj.Dict.IndexOf('/' + $check.Key) -ge 0) { Get-PdfKeyLines $obj.Dict $check $Objects $nameRegex }
        }
    }
}

function Get-PdfKeyLines {
    <#
    .SYNOPSIS
      The lines of one key in one dictionary: each string value, a list of strings (a multi-select field), or a
      value held in an object of its own (a string, or a text stream for /V and /RC).
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER Check
      The key, its label and its regexes (Get-PdfDictText).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .OUTPUTS
      [string] lines.
    #>
    param([string]$Dict, [hashtable]$Check, $Objects, $NameRegex)
    $values = @($Check.Value.Matches($Dict))
    $literals = @($values | Where-Object { $_.Groups[1].Value.StartsWith('(') }).Count
    if ([regex]::Matches($Dict, $Check.Open).Count -gt $literals) { throw (New-Object System.FormatException $script:PdfStringError) }   # not read
    foreach ($value in $values) {
        $text = ConvertFrom-PdfTextString $value.Groups[1].Value
        if ($Check.Key -eq 'RC') { $text = $text -replace '<[^>]+>', ' ' }   # rich text: XHTML tags dropped
        $label = $Check.Label
        if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $value.Index }
        Format-PdfDictLine $label $text
    }
    if ($values.Count -gt 0) { return }
    $list = [regex]::Match($Dict, '/' + $Check.Key + '\s*\[([^\]]*)\]')
    if ($list.Success) {
        $items = @(foreach ($item in (Get-PdfRegex $script:PdfStringPattern).Matches($list.Groups[1].Value)) { ConvertFrom-PdfTextString $item.Value })
        $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $list.Index }
        if ($items.Count -gt 0) { Format-PdfDictLine $label ($items -join ', ') }
        return
    }
    $ref = $Check.Ref.Match($Dict)
    if (-not $ref.Success -or -not $Objects.Contains($ref.Groups[1].Value)) { return }
    $held = $Objects[$ref.Groups[1].Value]
    $target = $held.Dict.Trim()
    $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $ref.Index }
    if ($target.StartsWith('(') -or ($target.StartsWith('<') -and -not $target.StartsWith('<<'))) { Format-PdfDictLine $label (ConvertFrom-PdfTextString $target) }
    elseif ($Check.Key -in @('V', 'RC') -and $held.Stream) { Format-PdfDictLine $label $held.Stream }   # a text stream
}

function Get-PdfXfaText {
    <#
    .SYNOPSIS
      The data of an XFA form (/AcroForm /XFA: one stream, or an array of name and stream pairs): XML, read as
      "element: value" lines, so a field name stays with its value (a dynamic XFA form shows only "Please wait..."
      on its page; its data is here).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] lines.
    #>
    param($Objects)
    $pair = Get-PdfRegex '<([\w:.-]+)\b[^<>]*>([^<]+)</\1>'
    foreach ($obj in @($Objects.Values)) {
        $xfa = [regex]::Match($obj.Dict, '/XFA\s*(\[[^\]]*\]|\d+\s+\d+\s+R)')
        if (-not $xfa.Success) { continue }
        foreach ($ref in [regex]::Matches($xfa.Groups[1].Value, '(\d+)\s+\d+\s+R')) {
            Assert-PdfTime
            $part = $Objects[$ref.Groups[1].Value]
            if (-not $part -or -not $part.Stream) { continue }   # a stream not decoded has its note already
            $lines = $pair.Replace($part.Stream, "`n`$1: `$2`n") -replace '<[^>]+>', ' '
            [System.Net.WebUtility]::HtmlDecode($lines)
        }
    }
}

# ==== PdfFonts.psm1 ====
# Fonts that draw pictures, not text (by name, with or without a subset prefix such as ABCDEF+): Wingdings 1-3,
# Webdings, (Zapf) Dingbats, Marlett and Symbol (Greek and maths signs). Their text is read as spaces: a row of symbols
# read as letters and digits gave Low hits that were plainly wrong (the Administrator, 2026-11-21).
# Standard glyph names (Adobe Glyph List) of the digits and common signs, for a font's /Differences; single letters
# name themselves. Any other name is a made-up one: its glyph cannot be read.
$script:GlyphNames = @{ space = ' '; zero = '0'; one = '1'; two = '2'; three = '3'; four = '4'; five = '5'; six = '6'; seven = '7'
    eight = '8'; nine = '9'; period = '.'; comma = ','; colon = ':'; semicolon = ';'; hyphen = '-'; minus = '-'; endash = [string][char]0x2013
    emdash = [string][char]0x2014; slash = '/'; backslash = '\'; parenleft = '('; parenright = ')'; bracketleft = '['; bracketright = ']'
    braceleft = '{'; braceright = '}'; at = '@'; numbersign = '#'; dollar = '$'; percent = '%'; ampersand = '&'; asterisk = '*'; plus = '+'
    equal = '='; less = '<'; greater = '>'; exclam = '!'; question = '?'; underscore = '_'; bar = '|'; asciitilde = '~'; quotesingle = "'"
    quotedbl = '"'; quoteleft = [string][char]0x2018; quoteright = [string][char]0x2019; quotedblleft = [string][char]0x201C
    quotedblright = [string][char]0x201D; bullet = [string][char]0x2022; degree = [string][char]0xB0; section = [string][char]0xA7
    periodcentered = [string][char]0xB7; germandbls = [string][char]0xDF; dotlessi = 'i'; fi = 'fi'; fl = 'fl'; ff = 'ff'; ffi = 'ffi'; ffl = 'ffl'
    nbspace = ' ' }
# Glyph names are case-sensitive (AE and ae are two letters), so the table is an ordinal dictionary.
$glyphTable = $script:GlyphNames
$script:GlyphNames = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
foreach ($name in $glyphTable.Keys) { $script:GlyphNames[$name] = $glyphTable[$name] }
foreach ($pair in @(@('AE', 0xC6), @('ae', 0xE6), @('Oslash', 0xD8), @('oslash', 0xF8), @('OE', 0x152), @('oe', 0x153))) {
    $script:GlyphNames[$pair[0]] = [string][char]$pair[1]
}

# The combining marks of accented glyph names (eacute = e + U+0301).
$script:GlyphMarks = @{ acute = 0x301; grave = 0x300; circumflex = 0x302; dieresis = 0x308; tilde = 0x303; ring = 0x30A; cedilla = 0x327; caron = 0x30C
    macron = 0x304; breve = 0x306; ogonek = 0x328; dotaccent = 0x307; hungarumlaut = 0x30B; commaaccent = 0x326 }
# A made-up glyph name: up to 5 letters, then a number (g12, a3, G1, cid34, glyph7, index5), with an optional suffix.
# Not the Adobe names that are letters and a number too: afii (Cyrillic and other letters), uni and u (code points),
# SF and H (box drawing and shapes).
# Case-sensitive (?-i): Adobe names are case-exact, so h12 or sf1 stay made-up names (review of PR 34).
$script:MadeUpGlyphName = '(?-i)^(?!afii|uni[0-9A-F]{4}|u[0-9A-F]{4,6}$|SF\d|H\d)[A-Za-z]{1,5}\d+(?:\..*)?$'

$script:PictureFontPattern = '(?i)wingdings|webdings|dingbats|marlett|(?:^|\+)symbol(?:mt)?(?:$|[,-])'

# Widths of the standard fonts (Adobe AFM, in thousandths of an em) for character codes 32-126; Courier is 600 for all.
# A simple font may leave out /Widths only when it is one of these (ISO 32000-1, 9.6.2.2).
$script:HelveticaWidths = @(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556,
    556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778,
    722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222,
    833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)
$script:TimesWidths = @(250, 333, 408, 500, 500, 833, 778, 180, 333, 333, 500, 564, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500,
    500, 500, 500, 278, 278, 564, 564, 564, 444, 921, 722, 667, 667, 722, 611, 556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722,
    667, 556, 611, 722, 722, 944, 722, 722, 611, 333, 278, 333, 469, 500, 333, 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500, 278,
    778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444, 480, 200, 480, 541)

# A /W array (a composite font's widths) is read for at most this many glyphs (a hostile array cannot fill the memory).
$script:MaxWidthEntries = 65536

function Get-PdfFonts {
    <#
    .SYNOPSIS
      The fonts the pages name (/F1 ...) -> their font records, the first of each name in the file (for streams
      that name no resources of their own).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param($Objects)
    $fonts = @{}
    foreach ($obj in $Objects.Values) {
        Assert-PdfTime
        $one = Get-PdfFontTable $obj.Dict $Objects
        foreach ($name in $one.Keys) { if (-not $fonts.ContainsKey($name)) { $fonts[$name] = $one[$name] } }
    }
    return $fonts
}

function Get-PdfFontTable {
    <#
    .SYNOPSIS
      The fonts that one dictionary's /Font resource names (/F1 -> its font record). A page uses its own table, so
      two pages that both name a font /F1 each get theirs; each font and each ToUnicode map is read once.
    .PARAMETER Dict
      A page or resources dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable]
    #>
    param([string]$Dict, $Objects)
    $fonts = @{}
    foreach ($resource in [regex]::Matches($Dict, '/Font\s*(?:<<(?<inline>[^<>]*)>>|(?<ref>\d+)\s+\d+\s+R)')) {
        $names = $resource.Groups['inline'].Value
        if ($resource.Groups['ref'].Success -and $Objects.Contains($resource.Groups['ref'].Value)) { $names = $Objects[$resource.Groups['ref'].Value].Dict }
        foreach ($font in [regex]::Matches($names, '/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R')) {
            $id = $font.Groups[2].Value
            if (-not $Objects.Contains($id)) { continue }
            $entry = $Objects[$id]
            if (-not $entry.Contains('Font')) { $entry['Font'] = New-PdfFontRecord $entry.Dict $Objects }
            $fonts[$font.Groups[1].Value] = $entry['Font']
        }
    }
    return $fonts
}

function New-PdfFontRecord {
    <#
    .SYNOPSIS
      One font's record: its ToUnicode map (CMap; a composite font without one gets an Unreadable marker: its codes
      are glyph numbers, never read as text), its code length, and its widths.
    .PARAMETER Dict
      The font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable] CMap, Bytes, Widths (code -> ems), DefaultWidth (ems, or -1), Standard (a width table, or $null),
      Vertical (vertical writing: each glyph advances 1 em down).
    #>
    param([string]$Dict, $Objects)
    $composite = $Dict -match '/Subtype\s*/Type0\b'
    $record = @{ CMap = $null; Bytes = 1; Widths = @{}; DefaultWidth = -1.0; Standard = $null; Vertical = ($Dict -match '/Encoding\s*/Identity-V\b')
        TextCache = (New-Object 'System.Collections.Generic.Dictionary[string,string]'); EmCache = (New-Object 'System.Collections.Generic.Dictionary[string,double]') }
    if ($composite) { $record.Bytes = 2 }
    # A picture font draws symbols, not letters: its codes read as text give runs of false letters and digits.
    $record.Picture = [regex]::Match($Dict, '/BaseFont\s*/([^\s/<>\[\]()]+)').Groups[1].Value -match $script:PictureFontPattern
    $unicode = [regex]::Match($Dict, '/ToUnicode\s+(\d+)\s+\d+\s+R')
    if ($unicode.Success -and $Objects.Contains($unicode.Groups[1].Value) -and $Objects[$unicode.Groups[1].Value].Stream) {
        $cmap = $Objects[$unicode.Groups[1].Value]   # read once, also when many fonts share it
        if (-not $cmap.Contains('CMap')) { $cmap['CMap'] = Get-PdfCMap $cmap.Stream }
        $record.CMap = $cmap['CMap']
        $record.Bytes = [Math]::Max(1, [int]$record.CMap.Bytes)
    }
    elseif ($composite) { $record.CMap = @{ Bytes = 2; Map = @{}; Ranges = @(); Unreadable = $true } }
    else { $record.Glyphs = Get-PdfDifferences $Dict $Objects }   # a simple font without a map: its own glyph names
    if ($composite) { Add-PdfCidWidths $record $Dict $Objects }
    else { Add-PdfSimpleWidths $record $Dict $Objects }
    return $record
}

function Get-PdfArrayText {
    <#
    .SYNOPSIS
      The text of an array value: written in place ([...], one level of inner arrays) or held in an object of its
      own (n 0 R); empty when the key is not there.
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER Key
      The key, without its slash.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([string]$Dict, [string]$Key, $Objects)
    $value = [regex]::Match($Dict, '/' + $Key + '\s*(\[[^\[\]]*(?:\[[^\[\]]*\][^\[\]]*)*\]|(\d+)\s+\d+\s+R)')
    if (-not $value.Success) { return '' }
    if ($value.Groups[2].Success) {
        if ($Objects.Contains($value.Groups[2].Value)) { return $Objects[$value.Groups[2].Value].Dict }
        return ''
    }
    return $value.Groups[1].Value
}

function Get-PdfDifferences {
    <#
    .SYNOPSIS
      The /Differences of a simple font's /Encoding (in the font, or in an encoding object of its own): code -> text by
      glyph name. A made-up name (letters and a number, such as /g12, /a3 or /cid34, as subset and Type3 fonts use) maps
      to $null: its glyph cannot be read. Any other name this table does not know is left out: its code reads as before. Without them, codes were read as Latin-1 and gave nonsense letters and
      signs (the Administrator, 2026-11-21: "Wingdings"-like text from a PDF with no Wingdings font).
    .PARAMETER Dict
      The font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable] code -> text or $null; $null when the font has no /Differences.
    #>
    param([string]$Dict, $Objects)
    $source = $Dict
    $reference = [regex]::Match($Dict, '/Encoding\s+(\d+)\s+\d+\s+R')
    if ($reference.Success -and $Objects.Contains($reference.Groups[1].Value)) { $source = $Objects[$reference.Groups[1].Value].Dict }
    $array = Get-PdfArrayText $source 'Differences' $Objects
    if (-not $array) { return $null }
    $glyphs = @{}
    $code = 0
    foreach ($token in [regex]::Matches($array, '(\d+)|/([^\s/\[\]()<>]+)')) {
        if ($token.Groups[1].Success) { $code = [int]$token.Groups[1].Value; continue }
        # A name can hold #xx escapes (/a#20b = "a b", ISO 32000-1, 7.3.5).
        $name = [regex]::Replace($token.Groups[2].Value, '#([0-9A-Fa-f]{2})',
            [System.Text.RegularExpressions.MatchEvaluator] { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
        if ($code -le 255) {
            $text = Get-PdfGlyphNameText $name
            if ($null -ne $text) { $glyphs[$code] = $text }
            elseif ($name -match $script:MadeUpGlyphName) { $glyphs[$code] = $null }
            else { [void]$glyphs.Remove($code) }   # a real name this table does not know: the code reads as Latin-1
        }
        $code++
    }
    return $glyphs
}

function Get-PdfGlyphNameText {
    <#
    .SYNOPSIS
      The text of a standard glyph name (Adobe Glyph List names for the letters, digits and common signs; accented
      letters; uniXXXX and uXXXX; a suffix such as .sc is dropped); $null for a name this table does not know.
    .PARAMETER Name
      The glyph name, without its slash.
    .OUTPUTS
      [string] or $null.
    #>
    param([string]$Name)
    $base = ($Name -split '\.')[0]
    if ($base -cmatch '^[A-Za-z]$') { return $base }
    if ($script:GlyphNames.ContainsKey($base)) { return $script:GlyphNames[$base] }
    # An accented letter: the letter and its mark (eacute = e + acute), joined into one character.
    $accent = [regex]::Match($base, '^([A-Za-z])(acute|grave|circumflex|dieresis|tilde|ring|cedilla|caron|macron|breve|ogonek|dotaccent|hungarumlaut|commaaccent)$')
    if ($accent.Success) { return ($accent.Groups[1].Value + [string][char]$script:GlyphMarks[$accent.Groups[2].Value]).Normalize() }
    if ($base -match '^uni([0-9A-Fa-f]{4})$' -or $base -match '^u([0-9A-Fa-f]{4,5})$') {
        $value = [Convert]::ToInt32($Matches[1], 16)
        if ($value -lt 0x10000 -and ($value -lt 0xD800 -or $value -gt 0xDFFF)) { return [string][char]$value }
    }
    return $null
}

function Add-PdfSimpleWidths {
    <#
    .SYNOPSIS
      A simple font's widths: /FirstChar and /Widths; without /Widths, the standard family's built-in widths.
    .PARAMETER Record
      The font record (filled in place).
    .PARAMETER Dict
      The font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([hashtable]$Record, [string]$Dict, $Objects)
    $first = 0
    $firstMatch = [regex]::Match($Dict, '/FirstChar\s+(\d+)')
    if ($firstMatch.Success) { $first = [int]$firstMatch.Groups[1].Value }
    # Widths are in glyph space: thousandths of an em, or for a Type3 font in units of its /FontMatrix.
    $unit = 0.001
    $matrix = [regex]::Match($Dict, '/FontMatrix\s*\[\s*(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))')
    if ($Dict -match '/Subtype\s*/Type3\b' -and $matrix.Success) {
        # The length of the matrix's first column: the glyph-space unit along the line, also for a rotated matrix.
        $a = [double]::Parse($matrix.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
        $b = [double]::Parse($matrix.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
        $unit = [Math]::Sqrt($a * $a + $b * $b)
    }
    $code = $first
    foreach ($number in [regex]::Matches((Get-PdfArrayText $Dict 'Widths' $Objects), '-?(?:\d+\.?\d*|\.\d+)')) {
        if ($code -gt 255) { break }
        $Record.Widths[$code] = [double]::Parse($number.Value, [Globalization.CultureInfo]::InvariantCulture) * $unit
        $code++
    }
    if ($Record.Widths.Count -gt 0) { return }
    $base = [regex]::Match($Dict, '/BaseFont\s*/([^\s/<>\[\]()]+)').Groups[1].Value
    if ($base -match 'Courier') { $Record.DefaultWidth = 0.6 }
    elseif ($base -match 'Times') { $Record.Standard = $script:TimesWidths }
    elseif ($base -match 'Helvetica|Arial') { $Record.Standard = $script:HelveticaWidths }
}

function Add-PdfCidWidths {
    <#
    .SYNOPSIS
      A composite font's widths, from its descendant font: /DW (the default) and /W, whose entries are either
      "first [w1 w2 ...]" or "first last w".
    .PARAMETER Record
      The font record (filled in place).
    .PARAMETER Dict
      The Type0 font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([hashtable]$Record, [string]$Dict, $Objects)
    $descendant = [regex]::Match((Get-PdfArrayText $Dict 'DescendantFonts' $Objects), '(\d+)\s+\d+\s+R')
    if (-not $descendant.Success -or -not $Objects.Contains($descendant.Groups[1].Value)) { return }
    $cid = $Objects[$descendant.Groups[1].Value].Dict
    $Record.DefaultWidth = 1.0
    $dw = [regex]::Match($cid, '/DW\s+(\d+(?:\.\d+)?)')
    if ($dw.Success) { $Record.DefaultWidth = [double]::Parse($dw.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture) / 1000 }
    # Read token by token (a hostile /W can hold millions): the entry cap and the clock are checked in every loop.
    $token = (Get-PdfRegex '\[|\]|-?(?:\d+\.?\d*|\.\d+)').Match((Get-PdfArrayText $cid 'W' $Objects))
    if ($token.Success) { $token = $token.NextMatch() }   # the outer [
    $numbers = New-Object System.Collections.Generic.List[double]
    $count = 0
    while ($token.Success) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        if ($token.Value -eq '[') {
            # first [w1 w2 ...]: the number before the bracket is the first glyph.
            if ($numbers.Count -eq 0) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }   # [ with no first glyph
            $glyph = [int]$numbers[$numbers.Count - 1]; $numbers.Clear()
            for ($token = $token.NextMatch(); $token.Success -and $token.Value -ne ']'; $token = $token.NextMatch()) {
                if ((++$count % 2000) -eq 0) { Assert-PdfTime }
                if ($Record.Widths.Count -ge $script:MaxWidthEntries) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }
                $Record.Widths[$glyph] = [double]$token.Value / 1000; $glyph++
            }
        }
        elseif ($token.Value -ne ']') {
            $numbers.Add([double]$token.Value)
            if ($numbers.Count -eq 3) {   # first last w
                for ($g = [int]$numbers[0]; $g -le [int]$numbers[1]; $g++) {
                    if ($Record.Widths.Count -ge $script:MaxWidthEntries) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }
                    $Record.Widths[$g] = $numbers[2] / 1000
                }
                $numbers.Clear()
            }
        }
        if ($token.Success) { $token = $token.NextMatch() }
    }
}

# ==== PdfShown.psm1 ====
# Per font: shown strings decoded and measured before (text and width in ems), up to this many each (speed).
$script:MaxFontCache = 20000
# Entries of what a line showed before they are summed (Compress-PdfShown).
$script:MaxShownItems = 1024
$script:MaxCacheKey = 256   # longer strings are rarely shown twice: not kept (memory)
# Shown hex strings decoded before (speed): short tokens only, emptied when full.
$script:ShownHex = New-Object 'System.Collections.Generic.Dictionary[string,string]'
$script:MaxShownHex = 50000

function ConvertFrom-PdfGlyphs {
    <#
    .SYNOPSIS
      The text of one shown string: each glyph code looked up in the font's ToUnicode map; Latin-1 without a map.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    .PARAMETER Font
      A font record, or $null.
    #>
    param([string]$Raw, $Font)
    # A picture font (Wingdings and the like): one space per glyph, so its symbols never read as PII (the Administrator, 2026-11-21).
    # First, before its glyph names: a picture font's codes are never letters.
    if ($Font -and $Font.ContainsKey('Picture') -and $Font.Picture) { return ' ' * [int][Math]::Ceiling($Raw.Length / [Math]::Max(1, [int]$Font.Bytes)) }
    if ($Font -and -not $Font.CMap -and $Font.ContainsKey('Glyphs') -and $Font.Glyphs) { return ConvertFrom-PdfNamedGlyphs $Raw $Font }
    if (-not $Font -or -not $Font.CMap) { return $Raw }
    $cmap = $Font.CMap
    if ($cmap.ContainsKey('Unreadable')) {
        if ($Raw.Length -gt 0) { Add-PdfNote 'A font in this PDF could not be read: check by hand' }
        return ''
    }
    # A font record lives for one PDF, so a note its first decode added is already on that PDF.
    $cache = $null; if ($Font.ContainsKey('TextCache')) { $cache = $Font.TextCache }
    if ($cache -and $cache.ContainsKey($Raw)) { return $cache[$Raw] }
    $out = New-Object System.Text.StringBuilder
    $step = [Math]::Max(1, [int]$cmap.Bytes)
    for ($i = 0; $i + $step -le $Raw.Length; $i += $step) {
        if ((($i / $step) % 1000) -eq 999) { Assert-PdfTime }
        [long]$number = 0
        for ($j = 0; $j -lt $step; $j++) { $number = $number * 256 + [int]$Raw[$i + $j] }
        $text = $cmap.Map[$number.ToString('X' + (2 * $step))]
        if ($null -eq $text) {
            foreach ($range in $cmap.Ranges) {
                if ($number -ge $range.Low -and $number -le $range.High) {
                    $text = ConvertFrom-Utf16Hex ($range.Prefix + (($range.Start + ($number - $range.Low)) -band 0xFFFF).ToString('X4'))
                    break
                }
            }
        }
        if ($null -ne $text) { [void]$out.Append($text) }
        elseif ($step -eq 1) { [void]$out.Append($Raw[$i]) }
        elseif ($number -ne 0) { Add-PdfNote 'Some characters in this PDF could not be read (a font without them in its map): check by hand' }
    }
    $result = $out.ToString()
    if ($cache -and $cache.Count -lt $script:MaxFontCache -and $Raw.Length -le $script:MaxCacheKey) { $cache[$Raw] = $result }
    return $result
}

function ConvertFrom-PdfNamedGlyphs {
    <#
    .SYNOPSIS
      The text of one shown string in a simple font with /Differences: a code it names gets its glyph's text (a
      glyph that cannot be read: a space and a "check by hand" note); any other code reads as Latin-1.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    .PARAMETER Font
      The font record (Glyphs set).
    #>
    param([string]$Raw, $Font)
    if ($Font.TextCache.ContainsKey($Raw)) { return $Font.TextCache[$Raw] }
    $out = New-Object System.Text.StringBuilder
    foreach ($char in $Raw.ToCharArray()) {
        $code = [int]$char
        if (-not $Font.Glyphs.ContainsKey($code)) { [void]$out.Append($char); continue }
        $text = $Font.Glyphs[$code]
        if ($null -eq $text) { Add-PdfNote 'A font in this PDF could not be read: check by hand'; [void]$out.Append(' ') }
        else { [void]$out.Append($text) }
    }
    $result = $out.ToString()
    if ($Font.TextCache.Count -lt $script:MaxFontCache -and $Raw.Length -le $script:MaxCacheKey) { $Font.TextCache[$Raw] = $result }
    return $result
}

function Get-PdfShownRaw {
    <#
    .SYNOPSIS
      The bytes of a shown string token: from a cache when decoded before (Chrome writes each glyph as its own <0024>), else decoded.
    .PARAMETER Token
      One PDF string token.
    #>
    param([string]$Token)
    $raw = $null
    if ($script:ShownHex.TryGetValue($Token, [ref]$raw)) { return $raw }
    $raw = ConvertFrom-PdfString $Token
    if ($Token.Length -le 66) {   # short hex and literal strings (a token decodes the same in every PDF)
        if ($script:ShownHex.Count -ge $script:MaxShownHex) { $script:ShownHex.Clear() }
        $script:ShownHex[$Token] = $raw
    }
    return $raw
}

function Get-PdfShownText {
    <#
    .SYNOPSIS
      The text of shown bytes in a font: from the font's cache when decoded before (ConvertFrom-PdfGlyphs fills it).
    .PARAMETER Raw
      The bytes as Latin-1 text.
    .PARAMETER Font
      The font record, or $null.
    #>
    param([string]$Raw, $Font)
    $text = $null
    if ($Font -and $Font.TextCache -and $Font.TextCache.TryGetValue($Raw, [ref]$text)) { return $text }
    return (ConvertFrom-PdfGlyphs $Raw $Font)
}

function Compress-PdfShown {
    <#
    .SYNOPSIS
      Keeps the list of what a line showed short: past $MaxShownItems entries, they become their width (one
      number), so a line of millions of strings cannot hold gigabytes (fresh review r14).
    .PARAMETER State
      The text state.
    #>
    param([hashtable]$State)
    if ($State.Shown.Count -lt $script:MaxShownItems) { return }
    $sum = Get-PdfAdvance $State
    $State.Shown.Clear()
    $State.Shown.Add([double]$sum)
}

function Get-PdfAdvance {
    <#
    .SYNOPSIS
      The width shown since the line's start, in text space units: the shown strings measured by their fonts, plus
      TJ adjustments (kept as numbers).
    .PARAMETER State
      The text state.
    .OUTPUTS
      [double]
    #>
    param([hashtable]$State)
    $sum = 0.0
    $count = 0
    foreach ($item in $State.Shown) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        if ($item -is [double]) { $sum += $item; continue }
        $font = $item[2]
        $ems = 0.0
        if ($font -and -not $font.Vertical -and $font.EmCache -and $font.EmCache.TryGetValue($item[0], [ref]$ems)) {
            # Measured before (Get-PdfShowWidth's cache): the same sum without the call (speed).
            $step = [Math]::Max(1, [int]$font.Bytes)
            $glyphs = [Math]::Max(1, [int]($item[0].Length / $step))
            $spaces = 0; if ($step -eq 1) { $spaces = $item[0].Length - $item[0].Replace(' ', '').Length }
            $sum += ($ems * $item[3] + $item[4] * $glyphs + $item[5] * $spaces) * $item[6]
        }
        else { $sum += Get-PdfShowWidth $item[0] $item[1] $item[2] $item[3] $item[4] $item[5] $item[6] }
    }
    return $sum
}

function Get-PdfShowWidth {
    <#
    .SYNOPSIS
      How far a shown string moves the text position, in text space units: each glyph's width from its font (an
      estimate from the text only where the font gives none) times the font size, plus character and word spacing
      (Tc, Tw), times the horizontal scaling (Tz).
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    .PARAMETER Text
      The decoded text (for the estimate).
    .PARAMETER Font
      The font record, or $null.
    .PARAMETER Size
      The font size.
    .PARAMETER CharSpace
      Tc.
    .PARAMETER WordSpace
      Tw.
    .PARAMETER Scale
      Tz as a fraction.
    .OUTPUTS
      [double]
    #>
    param([string]$Raw, [string]$Text, $Font, [double]$Size, [double]$CharSpace, [double]$WordSpace, [double]$Scale)
    $font = $Font
    $known = $font -and ($font.Widths.Count -gt 0 -or $font.DefaultWidth -ge 0 -or $font.Standard)
    $step = 1; if ($font) { $step = [Math]::Max(1, [int]$font.Bytes) }
    $glyphs = [Math]::Max(1, [int]($Raw.Length / $step))
    $spaces = 0; if ($step -eq 1) { $spaces = $Raw.Length - $Raw.Replace(' ', '').Length }   # no copy per character
    if ($font -and $font.Vertical) { $ems = [double]$glyphs }   # vertical writing: 1 em down per glyph (the /DW2 default)
    elseif (-not $known) { $ems = (Get-PdfTextWidth $Text 1.0) }
    elseif ($font.ContainsKey('EmCache') -and $font.EmCache.ContainsKey($Raw)) { $ems = $font.EmCache[$Raw] }
    else {
        $ems = 0.0
        for ($i = 0; $i + $step -le $Raw.Length; $i += $step) {
            if (($i % 8192) -eq 0) { Assert-PdfTime }
            $code = 0; for ($j = 0; $j -lt $step; $j++) { $code = $code * 256 + [int]$Raw[$i + $j] }
            if ($font.Widths.ContainsKey($code)) { $ems += $font.Widths[$code] }
            elseif ($font.Standard -and $code -ge 32 -and $code -le 126) { $ems += $font.Standard[$code - 32] / 1000 }
            elseif ($font.DefaultWidth -ge 0) { $ems += $font.DefaultWidth }
            else { $ems += 0.5 }
        }
        if ($font.ContainsKey('EmCache') -and $font.EmCache.Count -lt $script:MaxFontCache -and $Raw.Length -le $script:MaxCacheKey) { $font.EmCache[$Raw] = $ems }
    }
    if ($font -and $font.Vertical) { return ($ems * $Size + $CharSpace * $glyphs + $WordSpace * $spaces) }   # Tz scales across, not down
    return (($ems * $Size + $CharSpace * $glyphs + $WordSpace * $spaces) * $Scale)
}

function Get-PdfTextWidth {
    <#
    .SYNOPSIS
      The estimated width of shown text in ems times a size: average glyph widths (Helvetica-like), the last resort
      for a font that gives no widths.
    .PARAMETER Text
      The shown text.
    .PARAMETER Size
      The font size.
    .OUTPUTS
      [double]
    #>
    param([string]$Text, [double]$Size)
    $ems = 0.0
    foreach ($c in $Text.ToCharArray()) {
        if ([char]::IsDigit($c)) { $ems += 0.556 }
        elseif ($c -eq ' ' -or $c -eq '.' -or $c -eq ',' -or $c -eq ':' -or $c -eq ';') { $ems += 0.278 }
        elseif ($c -eq '-' -or $c -eq '(' -or $c -eq ')' -or $c -eq '/') { $ems += 0.333 }
        elseif ([char]::IsUpper($c)) { $ems += 0.667 }
        else { $ems += 0.5 }
    }
    return ($ems * $Size)
}

# ==== PdfLayout.psm1 ====
# Gaps along a line (a same-line Td move past the real end of the text before it, or a TJ adjustment), in ems: a
# space when wider than the block's letter spacing plus WordGapEms. The letter spacing is the median gap between two
# one-glyph pieces when a block has at least MinSpacingGaps of them and they are at least half of its gaps (gaps of
# ColumnGapEms or more are columns, not spacing), else 0 (fresh reviews, 2026-11-21).
$script:WordGapEms = 0.2
$script:MinSpacingGaps = 4
$script:ColumnGapEms = 1.0
# How near (a share of the block's letter spacing) a short line's letter gaps must be to take the block's spacing.
$script:SpacingMatch = 0.25

# Symbol-font and other private-use characters (Complete-PdfTextBlock), built once (speed).
$script:SymbolChars = [regex]'[\uF020-\uF0FF]'
$script:PrivateChars = [regex]'[\uE000-\uF8FF]'

function Get-PdfSpacing {
    <#
    .SYNOPSIS
      The letter spacing of some text pieces: the median gap between two one-glyph pieces (line breaks are not
      glyphs), with how many such gaps there are and how many gaps in all. Base is 0 unless there are at least
      $MinSpacingGaps letter gaps and they are at least half of the gaps.
    .PARAMETER Pieces
      Strings, and gaps as @{ Gap; Single }.
    .OUTPUTS
      [hashtable] Base, Letters, Gaps.
    #>
    param($Pieces)
    $letters = [System.Collections.Generic.List[double]]::new()
    $gaps = 0
    for ($k = 0; $k -lt $Pieces.Count; $k++) {
        if (($k % 2000) -eq 1999) { Assert-PdfTime }
        $piece = $Pieces[$k]
        if ($piece -isnot [hashtable]) { continue }
        $gaps++
        $after = $null; if ($k + 1 -lt $Pieces.Count) { $after = $Pieces[$k + 1] }
        $glyph = $after -is [string] -and $after.Length -eq 1 -and $after -ne "`n"
        if ($piece.Single -and $glyph -and $piece.Gap -lt $script:ColumnGapEms) { $letters.Add($piece.Gap) }
    }
    $base = 0.0
    if ($letters.Count -ge $script:MinSpacingGaps -and $letters.Count * 2 -ge $gaps) {
        $sorted = $letters.ToArray(); [Array]::Sort($sorted); $base = [Math]::Max(0.0, $sorted[[int]($sorted.Count / 2)])
    }
    return @{ Base = $base; Letters = $letters.Count; Gaps = $gaps; Values = $letters.ToArray() }
}

function Add-PdfLine {
    <#
    .SYNOPSIS
      Writes one line of a text block: its strings, and a space for each gap wider than its letter spacing plus
      $WordGapEms. The letter spacing is the line's own (Get-PdfSpacing); a short line with only a few letter gaps
      ("SSN:" spaced out) takes the block's, so it is not read as "S S N :"; a line without letter gaps has none, so
      a spaced heading never joins the words below it (fresh reviews r12, r14).
    .PARAMETER Line
      The line's pieces: strings, and gaps as @{ Gap; Single }.
    .PARAMETER Out
      The StringBuilder that receives the line.
    .PARAMETER BlockBase
      The letter spacing of the whole block.
    #>
    param($Line, $Out, [double]$BlockBase)
    $spacing = Get-PdfSpacing $Line
    $base = $spacing.Base
    # The block's spacing only when every letter gap of the line is near it (a few spaced letters, as in the block);
    # else a normal word gap next to two initials would stop being a space (review r15).
    if ($base -eq 0 -and $BlockBase -gt 0 -and $spacing.Letters -gt 0 -and $spacing.Letters * 2 -ge $spacing.Gaps) {
        $near = @($spacing.Values | Where-Object { [Math]::Abs($_ - $BlockBase) -le $BlockBase * $script:SpacingMatch }).Count
        if ($near -eq $spacing.Letters) { $base = $BlockBase }
    }
    $limit = $base + $script:WordGapEms
    $count = 0
    foreach ($piece in $Line) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        if ($piece -is [hashtable]) { if ($piece.Gap -gt $limit) { [void]$Out.Append(' ') } }
        else { [void]$Out.Append($piece) }
    }
}

function Complete-PdfTextBlock {
    <#
    .SYNOPSIS
      Ends a text block: its gaps become spaces where wider than the block's letter spacing plus $WordGapEms;
      symbol-font characters back to their plain ones; other private-use characters a space and a note; a line break.
    .PARAMETER State
      The text state of Convert-PdfContentToText.
    #>
    param([hashtable]$State)
    # Letter spacing is measured per line (a letter-spaced heading must not join the words of the body below it); a
    # short line whose few letter gaps match the block's takes the block's (Add-PdfLine). A gap is a space only when
    # wider than the letter spacing plus $WordGapEms.
    $gapped = $false
    $count = 0
    foreach ($piece in $State.Pieces) {
        if ((++$count % 20000) -eq 0) { Assert-PdfTime }
        if ($piece -is [hashtable]) { $gapped = $true; break }
    }
    if (-not $gapped) { $text = -join $State.Pieces }   # no gaps (most blocks): the strings as they are (speed)
    else {
        $out = [System.Text.StringBuilder]::new()
        $blockBase = (Get-PdfSpacing $State.Pieces).Base
        $line = [System.Collections.Generic.List[object]]::new()
        foreach ($piece in $State.Pieces) {
            if ((++$count % 2000) -eq 0) { Assert-PdfTime }
            if ($piece -is [string] -and $piece.StartsWith("`n")) { Add-PdfLine $line $out $blockBase; [void]$out.Append($piece); $line.Clear() }
            else { $line.Add($piece) }
        }
        Add-PdfLine $line $out $blockBase
        $text = $out.ToString()
    }
    # Microsoft symbol fonts map their characters to U+F020-U+F0FF (Word's bullet is U+F0B7): back to U+0020-U+00FF.
    if ($script:SymbolChars.IsMatch($text)) {
        $text = [regex]::Replace($text, '[\uF020-\uF0FF]', [System.Text.RegularExpressions.MatchEvaluator] { param($m) [string][char]([int]$m.Value[0] - 0xF000) })
    }
    # A font can map a glyph to a private-use character (U+E000-U+F8FF): it has no meaning (Chrome does this for some
    # punctuation), so it becomes a space and the PDF gets a note.
    if ($script:PrivateChars.IsMatch($text)) {
        Add-PdfNote 'Some characters in this PDF could not be read (a font without them in its map): check by hand'
        $text = $text -replace '[\uE000-\uF8FF]', ' '
    }
    [void]$State.Out.Append($text).Append("`n")
    $State.Pieces.Clear()
    $State.Shown.Clear()
}

# ==== PdfContent.psm1 ====
# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = (Get-PdfStringPattern) + '|%[^\r\n]*|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'



# Tokens per Matches walk (Read-PdfTextObject): bounds the memory a walk holds.
$script:WalkTokens = 5000
# Entries one [ ] array, or the operands of one operator, may hold (a TJ line has tens to thousands); more makes
# the PDF NOT SCANNED.
$script:MaxArrayItems = 100000

# Keywords that are values, not operators, inside an array (a marked-content property list: [(lbl) true]).
$script:ArrayValues = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($value in @('true', 'false', 'null')) { [void]$script:ArrayValues.Add($value) }

# The operators that change the text (Invoke-PdfTextOperator); every other operator only drops its operands.
$script:TextOperators = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($operator in @('Tf', 'Tc', 'Tw', 'Tz', 'TJ', 'Tj', "'", '"', 'Td', 'TD', 'T*', 'Tm', 'ET')) { [void]$script:TextOperators.Add($operator) }

# Number parsing: PDF numbers use a point, whatever the computer's culture.
$script:NumberStyle = [Globalization.NumberStyles]::Float
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
# The end of an inline image's bytes (BI <keys> ID <bytes> EI): white space, EI, then a delimiter or the end. After a
# BI operator (a token, never the word BI in a string) the reader steps over the bytes to here, so image bytes never
# join with page text into one token (fresh reviews, 2026-11-21).
$script:InlineImageEnd = '\sEI(?=[\s/\[<(%]|$)'
function Convert-PdfContentToText {
    <#
    .SYNOPSIS
      The text a content stream shows: Tj, TJ, ' and " strings in the current font; a line break at each line move
      and text block end; a space at a wide TJ gap and at a same-line Td move that leaves a gap (Add-PdfMove), so
      text placed piece by piece (Chrome splits at kerning pairs and draws glyph by glyph) joins again.
    .PARAMETER Content
      The (inflated) content stream.
    .PARAMETER Fonts
      Output of Get-PdfFonts or Get-PdfFontTable.
    .OUTPUTS
      [string]
    #>
    param([string]$Content, [hashtable]$Fonts)
    $state = @{ Out = (New-Object System.Text.StringBuilder); Pieces = (New-Object System.Collections.Generic.List[object]); Font = $null; Size = 12.0
        Shown = (New-Object System.Collections.Generic.List[object]); CharSpace = 0.0; WordSpace = 0.0; Scale = 1.0; LastLength = 0 }
    $tokens = Get-PdfRegex $script:PdfTokenPattern
    $imageEnd = Get-PdfRegex $script:InlineImageEnd
    # The whole stream is walked token by token (one Matches walk per text object: fast), also between text objects:
    # text state set there applies, and text shown outside BT/ET (against the spec) is read too (fresh review r13).
    $pos = 0
    while ($pos -lt $Content.Length) {
        $after = Read-PdfTextObject $Content $pos $state $Fonts $tokens $imageEnd
        if ($after -le $pos) { break }
        $pos = $after
    }
    Complete-PdfTextBlock $state
    return $state.Out.ToString()
}

function Read-PdfTextObject {
    <#
    .SYNOPSIS
      Reads the content stream token by token from Start to the next ET (or the stream's end), and returns where the
      reader goes on. A "(" between tokens is a string not read: it throws (Assert-PdfGap).
    .PARAMETER Content
      The content stream.
    .PARAMETER Start
      Where to start reading.
    .PARAMETER State
      The text state.
    .PARAMETER Fonts
      The font table.
    .PARAMETER Tokens
      The timed regex of $PdfTokenPattern.
    .PARAMETER ImageEnd
      The timed regex of $InlineImageEnd.
    .OUTPUTS
      [int] the position after ET.
    #>
    param([string]$Content, [int]$Start, [hashtable]$State, [hashtable]$Fonts, $Tokens, $ImageEnd)
    $stack = New-Object System.Collections.Generic.List[object]
    $array = $null; $count = 0; $next = $Start; $inImage = $false
    # One Matches walk per stretch (a Match call per token costs 40 times more; speed): it starts again past an inline
    # image, and every $WalkTokens tokens, because a match list keeps every match it gave out (memory, review r14).
    for ($walk = $true; $walk; ) {
    $walk = $false; $inWalk = 0
    foreach ($token in $Tokens.Matches($Content, $next)) {
        if ((++$inWalk) -gt $script:WalkTokens) { $walk = $true; break }   # this token is found again from $next
        if ((++$count % 100) -eq 0) { Assert-PdfTime }
        $t = $token.Value
        if ($token.Index -gt $next -and $Content.IndexOf('(', $next, $token.Index - $next) -ge 0) { Assert-PdfGap $Content $next $token.Index }
        $next = $token.Index + $token.Length
        $first = $t[0]; if ($first -eq '%') { continue }   # a comment
        if ($t -ceq 'BI') { $inImage = $true; $stack.Clear(); continue }
        if ($inImage) {
            if ($t -ceq 'ID') { $next = Skip-PdfInlineImage $Content $next $ImageEnd; $inImage = $false; $walk = $true; break }   # else: its keys
            continue
        }
        if ($t -eq '[') { Complete-PdfOpenArray $array $State; $array = New-Object System.Collections.Generic.List[string]; continue }
        if ($t -eq ']') { if ($null -ne $array) { $stack.Add($array) }; $array = $null; continue }
        if ($first -eq '(' -or $first -eq '<' -or $first -eq '/' -or ($first -ge '0' -and $first -le '9') -or $first -eq '-' -or $first -eq '+' -or $first -eq '.' -or
            ($null -ne $array -and $script:ArrayValues.Contains($t))) {   # true, false and null are values in an array
            if ($null -eq $array) { if ($stack.Count -ge $script:MaxArrayItems) { Stop-PdfHostile }; $stack.Add($t); continue }
            if ($array.Count -ge $script:MaxArrayItems) { Stop-PdfHostile }   # no real array is this long (r14)
            $array.Add($t); continue
        }
        # Only text operators change the text; drawing operators just drop their operands.
        if ($null -ne $array) { Complete-PdfOpenArray $array $State; $array = $null }   # an operator in an open [ ] (damaged)
        if ($t -ceq 'Td') { Add-PdfMove $stack $State }   # the most common operator: straight to its handler (speed)
        elseif ($script:TextOperators.Contains($t)) { Invoke-PdfTextOperator $t $stack $State $Fonts }
        $stack.Clear()
        if ($t -ceq 'ET') { return $next }
    }
    }
    if ($next -lt $Content.Length -and $Content.IndexOf('(', $next) -ge 0) { Assert-PdfGap $Content $next $Content.Length }
    Complete-PdfOpenArray $array $State
    return $Content.Length
}


function Complete-PdfOpenArray {
    <#
    .SYNOPSIS
      A [ not closed before the next operator, the next [ or the stream's end (damaged): its strings are shown where
      it stops, with a note, so no text is lost (reviews r15, r16).
    .PARAMETER Array
      The open array, or $null.
    .PARAMETER State
      The text state.
    #>
    param($Array, [hashtable]$State)
    if ($null -eq $Array) { return }
    $strings = $false
    foreach ($item in $Array) { if ($item[0] -eq '(' -or $item[0] -eq '<') { $strings = $true; break } }
    if (-not $strings) { return }   # no strings: nothing lost
    Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine; an unclosed [ ]): check by hand'
    Add-PdfArrayPieces $Array $State
}

function Add-PdfStackText {
    <#
    .SYNOPSIS
      For a show operator (Tj ' " TJ) whose operands are not the usual one string (Tj) or one array (TJ): shows every
      string and array on the stack, in order, with a "check by hand" note. Returns false (nothing done) for the
      usual case, which the caller handles faster.
    .PARAMETER Stack
      The operands.
    .PARAMETER State
      The text state.
    .PARAMETER Operator
      The show operator.
    .OUTPUTS
      [bool]
    #>
    param($Stack, [hashtable]$State, [string]$Operator)
    $texts = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Stack) {
        if ($item -is [System.Collections.Generic.List[string]]) { $texts.Add($item) }
        elseif ($item -is [string] -and $item.Length -ge 2 -and ($item[0] -eq '(' -or ($item[0] -eq '<' -and $item[1] -ne '<'))) { $texts.Add($item) }
    }
    if ($texts.Count -eq 0) { return $false }
    $usual = $texts.Count -eq 1 -and [object]::ReferenceEquals($texts[0], $Stack[$Stack.Count - 1]) -and (($Operator -ceq 'TJ') -eq ($texts[0] -is [System.Collections.Generic.List[string]]))
    if ($usual) { return $false }
    Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine; text with stray operands): check by hand'
    $gap = $false
    foreach ($text in $texts) {
        Assert-PdfTime   # up to $MaxArrayItems operands: the clock is checked for each (review r19)
        if ($gap) { $State.Pieces.Add(' ') }   # separate operands: never one word
        if ($text -is [System.Collections.Generic.List[string]]) { Add-PdfArrayPieces $text $State } else { Add-PdfArrayPieces @($text) $State }
        $gap = $true
    }
    return $true
}

function Invoke-PdfTextOperator {
    <#
    .SYNOPSIS
      Applies one content-stream operator to the text state: font change, shown text, line and word breaks.
    .PARAMETER Operator
      The operator token.
    .PARAMETER Stack
      Its operands.
    .PARAMETER State
      Out, Pieces (the text block in work: strings, and gaps as @{ Gap (ems); Single (after a one-glyph piece) }),
      Font, Size, Shown (what was shown since the line's start, measured only when a move needs it: Get-PdfAdvance),
      CharSpace, WordSpace, Scale (Tc, Tw, Tz), LastLength.
    .PARAMETER Fonts
      The font table.
    #>
    param([string]$Operator, $Stack, [hashtable]$State, [hashtable]$Fonts)
    $last = $null; if ($Stack.Count -gt 0) { $last = $Stack[$Stack.Count - 1] }
    $shown = $null   # the string a Tj, ' or " shows: only a string token counts
    if ($Operator -ceq '"' -and $Stack.Count -ge 3) {   # aw ac (string) ": word and character spacing first
        $State.WordSpace = ConvertTo-PdfNumber $Stack[$Stack.Count - 3]; $State.CharSpace = ConvertTo-PdfNumber $Stack[$Stack.Count - 2]
    }
    $quote = $Operator -ceq "'" -or $Operator -ceq '"'
    if ($quote) { $State.Pieces.Add("`n"); $State.Shown.Clear() }
    # A show operator with more than its one string or array (damaged): every string and array is read (review r18).
    $odd = $Stack.Count -gt 1 -or ($Operator -ceq 'TJ' -and $last -isnot [System.Collections.Generic.List[string]])   # a cheap test first
    if ($odd -and ($quote -or $Operator -ceq 'Tj' -or $Operator -ceq 'TJ') -and (Add-PdfStackText $Stack $State $Operator)) { return }
    if (($quote -or $Operator -ceq 'Tj') -and $last -is [string] -and $last.Length -ge 2 -and ($last[0] -eq '(' -or $last[0] -eq '<')) {
        # The font's text cache is looked up here first: a call less per shown string (speed).
        $raw = $null; $shown = $null
        $raw = Get-PdfShownRaw $last
        $font = $State.Font
        if (-not ($font -and $font.TextCache -and $font.TextCache.TryGetValue($raw, [ref]$shown))) { $shown = Get-PdfShownText $raw $font }
        $State.Shown.Add(@($raw, $shown, $State.Font, $State.Size, $State.CharSpace, $State.WordSpace, $State.Scale)); Compress-PdfShown $State
        $State.LastLength = $shown.Length
        $State.Pieces.Add($shown)
    }
    $pieces = $State.Pieces
    switch -CaseSensitive ($Operator) {
        'Tf' { if ($Stack.Count -ge 2) { $State.Font = $Fonts[([string]$Stack[$Stack.Count - 2]).TrimStart('/')]; $State.Size = [Math]::Abs((ConvertTo-PdfNumber $last)) } }
        'Tc' { $State.CharSpace = ConvertTo-PdfNumber $last }
        'Tw' { $State.WordSpace = ConvertTo-PdfNumber $last }
        'Tz' { $State.Scale = (ConvertTo-PdfNumber $last) / 100 }
        'TJ' { if ($last -is [System.Collections.Generic.List[string]]) { Add-PdfArrayPieces $last $State } }
        'Td' { Add-PdfMove $Stack $State }
        'TD' { $pieces.Add("`n"); $State.Shown.Clear() }
        'T*' { $pieces.Add("`n"); $State.Shown.Clear() }
        'Tm' { $pieces.Add("`n"); $State.Shown.Clear() }
        'ET' { Complete-PdfTextBlock $State }
    }
}

function Add-PdfMove {
    <#
    .SYNOPSIS
      A Td move: down or up starts a new line; along the line, a gap piece: how far (in ems) the move starts the next
      piece past the end of the text before it (its real width). Complete-PdfTextBlock decides which gaps are spaces.
    .PARAMETER Stack
      The operands (tx ty).
    .PARAMETER State
      The text state.
    #>
    param($Stack, [hashtable]$State)
    if ($Stack.Count -lt 2) { return }
    $tx = 0.0; $ty = 0.0   # a token that is no number counts as 0 (as ConvertTo-PdfNumber), parsed here without a call (speed)
    if (-not [double]::TryParse([string]$Stack[$Stack.Count - 2], $script:NumberStyle, $script:Invariant, [ref]$tx)) { $tx = 0.0 }
    if (-not [double]::TryParse([string]$Stack[$Stack.Count - 1], $script:NumberStyle, $script:Invariant, [ref]$ty)) { $ty = 0.0 }
    $advance = 0.0
    if ($State.Shown.Count -gt 0 -and (($State.Font -and $State.Font.Vertical -and $tx -eq 0) -or ($ty -eq 0 -and $tx -ge 0))) { $advance = Get-PdfAdvance $State }   # only when a gap is decided
    if ($State.Font -and $State.Font.Vertical) {
        # Vertical writing (Identity-V): the line runs down; a move across it starts a new line.
        if ($tx -ne 0) { $State.Pieces.Add("`n") }
        else { $State.Pieces.Add(@{ Gap = [double](([Math]::Abs($ty) - $advance) / [Math]::Max(1.0, $State.Size)); Single = ($State.LastLength -eq 1) }) }
    }
    elseif ($ty -ne 0) { $State.Pieces.Add("`n") }
    elseif ($tx -lt 0) { $State.Pieces.Add(@{ Gap = [double]::MaxValue; Single = $false }) }   # back along the line: a new column
    else { $State.Pieces.Add(@{ Gap = [double](($tx - $advance) / [Math]::Max(1.0, $State.Size)); Single = ($State.LastLength -eq 1) }) }
    $State.Shown.Clear()
}

function Skip-PdfInlineImage {
    <#
    .SYNOPSIS
      Where the reader goes on after an inline image's ID: past the image bytes, white space and EI. Without an EI the
      rest of the stream is image bytes: a note, and the end of the stream.
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the ID operator ended.
    .PARAMETER ImageEnd
      The timed regex of $InlineImageEnd.
    .OUTPUTS
      [int] the position after EI.
    #>
    param([string]$Content, [int]$From, $ImageEnd)
    $end = $ImageEnd.Match($Content, [Math]::Min($Content.Length, $From + 1))   # ID is followed by one white-space byte
    if ($end.Success) { return ($end.Index + $end.Length) }
    Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'
    return $Content.Length
}

function ConvertTo-PdfNumber {
    # A PDF number token as a number; 0 when it is not one.
    param($Token)
    $number = 0.0
    if (-not [double]::TryParse([string]$Token, $script:NumberStyle, $script:Invariant, [ref]$number)) { return 0.0 }
    return $number
}

function Add-PdfArrayPieces {
    <#
    .SYNOPSIS
      The pieces of a TJ array: each string in the font, and each adjustment that moves right (a negative number,
      in thousandths of an em) as a gap piece. Complete-PdfTextBlock decides which gaps are spaces.
    .PARAMETER Items
      The array's string and number tokens.
    .PARAMETER State
      The text state.
    #>
    param($Items, [hashtable]$State)
    $count = 0
    foreach ($item in $Items) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        if ($item.Length -ge 2 -and ($item[0] -eq '(' -or $item[0] -eq '<')) {
            $raw = Get-PdfShownRaw $item
            $text = Get-PdfShownText $raw $State.Font
            $State.Pieces.Add($text)
            $State.Shown.Add(@($raw, $text, $State.Font, $State.Size, $State.CharSpace, $State.WordSpace, $State.Scale)); Compress-PdfShown $State
            $State.LastLength = $text.Length
        }
        else {
            $move = -(ConvertTo-PdfNumber $item) / 1000   # in ems; positive moves on along the line
            if ($State.Font -and $State.Font.Vertical) { $move = -$move }   # vertical: a positive number moves down (forward)
            $scale = $State.Scale; if ($State.Font -and $State.Font.Vertical) { $scale = 1.0 }   # Tz scales across, not down
            $State.Shown.Add([double]($move * $State.Size * $scale)); Compress-PdfShown $State
            if ($move -gt 0) { $State.Pieces.Add(@{ Gap = [double]$move; Single = ($State.LastLength -eq 1) }) }
        }
    }
}

# ==== PdfText.psm1 ====
# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:PdfLatin1 = [System.Text.Encoding]::GetEncoding(28591)


$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'
# A scanned PDF's JPEG page images, found by Get-PdfAllText: read with Windows OCR when the PDF has no text.
$script:PdfScanImages = @()
# The scanned images of a PDF that also has text (the Administrator, 2026-11-26: plan step 3), read with Windows OCR: every image of
# a page with no text, and a page-size image of a page with visible text (a scan with a banner or a stamp). A page with
# invisible text (render mode 3: a "searchable" scan's text layer) is read from that text only, so its hits are not
# found twice and no OCR time is spent (reviews of step 3).
$script:PdfPageImages = @()
# On a page with visible text, only an image this big (pixels, both sides) is a scan.
$script:PdfPageImageMin = 800
# A smaller picture on such a page with a side of this many pixels or more, or of an unknown size, can be a small scan
# (an ID card): it is not read with OCR, and the file gets a note (the Administrator, 2026-11-26: "Merge, then a note").
$script:PdfMaybeScanMin = 300
$script:PdfMaybeScans = 0
$script:PdfMaybeScanNote = 'Scanned PDF: an image on a page with text was not read with OCR (it can be a scanned page): check by hand'
# Not-scanned reasons, each its own group on the Not scanned tab (the Administrator, 2026-11-21: tell labels from passwords, and
# a time limit from damage: many PDFs called damaged open fine). The limits are $PdfSecondsLimit and $PdfRegexTimeout of PdfExtract.psm1.
$script:PdfSlowReason = 'This PDF took longer than the time limit to read (2 minutes; the file may be fine): check by hand'
$script:PdfSearchSlowReason = 'One search in this PDF took longer than its time limit (2 seconds; the file may be fine): check by hand'
$script:PdfPartReason = 'Part of this PDF could not be read (an unusual string; the file may still open fine): check by hand'
$script:PdfEncryptedReason = 'Encrypted PDF (a password, or limits set by its author; it may open without one): check by hand'
# Microsoft Purview protection: an encrypted PDF whose security handler is Microsoft's, or a wrapper PDF that holds the
# protected document as an encrypted payload (ISO 32000-2; its own pages only say the file is protected).
$script:PdfLabelReason = 'Sensitivity label with encryption (Microsoft Purview): open it and check by hand'
$script:PdfLabelPattern = '/Filter\s*/MicrosoftIRMServices\b|/EncryptedPayload\b'

function Get-PdfText {
    <#
    .SYNOPSIS
      The text of a PDF: the text of every page and the values of its form fields. A PDF that is encrypted, has
      no text (a scanned image) or is not a PDF gets a reason instead. Fonts without a ToUnicode map and other
      unusual encodings can give partial text: what can be read is checked.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [hashtable] Text, Reason (empty when the text was read), Note (a part not read, for a "check by hand" row; or empty).
    #>
    param([byte[]]$Bytes)
    $text = $script:PdfLatin1.GetString($Bytes)
    if ($text.IndexOf('%PDF') -lt 0 -or $text.IndexOf('%PDF') -gt 1024) {
        return @{ Text = ''; Reason = 'Not a valid PDF (damaged, or another type with a .pdf name): check by hand' }
    }
    if ($text -match $script:PdfLabelPattern) { return @{ Text = ''; Reason = $script:PdfLabelReason } }
    if ($text -match '/Encrypt\s*(?:\d+\s+\d+\s+R|<<)') { return @{ Text = ''; Reason = $script:PdfEncryptedReason } }
    Reset-PdfOldObjects   # this PDF's own older versions only
    $script:PdfScanImages = @(); $script:PdfPageImages = @(); $script:PdfMaybeScans = 0   # and its own scanned pages
    Start-PdfBudget
    try { $all = Get-PdfAllText $text }
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSearchSlowReason } }
            if ($e -is [System.TimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
            if ($e -is [System.FormatException] -and $e.Message -eq (Get-PdfStringError)) { return @{ Text = ''; Reason = $script:PdfPartReason } }
        }
        throw   # any other error is the caller's (Get-TextFromBytes reports the file)
    }
    finally { $notes = @(Stop-PdfBudget) }
    $note = (@($notes) -join '; ')
    # Runs of spaces and tabs become one space: the detectors never need them, and a page of 15 million spaces would
    # otherwise cost the later checks a gigabyte (fresh review, 2026-11-21).
    $all = [regex]::Replace($all, '[ \t]{2,}', ' ')
    if (-not $all.Trim()) {
        if ($script:PdfScanImages.Count -gt 0) {
            $scan = Get-ScanPagesText $script:PdfScanImages $note
            if (-not $scan.Reason -and -not $scan.Text.Trim()) { $scan.Reason = $script:PdfNoTextReason }
            return $scan
        }
        $reason = $script:PdfNoTextReason; if ($note) { $reason = $note }
        return @{ Text = ''; Reason = $reason; Note = '' }
    }
    if ($script:PdfMaybeScans -gt 0) { $note = (@($note, $script:PdfMaybeScanNote) | Where-Object { $_ }) -join '; ' }
    if ($script:PdfPageImages.Count -gt 0) {   # text pages and scanned pages: OCR the pages with no text too
        $scan = Get-ScanPagesText $script:PdfPageImages ''
        if ($scan.Reason) {
            $why = $scan.Reason; $hand = ': check by hand'   # said once, at the end
            if ($why.EndsWith($hand)) { $why = $why.Substring(0, $why.Length - $hand.Length) }
            $notes = @($note, ($why + ' (its pages with no text)' + $hand))
        }
        else { $all = $all + "`n" + $scan.Text; $notes = @($note, $scan.Note) }
        $note = (@($notes | Where-Object { $_ }) -join '; ')
    }
    return @{ Text = $all; Reason = ''; Note = $note }
}

function Get-PdfPageResources {
    <#
    .SYNOPSIS
      The resources dictionary of a page: its own (inline or by reference), else the nearest one up its /Parent
      chain (a Pages node), so a page that inherits its fonts gets the right ones.
    .PARAMETER Page
      The page object.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] the dictionary text to look up fonts in.
    #>
    param($Page, $Objects)
    $node = $Page
    for ($level = 0; $level -lt 32 -and $node; $level++) {   # 32: deeper than any real page tree; stops a loop
        $ref = [regex]::Match($node.Dict, '/Resources\s+(\d+)\s+\d+\s+R')
        if ($ref.Success -and $Objects.Contains($ref.Groups[1].Value)) { return $Objects[$ref.Groups[1].Value].Dict }
        if ($node.Dict -match '/Resources\s*<<') { return $node.Dict }
        $parent = [regex]::Match($node.Dict, '/Parent\s+(\d+)\s+\d+\s+R')
        $node = $null
        if ($parent.Success -and $Objects.Contains($parent.Groups[1].Value)) { $node = $Objects[$parent.Groups[1].Value] }
    }
    return $Page.Dict
}

function Get-PdfLooseText {
    <#
    .SYNOPSIS
      The text of content streams no page or form reads: older page versions kept by an incremental update ("SSN"
      edited out of a page is still in the file) and streams of pages since removed. A stream counts as content when
      it has a text block (BT, Tj or TJ, ET) and is not XML. One that cannot be read gives a note, not a failure.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Fonts
      The fonts of the whole file.
    .PARAMETER Read
      The streams already read.
    .OUTPUTS
      [string] text, one item per stream.
    #>
    param($Objects, [hashtable]$Fonts, $Read)
    $loose = @($Objects.Values) + (Get-PdfOldObjects)
    foreach ($obj in $loose) {
        if (-not $obj.Stream -or $Read.Contains($obj) -or $obj.Dict -match '/Type\s*/(?:ObjStm|XRef)\b') { continue }
        $start = $obj.Stream.TrimStart()
        if ($start.StartsWith('<') -or $obj.Stream -notmatch '\bBT\b' -or $obj.Stream -notmatch '\bT[Jj]\b|\)\s*[''"]') { continue }
        try { Convert-PdfContentToText $obj.Stream $Fonts }
        catch {
            if ($_.Exception -is [System.FormatException] -and $_.Exception.Message -eq (Get-PdfStringError)) {
                Add-PdfNote 'Part of this PDF could not be read: check by hand'
            }
            else { throw }
        }
    }
}

function Get-PdfXObjects {
    <#
    .SYNOPSIS
      The objects that a resources dictionary names under /XObject (inline, or a dictionary by reference).
    .PARAMETER Resources
      The resources dictionary text.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [object[]]
    #>
    param([string]$Resources, $Objects)
    $list = ''
    $inline = [regex]::Match($Resources, '/XObject\s*<<([^>]*)>>')
    if ($inline.Success) { $list = $inline.Groups[1].Value }
    else {
        $ref = [regex]::Match($Resources, '/XObject\s+(\d+)\s+\d+\s+R')
        if ($ref.Success -and $Objects.Contains($ref.Groups[1].Value)) { $list = $Objects[$ref.Groups[1].Value].Dict }
    }
    foreach ($name in [regex]::Matches($list, '(\d+)\s+\d+\s+R')) {
        if ($Objects.Contains($name.Groups[1].Value)) { $Objects[$name.Groups[1].Value] }
    }
}

function Get-PdfTextKind {
    <#
    .SYNOPSIS
      What text a content stream gives: 'none'; 'hidden' when it draws invisible text (render mode 3: the text layer of
      a searchable scan); 'visible' for other text (a banner, a header, a stamp, a typed page).
    .PARAMETER Raw
      The content stream as written.
    .PARAMETER Text
      Its text (Convert-PdfContentToText).
    .OUTPUTS
      [string]
    #>
    param([string]$Raw, [string]$Text)
    if (-not $Text.Trim()) { return 'none' }
    if ($Raw -match '(^|\s)3\s+Tr\b') { return 'hidden' }
    return 'visible'
}

function Get-PdfFormKinds {
    <#
    .SYNOPSIS
      The text kinds (Get-PdfTextKind) of the form XObjects that a page's resources name, one level; each form is
      worked out once per PDF (a cache: a PDF of 500 pages with shared resources stays fast; review of step 3).
    .PARAMETER Resources
      The page's resources dictionary text.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER AllFonts
      The fonts of the whole file, for a form that names none.
    .PARAMETER Cache
      Form object -> kind (changed in place).
    .OUTPUTS
      [string[]]
    #>
    param([string]$Resources, $Objects, $AllFonts, $Cache)
    foreach ($form in @(Get-PdfXObjects $Resources $Objects)) {
        if (-not $form.Stream -or $form.Dict -notmatch '/Subtype\s*/Form\b') { continue }
        if (-not $Cache.ContainsKey($form)) {
            $fonts = Get-PdfFontTable (Get-PdfPageResources $form $Objects) $Objects
            if ($fonts.Count -eq 0) { $fonts = $AllFonts }
            $Cache[$form] = Get-PdfTextKind $form.Stream (Convert-PdfContentToText $form.Stream $fonts)
        }
        $Cache[$form]
    }
}

function Test-PdfPageImage {
    <#
    .SYNOPSIS
      True for an image of page size: /Width and /Height both at least $PdfPageImageMin pixels (a scan; a logo is
      smaller). Used only on a page with visible text, to tell a scan with a banner from a typed page with a logo.
    .PARAMETER Dict
      The image object's dictionary text.
    .OUTPUTS
      [bool]
    #>
    param([string]$Dict)
    $w = [regex]::Match($Dict, '/Width\s+(\d{1,6})\b'); $h = [regex]::Match($Dict, '/Height\s+(\d{1,6})\b')
    return ($w.Success -and $h.Success -and [int]$w.Groups[1].Value -ge $script:PdfPageImageMin -and [int]$h.Groups[1].Value -ge $script:PdfPageImageMin)
}

function Get-PdfResourceImages {
    <#
    .SYNOPSIS
      The JPEG image objects that a resources dictionary names under /XObject (inline or by reference), and those of a
      form XObject it names (one level: a scanner can wrap the page image in a form).
    .PARAMETER Resources
      The resources dictionary text (Get-PdfPageResources).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Depth
      0 for a page; 1 inside a form XObject (no deeper).
    .OUTPUTS
      [object[]] the image objects (each with Image).
    #>
    param([string]$Resources, $Objects, [int]$Depth = 0)
    foreach ($object in @(Get-PdfXObjects $Resources $Objects)) {
        if ($object.ContainsKey('Image') -and $object.Image) { $object; continue }
        if ($Depth -eq 0 -and $object.Dict -match '/Subtype\s*/Form\b') { Get-PdfResourceImages (Get-PdfPageResources $object $Objects) $Objects 1 }
    }
}

function Add-PdfScanImages {
    <#
    .SYNOPSIS
      Adds a page's scanned images to the OCR list: none for a page with a text layer (hidden text); for a page with
      visible text, only its page-size images (Test-PdfPageImage); for a page with no text, all its images.
    .PARAMETER Kinds
      The text kinds of the page and of the forms it names (Get-PdfTextKind, Get-PdfFormKinds).
    .PARAMETER Resources
      The page's resources dictionary text.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER List
      The image objects to read with OCR (changed in place; each image once).
    .PARAMETER Cache
      Resources text -> its images (changed in place).
    .PARAMETER Seen
      The images already in List (a HashSet; changed in place).
    #>
    param([string[]]$Kinds, [string]$Resources, $Objects, $List, $Cache, $Seen)
    if ($Kinds -contains 'hidden') { return }
    $anyText = $Kinds -contains 'visible'
    # Shared resources (one dictionary for 300 pages) are looked into once (review 4 of step 3: 415 s without this).
    if (-not $Cache.ContainsKey($Resources)) {   # each image once: forms that name the same shared resources repeat them
        $unique = New-Object 'System.Collections.Generic.HashSet[object]'
        $Cache[$Resources] = @(foreach ($image in @(Get-PdfResourceImages $Resources $Objects)) { if ($unique.Add($image)) { $image } })
    }
    foreach ($image in $Cache[$Resources]) {
        if ($anyText -and -not (Test-PdfPageImage $image.Dict)) {
            if (Test-PdfMaybeScan $image.Dict) { $script:PdfMaybeScans++ }
            continue
        }
        if ($Seen.Add($image)) { $List.Add($image) }   # a set: 300 pages x 300 shared images stay fast (review 4 of step 3)
    }
}

function Test-PdfMaybeScan {
    <#
    .SYNOPSIS
      True for a picture that can be a small scan: a side of $PdfMaybeScanMin pixels or more, or a size that is not
      written as a number (an indirect /Width). A logo or an icon is smaller.
    .PARAMETER Dict
      The image object's dictionary text.
    .OUTPUTS
      [bool]
    #>
    param([string]$Dict)
    $w = [regex]::Match($Dict, '/Width\s+(\d{1,6})\b(?!\s+\d+\s+R)'); $h = [regex]::Match($Dict, '/Height\s+(\d{1,6})\b(?!\s+\d+\s+R)')
    if (-not $w.Success -or -not $h.Success) { return $true }
    return ([int]$w.Groups[1].Value -ge $script:PdfMaybeScanMin -or [int]$h.Groups[1].Value -ge $script:PdfMaybeScanMin)
}

function Get-PdfAllText {
    <#
    .SYNOPSIS
      Page text (each page with its own fonts; the fonts of the whole file when a page names none, as when its
      resources are inherited), form XObjects, dictionaries (Get-PdfDictText), XFA data and content streams no page
      reads (Get-PdfLooseText), joined.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [string]
    #>
    param([string]$Text)
    $objects = Get-PdfObjects $Text
    $script:PdfScanImages = @($objects.Values | Where-Object { $_.ContainsKey('Image') -and $_.Image } | ForEach-Object { $_.Image })
    $pageImages = New-Object System.Collections.Generic.List[object]   # the scanned images of pages with no text layer
    $formKinds = New-Object 'System.Collections.Generic.Dictionary[object,string]'   # form object -> text kind (Get-PdfFormKinds)
    $resourceImages = New-Object 'System.Collections.Generic.Dictionary[string,object]'   # resources text -> its images; case counts
    $pageImageSet = New-Object 'System.Collections.Generic.HashSet[object]'   # the images already in $pageImages
    $resourceKinds = New-Object 'System.Collections.Generic.Dictionary[string,object]'   # resources text -> form kinds; case counts
    Assert-PdfTime
    $allFonts = Get-PdfFonts $objects
    Assert-PdfTime
    $parts = New-Object System.Collections.Generic.List[string]
    $read = New-Object System.Collections.Generic.HashSet[object]   # the content streams read as pages or forms
    foreach ($page in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/Page(?![s\w])' })) {
        $resources = Get-PdfPageResources $page $objects
        $first = $parts.Count; $pageRaw = ''
        $fonts = Get-PdfFontTable $resources $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $contents = [regex]::Match($page.Dict, '/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)').Groups[1].Value
        foreach ($content in [regex]::Matches($contents, '(\d+)\s+\d+\s+R')) {
            $stream = $objects[$content.Groups[1].Value]
            if (-not $stream) { continue }
            if ($stream.Stream) { [void]$read.Add($stream); $pageRaw += $stream.Stream + "`n"; $parts.Add((Convert-PdfContentToText $stream.Stream $fonts)); continue }
            # /Contents 5 0 R where object 5 is an array of streams ([6 0 R 7 0 R]): read each of them.
            foreach ($inner in [regex]::Matches($stream.Dict, '(\d+)\s+\d+\s+R')) {
                $part = $objects[$inner.Groups[1].Value]
                if ($part -and $part.Stream) { [void]$read.Add($part); $pageRaw += $part.Stream + "`n"; $parts.Add((Convert-PdfContentToText $part.Stream $fonts)) }
            }
        }
        $pageText = ''; for ($i = $first; $i -lt $parts.Count; $i++) { $pageText += $parts[$i] }
        if (-not $resourceKinds.ContainsKey($resources)) { $resourceKinds[$resources] = @(Get-PdfFormKinds $resources $objects $allFonts $formKinds) }
        $kinds = @(Get-PdfTextKind $pageRaw $pageText) + $resourceKinds[$resources]
        Add-PdfScanImages $kinds $resources $objects $pageImages $resourceImages $pageImageSet
    }
    $script:PdfPageImages = @(foreach ($image in $pageImages) { $image.Image })   # a loop: '@()' of a New-Object list throws (probe 6)
    # Form XObjects (/Subtype /Form, drawn with "Do"): reusable page parts that can hold text of their own. Read each
    # once; a form field's appearance can repeat its value, so a hit may show twice (better than a miss).
    foreach ($form in @($objects.Values | Where-Object { $_.Stream -and $_.Dict -match '/Subtype\s*/Form\b' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $form $objects) $objects   # also /Resources 10 0 R
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        [void]$read.Add($form)
        $parts.Add((Convert-PdfContentToText $form.Stream $fonts))
    }
    foreach ($line in @(Get-PdfDictText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfXfaText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfLooseText $objects $allFonts $read)) { $parts.Add($line) }
    return ($parts -join "`n")
}

Export-ModuleMember -Function Get-PdfRegex, Assert-PdfTime, Add-PdfNote, Start-PdfBudget, Stop-PdfBudget, Read-PdfStream, ConvertFrom-Utf16Hex, Get-PdfCMap, Get-PdfObjects, ConvertFrom-PdfString,
    ConvertFrom-PdfTextString, Assert-PdfGap, Get-PdfDictText, Get-PdfXfaText, Get-PdfStringPattern, Get-PdfStringError, Stop-PdfHostile, Reset-PdfOldObjects, Get-PdfOldObjects, Get-PdfFonts,
    Get-PdfFontTable, ConvertFrom-PdfGlyphs, Get-PdfShowWidth, Get-PdfShownRaw, Get-PdfShownText, Get-PdfAdvance, Compress-PdfShown, Complete-PdfTextBlock, Convert-PdfContentToText, Get-PdfText
