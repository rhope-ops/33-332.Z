<#
PdfCore.psm1 - the base of the PDF reader: its time and size limits, stream filters and ToUnicode font maps.

  Purpose:     So one damaged or hostile PDF cannot stall the scan or fill the memory (fresh reviews, 2026-11-20/21):
               a time limit per PDF and per regex search, caps on unpacked size, and "check by hand" notes for any
               part not read. Also the stream filters (Flate, ASCII85, ASCIIHex) and the ToUnicode maps of fonts.
               PdfText.psm1 builds the page text on top of it.
  Inputs:      PDF stream bytes and dictionaries, as Latin-1 text.
  Outputs:     decoded streams, font maps, the limit checks and the notes of one PDF (Start-/Stop-PdfBudget).
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

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
    # only 0x78: a header read as deflate data made good PDFs look damaged (the PAM's gov run, 2026-11-21).
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

Export-ModuleMember -Function Get-PdfRegex, Assert-PdfTime, Add-PdfNote, Start-PdfBudget, Stop-PdfBudget, Read-PdfStream, ConvertFrom-Utf16Hex, Get-PdfCMap
