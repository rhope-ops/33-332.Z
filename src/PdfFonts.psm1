<#
PdfFonts.psm1 - the fonts of a PDF: their character maps and their real glyph widths, with .NET only.

  Purpose:     A shown string becomes text through its font's ToUnicode map, and its width comes from the font's own
               widths (/Widths of a simple font, /W of a composite font, built-in widths for the standard Helvetica,
               Times and Courier families), so PdfText.psm1 can tell a word gap from a glyph placed one by one
               (fresh reviews, 2026-11-21: estimated widths split and joined words wrongly). A guess is the last resort.
  Inputs:      the PDF objects (PdfObjects.psm1 Get-PdfObjects) and shown strings.
  Outputs:     font records (CMap, Bytes, Widths, DefaultWidth, Standard), decoded text and widths.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (Get-PdfCMap, ConvertFrom-Utf16Hex, Assert-PdfTime, Add-PdfNote), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

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

# Per font: shown strings decoded and measured before (text and width in ems), up to this many each (speed).
$script:MaxFontCache = 20000
$script:MaxCacheKey = 256   # longer strings are rarely shown twice: not kept (memory)
# Shown hex strings decoded before (speed): short tokens only, emptied when full.
$script:ShownHex = New-Object 'System.Collections.Generic.Dictionary[string,string]'
$script:MaxShownHex = 50000

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
    $unicode = [regex]::Match($Dict, '/ToUnicode\s+(\d+)\s+\d+\s+R')
    if ($unicode.Success -and $Objects.Contains($unicode.Groups[1].Value) -and $Objects[$unicode.Groups[1].Value].Stream) {
        $cmap = $Objects[$unicode.Groups[1].Value]   # read once, also when many fonts share it
        if (-not $cmap.Contains('CMap')) { $cmap['CMap'] = Get-PdfCMap $cmap.Stream }
        $record.CMap = $cmap['CMap']
        $record.Bytes = [Math]::Max(1, [int]$record.CMap.Bytes)
    }
    elseif ($composite) { $record.CMap = @{ Bytes = 2; Map = @{}; Ranges = @(); Unreadable = $true } }
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

Export-ModuleMember -Function Get-PdfFonts, Get-PdfFontTable, ConvertFrom-PdfGlyphs, Get-PdfShowWidth, Get-PdfShownRaw, Get-PdfShownText
