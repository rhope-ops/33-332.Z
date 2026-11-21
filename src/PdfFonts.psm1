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

# Fonts that draw pictures, not text (by name, with or without a subset prefix such as ABCDEF+): Wingdings 1-3,
# Webdings, (Zapf) Dingbats, Marlett and Symbol (Greek and maths signs). Their text is read as spaces: a row of symbols
# read as letters and digits gave Low hits that were plainly wrong (the PAM, 2026-11-21).
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
# Not afii plus a number: the Adobe names of Cyrillic and other letters.
$script:MadeUpGlyphName = '^(?!afii)[A-Za-z]{1,5}\d+(?:\..*)?$'

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

# Per font: shown strings decoded and measured before (text and width in ems), up to this many each (speed).
$script:MaxFontCache = 20000
# Entries of what a line showed before they are summed (Compress-PdfShown).
$script:MaxShownItems = 1024
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
      signs (the PAM, 2026-11-21: "Wingdings"-like text from a PDF with no Wingdings font).
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
    # A picture font (Wingdings and the like): one space per glyph, so its symbols never read as PII (the PAM, 2026-11-21).
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

Export-ModuleMember -Function Get-PdfFonts, Get-PdfFontTable, ConvertFrom-PdfGlyphs, Get-PdfShowWidth, Get-PdfShownRaw, Get-PdfShownText, Get-PdfAdvance, Compress-PdfShown
