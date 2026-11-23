<#
PdfShown.psm1 - a PDF's shown strings: their text through the font's map, and their widths in ems.

  Purpose:     Decodes a shown string with its font record (ToUnicode map, glyph names, one- or two-byte codes),
               measures it with the font's widths, and sums the pieces a line showed, so PdfText.psm1 can tell a word
               gap from a glyph placed one by one. Caches per font (speed). Split from PdfFonts.psm1 (fleet v2 DEC 6,
               2026-11-23: 450 lines for hand-written code).
  Inputs:      font records (PdfFonts.psm1 Get-PdfFontTable) and shown strings.
  Outputs:     decoded text and widths.
  Depends on:  PdfCore.psm1 (Add-PdfNote, Assert-PdfTime, ConvertFrom-Utf16Hex), PdfObjects.psm1 (ConvertFrom-PdfString)
               and PdfFonts.psm1 (the font records it reads), imported first.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

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

Export-ModuleMember -Function ConvertFrom-PdfGlyphs, Get-PdfShowWidth, Get-PdfShownRaw, Get-PdfShownText, Get-PdfAdvance, Compress-PdfShown
