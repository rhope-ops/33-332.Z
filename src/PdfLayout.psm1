<#
PdfLayout.psm1 - spaces and line breaks of PDF page text: where a gap between two pieces of text is a word space.

  Purpose:     A PDF places text in pieces (Chrome glyph by glyph, Word by kerning pairs, LibreOffice with letter
               spacing); this module joins a text block's pieces back into words and lines, so the detectors see
               "SSN: <number>" and not "S S N :" or "SSN:<number>" (fresh reviews r9 to r16, 2026-11-21).
  Inputs:      a text block's pieces (strings, and gaps in ems) from PdfText.psm1.
  Outputs:     the block's text, added to the page text.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (Assert-PdfTime, Add-PdfNote), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

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

Export-ModuleMember -Function Complete-PdfTextBlock
