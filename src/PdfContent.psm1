<#
PdfContent.psm1 - a PDF content stream walked for its text: tokens, text operators, moves and shown strings.

  Purpose:     Reads one content stream's tokens (strings, arrays, names, numbers, operators), runs the text
               operators (Tf, Tc, Tw, Tz, TJ, Tj, ', ", Td, TD, T*, Tm, ET) on a text state, steps over inline image
               bytes, and returns the page text with its spaces and line breaks (ISO 32000-1, 9.4). Split from
               PdfText.psm1 (fleet v2 DEC 6, 2026-11-23: 450 lines for hand-written code).
  Inputs:      a decoded content stream and the page's fonts (PdfFonts.psm1 Get-PdfFontTable).
  Outputs:     the text of the stream (Convert-PdfContentToText).
  Depends on:  PdfCore.psm1 (Assert-PdfTime, Add-PdfNote), PdfObjects.psm1 (Get-PdfStringPattern, Assert-PdfGap),
               PdfShown.psm1 (shown text and widths) and PdfLayout.psm1 (Complete-PdfTextBlock), imported first.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

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

Export-ModuleMember -Function Convert-PdfContentToText
