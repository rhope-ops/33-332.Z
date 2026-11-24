<#
PiiMasking.psm1 - turns raw PII hits into masked report data. Nothing unmasked leaves this module.

  Purpose:     Find-PiiInText returns, for one text, each hit's type, confidence, masked value and a
               masked snippet of the text around it. Protect-PiiText masks PII inside a short text that
               the report shows as it is (file and folder names, error messages).
               The PAM, 2026-11-20: every value is masked on screen and in the report files.
  Inputs:      plain text; hits from PiiDetectors.psm1 (Get-PiiSpan).
  Outputs:     masked hit objects and masked strings.
  Depends on:  PiiDetectors.psm1 (Get-PiiSpan), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# Characters of text kept on each side of a hit in the report snippet. The author's choice,
# 2026-11-20: enough to see the label ("SSN:", "DOB") without copying a whole paragraph.
$script:SnippetRadiusChars = 30

# How many trailing characters a masked value keeps, so a reviewer can find it in the file. The author's
# choice, 2026-11-20, after common practice: receipts and DoD forms show at most the last 4 digits of an
# SSN or card number. Passports keep 3, as their numbers are shorter.
$script:KeepLastDigits = 4
$script:KeepLastPassportChars = 3

# An e-mail local part: everything before "@domain.tld" that is not a space, a delimiter or a path slash.
# Broad on purpose (apostrophes, accents, non-ASCII domains), so no part of a name in an address stays
# visible; slashes are excluded so the folders before an address in a path stay readable.
# It starts only at the start of a word (after a space, a delimiter, a slash or an '@'), so a long run of text is read
# once, not once per character.
$script:EmailLocalPartPattern = '(?<=^|[\s@<>()\[\],;:"/\\])[^\s@<>()\[\],;:"/\\]+(?=@[^\s@<>()\[\],;:"/\\.]+\.)'

# A CAC name in an e-signature ("DOE.JANE.Q."): in the Context column each part keeps only its first letter (approved
# 2026-11-21: "mask the CAC name in Context"). $CacCandidate finds only two or more dotted parts in a row (and the part
# after them), each part read once (atomic groups) and never from inside a run of name characters, so long text stays
# fast and a plain word never reaches Protect-CacName (fifth review of PR 81). A name starts at a part of two letters
# or more, so "e.g." and "i.e." stay readable; a dotted host or file name that looks the same is masked too.
$script:CacCandidate = [regex]"(?<![\p{L}\p{M}\p{Cs}_'-])(?>[\p{L}\p{M}\p{Cs}_'-]+\.){2,}(?>[\p{L}\p{M}\p{Cs}_'-]*)"
$script:NameLetter = [regex]'[\p{L}\p{M}\p{Cs}]'   # a letter, a mark, or half of a letter outside the BMP
# The first part of two letters or more (where a name starts), and the first letter of each part (it stays).
$script:NameStart = [regex]'(?:^|(?<=\.))(?=[^.\p{L}\p{Cs}]*[\p{L}\p{Cs}][^.\p{L}\p{Cs}]*[\p{L}\p{Cs}])'
$script:PartInitial = [regex]'(?:^|(?<=\.))[^.\p{L}\p{Cs}]*(\p{L}|\p{Cs}\p{Cs})'

# =====================================================================================
# Masking one value
# =====================================================================================

function Get-LastCharacters {
    # Returns the last $Count characters of a string (all of it when it is shorter).
    param([string]$Value, [int]$Count)
    return $Value.Substring([Math]::Max(0, $Value.Length - $Count))
}

function Get-MaskedValue {
    <#
    .SYNOPSIS
      Masks one found value so the report never holds it in full.
    .DESCRIPTION
      SSN    -> ***-**-1234        DoD ID   -> ******7890        Card    -> ************1111
      Phone  -> ***-***-1234       Passport -> *****567          E-mail  -> ***@domain
      DOB    -> [date masked]      Address  -> [address masked]  Name    -> initials (P*** S***)
    .PARAMETER Detector
      The PII type (as named in PiiDetectors.psm1).
    .PARAMETER Value
      The raw value.
    .OUTPUTS
      [string] the masked value.
    #>
    param([string]$Detector, [string]$Value)
    $digits = $Value -replace '\D', ''
    $lastDigits = Get-LastCharacters $digits $script:KeepLastDigits
    switch ($Detector) {
        'SSN' { return '***-**-' + $lastDigits }
        'DoD ID' { return '******' + $lastDigits }
        'Payment card' { return ('*' * ($digits.Length - $lastDigits.Length)) + $lastDigits }
        'Phone' { return '***-***-' + $lastDigits }
        'Personal e-mail' { return '***@' + (($Value -split '@')[-1] -replace '\d', '#') }   # domain kept, its digits hidden
        'Passport' {
            $tail = Get-LastCharacters $Value $script:KeepLastPassportChars
            return ('*' * ($Value.Length - $tail.Length)) + $tail
        }
        'Date of birth' { return '[date masked]' }
        'Home address' { return '[address masked]' }
        'Name' { return ((($Value -split '[\s,]+') | Where-Object { $_ } | ForEach-Object { $_.Substring(0, 1) + '***' }) -join ' ') }
        default { return '[masked]' }
    }
}

# =====================================================================================
# Masking a whole text, then cutting snippets from it
# =====================================================================================

function Protect-CacName {
    <#
    .SYNOPSIS
      One candidate ($CacCandidate) with its name parts masked: from the first part of two letters or more on, every
      letter and mark becomes '*' except the first letter of each part. The regexes do the per-letter work in .NET (one
      pass each); this loop only puts back one letter per part (fifth review of PR 81: speed). The length stays.
    .PARAMETER Text
      The candidate, for example "DOE.JANE.Q." or "DOE.J.Q.".
    .OUTPUTS
      [string]
    #>
    param([string]$Text)
    $start = $script:NameStart.Match($Text)
    if (-not $start.Success) { return $Text }   # no part of two letters: "e.g.", "U.S."
    $tail = $Text.Substring($start.Index)
    $masked = $script:NameLetter.Replace($tail, '*').ToCharArray()
    foreach ($m in $script:PartInitial.Matches($tail)) {
        $g = $m.Groups[1]
        for ($k = $g.Index; $k -lt $g.Index + $g.Length; $k++) { $masked[$k] = $tail[$k] }
    }
    return $Text.Substring(0, $start.Index) + (New-Object string (, $masked))
}

function Get-MaskedText {
    <#
    .SYNOPSIS
      Masks a whole text and keeps its length: every digit becomes '#', every e-mail local part
      becomes '*', a CAC name keeps only the first letter of each part (Protect-CacName), and every hit (names
      and addresses included) becomes '*'. Because the length
      stays the same, hit positions from the raw text still point to the right place.
    .PARAMETER Text
      The raw text.
    .PARAMETER Hits
      The hits found in it (Get-PiiSpan).
    .OUTPUTS
      [string] the masked text, same length as the input.
    #>
    param([string]$Text, [object[]]$Hits)
    $sameLengthStars = [System.Text.RegularExpressions.MatchEvaluator] { param($m) '*' * $m.Length }
    $masked = [regex]::Replace($Text, $script:EmailLocalPartPattern, $sameLengthStars) -replace '\d', '#'
    # One pass over the name: the first letter of each part stays, every other letter becomes '*' (linear: review of PR 81).
    $masked = $script:CacCandidate.Replace($masked, [System.Text.RegularExpressions.MatchEvaluator] { param($m) Protect-CacName $m.Value })
    $chars = $masked.ToCharArray()
    foreach ($hit in $Hits) {
        for ($i = $hit.Index; $i -lt $hit.Index + $hit.Length; $i++) {
            if (-not [char]::IsWhiteSpace($chars[$i])) { $chars[$i] = '*' }
        }
    }
    return New-Object string (, $chars)
}

function Get-MaskedSnippet {
    <#
    .SYNOPSIS
      A short piece of the MASKED text around one hit. It is cut from text that is already masked,
      so a cut through an address or number cannot expose part of it.
    .PARAMETER MaskedText
      Output of Get-MaskedText.
    .PARAMETER Index
      Start position of the hit.
    .PARAMETER Length
      Length of the hit.
    .OUTPUTS
      [string] the snippet, white space folded to single spaces.
    #>
    param([string]$MaskedText, [int]$Index, [int]$Length)
    $start = [Math]::Max(0, $Index - $script:SnippetRadiusChars)
    $end = [Math]::Min($MaskedText.Length, $Index + $Length + $script:SnippetRadiusChars)
    return ($MaskedText.Substring($start, $end - $start) -replace '\s+', ' ').Trim()
}

# =====================================================================================
# Public functions
# =====================================================================================

function Find-PiiInText {
    <#
    .SYNOPSIS
      Finds PII in one text and returns masked results only.
    .PARAMETER Text
      Plain text from a file or a post.
    .OUTPUTS
      [pscustomobject] Detector, Confidence, MaskedValue, Context (no raw values).
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return }
    $hits = @(Get-PiiSpan $Text)
    if ($hits.Count -eq 0) { return }
    $maskedText = Get-MaskedText $Text $hits
    foreach ($hit in $hits) {
        [pscustomobject]@{
            Detector    = $hit.Detector
            Confidence  = $hit.Confidence
            MaskedValue = Get-MaskedValue $hit.Detector $hit.Value
            Context     = Get-MaskedSnippet $maskedText $hit.Index $hit.Length
        }
    }
}

function Protect-PiiText {
    <#
    .SYNOPSIS
      Masks an error message that the screen or the report shows as it is. Strict on purpose, because a message
      can quote any text: Protect-PiiName, then EVERY digit becomes '#'.
      Example: "Roster_123-45-6789.docx" -> "Roster_***-**-####.docx"; "FY27 Plan.docx" -> "FY## Plan.docx".
    .PARAMETER Text
      The message.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    return ((Protect-PiiName $Text) -replace '\d', '#')
}

function Protect-PiiName {
    <#
    .SYNOPSIS
      Masks a site, folder or file name the person must find again: each found value becomes its masked value and
      every e-mail name becomes '***'; other digits stay (the PAM, 2026-11-20: "FY27 Plan.docx" stays as it is, so the
      file can be found). Example: "Roster_123-45-6789.docx" -> "Roster_***-**-6789.docx".
    .PARAMETER Text
      The name or path.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $result = $Text
    # Replace found values from the end, so the positions of earlier ones stay valid.
    foreach ($hit in (@(Get-PiiSpan $Text) | Sort-Object Index -Descending)) {
        $masked = Get-MaskedValue $hit.Detector $hit.Value
        $result = $result.Substring(0, $hit.Index) + $masked + $result.Substring($hit.Index + $hit.Length)
    }
    return [regex]::Replace($result, $script:EmailLocalPartPattern, '***')
}

Export-ModuleMember -Function Find-PiiInText, Protect-PiiText, Protect-PiiName, Get-MaskedValue
