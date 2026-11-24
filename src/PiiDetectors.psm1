<#
PiiDetectors.psm1 - finds text that looks like PII and returns where it is.

  Purpose:     Get-PiiSpan runs every detector (PiiPatterns.psm1, and the labelled-name detector here), removes
               overlaps, raises a name near an identifier, drops Low hits in unreadable text, and returns each hit
               with its position and raw value.
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     hit objects WITH the raw value. Only PiiMasking.psm1 reads them; it masks every value
               before anything is shown or saved (the PAM, 2026-11-20: all values masked).
  Depends on:  PiiPatterns.psm1 (the one-type detectors, New-PiiHit, Test-KeywordBefore, Test-DodIdLabelBefore), imported first.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings (named values; each says where it comes from)
# =====================================================================================

# A labelled name within this many characters of an identifier is raised to High. DoD privacy
# guidance treats a name together with an identifier (SSN, DoD ID, date of birth) as sensitive PII.
$script:NameToIdentifierChars = 80

# Unreadable text around a Low hit (Test-GarbledNear): the characters looked at on each side of the hit, the fewest
# visible ones needed to judge, and the share of unreadable ones that drops the hit. 30% and 20 characters: the
# author's choice, 2026-11-21 (normal prose has almost none; a run of symbol-font characters is nearly all).
$script:GarbledWindowChars = 40
$script:GarbledMinChars = 20
$script:GarbledShare = 0.3
$script:VisibleChar = [regex]'\S'
# Unreadable: private-use characters, the replacement character, control characters other than NUL, tab, CR and LF,
# and the Latin-1 signs a symbol font read as text gives (not the degree sign or the middle dot). Symbols that normal
# text uses (|, the cell break of a sheet; box drawing; maths; currency) are not counted (review of PR 31).
$script:GarbledChar = [regex]'[\p{Co}\uFFFD\x01-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\u00A4-\u00A6\u00A8-\u00AF\u00B1-\u00B6\u00B8-\u00BF]'


# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last.
$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone',
    'Personal e-mail', 'Home address', 'Name')

# Types that make a nearby name sensitive (see $NameToIdentifierChars).
$script:IdentifierTypes = @('SSN', 'DoD ID', 'Date of birth', 'Passport', 'Payment card')

# Confidence labels as numbers, for sorting (High first).
$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# E-signature DoD IDs that PiiPatterns.psm1 Find-DodIdHits misses (docs/reviews/2026-11-23-ion-esignature-layouts.md). $SignedName: a CAC name (line
# breaks in or after parts, up to 3 words a later part, 40 letters a part: linear time), then 10 digits with one space, tab,
# no-break space or line break between any two (or dashes right after the name), no digit after them, no phone shape.
$script:SignedPart = "[A-Z][A-Z'-]{0,39}(?:\r?\n[A-Z'-]{1,40})?"
$script:SignedId = '(?<id>\d(?:(?:[ \t\u00A0]|\r?\n)?\d){9})(?![\w-])(?![ \t\u00A0]\d)'
# The same on one line, space gaps, no small word after (the ID alone after a signature word), or with dashes (right after a CAC name).
$script:SignedIdLine = '(?<id>\d(?:[ \u00A0]?\d){9})(?![\w-])(?![ \t\u00A0]\d)(?![ \t\u00A0]+[a-z])'
$script:SignedIdDash = '(?<id>\d(?:[ \t\u00A0-]?\d){9})(?![\w-])(?![ \t\u00A0]\d)'
$script:NotPhone = '(?!\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{4}(?![\w-]))'
# Each pattern starts with a letter or digit, then looks back, so the engine skips ahead fast on long text (review of PR 60).
$script:SignedNameText = ("(?<name>[A-Z](?<=(?:^|[^\w.])[A-Z])(?:[A-Z'-]{1,39}(?:\r?\n[A-Z'-]{1,40})?|\r?\n[A-Z'-]{1,40})\." +
    '(?<!(?:^|[^\w.])(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)' +
    "(?:\r?\n)?(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)" +
    "$script:SignedPart(?: $script:SignedPart){0,2}\.(?:(?:\r?\n)?$script:SignedPart(?: $script:SignedPart){0,2}\.){0,2})" +
    "(?:$script:SignedIdDash|$script:SignedId|(?:[ \t\u00A0]+(?:\r?\n[ \t\u00A0]*)?|\r?\n[ \t\u00A0]*)$script:NotPhone$script:SignedId)")
$script:SignedName = [regex]$script:SignedNameText
$script:SignedNameAnyCase = [regex]('(?i)' + $script:SignedNameText)
$script:SignedNameAnyCaseAt = [regex]('(?i)\G(?:' + $script:SignedNameText + ')')   # \G: only at the given place
$script:SignedWords = '(?i:(?:sign(?:ed|ature)\b|CN\s*=|DN\s*:)(?<!\w(?:sign(?:ed|ature)|CN\s*=|DN\s*:)))'   # a word start, said last (speed)
# The 10 digits alone on one line after a signature word ("Signed_1234567890.pdf"): not a date, a phone or in a dotted name.
$script:SignedBareText = '(?<![\w.-])(?!(?:19|20)\d\d(?:[ \t\u00A0./-]|\r?\n|(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])))' + $script:NotPhone + $script:SignedIdLine
$script:SignedBareId = [regex]$script:SignedBareText
$script:SignedBareIdAt = [regex]('\G(?:' + $script:SignedBareText + ')')
# The 10 digits in a row just before a CAC name in capitals (the ID above the name; not "REF.NO." and the like).
$script:SignedIdThenName = [regex]('(?<id>\d(?<=(?:^|[^\w.-])\d)\d{9})(?![\w-])[ \t\u00A0]*(?:\r?\n[ \t\u00A0]*)?' +
    "(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)(?=[A-Z][A-Z'-]+\.[A-Z][A-Z'-]*\.)")

# =====================================================================================
# Small helper and the labelled-name detector (it calls Get-PiiSpan, so it stays here)
# =====================================================================================

function Get-ConfidenceRank {
    # Returns 3, 2 or 1 for High, Medium or Low (0 for anything else), for sorting.
    param([string]$Confidence)
    if ($script:ConfidenceRank.ContainsKey($Confidence)) { return $script:ConfidenceRank[$Confidence] }
    return 0
}

function Find-NameHits {
    <#
    .SYNOPSIS
      Names after a label (Name:, Full name:, Member:, Employee:, Spouse:, Dependent:, NOK:,
      Next of kin:, POC:). Added at the PAM's request, 2026-11-20.
    .DESCRIPTION
      Low by itself: a name alone is PII but rarely sensitive. Get-PiiSpan raises it to High when an
      identifier is near ($NameToIdentifierChars). A name without a label is not found: that would
      need a list of real names. Military "LAST, FIRST MI" form is found.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit) for the name only, not its label.
    #>
    param([string]$Text)
    $label = '(?:full\s+name|member\s+name|name|member|employee|spouse|dependent|next\s+of\s+kin|NOK|POC)\s*[:=-]\s*'
    # Labels of other data often follow a name on the same line; a name stops before them.
    $stopWords = '(?!(?:SSN|SSAN|DOB|DoD|EDIPI|CAC|Phone|Cell|Mobile|Home|Address|Email|E-mail|Rank|Grade|Unit|Date|Born)\b)'
    $name = "[A-Z][A-Za-z'\-]+(?:,?\s+$stopWords[A-Z](?:[A-Za-z'\-]+|\.))+"
    foreach ($m in [regex]::Matches($Text, '(?i:' + $label + ')(' + $name + ')')) {
        New-PiiHit 'Name' 'Low' $m.Groups[1].Index $m.Groups[1].Value
    }
}

# =====================================================================================
# All detectors together
# =====================================================================================

function Get-AllHits {
    <#
    .SYNOPSIS
      Runs every detector on the text and returns all hits, overlaps included.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    Find-SsnHits $Text
    Find-DodIdHits $Text
    Find-CardHits $Text
    Find-DobHits $Text
    Find-PassportHits $Text
    Find-PhoneHits $Text
    Find-PersonalEmailHits $Text
    Find-HomeAddressHits $Text
    Find-NameHits $Text
}

function Remove-OverlappingHits {
    <#
    .SYNOPSIS
      Keeps one hit per stretch of text: the type earlier in $DetectorOrder wins, then the higher
      confidence. Example: the 9 digits inside a 16-digit card number are not also an SSN.
    .DESCRIPTION
      The kept hits never overlap, so kept in order of their start they also end in order. A new hit overlaps a
      kept one only if it overlaps the kept hit that starts last before the new hit ends: one binary search per
      hit instead of a check against every kept hit. (Each kept hit is inserted in place, a fast memory move.)
    .PARAMETER Hits
      All hits from Get-AllHits.
    .OUTPUTS
      The kept hits (a list, in priority order).
    #>
    param([object[]]$Hits)
    $sorted = $Hits | Where-Object { $_ } | Sort-Object `
        @{ Expression = { [array]::IndexOf($script:DetectorOrder, $_.Detector) } },
        @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true },
        Index
    $kept = New-Object System.Collections.Generic.List[object]
    $starts = New-Object System.Collections.Generic.List[int]   # starts of the kept hits, ascending
    $ends = New-Object System.Collections.Generic.List[int]     # their ends, in the same order (also ascending)
    foreach ($hit in $sorted) {
        $end = $hit.Index + $hit.Length
        # The first kept hit that starts at or after this hit's end (BinarySearch gives its place as ~n if absent).
        $after = $starts.BinarySearch($end)
        if ($after -lt 0) { $after = -bnot $after }
        while ($after -gt 0 -and $starts[$after - 1] -ge $end) { $after-- }
        if ($after -gt 0 -and $ends[$after - 1] -gt $hit.Index) { continue }   # overlaps the kept hit before it
        $starts.Insert($after, $hit.Index)
        $ends.Insert($after, $end)
        $kept.Add($hit)
    }
    return , $kept
}

function Find-SignedIdHits {
    <#
    .SYNOPSIS
      E-signature DoD ID candidates (Low, the PAM, 2026-11-21) that the CAC-name rule of Find-DodIdHits misses: after a
      CAC name ($SignedName; small letters only after a signature word), alone after a signature word ($SignedBareId), or
      before a CAC name in capitals ($SignedIdThenName). Get-PiiSpan adds a candidate only where no other hit is.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), DoD ID / Low (High with a DoD ID label just before), each with Sure = $true when a
      signature word is just before it or the digits touch the name's last dot.
    #>
    param([string]$Text)
    if ($Text.IndexOfAny([char[]]'0123456789') -lt 0) { return }   # every rule needs the digits (speed)
    $seen = @{}
    $found = New-Object System.Collections.Generic.List[object]   # (match, needs a signature word)
    foreach ($m in $script:SignedName.Matches($Text)) { $found.Add(@($m, $false)) }
    foreach ($m in $script:SignedIdThenName.Matches($Text)) { $found.Add(@($m, $false)) }
    # The any-case name and the digits alone: only in the 400 characters from each signature word, each match then checked
    # on the whole text (review of PR 60: one word must not make long text slow).
    $ranges = New-Object System.Collections.Generic.List[int[]]   # the windows, overlapping ones joined (each part read once)
    foreach ($w in [regex]::Matches($Text, $script:SignedWords)) {
        $end = [Math]::Min($Text.Length, $w.Index + 400)
        if ($ranges.Count -gt 0 -and $w.Index -le $ranges[$ranges.Count - 1][1]) { $ranges[$ranges.Count - 1][1] = $end } else { $ranges.Add(@($w.Index, $end)) }
    }
    foreach ($range in $ranges) {
        $window = $Text.Substring($range[0], $range[1] - $range[0])
        foreach ($pair in @(($script:SignedNameAnyCase, $script:SignedNameAnyCaseAt), ($script:SignedBareId, $script:SignedBareIdAt))) {
            foreach ($m in $pair[0].Matches($window)) { $full = $pair[1].Match($Text, $range[0] + $m.Index); if ($full.Success) { $found.Add(@($full, $true)) } }
        }
    }
    foreach ($item in $found) {
        $m = $item[0]
        $digits = $m.Groups['id']
        if ($seen.ContainsKey($digits.Index)) { continue }
        $word = Test-KeywordBefore $Text $m.Index $script:SignedWords
        if ($item[1] -and -not $word) { continue }
        $seen[$digits.Index] = $true
        $confidence = if (Test-DodIdLabelBefore $Text $digits.Index) { 'High' } else { 'Low' }
        [pscustomobject]@{ Detector = 'DoD ID'; Confidence = $confidence; Index = $digits.Index; Length = $digits.Length; Value = $digits.Value
            Sure = ($word -or ($digits.Index -gt 0 -and $Text[$digits.Index - 1] -eq '.')) }
    }
}

function Add-SignedIdHits {
    <#
    .SYNOPSIS
      Adds the e-signature DoD IDs of Find-SignedIdHits that the other detectors did not see (the e-signature fix,
      2026-11-23). A candidate is added only where no kept hit is, so no other hit is lost or changed. One exception: SSN
      hits (not High, no dashes) wholly inside a sure candidate (a signature word before it, or digits touching the name) give way (the PAM,
      2026-11-21: e-signature DoD IDs were shown as SSNs). Added after the overlap removal, so it never outranks a hit.
    .PARAMETER Text
      The text searched (as the detectors saw it).
    .PARAMETER Kept
      The kept hits (Remove-OverlappingHits); changed in place.
    .OUTPUTS
      None.
    #>
    param([string]$Text, [System.Collections.Generic.List[object]]$Kept)
    $candidates = @(Find-SignedIdHits $Text)
    if ($candidates.Count -eq 0) { return }
    $owner = New-Object 'int[]' ($Text.Length + 1)   # which kept hit holds each character: its place + 1 (0: none)
    for ($k = 0; $k -lt $Kept.Count; $k++) { for ($i = $Kept[$k].Index; $i -lt $Kept[$k].Index + $Kept[$k].Length; $i++) { $owner[$i] = $k + 1 } }
    $giveWay = New-Object System.Collections.Generic.List[int]
    $garbledText = $script:GarbledChar.IsMatch($Text)   # the unreadable-text check only when the text has such a character
    foreach ($candidate in $candidates) {
        $end = $candidate.Index + $candidate.Length
        $held = @{}
        for ($i = $candidate.Index; $i -lt $end; $i++) { if ($owner[$i] -gt 0) { $held[$owner[$i] - 1] = $true } }
        $free = $true
        foreach ($k in $held.Keys) {
            $hit = $Kept[$k]
            $isSignedSsn = $candidate.Sure -and $hit.Detector -eq 'SSN' -and $hit.Confidence -ne 'High' -and
                $hit.Value -notmatch '-' -and $hit.Index -ge $candidate.Index -and $hit.Index + $hit.Length -le $end
            if (-not $isSignedSsn) { $free = $false }
        }
        # Not where the unreadable-text filter would drop it (a Low hit), or a swapped SSN would be lost (review of PR 60).
        if (-not $free -or ($garbledText -and $candidate.Confidence -eq 'Low' -and (Test-GarbledNear $Text $candidate.Index $candidate.Length))) { continue }
        foreach ($k in $held.Keys) { $giveWay.Add($k) }
        # Signed: never raises a name (Set-NameNearIdentifierHigh).
        $Kept.Add([pscustomobject]@{ Detector = 'DoD ID'; Confidence = $candidate.Confidence; Index = $candidate.Index
            Length = $candidate.Length; Value = $candidate.Value; Signed = $true })
        for ($i = $candidate.Index; $i -lt $end; $i++) { $owner[$i] = $Kept.Count }
    }
    foreach ($k in @($giveWay | Sort-Object -Unique -Descending)) { $Kept.RemoveAt($k) }
}

function Set-NameNearIdentifierHigh {
    <#
    .SYNOPSIS
      Raises a labelled name to High when an identifier hit is within $NameToIdentifierChars.
    .PARAMETER Hits
      The kept hits (changed in place).
    .OUTPUTS
      None.
    #>
    param([object[]]$Hits)
    foreach ($nameHit in @($Hits | Where-Object { $_.Detector -eq 'Name' })) {
        foreach ($other in $Hits) {
            $near = [Math]::Abs($other.Index - $nameHit.Index) -le $script:NameToIdentifierChars
            # An e-signature DoD ID (Low, or added by Add-SignedIdHits) does not raise a name: the name would be surer than the
            # ID (reviews of PR 50 and PR 60).
            $signatureId = $other.Detector -eq 'DoD ID' -and ($other.Confidence -eq 'Low' -or $other.PSObject.Properties['Signed'])
            if ($near -and -not $signatureId -and $script:IdentifierTypes -contains $other.Detector) { $nameHit.Confidence = 'High'; break }
        }
    }
}

function Test-GarbledNear {
    <#
    .SYNOPSIS
      True when the text around a hit is mostly unreadable: at least $GarbledShare of its visible characters are
      private-use, replacement or control characters or Latin-1 signs ($GarbledChar), as when a symbol font is read
      as text (the PAM, 2026-11-21: 50-100 characters of Wingdings flagged as Low hits).
    .PARAMETER Text
      The text searched.
    .PARAMETER Index
      The hit's start.
    .PARAMETER Length
      The hit's length.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, [int]$Index, [int]$Length)
    $start = [Math]::Max(0, $Index - $script:GarbledWindowChars)
    $end = [Math]::Min($Text.Length, $Index + $Length + $script:GarbledWindowChars)
    $around = $Text.Substring($start, $Index - $start) + $Text.Substring($Index + $Length, $end - $Index - $Length)
    $visible = @($script:VisibleChar.Matches($around)).Count
    if ($visible -lt $script:GarbledMinChars) { return $false }
    return (@($script:GarbledChar.Matches($around)).Count / $visible) -ge $script:GarbledShare
}

function Get-PiiSpan {
    <#
    .SYNOPSIS
      The one entry point of this module: every PII hit in a text, without overlaps, in text order.
      Hits hold raw values: pass them only to PiiMasking.psm1.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), sorted by position.
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    # Underscores join words in file names ("CAC_1234..."), which defeats word-boundary checks.
    # Seen as spaces by the detectors only; the length is the same, so positions still fit the text.
    $searchText = $Text.Replace('_', ' ')
    $kept = Remove-OverlappingHits @(Get-AllHits $searchText)
    Add-SignedIdHits $searchText $kept
    Set-NameNearIdentifierHigh $kept.ToArray()
    # A Low hit inside unreadable text (symbols, a font read wrong) is noise: it is dropped. High and Medium hits stay.
    $kept = @($kept | Where-Object { $_.Confidence -ne 'Low' -or -not (Test-GarbledNear $searchText $_.Index $_.Length) })
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

Export-ModuleMember -Function Get-PiiSpan, Get-ConfidenceRank
