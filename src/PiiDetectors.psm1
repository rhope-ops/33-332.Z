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
$script:DigitRun = [regex]'\d(?:(?:[ \t\u00A0-]|\r?\n)?\d){9}'   # 10 digits with gaps: where a window can find an ID
$script:NotDate = '(?!(?:19|20)\d\d(?:[ \t\u00A0./-]|\r?\n|(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))|\d{1,2}[ \t./-]\d{1,2}[ \t./-](?:19|20)\d\d)'
$script:NotPhone = '(?!\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{3}(?:[ \t\u00A0.-]|\r?\n)?\d{4}(?![\w-]))'   # 3-3-4 or 3-7: a phone
# Each pattern starts with a letter or digit, then looks back, so the engine skips ahead fast on long text (review of PR 60).
$script:SignedNameText = ("(?<name>[A-Z](?<=(?:^|[^\w.])[A-Z])(?:[A-Z'-]{1,39}(?:\r?\n[A-Z'-]{1,40})?|\r?\n[A-Z'-]{1,40})\." +
    '(?<!(?:^|[^\w.])(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)' +
    "(?:\r?\n)?(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)" +
    "$script:SignedPart(?: $script:SignedPart){0,2}\.(?:(?:\r?\n)?$script:SignedPart(?: $script:SignedPart){0,2}\.){0,2})" +
    "(?:$script:SignedIdDash|$script:SignedId|(?:[ \t\u00A0]+(?:\r?\n[ \t\u00A0]*)?|\r?\n[ \t\u00A0]*)$script:NotDate$script:NotPhone$script:SignedId)")
$script:SignedName = [regex]$script:SignedNameText
$script:SignedNameAnyCase = [regex]('(?i)' + $script:SignedNameText)
$script:SignedNameAnyCaseAt = [regex]('(?i)\G(?:' + $script:SignedNameText + ')')   # \G: only at the given place
$script:SignedWords = '(?i:(?:sign(?:ed|ature)\b|CN\s*=|DN\s*:)(?<!\w(?:sign(?:ed|ature)|CN\s*=|DN\s*:)))'   # a word start, said last (speed)
# The 10 digits alone on one line after a signature word ("Signed_1234567890.pdf"): not a date, a phone or in a dotted name.
$script:SignedBareText = '(?<![\w.-])(?<!\d[ \t\u00A0./])' + $script:NotDate + $script:NotPhone + $script:SignedIdLine
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

function Get-FreeDigitRun {
    <#
    .SYNOPSIS
      The starts of the 10-digit runs (with gaps) that no kept hit other than an SSN holds, sorted: the places where
      Find-SignedIdHits can find an ID (CAC-name text that main found gives none, so it reads nothing more: speed).
    .PARAMETER Text
      The text to search.
    .PARAMETER Starts
      The kept hits' starts, sorted, with Order (their places in Kept).
    .PARAMETER Order
      See Starts.
    .PARAMETER Kept
      The kept hits.
    .OUTPUTS
      The run starts, as int[].
    #>
    param([string]$Text, [int[]]$Starts, [int[]]$Order, $Kept)
    if ($Starts.Length -eq 0) { return , [int[]]@(foreach ($r in $script:DigitRun.Matches($Text)) { $r.Index }) }   # no hit: all (speed)
    $list = New-Object System.Collections.Generic.List[int]
    foreach ($r in $script:DigitRun.Matches($Text)) {
        $at = $r.Index; $stop = $r.Index + $r.Length
        while ($true) {
            $p = [Array]::BinarySearch($Starts, $at)
            if ($p -lt 0) { $p = (-bnot $p) - 1 }
            $hold = if ($p -ge 0) { $Kept[$Order[$p]] } else { $null }
            if ($null -eq $hold -or $hold.Detector -eq 'SSN' -or $hold.Index + $hold.Length -le $at) { $list.Add($at); break }
            # The run starts in a hit: look again just after the hit, so an ID one gap after a phone or ZIP keeps its run
            # (seventh review of PR 60). Only a short part of the text: Regex.Match with a start reads far on long text.
            $from = $hold.Index + $hold.Length
            if ($from -ge $stop) { break }   # the next run of Matches starts at or after $stop
            $tail = $script:DigitRun.Match($Text.Substring($from, [Math]::Min($Text.Length, $stop + 40) - $from))
            if (-not $tail.Success -or $from + $tail.Index -ge $stop) { break }
            $at = $from + $tail.Index
        }
    }
    return , $list.ToArray()
}

function Find-SignedIdHits {
    <#
    .SYNOPSIS
      E-signature DoD ID candidates (Low, the PAM, 2026-11-21) that the CAC-name rule of Find-DodIdHits misses: after a
      CAC name ($SignedName; small letters only after a signature word), alone after a signature word ($SignedBareId), or
      before a CAC name in capitals ($SignedIdThenName). Get-PiiSpan adds a candidate only where no other hit is.
    .PARAMETER Text
      The text to search.
    .PARAMETER Starts
      The kept hits' starts, sorted (Add-SignedIdHits), with Order (their places in Kept): a digit run a kept hit other
      than an SSN holds needs no reading.
    .PARAMETER Order
      See Starts.
    .PARAMETER Kept
      The kept hits.
    .OUTPUTS
      Found: the candidates in rule order, each a pair: the regex match (group "id": the digits) and $true when it counts
      only after a signature word. WordStarts, WordEnds: the signature words, in text order. Add-SignedIdHits checks the
      words and labels, only for candidates no other hit holds.
    #>
    param([string]$Text, [int[]]$Starts = @(), [int[]]$Order = @(), $Kept = @())
    $found = New-Object System.Collections.Generic.List[object]   # (match, needs a signature word)
    $result = [pscustomobject]@{ Found = $found; WordStarts = [int[]]@(); WordEnds = [int[]]@() }
    # Every rule needs 10 digits: one fast scan first; none, then nothing else is read (speed on prose: review of PR 60).
    $runs = Get-FreeDigitRun $Text $Starts $Order $Kept
    if ($runs.Length -eq 0) { return $result }
    $words = @([regex]::Matches($Text, $script:SignedWords))   # on the whole text, so no cut makes "assigned" a word
    $result.WordStarts = [int[]]@(foreach ($w in $words) { $w.Index })
    $result.WordEnds = [int[]]@(foreach ($w in $words) { $w.Index + $w.Length })
    foreach ($m in $script:SignedName.Matches($Text)) { $found.Add(@($m, $false)) }
    foreach ($m in $script:SignedIdThenName.Matches($Text)) { $found.Add(@($m, $false)) }
    # The any-case name and the digits alone: only in the 400 characters from each signature word, each match then checked
    # on the whole text (review of PR 60: one word must not make long text slow).
    $ranges = New-Object System.Collections.Generic.List[int[]]   # the windows, overlapping ones joined (each part read once)
    foreach ($w in $words) {   # only windows with 10 digits in them
        $next = [Array]::BinarySearch($runs, $w.Index)
        if ($next -lt 0) { $next = -bnot $next }
        if ($next -ge $runs.Length -or $runs[$next] -ge $w.Index + 400) { continue }
        $end = [Math]::Min($Text.Length, $w.Index + 400)
        if ($ranges.Count -gt 0 -and $w.Index -le $ranges[$ranges.Count - 1][1]) { $ranges[$ranges.Count - 1][1] = $end } else { $ranges.Add(@($w.Index, $end)) }
    }
    foreach ($range in $ranges) {
        $window = $Text.Substring($range[0], $range[1] - $range[0])
        foreach ($pair in @(($script:SignedNameAnyCase, $script:SignedNameAnyCaseAt), ($script:SignedBareId, $script:SignedBareIdAt))) {
            foreach ($m in $pair[0].Matches($window)) { $full = $pair[1].Match($Text, $range[0] + $m.Index); if ($full.Success) { $found.Add(@($full, $true)) } }
        }
    }
    return $result
}

function Add-SignedIdHits {
    <#
    .SYNOPSIS
      Adds Find-SignedIdHits' e-signature DoD IDs only where no kept hit is (after the overlap removal), except over SSNs
      (not High, no dashes) wholly inside, with a signature word before or digits touching the name (the PAM, 2026-11-21, -24).
    .PARAMETER Text
      The text searched (as the detectors saw it).
    .PARAMETER Kept
      The kept hits (Remove-OverlappingHits); changed in place.
    .OUTPUTS
      None.
    #>
    param([string]$Text, [System.Collections.Generic.List[object]]$Kept)
    # The kept hits by start (they never overlap): binary searches, not a mark per character (speed: review of PR 60).
    $starts = New-Object 'int[]' $Kept.Count
    $order = New-Object 'int[]' $Kept.Count
    for ($k = 0; $k -lt $Kept.Count; $k++) { $starts[$k] = $Kept[$k].Index; $order[$k] = $k }
    [Array]::Sort($starts, $order)
    $signed = Find-SignedIdHits $Text $starts $order $Kept
    if ($signed.Found.Count -eq 0) { return }
    $added = New-Object 'bool[]' ($Text.Length + 1)   # the characters of the IDs added here
    $giveWay = New-Object System.Collections.Generic.List[int]
    $garbledText = $script:GarbledChar.IsMatch($Text)   # the unreadable-text check only when the text has such a character
    $seen = @{}
    foreach ($item in $signed.Found) {
        $digits = $item[0].Groups['id']
        $end = $digits.Index + $digits.Length
        if ($seen.ContainsKey($digits.Index) -or $added[$digits.Index] -or $added[$end - 1]) { continue }
        # The kept hits over the digits: from the last one that starts at or before them, in text order.
        $p = [Array]::BinarySearch($starts, $digits.Index)
        if ($p -lt 0) { $p = [Math]::Max(0, (-bnot $p) - 1) }
        $held = New-Object System.Collections.Generic.List[int]
        for ($q = $p; $q -lt $starts.Length -and $starts[$q] -lt $end; $q++) {
            if ($Kept[$order[$q]].Index + $Kept[$order[$q]].Length -gt $digits.Index) { $held.Add($order[$q]) }
        }
        if (@(foreach ($k in $held) { if ($Kept[$k].Detector -ne 'SSN') { $k } }).Count) { continue }   # cheap test first
        # A signature word wholly in the 40 characters before: the first word that starts there ends first (binary search).
        $at = $item[0].Index
        $place = [Array]::BinarySearch($signed.WordStarts, $at - 40)
        if ($place -lt 0) { $place = -bnot $place }
        $word = $place -lt $signed.WordStarts.Length -and $signed.WordEnds[$place] -le $at
        if ($item[1] -and -not $word) { continue }
        $seen[$digits.Index] = $true   # only after the word test: a candidate that needs a word must not block a sure one
        $confidence = if (Test-DodIdLabelBefore $Text $digits.Index) { 'High' } else { 'Low' }; $sure = $word -or ($digits.Index -gt 0 -and $Text[$digits.Index - 1] -eq '.')
        $free = $true
        foreach ($k in $held) {
            $hit = $Kept[$k]
            if (-not ($sure -and $hit.Confidence -ne 'High' -and $hit.Value -notmatch '-' -and $hit.Index -ge $digits.Index -and
                $hit.Index + $hit.Length -le $end)) { $free = $false }
        }
        # Not where the unreadable-text filter would drop it (a Low hit), or a swapped SSN would be lost (review of PR 60).
        if (-not $free -or ($garbledText -and $confidence -eq 'Low' -and (Test-GarbledNear $Text $digits.Index $digits.Length))) { continue }
        foreach ($k in $held) { $giveWay.Add($k) }
        # Signed: never raises a name (Set-NameNearIdentifierHigh).
        $Kept.Add([pscustomobject]@{ Detector = 'DoD ID'; Confidence = $confidence; Index = $digits.Index; Length = $digits.Length
            Value = $digits.Value; Signed = $true })
        for ($i = $digits.Index; $i -lt $end; $i++) { $added[$i] = $true }
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
