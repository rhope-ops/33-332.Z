<#
PiiSignedIds.psm1 - e-signature DoD IDs, and the unreadable-text test.

  Purpose:     Add-SignedIdHits adds the e-signature DoD IDs that PiiPatterns.psm1 Find-DodIdHits misses (Low, the PAM,
               2026-11-21 and -24), only where no other hit is. Test-GarbledNear tells when a hit sits in unreadable text.
               Split from PiiDetectors.psm1 (fleet v2 DEC 6, 2026-11-23: 450 lines for hand-written code; the PAM,
               2026-11-24: split first, then the PR 69 MINORs).
  Inputs:      plain text and the kept hits of PiiDetectors.psm1 (Remove-OverlappingHits).
  Outputs:     hit objects WITH the raw value, added to the kept hits. Only PiiMasking.psm1 reads them; it masks
               every value before anything is shown or saved (the PAM, 2026-11-20: all values masked).
  Depends on:  PiiPatterns.psm1 (Test-DodIdLabelBefore), imported first.
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
# After a CAC name only a 3-3-4 phone is no ID: a 3-7 split there is an ID (the PAM, 2026-11-24, "1a": narrow the guard).
$script:NotPhoneAfterName = '(?!\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{4}(?![\w-]))'
# Before bare digits after a signature word: a date, a time or a page number (a year alone: $SignedYearLeadIn), then one gap, is no part of the ID
# (the PAM, 2026-11-24, "1a"; the guard of 013ca82 blocked every digit and gap: 0 of 400 such IDs found).
# A date with an English month name or its 3 letters (any case, a dot or not) is a date: "20 Nov yyyy", "20NOVyyyy",
# "Nov 20, yyyy", "yyyy-Nov-20" (the PAM, 2026-11-24, no-persistence plan step 1). Also (the PAM, 2026-11-24, question
# dialog): "Nov 20 yyyy"; a 2-digit year in "20NOVyy", "20 Nov yy", "Nov 20, yy"; "20-Nov-yyyy", "20-Nov-yy"; ordinals
# (st, nd, rd, th, any case) in "Nov 20th, yyyy", "20th Nov yyyy", "Nov 1st yyyy". Gaps (space, tab, no-break space;
# optional in "20 Nov yy" as in "20 Nov yyyy"), the month's dot and the unbounded day are as in main's forms; "Nov 20 yy",
# "20th Nov yy", "Nov 20th, yy", "20th-Nov-yyyy", "20thNovyyyy" and "Nov-20-yyyy" stay no date.
$script:MonthName = '(?i:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?'
$script:Ordinal = '(?i:st|nd|rd|th)'
$script:SignedLeadIn = '(?:(?:19|20)\d\d[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2}[./-](?:(?:19|20)\d\d|\d\d)|\d{1,2}:\d\d(?::\d\d)?|' +
    '(?i:page)[ \t\u00A0]+\d{1,4}|\d{1,2}[ \t\u00A0]?' + $script:MonthName + '[ \t\u00A0]?(?:(?:19|20)\d\d|\d\d)|' +
    '\d{1,2}-' + $script:MonthName + '-(?:(?:19|20)\d\d|\d\d)|\d{1,2}' + $script:Ordinal + '[ \t\u00A0]' + $script:MonthName + '[ \t\u00A0](?:19|20)\d\d|' +
    '(?<![A-Za-z])' + $script:MonthName + '[ \t\u00A0]\d{1,2}(?:' + $script:Ordinal + '?,?[ \t\u00A0](?:19|20)\d\d|,[ \t\u00A0]\d\d)|' +
    '(?:19|20)\d\d-' + $script:MonthName + '-\d{1,2})[ \t\u00A0]'
# After a year alone only 10 digits in a row or the 3-2-4-1 shape are an ID ("Signed 2012 #### #### ##": no ID; the PAM, 2026-11-24, PR 69 MINORs).
$script:SignedYearLeadIn = '(?<![\w.-])(?:19|20)\d\d[ \t\u00A0]'
$script:SignedYearShape = '(?=(?:\d{10}|\d{3}[ \u00A0]\d{2}[ \u00A0]\d{4}[ \u00A0]\d)(?![\w-]))'
# "signed" in prose, not a stamp: "signed." or "Signed." after a form of "be", ending a sentence, a line break allowed as
# its gap ("was signed. Tracking"; PRs 57, 69; the PAM, 2026-11-24, PR 69 MINORs). Other words and "?" or "!" count.
$script:ProseBefore = [regex]'(?i)\b(?:was|were|is|are|been|be|being)[ \t\u00A0]+$'
$script:SentenceEnd = [regex]'^[Ss]igned\.(?:[ \t\u00A0]+(?:\r?\n[ \t\u00A0]*)?|\r?\n[ \t\u00A0]*)[A-Z][a-z]{2}'   # one line break at most
# A signature word counts when it starts at most this many characters before a candidate (PR 60: the "Rule" section
# of docs/reviews/2026-11-23-ion-esignature-layouts.md, "within 40 characters before").
$script:SignatureWordChars = 40
# The characters before "signed" read for a form of "be" ($ProseBefore) (PR 69, commit 037a675).
$script:ProseBeforeChars = 12
# A dashed date (y-m-d, m-d-y) then a gap, right after a CAC name: no ID; a dashed ID is no date shape (PR 57 and 69).
$script:NotDashDate = '(?!(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}-\d{1,2}-(?:\d{4}|\d{2}))(?:[ \t\u00A0]|\r?\n|$))'
# Each pattern starts with a letter or digit, then looks back, so the engine skips ahead fast on long text (review of PR 60).
$script:SignedNameText = ("(?<name>[A-Z](?<=(?:^|[^\w.])[A-Z])(?:[A-Z'-]{1,39}(?:\r?\n[A-Z'-]{1,40})?|\r?\n[A-Z'-]{1,40})\." +
    '(?<!(?:^|[^\w.])(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)' +
    "(?:\r?\n)?(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)" +
    "$script:SignedPart(?: $script:SignedPart){0,2}\.(?:(?:\r?\n)?$script:SignedPart(?: $script:SignedPart){0,2}\.){0,2})" +
    "$script:NotDashDate(?:$script:SignedIdDash|$script:SignedId|(?:[ \t\u00A0]+(?:\r?\n[ \t\u00A0]*)?|\r?\n[ \t\u00A0]*)$script:NotDate$script:NotPhoneAfterName$script:SignedId)")
$script:SignedName = [regex]$script:SignedNameText
$script:SignedNameAnyCase = [regex]('(?i)' + $script:SignedNameText)
$script:SignedNameAnyCaseAt = [regex]('(?i)\G(?:' + $script:SignedNameText + ')')   # \G: only at the given place
$script:SignedWords = '(?i:(?:sign(?:ed|ature)\b|CN\s*=|DN\s*:)(?<!\w(?:sign(?:ed|ature)|CN\s*=|DN\s*:)))'   # a word start, said last (speed)
# The 10 digits alone on one line after a signature word ("Signed_1234567890.pdf"): not a date, a phone or in a dotted name.
$script:SignedBareText = '(?<![\w.-])(?:(?<=' + $script:SignedLeadIn + ')|(?<=' + $script:SignedYearLeadIn + ')' + $script:SignedYearShape +
    '|(?<!\d[ \t\u00A0./]))' + $script:NotDate + $script:NotPhone + $script:SignedIdLine
$script:SignedBareId = [regex]$script:SignedBareText
$script:SignedBareIdAt = [regex]('\G(?:' + $script:SignedBareText + ')')
# The 10 digits in a row just before a CAC name in capitals (the ID above the name; not "REF.NO." and the like).
$script:SignedIdThenName = [regex]('(?<id>\d(?<=(?:^|[^\w.-])\d)\d{9})(?![\w-])[ \t\u00A0]*(?:\r?\n[ \t\u00A0]*)?' +
    "(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)(?=[A-Z][A-Z'-]+\.[A-Z][A-Z'-]*\.)")
# The PAM, 2026-11-24: a capitals name without dots on one line ("DOE JANE Q", "DOE, JANE Q."), only before an SSN shape and
# one digit ($GapShape), read in the 80 characters before it (speed; no headings: review of PR 69).
$script:NotNameWord = '\b(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\b'   # as the dotted rule's list
$script:GapShape = [regex]'(?<![\w-])(?:\d{3} \d{2} \d{4}|\d{9}) \d(?![\w-])(?![ \t\u00A0]\d)'
$script:SignedNameSpaced = [regex]("(?<name>[A-Z](?<=(?:^|[^\w.,])[A-Z])[A-Z'\u2019-]{1,39},?(?: [A-Z][A-Z'\u2019-]{0,39}){1,3}\.?)" +
    "[ \t\u00A0]{1,3}(?<id>(?:\d{3} \d{2} \d{4}|\d{9}) \d)$")

# =====================================================================================
# The e-signature DoD ID and the unreadable-text test
# =====================================================================================

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
      CAC name ($SignedName; small letters only after a signature word), a name in capitals without dots before an SSN shape
      and one digit ($SignedNameSpaced), alone after a signature word ($SignedBareId), or before a CAC name in capitals
      ($SignedIdThenName). Get-PiiSpan adds a candidate only where no other hit is.
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
      Found: the candidates in rule order, each three items: the match (group "id": the digits), $true when it counts
      only after a signature word, and $true when it follows a CAC name (an SSN of its first 9 digits gives way). WordStarts, WordEnds: the signature words, in text order. Add-SignedIdHits checks the
      words and labels, only for candidates no other hit holds.
    #>
    param([string]$Text, [int[]]$Starts = @(), [int[]]$Order = @(), $Kept = @())
    $found = New-Object System.Collections.Generic.List[object]   # (match, needs a signature word, follows a CAC name)
    $result = [pscustomobject]@{ Found = $found; WordStarts = [int[]]@(); WordEnds = [int[]]@() }
    # Every rule needs 10 digits: one fast scan first; none, then nothing else is read (speed on prose: review of PR 60).
    $runs = Get-FreeDigitRun $Text $Starts $Order $Kept
    if ($runs.Length -eq 0) { return $result }
    $words = @([regex]::Matches($Text, $script:SignedWords))   # on the whole text, so no cut makes "assigned" a word
    $result.WordStarts = [int[]]@(foreach ($w in $words) { $w.Index })
    $result.WordEnds = [int[]]@(foreach ($w in $words) { $w.Index + $w.Length })
    # Each candidate: (match, needs a signature word, follows a CAC name: an SSN in it gives way, the PAM's gap rule 2026-11-24).
    foreach ($m in $script:SignedName.Matches($Text)) { $found.Add(@($m, $false, $true)) }
    foreach ($m in $script:SignedIdThenName.Matches($Text)) { $found.Add(@($m, $false, $true)) }
    foreach ($g in $script:GapShape.Matches($Text)) {   # a name without dots: only just before an SSN shape and one digit
        $from = [Math]::Max(0, $g.Index - 80)
        $m = $script:SignedNameSpaced.Match($Text.Substring($from, $g.Index + $g.Length - $from))
        if (-not $m.Success -or ($m.Index -eq 0 -and $from -gt 0 -and $Text[$from - 1] -match '[\w.,]')) { continue }   # the name runs on before the window
        if ($m.Groups['name'].Value -match $script:NotNameWord) { continue }   # "ACCT NO", "REF NO", "TEL NR": no name
        $id = $m.Groups['id']   # the same fields as a regex match on the whole text (Add-SignedIdHits reads only these)
        $found.Add(@([pscustomobject]@{ Index = $from + $m.Index; Groups = @{ id = [pscustomobject]@{ Index = $from + $id.Index
                            Value = $id.Value; Length = $id.Length } } }, $false, $true))
    }
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
        foreach ($pair in @(($script:SignedNameAnyCase, $script:SignedNameAnyCaseAt, $true), ($script:SignedBareId, $script:SignedBareIdAt, $false))) {
            foreach ($m in $pair[0].Matches($window)) { $full = $pair[1].Match($Text, $range[0] + $m.Index); if ($full.Success) { $found.Add(@($full, $true, $pair[2])) } }
        }
    }
    return $result
}

function Test-SignatureWordBefore {
    <#
    .SYNOPSIS
      True when a signature word that ends at or before At starts in the $SignatureWordChars characters before it. With
      -Prose, a word in prose ("was signed." or "was Signed." ending a sentence: $ProseBefore, $SentenceEnd) does not count.
    .PARAMETER Text
      The text searched.
    .PARAMETER Signed
      Output of Find-SignedIdHits (WordStarts and WordEnds, in text order).
    .PARAMETER At
      Where the candidate starts.
    .PARAMETER Prose
      Leave out a signature word in prose.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, $Signed, [int]$At, [switch]$Prose)
    $place = [Array]::BinarySearch($Signed.WordStarts, $At)
    if ($place -lt 0) { $place = -bnot $place }
    $place--
    while ($place -ge 0 -and $Signed.WordEnds[$place] -gt $At) { $place-- }
    for (; $place -ge 0 -and $Signed.WordStarts[$place] -ge $At - $script:SignatureWordChars; $place--) {   # the words in the window, last first
        $start = $Signed.WordStarts[$place]
        $from = [Math]::Max(0, $start - $script:ProseBeforeChars)
        if (-not $Prose -or -not $script:SentenceEnd.IsMatch($Text.Substring($start, $At - $start)) -or
            -not $script:ProseBefore.IsMatch($Text.Substring($from, $start - $from))) { return $true }
    }
    return $false
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
        $word = Test-SignatureWordBefore $Text $signed $item[0].Index
        if ($item[1] -and -not (Test-SignatureWordBefore $Text $signed $item[0].Index -Prose)) { continue }
        $seen[$digits.Index] = $true   # only after the word test: a candidate that needs a word must not block a sure one
        $confidence = if (Test-DodIdLabelBefore $Text $digits.Index) { 'High' } else { 'Low' }; $sure = $word -or ($digits.Index -gt 0 -and $Text[$digits.Index - 1] -eq '.')
        $free = $true
        foreach ($k in $held) {
            $hit = $Kept[$k]
            # After a CAC name and a gap only an SSN of the first 9 digits gives way (9+1: the PAM's gap rule; 1+9: review of PR 69).
            if (-not (($sure -or ($item[2] -and $hit.Index -eq $digits.Index)) -and $hit.Confidence -ne 'High' -and $hit.Value -notmatch '-' -and $hit.Index -ge $digits.Index -and
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

Export-ModuleMember -Function Add-SignedIdHits, Test-GarbledNear
