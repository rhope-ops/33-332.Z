<#
PiiPatterns.psm1 - the patterns and the one-type detectors that PiiDetectors.psm1 runs together.

  Purpose:     One detector function per PII type (SSN, DoD ID, date of birth, payment card, passport, phone,
               personal e-mail, home address), the label test before a match, the SSN and Luhn checks, and the
               hit object. Split from PiiDetectors.psm1 (fleet v2 DEC 6, 2026-11-23: 450 lines for hand-written code).
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     hit objects WITH the raw value. Only PiiMasking.psm1 reads them; it masks every value
               before anything is shown or saved (the PAM, 2026-11-20: all values masked).
  Depends on:  nothing (PiiDetectors.psm1 imports it next and runs these detectors).
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

# How far before a match to look for a label such as "SSN:" (characters). The author's choice,
# 2026-11-20: long enough for a label and a short phrase ("SSN of the member:"), short enough
# that a label in another table row does not count.
$script:KeywordWindowChars = 40

# A shorter window for labels that sit right next to their value ("DOB 01/02/1990", "cell 555-...").
$script:ShortKeywordWindowChars = 30

# Regular expressions: .NET keeps the last 15 patterns compiled; the detectors and the masking use more, so each file
# re-parsed them. A larger cache keeps them all (a speed-up for many files; the author's choice, 2026-11-20).
[regex]::CacheSize = 64

# A CAC name as certificates and digital signatures write it: 2 to 4 name parts, each ending in a dot. $CacNameEnd is the
# name alone, at the end of the text before an SSN-shaped match; $CacNameChars is how far back it is looked for.
# Certificates write the name in capitals: the last name has 2 letters or more, so initials ("J.A.", "U.S.") and lower-case
# words ("e.g.", "Acct.No.", "backup.config.") are no CAC name; nor are capital abbreviations such as "REF.NO." (review of PR 49).
$script:CacNamePrefix = "(?<![\w.])(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)[A-Z][A-Z'-]+\." +
    "(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)[A-Z][A-Z'-]*\.(?:[A-Z][A-Z'-]*\.){0,2}"
# Words that put a number in a digital signature (where a gapped DoD ID is drawn), looked for just before it.
$script:SignatureWords = '(?i:\bsigned\b|\bsignature\b|\bCN\s*=|\bDN\s*:)'
$script:CacNameEnd = $script:CacNamePrefix + '$'
$script:CacNameChars = 120

# E-signature DoD IDs as readers give them (2026-11-23; docs/reviews/2026-11-23-ion-esignature-layouts.md). $SignedName: a CAC
# name (any case; small letters only after a signature word), a line break allowed in or after a part, up to 3 words in a
# later part ("DOE.MARY ANN."); then 10 digits, one space, tab, no-break space or line break allowed between any two, and no
# digit after them on the line. Right after the last dot, dashes too (no line break then: a dashed SSN stays an SSN). After
# white space, a phone shape (3-3-4) is no ID (review of PR 60). A part has at most 40 letters and 1 line break (linear time).
$script:SignedPart = "[A-Z][A-Z'-]{0,39}(?:\r?\n[A-Z'-]{1,40})?"
$script:SignedId = '(?<id>\d(?:(?:[ \t\u00A0]|\r?\n)?\d){9})(?![\w-])(?![ \t\u00A0]\d)'
# The same on one line only (an ID alone after a signature word), and with dashes too (right after a CAC name).
$script:SignedIdLine = '(?<id>\d(?:[ \t\u00A0]?\d){9})(?![\w-])(?![ \t\u00A0]\d)'
$script:SignedIdDash = '(?<id>\d(?:[ \t\u00A0-]?\d){9})(?![\w-])(?![ \t\u00A0]\d)'
# A phone number: 3-3-4 digits with one space, tab, no-break space, dot, dash or line break between the groups.
$script:NotPhone = '(?!\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{4}(?![\w-]))'
$script:SignedName = [regex]('(?i)(?<![\w.])(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)(?<name>' +
    "[A-Z](?:[A-Z'-]{1,39}(?:\r?\n[A-Z'-]{1,40})?|\r?\n[A-Z'-]{1,40})\.(?:\r?\n)?(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)" +
    "$script:SignedPart(?: $script:SignedPart){0,2}\.(?:(?:\r?\n)?$script:SignedPart(?: $script:SignedPart){0,2}\.){0,2})" +
    "(?:$script:SignedIdDash|$script:SignedId|(?:[ \t\u00A0]+(?:\r?\n[ \t\u00A0]*)?|\r?\n[ \t\u00A0]*)$script:NotPhone$script:SignedId)")
# The 10 digits alone, on one line, after a signature word ("Digitally signed by 123 45 6789 0", "Signed_1234567890.pdf");
# not a date (2026 11 20 ...), not a phone, and not inside a dotted name or number (those are $SignedName's).
$script:SignedBareId = [regex]('(?<![\w.-])(?!(?:19|20)\d\d(?:[ \t\u00A0./-]|\r?\n))' + $script:NotPhone + $script:SignedIdLine)
# The 10 digits in a row just before a CAC name in capitals (a signature box that shows the ID on the line above the name);
# not before an abbreviation such as "REF.NO." (review of PR 60).
$script:SignedIdThenName = [regex]('(?<![\w.-])(?<id>\d{10})(?![\w-])[ \t\u00A0]*(?:\r?\n[ \t\u00A0]*)?' +
    "(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)(?=[A-Z][A-Z'-]+\.[A-Z][A-Z'-]*\.)")

# E-mail top-level labels that mark a government address (not reported as personal e-mail).
$script:GovernmentMailLabels = @('gov', 'mil')

# Personal mail providers (AFI 33-332 A3's breach example names Yahoo, Gmail and Hotmail): an address there is Medium;
# an address at any other domain (a company, a vendor, a no-reply sender) is Low, to cut noise (the PAM, 2026-11-20).
# The whole domain must be the provider and a top-level label (gmail.com, yahoo.co.uk): mail.<company>.com is not one.
$script:PersonalMailDomains = '^(?:gmail|googlemail|yahoo|ymail|rocketmail|hotmail|outlook|live|msn|aol|icloud|me|mac|' +
    'proton|protonmail|pm|gmx|mail|zoho|yandex|comcast|verizon|att|sbcglobal|cox|charter|spectrum|earthlink|juno|netzero)' +
    '\.(?:[a-z]{2,6}|co\.[a-z]{2}|com\.[a-z]{2})$'

# Payment card issuer prefixes and lengths (ISO/IEC 7812 IINs as the networks publish them): Visa 4 (13, 16, 19 digits);
# Mastercard 51-55 and 2221-2720 (16); American Express 34, 37 (15); Discover 6011, 644-649, 65 (16-19); JCB 3528-3589
# (16-19); Diners Club 300-305, 3095, 36, 38, 39 (14-19); Discover also 622126-622925 (16-19). A Luhn-valid number outside these is a document or tracking number.
$script:CardIssuerPattern = '^(?:4(?:\d{12}|\d{15}|\d{18})|(?:5[1-5]\d{2}|222[1-9]|22[3-9]\d|2[3-6]\d{2}|27[01]\d|2720)\d{12}|' +
    '3[47]\d{13}|(?:6011|64[4-9]\d|65\d{2})\d{12,15}|62(?:212[6-9]|21[3-9]\d|2[2-8]\d{2}|29[01]\d|292[0-5])\d{10,13}|(?:352[89]|35[3-8]\d)\d{12,15}|(?:30[0-5]\d|3095|3[689]\d{2})\d{10,15})$'

# =====================================================================================
# Small helpers
# =====================================================================================

function Test-KeywordBefore {
    <#
    .SYNOPSIS
      True when a label pattern (for example "SSN") appears just before a match.
    .PARAMETER Text
      The whole text.
    .PARAMETER Index
      Start position of the match in the text.
    .PARAMETER Pattern
      Regular expression of the labels (case is ignored).
    .PARAMETER Window
      How many characters before the match to search.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, [int]$Index, [string]$Pattern, [int]$Window = $script:KeywordWindowChars)
    $start = [Math]::Max(0, $Index - $Window)
    $before = $Text.Substring($start, $Index - $start)
    return [regex]::IsMatch($before, $Pattern, 'IgnoreCase')
}

function Test-SsnValid {
    <#
    .SYNOPSIS
      Applies the SSA issuing rules. The SSA never issues area 000, 666 or 900-999, group 00
      or serial 0000, so such a number is not an SSN.
    .PARAMETER Area
      The first 3 digits.
    .PARAMETER Group
      The middle 2 digits.
    .PARAMETER Serial
      The last 4 digits.
    .OUTPUTS
      [bool]
    #>
    param([string]$Area, [string]$Group, [string]$Serial)
    $areaNumber = [int]$Area
    if ($areaNumber -eq 0 -or $areaNumber -eq 666 -or $areaNumber -ge 900) { return $false }
    if ([int]$Group -eq 0) { return $false }
    if ([int]$Serial -eq 0) { return $false }
    return $true
}

function Test-Luhn {
    <#
    .SYNOPSIS
      Luhn (mod 10) check digit test, used by every payment card number (ISO/IEC 7812).
      It removes most random digit strings that only look like card numbers.
    .PARAMETER Digits
      Digits only, no spaces or dashes.
    .OUTPUTS
      [bool]
    #>
    param([string]$Digits)
    $sum = 0
    $doubleThis = $false
    # Walk from the right; double every second digit; subtract 9 from results over 9.
    for ($i = $Digits.Length - 1; $i -ge 0; $i--) {
        $digit = [int][string]$Digits[$i]
        if ($doubleThis) {
            $digit = $digit * 2
            if ($digit -gt 9) { $digit = $digit - 9 }
        }
        $sum += $digit
        $doubleThis = -not $doubleThis
    }
    return ($sum % 10 -eq 0)
}

function New-PiiHit {
    <#
    .SYNOPSIS
      Builds one hit object. It holds the raw value for overlap checks and masking;
      PiiMasking.psm1 removes the raw value before anything leaves the scan.
    .PARAMETER Detector
      The PII type, as named in $DetectorOrder.
    .PARAMETER Confidence
      High, Medium or Low.
    .PARAMETER Index
      Start position of the value in the text.
    .PARAMETER Value
      The raw value found.
    .OUTPUTS
      [pscustomobject] Detector, Confidence, Index, Length, Value.
    #>
    param([string]$Detector, [string]$Confidence, [int]$Index, [string]$Value)
    [pscustomobject]@{ Detector = $Detector; Confidence = $Confidence; Index = $Index; Length = $Value.Length; Value = $Value }
}

# =====================================================================================
# Detectors: one per PII type. Each returns zero or more hits (New-PiiHit).
# =====================================================================================

function Test-SsnPartOfLonger {
    <#
    .SYNOPSIS
      True when an SSN-shaped match without dashes is part of a gapped 10-digit DoD ID: one more digit alone after one
      space ("123 45 6789 0") with a signature word just before, or a CAC name ("LAST.FIRST." or "LAST.FIRST.M.") right before it, with no space.
      A dashed SSN, and one followed by a longer number (a date, a count), stays an SSN (review of PR 49).
    .PARAMETER Text
      The text searched.
    .PARAMETER Match
      The SSN regex match.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, [System.Text.RegularExpressions.Match]$Match)
    if ($Match.Groups[2].Value -eq '-') { return $false }
    $end = $Match.Index + $Match.Length
    $loneDigit = $end + 1 -lt $Text.Length -and $Text[$end] -eq ' ' -and [char]::IsDigit($Text[$end + 1]) -and
        ($end + 2 -ge $Text.Length -or -not [char]::IsLetterOrDigit($Text[$end + 2]))
    # A lone digit after it ("... 2 copies") is a gapped DoD ID only in a signature ("Digitally signed by", "CN=").
    if ($loneDigit -and (Test-KeywordBefore $Text $Match.Index $script:SignatureWords)) { return $true }
    $start = [Math]::Max(0, $Match.Index - $script:CacNameChars)
    return [regex]::IsMatch($Text.Substring($start, $Match.Index - $start), $script:CacNameEnd)
}

function Find-SignedIdHits {
    <#
    .SYNOPSIS
      E-signature DoD IDs, Low (the PAM, 2026-11-21): after a CAC name ($SignedName), alone after a signature word
      ($SignedBareId), or before a CAC name in capitals ($SignedIdThenName). Find-SsnHits drops an unlabelled SSN inside one.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), DoD ID / Low, one per ID.
    #>
    param([string]$Text)
    $seen = @{}
    foreach ($rule in @(($script:SignedName, 'name'), ($script:SignedBareId, 'word'), ($script:SignedIdThenName, ''))) {
        foreach ($m in $rule[0].Matches($Text)) {
            $digits = $m.Groups['id']
            # A name in small letters, and the digits alone, count only after a signature word.
            $needsWord = $rule[1] -eq 'word' -or ($rule[1] -eq 'name' -and $m.Groups['name'].Value -cmatch '[a-z]')
            if ($seen.ContainsKey($digits.Index) -or ($needsWord -and -not (Test-KeywordBefore $Text $m.Index $script:SignatureWords))) { continue }
            $seen[$digits.Index] = $true
            New-PiiHit 'DoD ID' 'Low' $digits.Index $digits.Value
        }
    }
}

function Find-SsnHits {
    <#
    .SYNOPSIS
      Social Security numbers: 3-2-4 digits with the same separator (dash, space or none),
      valid by the SSA rules (Test-SsnValid).
    .DESCRIPTION
      High   = a label (SSN, SSAN, social security) is near.
      Medium = dashes or spaces, no label.
      Low    = 9 bare digits, no label (many false matches: check by hand).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = 'SSN|SSAN|social\s*security|soc\.?\s*sec'
    # No letter, digit, underscore or dash may touch the number: Word bookmarks (_Toc123456789) and IDs (ABC123456789)
    # are not SSNs (the PAM, 2026-11-20: trim false positives).
    $pattern = '(?<![\w-])(\d{3})([- ]?)(\d{2})\2(\d{4})(?![\w-])'
    # Which e-signature DoD ID covers each character (0: none), so each SSN match is checked in constant time.
    $signedIds = @(Find-SignedIdHits $Text)
    $signedAt = $null
    if ($signedIds.Count -gt 0) { $signedAt = New-Object 'int[]' ($Text.Length + 1) }
    for ($k = 0; $k -lt $signedIds.Count; $k++) { foreach ($i in 0..($signedIds[$k].Length - 1)) { $signedAt[$signedIds[$k].Index + $i] = $k + 1 } }
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if (-not (Test-SsnValid $m.Groups[1].Value $m.Groups[3].Value $m.Groups[4].Value)) { continue }
        $labelled = Test-KeywordBefore $Text $m.Index $labels
        # Without a label, 9 digits inside an e-signature DoD ID (2026-11-23), or part of a gapped DoD ID (Test-SsnPartOfLonger;
        # the PAM's gov run, 2026-11-21: DoD IDs in PDF signature blocks were shown as SSNs), are no SSN.
        $inSignedId = $null -ne $signedAt -and $signedAt[$m.Index] -gt 0 -and $signedAt[$m.Index] -eq $signedAt[$m.Index + $m.Length - 1]
        if (-not $labelled -and ($inSignedId -or (Test-SsnPartOfLonger $Text $m))) { continue }
        $confidence = 'Low'
        if ($m.Groups[2].Value -ne '') { $confidence = 'Medium' }
        if ($labelled) { $confidence = 'High' }
        New-PiiHit 'SSN' $confidence $m.Index $m.Value
    }
}

function Find-DodIdHits {
    <#
    .SYNOPSIS
      DoD ID numbers (EDIPI, 10 digits). Reported only with a label near (DoD ID, DODID, EDIPI, CAC),
      because 10 bare digits are far more often something else.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit): High with a label; Low at the end of a CAC name (an e-signature).
    #>
    param([string]$Text)
    $labels = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'
    $labelledAt = @{}
    foreach ($m in [regex]::Matches($Text, '(?<!\w)\d{10}(?!\w)')) {
        if (Test-KeywordBefore $Text $m.Index $labels) { $labelledAt[$m.Index] = $true; New-PiiHit 'DoD ID' 'High' $m.Index $m.Value }
    }
    # An e-signature (Find-SignedIdHits): Low, the only Low DoD ID: no version history read (the PAM, 2026-11-21).
    foreach ($hit in @(Find-SignedIdHits $Text)) {
        if (Test-KeywordBefore $Text $hit.Index $labels) {
            if ($labelledAt.ContainsKey($hit.Index)) { continue }   # already found above, with its label
            $hit.Confidence = 'High'   # a label, and gaps in the digits ("Signed, EDIPI" then 3-2-4-1; review of PR 60)
        }
        $hit
    }
}

function Find-CardHits {
    <#
    .SYNOPSIS
      Payment card numbers: 13-19 digits (ISO/IEC 7812 lengths), spaces or dashes allowed,
      that pass the Luhn check.
    .DESCRIPTION
      High with a label near (card, Visa, Mastercard, Amex, Discover, GTC, GPC), else Medium.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = 'card|visa|master\s*card|amex|american\s*express|discover|\bGTC\b|\bGPC\b'
    foreach ($m in [regex]::Matches($Text, '(?<![\w-])\d(?:[ -]?\d){12,18}(?![\w-])')) {
        $digits = $m.Value -replace '\D', ''
        if (-not (Test-Luhn $digits)) { continue }
        if ($digits -notmatch $script:CardIssuerPattern) { continue }   # Luhn alone passes 1 in 10 random numbers
        $confidence = 'Medium'
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
        New-PiiHit 'Payment card' $confidence $m.Index $m.Value
    }
}

function Find-DobHits {
    <#
    .SYNOPSIS
      Dates of birth: a date with a label near (DOB, date of birth, birth date, born).
      A date without a label is not reported (most dates are suspenses and events).
    .DESCRIPTION
      Date forms: 1/2/1990, 01-02-90, 1990-01-02, 2 Jan 1990, 02JAN1990, January 2, 1990.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always High.
    #>
    param([string]$Text)
    $labels = '\bDOB\b|date\s*of\s*birth|birth\s*date|birthdate|\bborn\b|\bD\.O\.B\.?'
    $month = '(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*'
    $numeric = '\d{1,2}[/-]\d{1,2}[/-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2}'
    $military = '\d{1,2}\s?' + $month + '\s?\d{2,4}'
    $written = $month + '\s+\d{1,2},?\s+\d{4}'
    $pattern = '\b(?:' + $numeric + '|' + $military + '|' + $written + ')\b'
    foreach ($m in [regex]::Matches($Text, $pattern, 'IgnoreCase')) {
        if (Test-KeywordBefore $Text $m.Index $labels $script:ShortKeywordWindowChars) { New-PiiHit 'Date of birth' 'High' $m.Index $m.Value }
    }
}

function Find-PassportHits {
    <#
    .SYNOPSIS
      Passport numbers: 6-9 capital letters or digits (at least 5 digits) with "passport" near.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always Medium.
    #>
    param([string]$Text)
    foreach ($m in [regex]::Matches($Text, '\b[A-Z0-9]{6,9}\b')) {
        if (($m.Value -replace '\D', '').Length -lt 5) { continue }
        if (Test-KeywordBefore $Text $m.Index 'passport' $script:ShortKeywordWindowChars) { New-PiiHit 'Passport' 'Medium' $m.Index $m.Value }
    }
}

function Find-PhoneHits {
    <#
    .SYNOPSIS
      US phone numbers (10 digits, optional +1). Medium with a personal label near
      (cell, mobile, home, personal, my number), else Low (office numbers are common).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = '\bcell\b|mobile|\bhome\b|personal|\bmy\s+(?:number|phone)'
    $pattern = '(?<![\w-])(?:\+?1[ .-]?)?(?:\(\d{3}\)\s?|\d{3}[ .-])\d{3}[ .-]\d{4}(?![\w-])'
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        $confidence = 'Low'
        if (Test-KeywordBefore $Text $m.Index $labels $script:ShortKeywordWindowChars) { $confidence = 'Medium' }
        New-PiiHit 'Phone' $confidence $m.Index $m.Value
    }
}

function Find-PersonalEmailHits {
    <#
    .SYNOPSIS
      E-mail addresses outside government domains (last label not in $GovernmentMailLabels). Medium at a personal
      mail provider ($PersonalMailDomains), else Low.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), Medium or Low.
    #>
    param([string]$Text)
    # Local part: any characters but spaces, delimiters and path slashes (accents and apostrophes included),
    # starting after a space, a delimiter, a slash, an '@' or the start of the text. Domain: any such
    # characters too (non-ASCII domains), ending in a top-level label of 2 or more letters. No further '@' may
    # follow before the next space or delimiter: in a run of two addresses glued by '@', the first part is not
    # taken for an address whose 'domain' is the second address's local part. Any other trailing character
    # (? ! quotes ... and so on) still ends the address normally. The top-level label is atomic ((?>...)): without it,
    # a very long run of letters with no space makes the end check retry at every letter (seconds per 100,000 letters).
    # The domain uses the internet's own limits (RFC 1035: a label is at most 63 characters; at most 10 labels here;
    # a top-level label of 2 to 24 letters), so a long dotted run without spaces is tried a fixed number of ways,
    # not once per label: the time grows with the text, not with its square. The address ends where its domain ends:
    # no letter or digit right after the top-level label, and no letter or digit right after a final period (so an
    # address that ends a sentence, then a quote, a dash or a star, is still found). The no-further-'@' look-ahead
    # stops at the next '@', so a long run of glued addresses is read once, not again for every address in it.
    $notInAddress = '\s@<>()\[\],;:"/\\'
    $domain = '((?:[^' + $notInAddress + '.]{1,63}\.){1,10}((?>\p{L}{2,24})))'
    $endsAddress = '(?![\p{L}\p{N}])(?!\.[\p{L}\p{N}])(?![^\s<>()\[\],;:"/\\@]*@)'
    $pattern = '(?<=^|[' + $notInAddress + '])[^' + $notInAddress + ']+@' + $domain + $endsAddress
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if ($script:GovernmentMailLabels -contains $m.Groups[2].Value.ToLower()) { continue }
        # A personal mail provider is Medium; any other domain (a company, a vendor, a no-reply address) is Low.
        $confidence = 'Low'
        if ($m.Groups[1].Value.ToLower() -match $script:PersonalMailDomains) { $confidence = 'Medium' }
        New-PiiHit 'Personal e-mail' $confidence $m.Index $m.Value
    }
}

function Find-HomeAddressHits {
    <#
    .SYNOPSIS
      US street addresses: house number, 1-4 capitalised words and a street type (St, Ave, Rd ...),
      with an optional unit, city, state and ZIP code. Added at the PAM's request, 2026-11-20.
    .DESCRIPTION
      High   = a label near (home address, residence, mailing address, lives at, home).
      Medium = ends with a 2-letter state and a ZIP code.
      Low    = street only (it can be a work address: check by hand).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $streetTypes = 'Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Court|Ct|Boulevard|Blvd|Way|Circle|Cir|' +
                   'Place|Pl|Terrace|Ter|Parkway|Pkwy|Highway|Hwy|Trail|Trl'
    $street = '\b\d{1,6}\s+(?:[NSEW]\.?\s+)?(?:[A-Z][A-Za-z]+\s+){1,4}(?:' + $streetTypes + ')\b\.?'
    $unit = '(?:(?:Apt|Unit|#)\s*\w+,?\s*)?'
    $cityStateZip = '(?:,?\s*' + $unit + '[A-Z][A-Za-z]+(?:\s[A-Z][A-Za-z]+)?,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?)?'
    $labels = 'home\s*address|residence|residential|mailing\s*address|lives\s+at|\bhome\b'
    foreach ($m in [regex]::Matches($Text, $street + $cityStateZip)) {
        $confidence = 'Low'
        if ($m.Value -match '\b[A-Z]{2}\s+\d{5}') { $confidence = 'Medium' }
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
        New-PiiHit 'Home address' $confidence $m.Index $m.Value
    }
}

Export-ModuleMember -Function New-PiiHit, Test-Luhn, Test-SsnValid, Find-SsnHits, Find-DodIdHits, Find-CardHits, Find-DobHits,
    Find-PassportHits, Find-PhoneHits, Find-PersonalEmailHits, Find-HomeAddressHits
