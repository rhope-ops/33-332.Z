<#
PiiEngine.psm1 - PII patterns, e-signature DoD IDs, detector orchestration, and masking.

  Purpose:     Finds SSNs, DoD IDs (text and e-signatures), dates of birth, payment cards, passports, phones,
               personal e-mails, home addresses, and labelled names; removes overlaps and masks every value
               and snippet before anything is shown or saved (the Administrator, 2026-11-20: all values masked).
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     masked hit records (Find-PiiInText), masked strings and names (Protect-PiiText, Protect-PiiName).
  Depends on:  nothing.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# ==== PiiPatterns.psm1 ====
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

# A CAC name as certificates and digital signatures write it: 2 to 4 name parts, each ending in a dot, then the 10-digit
# DoD ID (single spaces allowed between digits: PDF signature blocks are often drawn with gaps). $CacNameEnd is the
# same name part alone, at the end of the text before an SSN-shaped match; $CacNameChars is how far back it is looked for.
# Certificates write the name in capitals: the last name has 2 letters or more, so initials ("J.A.", "U.S.") and lower-case
# words ("e.g.", "Acct.No.", "backup.config.") are no CAC name; nor are capital abbreviations such as "REF.NO." (review of PR 49).
$script:CacNamePrefix = "(?<![\w.])(?!(?:NO|NR|NUM|REF|ACCT|ACC|TEL|PH|FAX|ID|DOC|REV|VER)\.)[A-Z][A-Z'-]+\." +
    "(?!(?:NO|NR|NUM|REF|ACCT|TEL|PH|FAX|ID)\.)[A-Z][A-Z'-]*\.(?:[A-Z][A-Z'-]*\.){0,2}"
# Words that put a number in a digital signature (where a gapped DoD ID is drawn), looked for just before it.
$script:SignatureWords = '(?i:\bsigned\b|\bCN\s*=|\bDN\s*:)'
$script:CacNamePattern = $script:CacNamePrefix + '(?<id>\d(?: ?\d){9})(?![\w-]| \d)'
$script:CacNameEnd = $script:CacNamePrefix + '$'
$script:CacNameChars = 120
$script:DodIdLabels = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'


# E-mail top-level labels that mark a government address (not reported as personal e-mail).
$script:GovernmentMailLabels = @('gov', 'mil')

# Personal mail providers (AFI 33-332 A3's breach example names Yahoo, Gmail and Hotmail): an address there is Medium;
# an address at any other domain (a company, a vendor, a no-reply sender) is Low, to cut noise (the Administrator, 2026-11-20).
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

function Test-DodIdLabelBefore {
    # True when a DoD ID label ($DodIdLabels) is just before the given place (PiiEngine.psm1 Find-SignedIdHits).
    param([string]$Text, [int]$Index)
    return (Test-KeywordBefore $Text $Index $script:DodIdLabels)
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
      PiiEngine.psm1 removes the raw value before anything leaves the scan.
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
    # are not SSNs (the Administrator, 2026-11-20: trim false positives).
    $pattern = '(?<![\w-])(\d{3})([- ]?)(\d{2})\2(\d{4})(?![\w-])'
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if (-not (Test-SsnValid $m.Groups[1].Value $m.Groups[3].Value $m.Groups[4].Value)) { continue }
        $labelled = Test-KeywordBefore $Text $m.Index $labels
        # Without a label, 9 digits that are part of a gapped DoD ID are no SSN (Test-SsnPartOfLonger). The Administrator's
        # gov run, 2026-11-21: DoD IDs in PDF signature blocks were shown as SSNs.
        if (-not $labelled -and (Test-SsnPartOfLonger $Text $m)) { continue }
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
    $labels = $script:DodIdLabels
    foreach ($m in [regex]::Matches($Text, '(?<!\w)\d{10}(?!\w)')) {
        if (Test-KeywordBefore $Text $m.Index $labels) { New-PiiHit 'DoD ID' 'High' $m.Index $m.Value }
    }
    # A CAC name, as a digital signature shows it: LAST.FIRST.MIDDLE. then the DoD ID (the digits may be drawn with single
    # spaces between them): Low, the only Low DoD ID, so the scan reads no version history for it (the Administrator, 2026-11-21:
    # "Low tab, no history").
    foreach ($m in [regex]::Matches($Text, $script:CacNamePattern)) {
        $digits = $m.Groups['id']
        if (Test-KeywordBefore $Text $digits.Index $labels) { continue }   # already found above, with its label
        New-PiiHit 'DoD ID' 'Low' $digits.Index $digits.Value
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
      with an optional unit, city, state and ZIP code. Added at the Administrator's request, 2026-11-20.
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

# ==== PiiSignedIds.psm1 ====
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

# E-signature DoD IDs that PiiEngine.psm1 Find-DodIdHits misses (docs/reviews/2026-11-23-ion-esignature-layouts.md). $SignedName: a CAC name (line
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
# After a CAC name only a 3-3-4 phone is no ID: a 3-7 split there is an ID (the Administrator, 2026-11-24, "1a": narrow the guard).
$script:NotPhoneAfterName = '(?!\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{3}(?:[ \t\u00A0.-]|\r?\n)\d{4}(?![\w-]))'
# Before bare digits after a signature word: a date, a time or a page number (a year alone: $SignedYearLeadIn), then one gap, is no part of the ID
# (the Administrator, 2026-11-24, "1a"; the guard of 013ca82 blocked every digit and gap: 0 of 400 such IDs found).
# A date with an English month name or its 3 letters (any case, a dot or not) is a date: "20 Nov yyyy", "20NOVyyyy",
# "Nov 20, yyyy", "yyyy-Nov-20" (the Administrator, 2026-11-24, no-persistence plan step 1). Also (the Administrator, 2026-11-24, question
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
# After a year alone only 10 digits in a row or the 3-2-4-1 shape are an ID ("Signed 2012 #### #### ##": no ID; the Administrator, 2026-11-24, PR 69 MINORs).
$script:SignedYearLeadIn = '(?<![\w.-])(?:19|20)\d\d[ \t\u00A0]'
$script:SignedYearShape = '(?=(?:\d{10}|\d{3}[ \u00A0]\d{2}[ \u00A0]\d{4}[ \u00A0]\d)(?![\w-]))'
# "signed" in prose, not a stamp: "signed." or "Signed." after a form of "be", ending a sentence, a line break allowed as
# its gap ("was signed. Tracking"; PRs 57, 69; the Administrator, 2026-11-24, PR 69 MINORs). Other words and "?" or "!" count.
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
# The Administrator, 2026-11-24: a capitals name without dots on one line ("DOE JANE Q", "DOE, JANE Q."), only before an SSN shape and
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
      E-signature DoD ID candidates (Low, the Administrator, 2026-11-21) that the CAC-name rule of Find-DodIdHits misses: after a
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
    # Each candidate: (match, needs a signature word, follows a CAC name: an SSN in it gives way, the Administrator's gap rule 2026-11-24).
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
      (not High, no dashes) wholly inside, with a signature word before or digits touching the name (the Administrator, 2026-11-21, -24).
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
            # After a CAC name and a gap only an SSN of the first 9 digits gives way (9+1: the Administrator's gap rule; 1+9: review of PR 69).
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
      as text (the Administrator, 2026-11-21: 50-100 characters of Wingdings flagged as Low hits).
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

# ==== PiiDetectors.psm1 ====
# =====================================================================================
# Settings (named values; each says where it comes from)
# =====================================================================================

# A labelled name within this many characters of an identifier is raised to High. DoD privacy
# guidance treats a name together with an identifier (SSN, DoD ID, date of birth) as sensitive PII.
$script:NameToIdentifierChars = 80

# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last. One exception: Remove-SsnUnderPassport.
$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone',
    'Personal e-mail', 'Home address', 'Name')

# Types that make a nearby name sensitive (see $NameToIdentifierChars).
$script:IdentifierTypes = @('SSN', 'DoD ID', 'Date of birth', 'Passport', 'Payment card')

# Confidence labels as numbers, for sorting (High first).
$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

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
      Next of kin:, POC:). Added at the Administrator's request, 2026-11-20.
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

function Remove-SsnUnderPassport {
    <#
    .SYNOPSIS
      Drops an SSN Low hit (9 bare digits, no SSN label) that covers the same characters as a Passport hit, so the
      Passport hit is kept (the Administrator, 2026-11-24: a 9-digit value right after a Passport label is a Passport).
    .DESCRIPTION
      The one exception to $DetectorOrder. A Passport hit exists only with a Passport label in the label window
      (Find-PassportHits). An SSN with an SSN label (High) or with dashes or spaces (Medium) stays an SSN.
    .PARAMETER Hits
      All hits from Get-AllHits.
    .OUTPUTS
      The hits without those SSN hits.
    #>
    param([object[]]$Hits)
    $passportPlaces = @{}
    foreach ($hit in $Hits) {
        if ($hit -and $hit.Detector -eq 'Passport') { $passportPlaces['{0}+{1}' -f $hit.Index, $hit.Length] = $true }
    }
    if ($passportPlaces.Count -eq 0) { return $Hits }
    return @($Hits | Where-Object {
        -not ($_ -and $_.Detector -eq 'SSN' -and $_.Confidence -eq 'Low' -and $passportPlaces.ContainsKey(('{0}+{1}' -f $_.Index, $_.Length)))
    })
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

function Get-PiiSpan {
    <#
    .SYNOPSIS
      The one entry point of this module: every PII hit in a text, without overlaps, in text order.
      Hits hold raw values: pass them only to PiiEngine.psm1.
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
    $kept = Remove-OverlappingHits @(Remove-SsnUnderPassport @(Get-AllHits $searchText))
    Add-SignedIdHits $searchText $kept
    Set-NameNearIdentifierHigh $kept.ToArray()
    # A Low hit inside unreadable text (symbols, a font read wrong) is noise: it is dropped. High and Medium hits stay.
    $kept = @($kept | Where-Object { $_.Confidence -ne 'Low' -or -not (Test-GarbledNear $searchText $_.Index $_.Length) })
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

# ==== PiiMasking.psm1 ====
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
# 2026-11-21: "mask the CAC name in Context"). $CacCandidate finds two or more dotted parts in a row (and the part
# after them) that hold a part of two letters; a plain word, "e.g." or "U.S." never reaches Protect-CacName. A run
# starts only at its first character (or right after a dot that follows a non-name character), each start reads its
# run once (atomic groups, the two-letter test part by part), and no part length is capped, so a long part cannot
# hide the name after it (reviews of PR 81: speed and leaks). A dotted host or file name is masked too.
$script:CacCandidate = [regex]("(?:(?<![\p{L}\p{M}\p{Cs}_'#.-])|(?<=(?<![\p{L}\p{M}\p{Cs}_'#-])\.))" +
    "(?=(?:[\p{L}\p{M}\p{Cs}_'#-]+\.)*?(?>[\p{M}_'#-]*)[\p{L}\p{Cs}](?>[\p{M}_'#-]*)[\p{L}\p{Cs}])" +
    "(?>[\p{L}\p{M}\p{Cs}_'#-]+\.){2,}(?>[\p{L}\p{M}\p{Cs}_'#-]*)")
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
      The PII type (as named in PiiEngine.psm1).
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

# An SSN shape in a name, also next to a dash ("TeamA-123-45-6789-x"): 3, 2 and 4 digits with dashes, no digit on
# either side. Only for names (Protect-PiiName), never a finding: the detectors keep their own rules.
$script:NameSsnShapePattern = '(?<![0-9])[0-9]{3}-[0-9]{2}-(?<last>[0-9]{4})(?![0-9])'

function Protect-PiiName {
    <#
    .SYNOPSIS
      Masks a site, folder or file name the person must find again: each found value becomes its masked value and
      every e-mail name becomes '***'; other digits stay (the Administrator, 2026-11-20: "FY27 Plan.docx" stays as it is, so the
      file can be found). Example: "Roster_123-45-6789.docx" -> "Roster_***-**-6789.docx". An SSN shape that the
      detectors skip because a dash or a digit group joins it (a site name "TeamA-123-45-6789") is masked too (the
      Administrator, 2026-11-24): a name is masked more, never less.
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
    $result = [regex]::Replace($result, $script:EmailLocalPartPattern, '***')
    return [regex]::Replace($result, $script:NameSsnShapePattern, '***-**-${last}')
}

Export-ModuleMember -Function New-PiiHit, Test-KeywordBefore, Test-DodIdLabelBefore, Test-Luhn, Test-SsnValid, Find-SsnHits, Find-DodIdHits, Find-CardHits, Find-DobHits, Find-PassportHits,
    Find-PhoneHits, Find-PersonalEmailHits, Find-HomeAddressHits, Add-SignedIdHits, Test-GarbledNear, Get-PiiSpan, Get-ConfidenceRank, Find-PiiInText, Protect-PiiText, Protect-PiiName,
    Get-MaskedValue
