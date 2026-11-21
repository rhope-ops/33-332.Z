<#
PiiDetectors.psm1 - finds text that looks like PII and returns where it is.

  Purpose:     One detector function per PII type (SSN, DoD ID, date of birth, payment card, passport,
               phone, personal e-mail, home address, labelled name). Get-PiiSpan runs them all, removes
               overlaps and returns each hit with its position and raw value.
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     hit objects WITH the raw value. Only PiiMasking.psm1 reads them; it masks every value
               before anything is shown or saved (the PAM, 2026-11-20: all values masked).
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


# Regular expressions: .NET keeps the last 15 patterns compiled; the detectors and the masking use more, so each file
# re-parsed them. A larger cache keeps them all (a speed-up for many files; the author's choice, 2026-11-20).
[regex]::CacheSize = 64

# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last.
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

$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone',
    'Personal e-mail', 'Home address', 'Name')

# Types that make a nearby name sensitive (see $NameToIdentifierChars).
$script:IdentifierTypes = @('SSN', 'DoD ID', 'Date of birth', 'Passport', 'Payment card')

# Confidence labels as numbers, for sorting (High first).
$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

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

function Get-ConfidenceRank {
    # Returns 3, 2 or 1 for High, Medium or Low (0 for anything else), for sorting.
    param([string]$Confidence)
    if ($script:ConfidenceRank.ContainsKey($Confidence)) { return $script:ConfidenceRank[$Confidence] }
    return 0
}

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
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if (-not (Test-SsnValid $m.Groups[1].Value $m.Groups[3].Value $m.Groups[4].Value)) { continue }
        $labelled = Test-KeywordBefore $Text $m.Index $labels
        # Without a label, 9 digits that are part of a gapped DoD ID are no SSN (Test-SsnPartOfLonger). The PAM's
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
    $labels = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'
    foreach ($m in [regex]::Matches($Text, '(?<!\w)\d{10}(?!\w)')) {
        if (Test-KeywordBefore $Text $m.Index $labels) { New-PiiHit 'DoD ID' 'High' $m.Index $m.Value }
    }
    # A CAC name, as a digital signature shows it: LAST.FIRST.MIDDLE. then the DoD ID (the digits may be drawn with single
    # spaces between them): Low, the only Low DoD ID, so the scan reads no version history for it (the PAM, 2026-11-21:
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
            if ($near -and $script:IdentifierTypes -contains $other.Detector) { $nameHit.Confidence = 'High'; break }
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
    Set-NameNearIdentifierHigh $kept.ToArray()
    # A Low hit inside unreadable text (symbols, a font read wrong) is noise: it is dropped. High and Medium hits stay.
    $kept = @($kept | Where-Object { $_.Confidence -ne 'Low' -or -not (Test-GarbledNear $searchText $_.Index $_.Length) })
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

Export-ModuleMember -Function Get-PiiSpan, Get-ConfidenceRank, Test-Luhn, Test-SsnValid
