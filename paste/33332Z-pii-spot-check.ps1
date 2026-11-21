<#
  Ion - DoD365 PII Reporter (33-332.Z; GENERATED from src/ by scripts/build_paste.py: do not edit this text)

  HOW TO RUN
    1. Optional: read this text first (the Open the script button saves it as a .txt file for Notepad).
    2. Open Windows PowerShell ISE and choose File > New.
    3. Paste ALL of this text (Ctrl+V) and press F5.
    4. When asked, paste a Teams link (... next to the team or channel name > Copy link) or a SharePoint
       site address (in Teams: the Shared tab > ... > Open in SharePoint). Sign in with your CAC.
       If you do not see a sign-in window, press Alt+Tab.
    5. Watch the two progress bars (the whole scan, and the file in work) and the line per folder.
       At the end, read the summary of hits per PII type. The results workbook is saved in
       Documents\Ion\<site>\<date>: a Summary tab, then one tab per PII type (worst first).
       Handle it as CUI: it shows where PII is.

  It scans the site's Documents library: every folder (in a team, one per standard channel) and the files at
  its top. One sign-in only: channel posts are not read.

  Built from these files; to change it, change them and rebuild with: python3 scripts/build_paste.py
    src/PiiDetectors.psm1, sha256 ea19831b4d68
    src/PiiMasking.psm1, sha256 752c251a94b2
    src/OcrText.psm1, sha256 a39acab441d4
    src/OleText.psm1, sha256 0b1924bcf524
    src/ContainerText.psm1, sha256 12e3acaa974e
    src/PdfCore.psm1, sha256 da11bdba883c
    src/PdfObjects.psm1, sha256 5ed1e4ab471c
    src/PdfFonts.psm1, sha256 f34800d85a7f
    src/PdfText.psm1, sha256 b76e2af725f1
    src/TextExtract.psm1, sha256 f27467e1e75a
    src/ReaderPool.psm1, sha256 0b40c1fe0897
    src/XlsxWriter.psm1, sha256 e29becdc6a65
    src/Report.psm1, sha256 4c5c871cb5c5
    src/ScanMeter.psm1, sha256 2c2a5a750104
    src/ScanQueue.psm1, sha256 1ec233b97fc7
    src/Accountability.psm1, sha256 2cced6e1f2d1
    src/ChannelScan.psm1, sha256 0a32b46dcc63
    src/DriveScan.psm1, sha256 097672c6cd08
    src/Find-ChannelPII.ps1, sha256 ffaee6b7c116

  Read-only: it never changes, moves or deletes anything in Teams or SharePoint. Every value found is masked.
  Run target: Windows PowerShell 5.1 with PnP.PowerShell 1.12.0; no admin rights, no file to unblock.
  Network: only the one site given.
  Writes: only the results workbook under $OutputFolder (default Documents\Ion).
  Credentials: your own CAC sign-in in the PnP sign-in window; nothing is stored.

  Contents (line numbers of this file):
    line    72  Settings
    line    83  Reader block
    line    89  Module PiiDetectors.psm1
    line   577  Module PiiMasking.psm1
    line   777  Module OcrText.psm1
    line   930  Module OleText.psm1
    line  1034  Module ContainerText.psm1
    line  1274  Module PdfCore.psm1
    line  1633  Module PdfObjects.psm1
    line  1985  Module PdfFonts.psm1
    line  2359  Module PdfText.psm1
    line  2951  Module TextExtract.psm1
    line  3334  Module ReaderPool.psm1
    line  3581  Reader block end
    line  3585  Module XlsxWriter.psm1
    line  3856  Module Report.psm1
    line  4419  Module ScanMeter.psm1
    line  4684  Module ScanQueue.psm1
    line  4900  Module Accountability.psm1
    line  5052  Module ChannelScan.psm1
    line  5607  Module DriveScan.psm1
    line  5837  Body of Find-ChannelPII.ps1
    line  5946  Clean-up
#>

# ==== Settings ==== (optional: edit the values between the quotes; empty = the script asks)
$SiteUrl = ''            # the team's Teams link, or its site https://usaf.dps.mil/teams/<team-site>; empty = asked
$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Ion')   # results: a folder per site, then per day
$MaxFileMB = 25          # files larger than this are reported as NOT SCANNED

# ISE runs pasted text in your session: keep its error setting so the Clean-up block at the end can put it back.
$previousErrorAction = $ErrorActionPreference
try {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ==== Reader modules ==== (one script block: this script loads it as an in-memory module, and each reader
# worker loads the same text; nothing is read from a file)
$IonReaderBlock = {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ==== Module PiiDetectors.psm1 ====
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

# Regular expressions: .NET keeps the last 15 patterns compiled; the detectors and the masking use more, so each file
# re-parsed them. A larger cache keeps them all (a speed-up for many files; the author's choice, 2026-11-20).
[regex]::CacheSize = 64

# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last.
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
        $confidence = 'Low'
        if ($m.Groups[2].Value -ne '') { $confidence = 'Medium' }
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
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
      Hit objects (New-PiiHit), always High.
    #>
    param([string]$Text)
    $labels = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'
    foreach ($m in [regex]::Matches($Text, '(?<!\w)\d{10}(?!\w)')) {
        if (Test-KeywordBefore $Text $m.Index $labels) { New-PiiHit 'DoD ID' 'High' $m.Index $m.Value }
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
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

# ==== Module PiiMasking.psm1 ====
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

function Get-MaskedText {
    <#
    .SYNOPSIS
      Masks a whole text and keeps its length: every digit becomes '#', every e-mail local part
      becomes '*', and every hit (names and addresses included) becomes '*'. Because the length
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

# ==== Module OcrText.psm1 ====
<#
OcrText.psm1 - text in images with Windows' built-in OCR (Windows.Media.Ocr), in memory, without a module.

  Purpose:     Reads the text of .jpg, .png, .tif, .bmp and .gif files and of scanned PDF pages, so the detectors can
               check them (the PAM, 2026-11-20: round 3 readers). Uses the OCR engine of Windows 10/11 through
               WinRT from Windows PowerShell 5.1; where it is missing (no OCR language, another OS), every image
               gets a "check by hand" reason instead.
  Inputs:      image bytes.
  Outputs:     a hashtable with Text, Reason (empty when the text was read) and Note (pages not read).
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1 and TextExtract.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 on Windows 10/11 (WinRT OCR); elsewhere images get a reason.
  Network:     none (the OCR engine runs on the computer).
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>


# Seconds one WinRT step (decode, read) may take; frames read from one multi-page TIFF.
$script:OcrStepMilliseconds = 30000
$script:MaxImageFrames = 30
# A scanned PDF (no text layer): its JPEG page images are read, at most this many (Get-ScanPagesText).
$script:MaxScanPages = 30
$script:ScanLatin1 = [System.Text.Encoding]::GetEncoding(28591)

$script:OcrMissingReason = 'Image: Windows OCR is not available on this computer (no OCR language?): check by hand'
$script:OcrFailedReason = 'Image: its text could not be read (damaged, or a format Windows cannot open): check by hand'

# The engine, found once per run: Checked, Engine, AsTask (the generic AsTask method) and the reason when missing.
$script:Ocr = @{ Checked = $false; Engine = $null; AsTask = $null }

# Tests (Linux, no WinRT) put a scriptblock here: param([byte[]]) -> hashtable Text, Reason.
$script:OcrOverride = $null

function Initialize-Ocr {
    <#
    .SYNOPSIS
      Finds the Windows OCR engine (WinRT, built into Windows 10/11) once per run; leaves Engine $null when it is not
      there, so every image then gets $OcrMissingReason.
    #>
    if ($script:Ocr.Checked) { return }
    $script:Ocr.Checked = $true
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $script:Ocr.AsTask = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
            $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $script:Ocr.Engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    }
    catch { $script:Ocr.Engine = $null }   # no WinRT here: every image gets $OcrMissingReason
}

function Wait-WinRt {
    <#
    .SYNOPSIS
      Waits for one WinRT async operation and returns its result; throws when it takes longer than the step limit.
    .PARAMETER Operation
      The IAsyncOperation.
    .PARAMETER Type
      Its result type.
    #>
    param($Operation, [type]$Type)
    $task = $script:Ocr.AsTask.MakeGenericMethod($Type).Invoke($null, @($Operation))
    if (-not $task.Wait($script:OcrStepMilliseconds)) { throw (New-Object System.TimeoutException 'OCR step') }
    return $task.Result
}

function Read-OcrBitmap {
    <#
    .SYNOPSIS
      The text of one decoded frame, its lines joined with line breaks; scaled down first when it is larger than
      the engine reads (OcrEngine.MaxImageDimension).
    .PARAMETER Frame
      A BitmapDecoder or BitmapFrame.
    .OUTPUTS
      [string]
    #>
    param($Frame)
    $max = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
    $width = [double]$Frame.PixelWidth; $height = [double]$Frame.PixelHeight
    if ($width -gt $max -or $height -gt $max) {
        $scale = [Math]::Min($max / $width, $max / $height)
        $transform = New-Object Windows.Graphics.Imaging.BitmapTransform
        $transform.ScaledWidth = [uint32][Math]::Floor($width * $scale)
        $transform.ScaledHeight = [uint32][Math]::Floor($height * $scale)
        $bitmap = Wait-WinRt ($Frame.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied, $transform,
            [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
            [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
    }
    else { $bitmap = Wait-WinRt ($Frame.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap]) }
    $result = Wait-WinRt ($script:Ocr.Engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    return (@($result.Lines | ForEach-Object { $_.Text }) -join "`n")
}

function Get-ImageText {
    <#
    .SYNOPSIS
      The text in one image (every frame of a multi-page TIFF, up to $MaxImageFrames), or why it could not be read.
      An image without text is read and gives empty text.
    .PARAMETER Bytes
      The image bytes (JPEG, PNG, TIFF, BMP or GIF).
    .OUTPUTS
      [hashtable] Text, Reason, Note.
    #>
    param([byte[]]$Bytes)
    if ($script:OcrOverride) { return (& $script:OcrOverride $Bytes) }
    Initialize-Ocr
    if (-not $script:Ocr.Engine) { return @{ Text = ''; Reason = $script:OcrMissingReason } }
    $memory = New-Object System.IO.MemoryStream(, $Bytes)
    try {
        $stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($memory)
        $decoder = Wait-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $pages = New-Object System.Collections.Generic.List[string]
        $count = [Math]::Min([int]$decoder.FrameCount, $script:MaxImageFrames)
        for ($i = 0; $i -lt $count; $i++) {
            $frame = Wait-WinRt ($decoder.GetFrameAsync([uint32]$i)) ([Windows.Graphics.Imaging.BitmapFrame])
            $pages.Add((Read-OcrBitmap $frame))
        }
        $note = ''
        if ([int]$decoder.FrameCount -gt $script:MaxImageFrames) { $note = "Only the first $script:MaxImageFrames pages of this image were read: check by hand" }
        return @{ Text = ($pages -join "`n"); Reason = ''; Note = $note }
    }
    catch { return @{ Text = ''; Reason = $script:OcrFailedReason } }
    finally { $memory.Dispose() }
}

function Get-ScanPagesText {
    <#
    .SYNOPSIS
      The text of a scanned PDF (JPEG page images, no text layer) read with Windows OCR, page by page, at most
      $MaxScanPages pages; more pages give a note, OCR that is not available a reason.
    .PARAMETER Images
      The JPEG image streams (Latin-1 text).
    .PARAMETER Note
      Notes already made for this PDF.
    .OUTPUTS
      [hashtable] Text, Reason, Note.
    #>
    param([string[]]$Images, [string]$Note)
    $pages = New-Object System.Collections.Generic.List[string]
    $notes = @($Note | Where-Object { $_ })
    for ($i = 0; $i -lt [Math]::Min($Images.Count, $script:MaxScanPages); $i++) {
        $read = Get-ImageText ($script:ScanLatin1.GetBytes($Images[$i]))
        if ($read.Reason) { return @{ Text = ''; Reason = 'Scanned PDF: ' + $read.Reason; Note = '' } }
        $pages.Add($read.Text)
    }
    if ($Images.Count -gt $script:MaxScanPages) { $notes += "Scanned PDF: only the first $script:MaxScanPages page images were read: check by hand" }
    return @{ Text = ($pages -join "`n"); Reason = ''; Note = ($notes -join '; ') }
}

# ==== Module OleText.psm1 ====
<#
OleText.psm1 - text from old binary Office files (.doc, .xls, .ppt) and Outlook messages (.msg), in memory.

  Purpose:     These files are OLE compound files, not zips. Their text is stored as runs of UTF-16 or 8-bit
               characters, which this module pulls out (the PAM, 2026-11-20: round 3 readers). Partial for .xls:
               numbers kept as binary values are not read, and the report says so for each .xls. Encrypted files
               are found and named.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text, Reason (not read, and why) and Note (read in part).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>


# Bytes as characters one to one, so a binary file can be searched with regular expressions.
$script:OleLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# First bytes of an OLE compound file (MS-CFB).
$script:OleSignature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)

# Text runs: at least 4 UTF-16LE characters, or at least 6 printable 8-bit characters (shorter runs are mostly
# noise from the binary structure).
$script:Utf16RunPattern = '(?:[\x09\x0A\x0D\x20-\x7E\xA0-\xFF]\x00){4,}'
$script:AnsiRunPattern = '[\x09\x0A\x0D\x20-\x7E]{6,}'

# Stream names that only an encrypted file has (UTF-16LE in the directory): an encrypted Office 2007+ file, and an
# encrypted PowerPoint 97-2003 file.
$script:EncryptedStreamNames = @('EncryptedPackage', 'EncryptedSummary')

# An old .xls keeps numbers as binary values, not text: a number typed into a cell (an SSN without dashes) is not
# read, so every .xls gets this note. Old .doc, .ppt and .msg files keep their text as text.
$script:XlsNote = 'Old Excel file: numbers stored as values are not read: check by hand'

function Test-OleSignature {
    # True when the bytes start like an OLE compound file.
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 8) { return $false }
    for ($i = 0; $i -lt 8; $i++) { if ($Bytes[$i] -ne $script:OleSignature[$i]) { return $false } }
    return $true
}

function Test-OleEncrypted {
    <#
    .SYNOPSIS
      True when an OLE file is encrypted: an encryption stream name in its directory, a Word file whose FIB says
      fEncrypted, or an Excel workbook with a FILEPASS record right after its first BOF record.
    .PARAMETER Text
      The file bytes as Latin-1 text.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, [byte[]]$Bytes)
    foreach ($name in $script:EncryptedStreamNames) {
        $utf16 = -join @($name.ToCharArray() | ForEach-Object { [string]$_ + [char]0 })
        if ($Text.IndexOf($utf16, [StringComparison]::Ordinal) -ge 0) { return $true }
    }
    # Streams start on 512-byte sectors. Word: the FIB starts with wIdent 0xA5EC; flag bit 0x0100 at offset 0x0A.
    # Excel: a BIFF8 BOF record (09 08 10 00) followed by FILEPASS (2F 00).
    for ($at = 512; $at + 32 -le $Bytes.Length; $at += 512) {
        if ($Bytes[$at] -eq 0xEC -and $Bytes[$at + 1] -eq 0xA5) { return (($Bytes[$at + 0x0B] -band 0x01) -ne 0) }
        if ($Bytes[$at] -eq 0x09 -and $Bytes[$at + 1] -eq 0x08 -and $Bytes[$at + 2] -eq 0x10 -and $Bytes[$at + 3] -eq 0x00) {
            $next = $at + 4 + 16
            return ($next + 2 -le $Bytes.Length -and $Bytes[$next] -eq 0x2F -and $Bytes[$next + 1] -eq 0x00)
        }
    }
    return $false
}

function Get-OleText {
    <#
    .SYNOPSIS
      The text runs of an old Office or Outlook file (an .xls with a note: its numbers are not read); a reason when
      the file is not an OLE file, is encrypted, or holds no text run.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER Extension
      The file extension (.doc, .xls, .ppt or .msg).
    .OUTPUTS
      [hashtable] Text, Reason, Note.
    #>
    param([byte[]]$Bytes, [string]$Extension)
    if (-not (Test-OleSignature $Bytes)) {
        return @{ Text = ''; Reason = 'Not a valid old Office or Outlook file (damaged, or another type with its name): check by hand'; Note = '' }
    }
    $text = $script:OleLatin1.GetString($Bytes)
    if (Test-OleEncrypted $text $Bytes) {
        return @{ Text = ''; Reason = 'Encrypted, password-protected or labelled with encryption: check by hand'; Note = '' }
    }
    $runs = New-Object System.Collections.Generic.List[string]
    foreach ($run in [regex]::Matches($text, $script:Utf16RunPattern)) {
        $runs.Add([System.Text.Encoding]::Unicode.GetString($script:OleLatin1.GetBytes($run.Value)))
    }
    foreach ($run in [regex]::Matches($text, $script:AnsiRunPattern)) { $runs.Add($run.Value) }
    if ($runs.Count -eq 0) { return @{ Text = ''; Reason = 'No readable text in this old Office or Outlook file: check by hand'; Note = '' } }
    $note = ''; if ($Extension -eq '.xls') { $note = $script:XlsNote }
    return @{ Text = ($runs -join "`n"); Reason = ''; Note = $note }
}

# ==== Module ContainerText.psm1 ====
<#
ContainerText.psm1 - unpacks files that hold other files: zip archives and e-mails (.eml), in memory.

  Purpose:     The parts of a .zip (each file inside) and of an .eml (its headers and text, and each attachment),
               so TextExtract.psm1 reads each part with its own reader and the report names it as
               "archive.zip/inner/file" (the PAM, 2026-11-20: round 3 readers). Limits on the number of files and
               the unpacked size guard against zip bombs; a part over a limit gets a reason instead of its bytes.
  Inputs:      the file bytes.
  Outputs:     parts: Name, and Bytes (a file to read), Text (text already read) or Reason (not read, and why).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Add-Type -AssemblyName System.IO.Compression

# Zip limits: files read from one archive, the size of one unpacked file, and of all of them together.
$script:MaxArchiveEntries = 2000
$script:MaxEntryBytes = 25MB
$script:MaxArchiveBytes = 200MB

# An e-mail's MIME parts nest at most this deep (a hostile message cannot recurse without end).
$script:MaxMimeDepth = 5

# Bytes as characters one to one, so a message can be cut with string functions and decoded part by part.
$script:MimeLatin1 = [System.Text.Encoding]::GetEncoding(28591)

function New-Part {
    # One part of a container: a file to read (Bytes), text already read (Text), or a reason it was not read.
    # $Text is not typed: [string] would turn "no text" ($null) into '' and a file would look like empty text.
    param([string]$Name, [byte[]]$Bytes = $null, $Text = $null, [string]$Reason = '')
    [pscustomobject]@{ Name = $Name; Bytes = $Bytes; Text = $Text; Reason = $Reason }
}

function Read-LimitedStream {
    <#
    .SYNOPSIS
      Reads a stream into memory up to a limit; $null when it holds more (a zip can state a false size).
    .PARAMETER Stream
      The open stream (disposed here).
    .PARAMETER Limit
      The most bytes to accept.
    .OUTPUTS
      [byte[]] or $null.
    #>
    param($Stream, [long]$Limit)
    $memory = New-Object System.IO.MemoryStream
    $buffer = New-Object byte[] 65536
    try {
        while ($true) {
            $read = $Stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            if ($memory.Length + $read -gt $Limit) { return $null }
            $memory.Write($buffer, 0, $read)
        }
        return , $memory.ToArray()
    }
    finally { $Stream.Dispose(); $memory.Dispose() }
}

function Get-ArchiveParts {
    <#
    .SYNOPSIS
      The files inside a zip archive, each as a part to read. Folders are skipped; a file that is too large, past
      the file limit or cannot be unpacked (encrypted, damaged) is a part with a reason.
    .PARAMETER Bytes
      The archive bytes.
    .OUTPUTS
      [pscustomobject] parts (New-Part).
    #>
    param([byte[]]$Bytes)
    $memory = New-Object System.IO.MemoryStream(, $Bytes)
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $count = 0
        $total = 0L
        foreach ($entry in $zip.Entries) {
            if (-not $entry.Name) { continue }   # a folder
            $count++
            if ($count -gt $script:MaxArchiveEntries) {
                New-Part '(more files)' -Reason "More than $script:MaxArchiveEntries files in this archive: the rest is not scanned: check by hand"
                break
            }
            if ($entry.Length -gt $script:MaxEntryBytes -or $total + $entry.Length -gt $script:MaxArchiveBytes) {
                New-Part $entry.FullName -Reason 'Too large to unpack in memory: check by hand'
                continue
            }
            $data = $null
            try { $data = Read-LimitedStream $entry.Open() $script:MaxEntryBytes }
            catch { New-Part $entry.FullName -Reason 'Could not unpack this file (encrypted or damaged?): check by hand'; continue }
            if ($null -eq $data) { New-Part $entry.FullName -Reason 'Too large to unpack in memory: check by hand'; continue }
            $total += $data.Length
            New-Part $entry.FullName -Bytes $data
        }
    }
    finally { $zip.Dispose(); $memory.Dispose() }
}

function Split-MimeEntity {
    <#
    .SYNOPSIS
      A MIME entity (a message or one part) as its headers (unfolded, names in lower case) and its body text.
    .PARAMETER Text
      The entity's text (Latin-1 characters, one per byte).
    .OUTPUTS
      [hashtable] Headers, Body.
    #>
    param([string]$Text)
    $cut = [regex]::Match($Text, '\r?\n\r?\n')
    $head = $Text; $body = ''
    if ($cut.Success) { $head = $Text.Substring(0, $cut.Index); $body = $Text.Substring($cut.Index + $cut.Length) }
    $headers = @{}
    foreach ($line in (($head -replace '\r?\n[ \t]+', ' ') -split '\r?\n')) {
        $colon = $line.IndexOf(':')
        if ($colon -gt 0) { $headers[$line.Substring(0, $colon).Trim().ToLower()] = $line.Substring($colon + 1).Trim() }
    }
    return @{ Headers = $headers; Body = $body }
}

function Get-HeaderParameter {
    # One parameter of a header value (boundary, charset, name, filename), quoted or not; empty when absent.
    param([string]$Value, [string]$Name)
    $match = [regex]::Match($Value, '(?i)\b' + $Name + '\*?\s*=\s*(?:"([^"]*)"|([^;\s]+))')
    if (-not $match.Success) { return '' }
    return ($match.Groups[1].Value + $match.Groups[2].Value)
}

function ConvertFrom-TransferEncoding {
    <#
    .SYNOPSIS
      The bytes of a MIME body: base64 and quoted-printable undone; 7bit, 8bit and binary taken as they are.
    .PARAMETER Body
      The body text (Latin-1 characters, one per byte).
    .PARAMETER Encoding
      The Content-Transfer-Encoding value.
    .OUTPUTS
      [byte[]]
    #>
    param([string]$Body, [string]$Encoding)
    switch -Regex ($Encoding.ToLower()) {
        'base64' {
            $clean = $Body -replace '[^A-Za-z0-9+/=]', ''
            $clean = $clean.Substring(0, $clean.Length - ($clean.Length % 4))
            return , [Convert]::FromBase64String($clean)
        }
        'quoted-printable' {
            $soft = $Body -replace '=\r?\n', ''
            $evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) }
            return , $script:MimeLatin1.GetBytes([regex]::Replace($soft, '=([0-9A-Fa-f]{2})', $evaluator))
        }
        default { return , $script:MimeLatin1.GetBytes($Body) }
    }
}

function ConvertFrom-EncodedWords {
    <#
    .SYNOPSIS
      A header with RFC 2047 encoded words (=?utf-8?B?...?= or =?utf-8?Q?...?=) as plain text; an unknown
      character set falls back to Latin-1.
    .PARAMETER Value
      The header value.
    .OUTPUTS
      [string]
    #>
    param([string]$Value)
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $bytes = $null
        if ($m.Groups[2].Value -ieq 'B') { $bytes = ConvertFrom-TransferEncoding $m.Groups[3].Value 'base64' }
        else { $bytes = ConvertFrom-TransferEncoding ($m.Groups[3].Value -replace '_', ' ') 'quoted-printable' }
        try { return [System.Text.Encoding]::GetEncoding($m.Groups[1].Value).GetString($bytes) } catch { return $script:MimeLatin1.GetString($bytes) }
    }
    return [regex]::Replace($Value, '=\?([^?]+)\?([BbQq])\?([^?]*)\?=', $evaluator)
}

function Get-MimeParts {
    <#
    .SYNOPSIS
      The text and the attachments of one MIME entity, depth first: a text part becomes text, an attachment (a file
      name, or a type that is not text) becomes a file to read.
    .PARAMETER Text
      The entity (Latin-1 characters, one per byte).
    .PARAMETER Depth
      How deep this entity sits (multipart inside multipart).
    .OUTPUTS
      [pscustomobject] parts (New-Part).
    #>
    param([string]$Text, [int]$Depth)
    $entity = Split-MimeEntity $Text
    $type = [string]$entity.Headers['content-type']
    $encoding = [string]$entity.Headers['content-transfer-encoding']
    if ($type -match '(?i)^\s*multipart/') {
        $boundary = Get-HeaderParameter $type 'boundary'
        if (-not $boundary -or $Depth -ge $script:MaxMimeDepth) { return }
        $pieces = $entity.Body -split ('(?m)^--' + [regex]::Escape($boundary) + '(?:--)?[ \t]*\r?$')
        for ($i = 1; $i -lt $pieces.Count; $i++) {   # piece 0 is the preamble
            $piece = $pieces[$i] -replace '^\r?\n', ''
            if ($piece.Trim()) { Get-MimeParts $piece ($Depth + 1) }
        }
        return
    }
    $file = Get-HeaderParameter ([string]$entity.Headers['content-disposition']) 'filename'
    if (-not $file) { $file = Get-HeaderParameter $type 'name' }
    $bytes = ConvertFrom-TransferEncoding $entity.Body $encoding
    if (-not $file -and ($type -eq '' -or $type -match '(?i)^\s*text/')) {
        $charset = Get-HeaderParameter $type 'charset'
        $decoded = $null
        try { if ($charset) { $decoded = [System.Text.Encoding]::GetEncoding($charset).GetString($bytes) } } catch { Write-Verbose 'Unknown charset: read as Latin-1.' }
        if ($null -eq $decoded) { $decoded = $script:MimeLatin1.GetString($bytes) }
        $name = '(message text)'; if ($type -match '(?i)text/html') { $name = '(message text).html' }
        New-Part $name -Text $decoded
        return
    }
    if (-not $file) { $file = 'attachment' }
    New-Part (ConvertFrom-EncodedWords $file) -Bytes $bytes
}

function Get-EmailParts {
    <#
    .SYNOPSIS
      The parts of an e-mail (.eml, RFC 5322 with MIME): its main headers as text (Subject, From, To, Cc, Date), its
      text bodies, and each attachment as a file to read.
    .PARAMETER Bytes
      The message bytes.
    .OUTPUTS
      [pscustomobject] parts (New-Part).
    #>
    param([byte[]]$Bytes)
    $text = $script:MimeLatin1.GetString($Bytes)
    $entity = Split-MimeEntity $text
    $lines = foreach ($name in @('subject', 'from', 'to', 'cc', 'date')) {
        if ($entity.Headers.ContainsKey($name)) { (Get-Culture).TextInfo.ToTitleCase($name) + ': ' + (ConvertFrom-EncodedWords $entity.Headers[$name]) }
    }
    New-Part '(message headers)' -Text (@($lines) -join "`n")
    Get-MimeParts $text 0
}

# ==== Module PdfCore.psm1 ====
<#
PdfCore.psm1 - the base of the PDF reader: its time and size limits, stream filters and ToUnicode font maps.

  Purpose:     So one damaged or hostile PDF cannot stall the scan or fill the memory (fresh reviews, 2026-11-20/21):
               a time limit per PDF and per regex search, caps on unpacked size, and "check by hand" notes for any
               part not read. Also the stream filters (Flate, ASCII85, ASCIIHex) and the ToUnicode maps of fonts.
               PdfText.psm1 builds the page text on top of it.
  Inputs:      PDF stream bytes and dictionaries, as Latin-1 text.
  Outputs:     decoded streams, font maps, the limit checks and the notes of one PDF (Start-/Stop-PdfBudget).
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Add-Type -AssemblyName System.IO.Compression

# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)

# A ToUnicode map uses at most this many ranges (real maps use far fewer); ranges are looked up per glyph, not expanded.
# A map with more is read in part, and the PDF gets a "check by hand" note (Add-PdfNote).
$script:MaxCMapRanges = 256

# A range's target is a few UTF-16 code units in hex (<0041>, or a ligature <00660069>); a longer one is not used.
$script:MaxRangeTarget = 66

# Limits, so one damaged or hostile PDF cannot stall the scan (fresh review, 2026-11-20): the time for one PDF, the
# time for one regex search, the size one stream may inflate to, and the inflated size of one PDF in all.
# 120 s: a 76-page table-heavy PDF took 18 s on PowerShell 7 (fresh review, 2026-11-21); Windows PowerShell 5.1 is slower.
$script:PdfSecondsLimit = 120
$script:PdfRegexTimeout = [TimeSpan]::FromSeconds(2)
$script:MaxStreamBytes = 16MB
$script:MaxInflatedBytes = 64MB
$script:PdfBudget = @{ Clock = $null; Inflated = 0L; Notes = $null }

# Streams that hold no page text and are not inflated: images and embedded font programs.
$script:NotTextStream = '/Subtype\s*/(?:Image|Type1C|CIDFontType0C|OpenType)\b|/Length[123]\s|/Type\s*/(?:EmbeddedFile|XRef|Metadata)\b'

function Start-PdfBudget {
    # Starts the clock, the unpack count and the notes of one PDF (Get-PdfText).
    $script:PdfBudget = @{ Clock = [Diagnostics.Stopwatch]::StartNew(); Inflated = 0L; Notes = (New-Object System.Collections.Generic.List[string]) }
}

function Stop-PdfBudget {
    # Ends one PDF's budget and returns its notes (parts not read).
    $notes = $script:PdfBudget.Notes
    $script:PdfBudget = @{ Clock = $null; Inflated = 0L; Notes = $null }
    if ($notes) { return $notes.ToArray() }
}

function ConvertFrom-HexText {
    <#
    .SYNOPSIS
      Hex digits (ASCIIHexDecode data, other characters ignored) as their bytes in Latin-1 characters.
    .PARAMETER Text
      The hex text.
    #>
    param([string]$Text)
    $hex = $Text -replace '[^0-9A-Fa-f]', ''
    if ($hex.Length % 2) { $hex += '0' }
    return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
        if (($i % 200000) -eq 0) { Assert-PdfTime }
        [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
    })
}

function Get-PdfRegex {
    # A regex with a time limit per search ($PdfRegexTimeout); a search over it throws instead of running on.
    param([string]$Pattern)
    return New-Object System.Text.RegularExpressions.Regex($Pattern, [System.Text.RegularExpressions.RegexOptions]::None, $script:PdfRegexTimeout)
}

function Assert-PdfTime {
    # Throws when this PDF has taken longer than $PdfSecondsLimit (Get-PdfText then reports it as too slow to read).
    if ($script:PdfBudget.Clock -and $script:PdfBudget.Clock.Elapsed.TotalSeconds -gt $script:PdfSecondsLimit) {
        throw (New-Object System.TimeoutException 'PDF time limit')
    }
}

function Add-PdfNote {
    # Records a part of this PDF that was not read (a cap, damage, attachments): the report then shows a "check by
    # hand" row next to the hits, so no part is left out without a word (fresh review, 2026-11-21).
    param([string]$Note)
    if (-not $script:PdfBudget.ContainsKey('Notes') -or $null -eq $script:PdfBudget.Notes) { return }   # outside Get-PdfText
    if (-not $script:PdfBudget.Notes.Contains($Note)) { $script:PdfBudget.Notes.Add($Note) }
}

function Expand-PdfStream {
    <#
    .SYNOPSIS
      A FlateDecode stream (zlib: a 2-byte header, then deflate) inflated; what was read before an error is kept.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    #>
    param([string]$Raw)
    $bytes = $script:Latin1.GetBytes($Raw)
    $skip = 0; if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x78) { $skip = 2 }
    $source = New-Object System.IO.MemoryStream -ArgumentList $bytes, $skip, ($bytes.Length - $skip)
    $inflate = New-Object System.IO.Compression.DeflateStream($source, [System.IO.Compression.CompressionMode]::Decompress)
    $output = New-Object System.IO.MemoryStream
    $buffer = New-Object byte[] 65536
    try {
        # Inflate in steps and stop at the caps: a small stream can unpack to gigabytes.
        while ($true) {
            if ($output.Length -ge $script:MaxStreamBytes -or $script:PdfBudget.Inflated -ge $script:MaxInflatedBytes) {
                # At a cap: a note only when there is more to unpack (a stream of exactly the cap is whole).
                if ($inflate.Read($buffer, 0, 1) -gt 0) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand' }
                break
            }
            $read = $inflate.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $output.Write($buffer, 0, $read)
            $script:PdfBudget.Inflated += $read
        }
    }
    catch { Add-PdfNote 'Part of this PDF is damaged: check by hand' }   # the text read before the damage is kept
    finally { $inflate.Dispose(); $source.Dispose() }
    return $script:Latin1.GetString($output.ToArray())
}

function Test-PdfUnpackCap {
    # True when a stream being decoded by hand (LZW, RunLength) reached the cap of one stream or of the whole PDF.
    param([long]$Length)
    return ($Length -ge $script:MaxStreamBytes -or $script:PdfBudget.Inflated + $Length -ge $script:MaxInflatedBytes)
}

function Close-PdfUnpack {
    <#
    .SYNOPSIS
      Ends a stream decoded by hand: its size counts toward the PDF's unpack total (as Flate output does), and a
      stream stopped at a cap gets a "check by hand" note.
    .PARAMETER Out
      The StringBuilder with the decoded bytes.
    .PARAMETER Whole
      The stream ended by itself (not at a cap).
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param($Out, [switch]$Whole)
    $script:PdfBudget.Inflated += $Out.Length
    if (-not $Whole) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand' }
    return $Out.ToString()
}

function Expand-LzwStream {
    <#
    .SYNOPSIS
      An LZWDecode stream (PDF and TIFF LZW: 9 to 12 bit codes, most significant bit first, 256 = clear table,
      257 = end) decoded, up to $MaxStreamBytes; damaged data ends the stream with a note.
    .PARAMETER Raw
      The stream bytes as Latin-1 text.
    .PARAMETER EarlyChange
      True (the default): the code width grows one code early, as the PDF default /EarlyChange 1 says.
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param([string]$Raw, [bool]$EarlyChange = $true)
    $bytes = $script:Latin1.GetBytes($Raw)
    $out = New-Object System.Text.StringBuilder
    $table = New-Object System.Collections.Generic.List[string]
    # Empties the code table back to its 258 starting entries (bytes 0-255, clear, end).
    $reset = { $table.Clear(); for ($c = 0; $c -lt 256; $c++) { $table.Add([string][char]$c) }; $table.Add(''); $table.Add('') }
    & $reset
    $width = 9; $bits = 0L; $count = 0; $previous = $null; $n = 0
    foreach ($byte in $bytes) {
        $bits = ($bits -shl 8) -bor $byte; $count += 8
        while ($count -ge $width) {
            if ((++$n % 10000) -eq 0) { Assert-PdfTime }
            $count -= $width
            $code = [int](($bits -shr $count) -band ((1 -shl $width) - 1))
            $bits = $bits -band ((1L -shl $count) - 1)
            if ($code -eq 256) { & $reset; $width = 9; $previous = $null; continue }
            if ($code -eq 257) { return (Close-PdfUnpack $out -Whole) }
            if ($code -lt $table.Count) { $entry = $table[$code] }
            elseif ($code -eq $table.Count -and $null -ne $previous) { $entry = $previous + $previous[0] }
            else { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return (Close-PdfUnpack $out -Whole) }
            [void]$out.Append($entry)
            if (Test-PdfUnpackCap $out.Length) { return (Close-PdfUnpack $out) }
            if ($null -ne $previous -and $table.Count -lt 4096) { $table.Add($previous + $entry[0]) }
            $previous = $entry
            $limit = $table.Count; if ($EarlyChange) { $limit++ }
            if ($limit -ge (1 -shl $width) -and $width -lt 12) { $width++ }
        }
    }
    return (Close-PdfUnpack $out -Whole)
}

function Expand-RunLengthStream {
    <#
    .SYNOPSIS
      A RunLengthDecode stream: a length byte 0-127 copies the next 1-128 bytes, 129-255 repeats the next byte
      2-128 times, 128 ends it; up to $MaxStreamBytes.
    .PARAMETER Raw
      The stream bytes as Latin-1 text.
    .OUTPUTS
      [string] the decoded bytes as Latin-1 text.
    #>
    param([string]$Raw)
    $out = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Raw.Length) {
        if (($out.Length % 65536) -eq 0) { Assert-PdfTime }
        $length = [int]$Raw[$i]
        if ($length -eq 128) { break }
        if ($length -lt 128) { $take = [Math]::Min($length + 1, $Raw.Length - $i - 1); [void]$out.Append($Raw, $i + 1, $take); $i += 1 + $take }
        elseif ($i + 1 -lt $Raw.Length) { [void]$out.Append($Raw[$i + 1], 257 - $length); $i += 2 }
        else { break }
        if (Test-PdfUnpackCap $out.Length) { return (Close-PdfUnpack $out) }
    }
    return (Close-PdfUnpack $out -Whole)
}

function Read-PdfStream {
    <#
    .SYNOPSIS
      A stream's data with its filters undone in order: FlateDecode, ASCII85Decode, ASCIIHexDecode, LZWDecode and
      RunLengthDecode (often chained, such as [/ASCII85Decode /FlateDecode]). Images and fonts are not text: $null.
      Any other filter: $null and a "check by hand" note, never a silent miss.
    .PARAMETER Dict
      The stream's dictionary.
    .PARAMETER Raw
      The stream's bytes as Latin-1 text.
    .OUTPUTS
      [string] or $null.
    #>
    param([string]$Dict, [string]$Raw)
    if ($Dict -match '/Type\s*/EmbeddedFile\b') { Add-PdfNote 'Files attached inside this PDF are not scanned: check by hand' }
    if ($Dict -match $script:NotTextStream) { return $null }
    $filter = [regex]::Match($Dict, '/Filter\s*(\[[^\]]*\]|/\w+)')
    $data = $Raw
    foreach ($name in @([regex]::Matches($filter.Groups[1].Value, '/(\w+)') | ForEach-Object { $_.Groups[1].Value })) {
        switch ($name) {
            'FlateDecode' { $data = Expand-PdfStream $data }
            'ASCII85Decode' { $data = ConvertFrom-Ascii85 $data }
            'ASCIIHexDecode' { $data = ConvertFrom-HexText $data }
            'LZWDecode' { $data = Expand-LzwStream $data ($Dict -notmatch '/EarlyChange\s+0\b') }
            'RunLengthDecode' { $data = Expand-RunLengthStream $data }
            default { Add-PdfNote 'Part of this PDF uses a compression this reader cannot open: check by hand'; return $null }
        }
    }
    return $data
}

function ConvertFrom-Ascii85 {
    <#
    .SYNOPSIS
      ASCII85 text (PDF ASCII85Decode, ending in ~>) to its bytes as Latin-1 text; z = four zero bytes.
    .PARAMETER Text
      The text.
    #>
    param([string]$Text)
    $end = $Text.IndexOf('~>'); if ($end -ge 0) { $Text = $Text.Substring(0, $end) }
    $Text = ($Text -replace '\s', '') -creplace 'z', '!!!!!'   # -creplace: only a small z (a capital Z is data)
    $out = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Text.Length; $i += 5) {
        if (($i % 100000) -eq 0) { Assert-PdfTime }
        $group = $Text.Substring($i, [Math]::Min(5, $Text.Length - $i))
        $count = $group.Length - 1
        if ($count -lt 1) { break }
        $group = $group.PadRight(5, 'u')
        [long]$value = 0
        foreach ($c in $group.ToCharArray()) { $value = $value * 85 + ([int]$c - 33) }
        for ($k = 0; $k -lt $count; $k++) { [void]$out.Append([char](($value -shr (24 - 8 * $k)) -band 0xFF)) }
    }
    return $out.ToString()
}

function ConvertFrom-Utf16Hex {
    # The text of a ToUnicode target: UTF-16BE given as hex.
    param([string]$Hex)
    $bytes = [byte[]]@(for ($i = 0; $i + 1 -lt $Hex.Length; $i += 2) { [Convert]::ToByte($Hex.Substring($i, 2), 16) })
    return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
}

function Get-PdfBlocks {
    <#
    .SYNOPSIS
      The text between each begin word and the end word after it (bfchar or bfrange blocks), found in one linear pass:
      a block without its end word runs to the next begin word, so a damaged map cannot make the search quadratic.
    .PARAMETER Text
      The CMap stream.
    .PARAMETER Begin
      The begin word, for example 'beginbfchar'.
    .PARAMETER End
      The end word, for example 'endbfchar'.
    .OUTPUTS
      [string] one per block.
    #>
    param([string]$Text, [string]$Begin, [string]$End)
    $starts = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($Begin))).Matches($Text)) { $m.Index + $m.Length })
    $stops = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($End))).Matches($Text)) { $m.Index })
    $e = 0
    for ($b = 0; $b -lt $starts.Count; $b++) {
        $stop = $Text.Length; if ($b + 1 -lt $starts.Count) { $stop = $starts[$b + 1] - $Begin.Length }
        while ($e -lt $stops.Count -and $stops[$e] -lt $starts[$b]) { $e++ }
        if ($e -lt $stops.Count -and $stops[$e] -lt $stop) { $stop = $stops[$e] }
        $Text.Substring($starts[$b], $stop - $starts[$b])
    }
}

function Get-PdfCMap {
    <#
    .SYNOPSIS
      A font's ToUnicode map: glyph code (hex, upper case) -> text from its bfchar entries and listed ranges, plus
      its plain ranges (looked up by ConvertFrom-PdfGlyphs, not expanded).
    .PARAMETER Text
      The CMap stream.
    .OUTPUTS
      [hashtable] Bytes (bytes per code), Map, Ranges.
    #>
    param([string]$Text)
    $map = @{}
    $ranges = New-Object System.Collections.Generic.List[object]
    $bytes = 1
    $pairRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>'
    $rangeRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])'
    $listRegex = Get-PdfRegex '<([0-9A-Fa-f]*)>'
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfchar' 'endbfchar')) {
        Assert-PdfTime
        $n = 0
        foreach ($pair in $pairRegex.Matches($block)) {
            if ((++$n % 1000) -eq 0) { Assert-PdfTime }
            $map[$pair.Groups[1].Value.ToUpper()] = ConvertFrom-Utf16Hex $pair.Groups[2].Value
            $bytes = [int]($pair.Groups[1].Value.Length / 2)
        }
    }
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfrange' 'endbfrange')) {
        Assert-PdfTime
        foreach ($range in $rangeRegex.Matches($block)) {
            $width = $range.Groups[1].Value.Length
            $bytes = [int]($width / 2)
            $low = [Convert]::ToInt64($range.Groups[1].Value, 16)
            $high = [Convert]::ToInt64($range.Groups[2].Value, 16)
            $target = $range.Groups[3].Value
            if ($high -lt $low) { continue }
            if ($target.StartsWith('[')) {
                # A listed range names each target: as many entries as the map's own text holds, so no expansion cost.
                $code = $low
                foreach ($item in $listRegex.Matches($target)) {
                    if ($code -gt $high) { break }
                    if ((($code - $low) % 1000) -eq 999) { Assert-PdfTime }
                    $map[$code.ToString('X' + $width)] = ConvertFrom-Utf16Hex $item.Groups[1].Value
                    $code++
                }
            }
            elseif ($ranges.Count -ge $script:MaxCMapRanges) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand' }
            elseif ($target.Length -gt $script:MaxRangeTarget) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand' }
            elseif ($target.Length -ge 6) {
                $hex = $target.Substring(1, $target.Length - 2)
                $ranges.Add(@{ Low = $low; High = $high; Prefix = $hex.Substring(0, $hex.Length - 4); Start = [Convert]::ToInt32($hex.Substring($hex.Length - 4), 16) })
            }
        }
    }
    return @{ Bytes = $bytes; Map = $map; Ranges = $ranges }
}

# ==== Module PdfObjects.psm1 ====
<#
PdfObjects.psm1 - the objects of a PDF and the text kept in its dictionaries, with .NET only (the PAM, 2026-11-20).

  Purpose:     Finds a PDF's objects (also in object streams and in incremental updates, older versions kept), decodes
               PDF strings, and reads the text kept in dictionaries: form-field values, comments, file info and XFA
               form data. A string it cannot read throws (Assert-PdfGap), so a PDF is never read in part without a word.
               PdfText.psm1 reads the page text on top of it.
  Inputs:      the PDF bytes as Latin-1 text.
  Outputs:     objects (Get-PdfObjects), decoded strings, dictionary and XFA text lines.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:ObjLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary). A literal holds at most
# 10,000 parts, so an unclosed bracket fails after a bounded scan. A "(" that no string token takes means a string
# was not read: the PDF is then NOT SCANNED, never read in part without a word (Assert-PdfGap).
$script:PdfStringPattern = '\((?>[^()\\]+|\\[\s\S]|\((?<d>)|\)(?<-d>)){0,10000}(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The message of the error a string not read throws (Get-PdfText turns it into "part of this PDF could not be read").
$script:PdfStringError = 'PDF string not read'

# The escapes of a literal string (\n, \053, a line break ...) and how each is undone, built once: building a
# PowerShell delegate per string cost about 450 microseconds (speed, 2026-11-21).
$script:EscapePattern = New-Object System.Text.RegularExpressions.Regex('\\(?:[0-7]{1,3}|\r?\n|.)')
$script:UndoEscape = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    $code = $m.Value.Substring(1)
    if ($code -match '^[0-7]+$') { return [string][char]([Convert]::ToInt32($code, 8) -band 0xFF) }
    switch -CaseSensitive ($code) { 'n' { "`n" } 'r' { "`r" } 't' { "`t" } 'b' { ' ' } 'f' { ' ' } default { $code.Trim("`r", "`n") } }
}

# An object stream (PDF 1.5) is unpacked only within these limits: its header of object numbers and offsets, and its
# object count (real files hold a few hundred objects per stream). More gives a "check by hand" note.
$script:MaxPackedHeader = 1MB

# Earlier versions of objects defined again further on (incremental updates): their text is still read.
$script:PdfOldObjects = New-Object System.Collections.Generic.List[object]
$script:MaxPackedObjects = 100000

# Dictionary keys whose string values are text a person wrote (Get-PdfDictText) -> the label put before the value:
# a form field's value (labelled with the field's name), a comment or note and its rich text, the file's info.
$script:PdfTextKeys = [ordered]@{ 'V' = ''; 'Contents' = 'Comment'; 'RC' = 'Comment'; 'Author' = 'Author'; 'Title' = 'Title'
    'Subject' = 'Subject'; 'Keywords' = 'Keywords' }

function Get-PdfStringPattern {
    # The pattern of one PDF string token (PdfText.psm1 builds its token pattern on it).
    return $script:PdfStringPattern
}

function Get-PdfStringError {
    # The message of the error a string not read throws.
    return $script:PdfStringError
}

function Reset-PdfOldObjects {
    # Forgets the older object versions of the last PDF (Get-PdfText, before each PDF).
    $script:PdfOldObjects.Clear()
}

function Get-PdfOldObjects {
    # The older versions of objects defined again further on in this PDF (an incremental update).
    return , $script:PdfOldObjects.ToArray()
}

function Get-PdfObjects {
    <#
    .SYNOPSIS
      The objects of a PDF by number: each its dictionary text and its stream (inflated when FlateDecode; $null for
      other filters, such as images). Objects packed in object streams (PDF 1.5) are unpacked too.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary] number -> @{ Dict; Stream; Pos; Image (JPEG data or $null) }, in file order.
    #>
    param([string]$Text)
    $objects = New-Object System.Collections.Specialized.OrderedDictionary
    # Each object runs from its header to the first "endobj" after it, or to the next header when "endobj" is missing.
    # Two linear passes, so a damaged file without "endobj" cannot make the search quadratic.
    $heads = @((Get-PdfRegex '(\d+)\s+\d+\s+obj\b').Matches($Text))
    $ends = @((Get-PdfRegex '\bendobj\b').Matches($Text) | ForEach-Object { $_.Index })
    $e = 0
    for ($h = 0; $h -lt $heads.Count; $h++) {
        Assert-PdfTime
        $start = $heads[$h].Index + $heads[$h].Length
        $stop = $Text.Length; if ($h + 1 -lt $heads.Count) { $stop = $heads[$h + 1].Index }
        while ($e -lt $ends.Count -and $ends[$e] -lt $start) { $e++ }
        if ($e -lt $ends.Count -and $ends[$e] -lt $stop) { $stop = $ends[$e] }
        $body = $Text.Substring($start, $stop - $start)
        $stream = $null
        $parts = Split-PdfStreamBody $body
        $image = $null
        if ($parts) {
            $body = $parts.Dict; $stream = Read-PdfStream $body $parts.Data
            # A JPEG page image (a scanned page): kept as it is, for OCR when the PDF has no text (Get-ScanPagesText).
            if ($body -match '/Subtype\s*/Image\b' -and $body -match '/Filter\s*\[?\s*/DCTDecode\s*\]?') { $image = $parts.Data }
        }
        $number = $heads[$h].Groups[1].Value
        # A file saved with added edits (an incremental update) defines an object again further on: the later one is
        # the one in use; the earlier one's text is still in the file, so it is kept (Get-PdfDictText, Get-PdfLooseText).
        if ($objects.Contains($number)) { $script:PdfOldObjects.Add($objects[$number]) }
        $objects[$number] = @{ Dict = $body; Stream = $stream; Pos = $heads[$h].Index; Image = $image }
    }
    foreach ($packed in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/ObjStm' -and $_.Stream })) { Add-PackedObjects $objects $packed }
    return $objects
}

function Split-PdfStreamBody {
    <#
    .SYNOPSIS
      An object body with a stream, split into its dictionary and its data: from the "stream" keyword (then CR LF, LF
      or CR) to the last "endstream". Plain string searches, so a stream of any size is split in one linear pass.
      $null for an object without a stream; a stream without its end gives a note.
    .PARAMETER Body
      The object's text from its header to endobj.
    .OUTPUTS
      [hashtable] Dict, Data; or $null.
    #>
    param([string]$Body)
    $at = -1
    for ($from = 0; ($at = $Body.IndexOf('stream', $from, [StringComparison]::Ordinal)) -ge 0; $from = $at + 6) {
        $after = $at + 6
        $word = $at -eq 0 -or -not [char]::IsLetter($Body[$at - 1])   # not the end of "endstream"
        if ($word -and $after -lt $Body.Length -and ($Body[$after] -eq "`n" -or $Body[$after] -eq "`r")) { break }
    }
    if ($at -lt 0) { return $null }
    $start = $at + 6
    if ($Body[$start] -eq "`r") { $start++ }
    if ($start -lt $Body.Length -and $Body[$start] -eq "`n") { $start++ }
    $end = $Body.LastIndexOf('endstream', [StringComparison]::Ordinal)
    if ($end -lt $start) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return $null }   # cut short
    $stop = $end
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`n") { $stop-- }
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`r") { $stop-- }
    return @{ Dict = $Body.Substring(0, $at); Data = $Body.Substring($start, $stop - $start) }
}

function Add-PackedObjects {
    <#
    .SYNOPSIS
      Adds the objects of one object stream (/Type /ObjStm: N pairs "number offset", then the objects from /First).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Packed
      One object stream of Get-PdfObjects.
    #>
    param($Objects, $Packed)
    $count = [int][regex]::Match($Packed.Dict, '/N\s+(\d+)').Groups[1].Value
    $first = [int][regex]::Match($Packed.Dict, '/First\s+(\d+)').Groups[1].Value
    if ($first -le 0 -or $first -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return }
    if ($first -gt $script:MaxPackedHeader -or $count -gt $script:MaxPackedObjects) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand'; return }
    $pairs = @($Packed.Stream.Substring(0, $first).Trim() -split '\s+')
    $number = 0
    for ($k = 0; $k -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        if (-not [int]::TryParse($pairs[$k], [ref]$number)) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return }
    }
    for ($k = 0; $k -lt $count -and 2 * $k + 1 -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        $start = $first + [int]$pairs[2 * $k + 1]
        $end = $Packed.Stream.Length
        if (2 * $k + 3 -lt $pairs.Count) { $end = $first + [int]$pairs[2 * $k + 3] }
        if ($start -ge $end -or $end -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; continue }
        $dict = $Packed.Stream.Substring($start, $end - $start)
        $number = $pairs[2 * $k]
        # The later definition in the file wins (an incremental update); the other one's text is kept too.
        if ($Objects.Contains($number) -and $Objects[$number].Pos -gt $Packed.Pos) { $script:PdfOldObjects.Add(@{ Dict = $dict; Stream = $null }); continue }
        if ($Objects.Contains($number)) { $script:PdfOldObjects.Add($Objects[$number]) }
        $Objects[$number] = @{ Dict = $dict; Stream = $null; Pos = $Packed.Pos }
    }
}

function ConvertFrom-PdfString {
    <#
    .SYNOPSIS
      The bytes of one PDF string token, as Latin-1 characters: (literal) with its escapes undone, or <hex>.
    .PARAMETER Token
      One PDF string token.
    #>
    param([string]$Token)
    if ($Token.StartsWith('<')) {
        $hex = $Token.Substring(1, $Token.Length - 2) -replace '\s', ''
        if ($hex.Length % 2) { $hex += '0' }
        return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
            if (($i % 200000) -eq 0) { Assert-PdfTime }
            [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
        })
    }
    $inner = $Token.Substring(1, $Token.Length - 2)
    if ($inner.IndexOf([char]92) -lt 0) { return $inner }   # no escapes (most strings): as it is (speed)
    return $script:EscapePattern.Replace($inner, $script:UndoEscape)
}

function ConvertFrom-PdfTextString {
    # A PDF text string (a form field's name or value): UTF-16 when it starts with the FE FF mark, else Latin-1.
    param([string]$Token)
    $raw = ConvertFrom-PdfString $Token
    if ($raw.Length -ge 2 -and [int]$raw[0] -eq 0xFE -and [int]$raw[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($script:ObjLatin1.GetBytes($raw.Substring(2)))
    }
    return $raw
}

function Assert-PdfGap {
    <#
    .SYNOPSIS
      Throws when the text between two tokens holds a "(": a string there was not read (too long, or damaged), so
      the PDF is reported NOT SCANNED rather than read in part without a word (fresh review, 2026-11-20).
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the gap starts.
    .PARAMETER To
      Where the gap ends.
    #>
    param([string]$Content, [int]$From, [int]$To)
    if ($To -gt $From -and $Content.IndexOf('(', $From, $To - $From) -ge 0) { throw (New-Object System.FormatException $script:PdfStringError) }
}

function Get-PdfFieldName {
    <#
    .SYNOPSIS
      A form field's name (/T) as the label of a value: the last name before the value in the dictionary (a
      dictionary can hold several fields inline), else the first; empty when there is none.
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .PARAMETER Before
      Where the value starts in the dictionary.
    #>
    param([string]$Dict, $NameRegex, [int]$Before)
    $names = @($NameRegex.Matches($Dict))
    if ($names.Count -eq 0) { return '' }
    $pick = $names[0]
    foreach ($name in $names) { if ($name.Index -lt $Before) { $pick = $name } }
    return (ConvertFrom-PdfTextString $pick.Groups[1].Value)
}

function Format-PdfDictLine {
    # One line of dictionary text: "label: text", or the text alone.
    param([string]$Label, [string]$Text)
    if ($Label) { return ($Label + ': ' + $Text) }
    return $Text
}

function Get-PdfDictText {
    <#
    .SYNOPSIS
      Text kept in dictionaries, not on a page: form-field values ("name: value", so labels stay with values),
      comments and sticky notes (/Contents of an annotation, /RC rich text) and the file's info (author, title,
      subject, keywords), in current and older object versions; every value of a key, also several in one
      dictionary. A literal value that is not read throws (Get-PdfText: "part of this PDF could not be read").
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] one line per value.
    #>
    param($Objects)
    $nameRegex = Get-PdfRegex ('/T\s*(' + $script:PdfStringPattern + ')')
    $checks = foreach ($key in $script:PdfTextKeys.Keys) {
        @{ Key = $key; Label = $script:PdfTextKeys[$key]; Value = (Get-PdfRegex ('/' + $key + '\s*(' + $script:PdfStringPattern + ')'))
           Ref = (Get-PdfRegex ('/' + $key + '\s+(\d+)\s+\d+\s+R')); Open = ('/' + $key + '\s*\(') }
    }
    # Older versions too. ToArray(), not @(): @() around an empty generic list fails ("Argument types do not match").
    $all = @($Objects.Values) + $script:PdfOldObjects.ToArray()
    foreach ($obj in $all) {
        Assert-PdfTime
        foreach ($check in $checks) {
            if ($obj.Dict.IndexOf('/' + $check.Key) -ge 0) { Get-PdfKeyLines $obj.Dict $check $Objects $nameRegex }
        }
    }
}

function Get-PdfKeyLines {
    <#
    .SYNOPSIS
      The lines of one key in one dictionary: each string value, a list of strings (a multi-select field), or a
      value held in an object of its own (a string, or a text stream for /V and /RC).
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER Check
      The key, its label and its regexes (Get-PdfDictText).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .OUTPUTS
      [string] lines.
    #>
    param([string]$Dict, [hashtable]$Check, $Objects, $NameRegex)
    $values = @($Check.Value.Matches($Dict))
    $literals = @($values | Where-Object { $_.Groups[1].Value.StartsWith('(') }).Count
    if ([regex]::Matches($Dict, $Check.Open).Count -gt $literals) { throw (New-Object System.FormatException $script:PdfStringError) }   # not read
    foreach ($value in $values) {
        $text = ConvertFrom-PdfTextString $value.Groups[1].Value
        if ($Check.Key -eq 'RC') { $text = $text -replace '<[^>]+>', ' ' }   # rich text: XHTML tags dropped
        $label = $Check.Label
        if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $value.Index }
        Format-PdfDictLine $label $text
    }
    if ($values.Count -gt 0) { return }
    $list = [regex]::Match($Dict, '/' + $Check.Key + '\s*\[([^\]]*)\]')
    if ($list.Success) {
        $items = @(foreach ($item in (Get-PdfRegex $script:PdfStringPattern).Matches($list.Groups[1].Value)) { ConvertFrom-PdfTextString $item.Value })
        $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $list.Index }
        if ($items.Count -gt 0) { Format-PdfDictLine $label ($items -join ', ') }
        return
    }
    $ref = $Check.Ref.Match($Dict)
    if (-not $ref.Success -or -not $Objects.Contains($ref.Groups[1].Value)) { return }
    $held = $Objects[$ref.Groups[1].Value]
    $target = $held.Dict.Trim()
    $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $ref.Index }
    if ($target.StartsWith('(') -or ($target.StartsWith('<') -and -not $target.StartsWith('<<'))) { Format-PdfDictLine $label (ConvertFrom-PdfTextString $target) }
    elseif ($Check.Key -in @('V', 'RC') -and $held.Stream) { Format-PdfDictLine $label $held.Stream }   # a text stream
}

function Get-PdfXfaText {
    <#
    .SYNOPSIS
      The data of an XFA form (/AcroForm /XFA: one stream, or an array of name and stream pairs): XML, read as
      "element: value" lines, so a field name stays with its value (a dynamic XFA form shows only "Please wait..."
      on its page; its data is here).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] lines.
    #>
    param($Objects)
    $pair = Get-PdfRegex '<([\w:.-]+)\b[^<>]*>([^<]+)</\1>'
    foreach ($obj in @($Objects.Values)) {
        $xfa = [regex]::Match($obj.Dict, '/XFA\s*(\[[^\]]*\]|\d+\s+\d+\s+R)')
        if (-not $xfa.Success) { continue }
        foreach ($ref in [regex]::Matches($xfa.Groups[1].Value, '(\d+)\s+\d+\s+R')) {
            Assert-PdfTime
            $part = $Objects[$ref.Groups[1].Value]
            if (-not $part -or -not $part.Stream) { continue }   # a stream not decoded has its note already
            $lines = $pair.Replace($part.Stream, "`n`$1: `$2`n") -replace '<[^>]+>', ' '
            [System.Net.WebUtility]::HtmlDecode($lines)
        }
    }
}

# ==== Module PdfFonts.psm1 ====
<#
PdfFonts.psm1 - the fonts of a PDF: their character maps and their real glyph widths, with .NET only.

  Purpose:     A shown string becomes text through its font's ToUnicode map, and its width comes from the font's own
               widths (/Widths of a simple font, /W of a composite font, built-in widths for the standard Helvetica,
               Times and Courier families), so PdfText.psm1 can tell a word gap from a glyph placed one by one
               (fresh reviews, 2026-11-21: estimated widths split and joined words wrongly). A guess is the last resort.
  Inputs:      the PDF objects (PdfObjects.psm1 Get-PdfObjects) and shown strings.
  Outputs:     font records (CMap, Bytes, Widths, DefaultWidth, Standard), decoded text and widths.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (Get-PdfCMap, ConvertFrom-Utf16Hex, Assert-PdfTime, Add-PdfNote), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>


# Widths of the standard fonts (Adobe AFM, in thousandths of an em) for character codes 32-126; Courier is 600 for all.
# A simple font may leave out /Widths only when it is one of these (ISO 32000-1, 9.6.2.2).
$script:HelveticaWidths = @(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556,
    556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778,
    722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222,
    833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)
$script:TimesWidths = @(250, 333, 408, 500, 500, 833, 778, 180, 333, 333, 500, 564, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500,
    500, 500, 500, 278, 278, 564, 564, 564, 444, 921, 722, 667, 667, 722, 611, 556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722,
    667, 556, 611, 722, 722, 944, 722, 722, 611, 333, 278, 333, 469, 500, 333, 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500, 278,
    778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444, 480, 200, 480, 541)

# A /W array (a composite font's widths) is read for at most this many glyphs (a hostile array cannot fill the memory).
$script:MaxWidthEntries = 65536

# Per font: shown strings decoded and measured before (text and width in ems), up to this many each (speed).
$script:MaxFontCache = 20000
$script:MaxCacheKey = 256   # longer strings are rarely shown twice: not kept (memory)
# Shown hex strings decoded before (speed): short tokens only, emptied when full.
$script:ShownHex = New-Object 'System.Collections.Generic.Dictionary[string,string]'
$script:MaxShownHex = 50000

function Get-PdfFonts {
    <#
    .SYNOPSIS
      The fonts the pages name (/F1 ...) -> their font records, the first of each name in the file (for streams
      that name no resources of their own).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param($Objects)
    $fonts = @{}
    foreach ($obj in $Objects.Values) {
        Assert-PdfTime
        $one = Get-PdfFontTable $obj.Dict $Objects
        foreach ($name in $one.Keys) { if (-not $fonts.ContainsKey($name)) { $fonts[$name] = $one[$name] } }
    }
    return $fonts
}

function Get-PdfFontTable {
    <#
    .SYNOPSIS
      The fonts that one dictionary's /Font resource names (/F1 -> its font record). A page uses its own table, so
      two pages that both name a font /F1 each get theirs; each font and each ToUnicode map is read once.
    .PARAMETER Dict
      A page or resources dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable]
    #>
    param([string]$Dict, $Objects)
    $fonts = @{}
    foreach ($resource in [regex]::Matches($Dict, '/Font\s*(?:<<(?<inline>[^<>]*)>>|(?<ref>\d+)\s+\d+\s+R)')) {
        $names = $resource.Groups['inline'].Value
        if ($resource.Groups['ref'].Success -and $Objects.Contains($resource.Groups['ref'].Value)) { $names = $Objects[$resource.Groups['ref'].Value].Dict }
        foreach ($font in [regex]::Matches($names, '/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R')) {
            $id = $font.Groups[2].Value
            if (-not $Objects.Contains($id)) { continue }
            $entry = $Objects[$id]
            if (-not $entry.Contains('Font')) { $entry['Font'] = New-PdfFontRecord $entry.Dict $Objects }
            $fonts[$font.Groups[1].Value] = $entry['Font']
        }
    }
    return $fonts
}

function New-PdfFontRecord {
    <#
    .SYNOPSIS
      One font's record: its ToUnicode map (CMap; a composite font without one gets an Unreadable marker: its codes
      are glyph numbers, never read as text), its code length, and its widths.
    .PARAMETER Dict
      The font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable] CMap, Bytes, Widths (code -> ems), DefaultWidth (ems, or -1), Standard (a width table, or $null),
      Vertical (vertical writing: each glyph advances 1 em down).
    #>
    param([string]$Dict, $Objects)
    $composite = $Dict -match '/Subtype\s*/Type0\b'
    $record = @{ CMap = $null; Bytes = 1; Widths = @{}; DefaultWidth = -1.0; Standard = $null; Vertical = ($Dict -match '/Encoding\s*/Identity-V\b')
        TextCache = (New-Object 'System.Collections.Generic.Dictionary[string,string]'); EmCache = (New-Object 'System.Collections.Generic.Dictionary[string,double]') }
    if ($composite) { $record.Bytes = 2 }
    $unicode = [regex]::Match($Dict, '/ToUnicode\s+(\d+)\s+\d+\s+R')
    if ($unicode.Success -and $Objects.Contains($unicode.Groups[1].Value) -and $Objects[$unicode.Groups[1].Value].Stream) {
        $cmap = $Objects[$unicode.Groups[1].Value]   # read once, also when many fonts share it
        if (-not $cmap.Contains('CMap')) { $cmap['CMap'] = Get-PdfCMap $cmap.Stream }
        $record.CMap = $cmap['CMap']
        $record.Bytes = [Math]::Max(1, [int]$record.CMap.Bytes)
    }
    elseif ($composite) { $record.CMap = @{ Bytes = 2; Map = @{}; Ranges = @(); Unreadable = $true } }
    if ($composite) { Add-PdfCidWidths $record $Dict $Objects }
    else { Add-PdfSimpleWidths $record $Dict $Objects }
    return $record
}

function Get-PdfArrayText {
    <#
    .SYNOPSIS
      The text of an array value: written in place ([...], one level of inner arrays) or held in an object of its
      own (n 0 R); empty when the key is not there.
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER Key
      The key, without its slash.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([string]$Dict, [string]$Key, $Objects)
    $value = [regex]::Match($Dict, '/' + $Key + '\s*(\[[^\[\]]*(?:\[[^\[\]]*\][^\[\]]*)*\]|(\d+)\s+\d+\s+R)')
    if (-not $value.Success) { return '' }
    if ($value.Groups[2].Success) {
        if ($Objects.Contains($value.Groups[2].Value)) { return $Objects[$value.Groups[2].Value].Dict }
        return ''
    }
    return $value.Groups[1].Value
}

function Add-PdfSimpleWidths {
    <#
    .SYNOPSIS
      A simple font's widths: /FirstChar and /Widths; without /Widths, the standard family's built-in widths.
    .PARAMETER Record
      The font record (filled in place).
    .PARAMETER Dict
      The font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([hashtable]$Record, [string]$Dict, $Objects)
    $first = 0
    $firstMatch = [regex]::Match($Dict, '/FirstChar\s+(\d+)')
    if ($firstMatch.Success) { $first = [int]$firstMatch.Groups[1].Value }
    # Widths are in glyph space: thousandths of an em, or for a Type3 font in units of its /FontMatrix.
    $unit = 0.001
    $matrix = [regex]::Match($Dict, '/FontMatrix\s*\[\s*(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))')
    if ($Dict -match '/Subtype\s*/Type3\b' -and $matrix.Success) {
        # The length of the matrix's first column: the glyph-space unit along the line, also for a rotated matrix.
        $a = [double]::Parse($matrix.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
        $b = [double]::Parse($matrix.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
        $unit = [Math]::Sqrt($a * $a + $b * $b)
    }
    $code = $first
    foreach ($number in [regex]::Matches((Get-PdfArrayText $Dict 'Widths' $Objects), '-?(?:\d+\.?\d*|\.\d+)')) {
        if ($code -gt 255) { break }
        $Record.Widths[$code] = [double]::Parse($number.Value, [Globalization.CultureInfo]::InvariantCulture) * $unit
        $code++
    }
    if ($Record.Widths.Count -gt 0) { return }
    $base = [regex]::Match($Dict, '/BaseFont\s*/([^\s/<>\[\]()]+)').Groups[1].Value
    if ($base -match 'Courier') { $Record.DefaultWidth = 0.6 }
    elseif ($base -match 'Times') { $Record.Standard = $script:TimesWidths }
    elseif ($base -match 'Helvetica|Arial') { $Record.Standard = $script:HelveticaWidths }
}

function Add-PdfCidWidths {
    <#
    .SYNOPSIS
      A composite font's widths, from its descendant font: /DW (the default) and /W, whose entries are either
      "first [w1 w2 ...]" or "first last w".
    .PARAMETER Record
      The font record (filled in place).
    .PARAMETER Dict
      The Type0 font dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param([hashtable]$Record, [string]$Dict, $Objects)
    $descendant = [regex]::Match((Get-PdfArrayText $Dict 'DescendantFonts' $Objects), '(\d+)\s+\d+\s+R')
    if (-not $descendant.Success -or -not $Objects.Contains($descendant.Groups[1].Value)) { return }
    $cid = $Objects[$descendant.Groups[1].Value].Dict
    $Record.DefaultWidth = 1.0
    $dw = [regex]::Match($cid, '/DW\s+(\d+(?:\.\d+)?)')
    if ($dw.Success) { $Record.DefaultWidth = [double]::Parse($dw.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture) / 1000 }
    # Read token by token (a hostile /W can hold millions): the entry cap and the clock are checked in every loop.
    $token = (Get-PdfRegex '\[|\]|-?(?:\d+\.?\d*|\.\d+)').Match((Get-PdfArrayText $cid 'W' $Objects))
    if ($token.Success) { $token = $token.NextMatch() }   # the outer [
    $numbers = New-Object System.Collections.Generic.List[double]
    $count = 0
    while ($token.Success) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        if ($token.Value -eq '[') {
            # first [w1 w2 ...]: the number before the bracket is the first glyph.
            if ($numbers.Count -eq 0) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }   # [ with no first glyph
            $glyph = [int]$numbers[$numbers.Count - 1]; $numbers.Clear()
            for ($token = $token.NextMatch(); $token.Success -and $token.Value -ne ']'; $token = $token.NextMatch()) {
                if ((++$count % 2000) -eq 0) { Assert-PdfTime }
                if ($Record.Widths.Count -ge $script:MaxWidthEntries) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }
                $Record.Widths[$glyph] = [double]$token.Value / 1000; $glyph++
            }
        }
        elseif ($token.Value -ne ']') {
            $numbers.Add([double]$token.Value)
            if ($numbers.Count -eq 3) {   # first last w
                for ($g = [int]$numbers[0]; $g -le [int]$numbers[1]; $g++) {
                    if ($Record.Widths.Count -ge $script:MaxWidthEntries) { Add-PdfNote 'A font in this PDF could not be fully read: check by hand'; return }
                    $Record.Widths[$g] = $numbers[2] / 1000
                }
                $numbers.Clear()
            }
        }
        if ($token.Success) { $token = $token.NextMatch() }
    }
}

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

# ==== Module PdfText.psm1 ====
<#
PdfText.psm1 - the text of a PDF file, in memory, with .NET only (no module; the PAM, 2026-11-20).

  Purpose:     Page text (Tj/TJ text in Flate, ASCII85 or hex streams, with ToUnicode font maps), filled-in form-field
               values, comments and file info, so the detectors can check a PDF (ISO 32000). An encrypted PDF,
               a PDF with no text (a scanned image) or a file that is not a PDF gets a reason instead.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text and Reason (Get-PdfText).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters), PdfObjects.psm1 (objects, strings, dictionary text) and
               PdfFonts.psm1 (font maps and widths), and OcrText.psm1 (scanned pages), imported first by
               Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Add-Type -AssemblyName System.IO.Compression


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:PdfLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = (Get-PdfStringPattern) + '|%[^\r\n]*|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'


# Gaps along a line (a same-line Td move past the real end of the text before it, or a TJ adjustment), in ems: a
# space when wider than the block's letter spacing plus WordGapEms. The letter spacing is the median gap between two
# one-glyph pieces when a block has at least MinSpacingGaps of them and they are at least half of its gaps (gaps of
# ColumnGapEms or more are columns, not spacing), else 0 (fresh reviews, 2026-11-21).
$script:WordGapEms = 0.2
$script:MinSpacingGaps = 4
$script:ColumnGapEms = 1.0

# The operators that change the text (Invoke-PdfTextOperator); every other operator only drops its operands.
# Tokens per Matches walk (Read-PdfTextObject): bounds the memory a walk holds.
$script:WalkTokens = 5000
# Entries one [ ] array may hold (a TJ line has tens to thousands); more makes the PDF NOT SCANNED.
$script:MaxArrayItems = 100000
# Entries of what a line showed before they are summed (Compress-PdfShown).
$script:MaxShownItems = 1024

$script:TextOperators = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($operator in @('Tf', 'Tc', 'Tw', 'Tz', 'TJ', 'Tj', "'", '"', 'Td', 'TD', 'T*', 'Tm', 'ET')) { [void]$script:TextOperators.Add($operator) }

# Number parsing: PDF numbers use a point, whatever the computer's culture.
$script:NumberStyle = [Globalization.NumberStyles]::Float
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
# Symbol-font and other private-use characters (Complete-PdfTextBlock), built once (speed).
$script:SymbolChars = [regex]'[\uF020-\uF0FF]'
$script:PrivateChars = [regex]'[\uE000-\uF8FF]'
# The end of an inline image's bytes (BI <keys> ID <bytes> EI): white space, EI, then a delimiter or the end. After a
# BI operator (a token, never the word BI in a string) the reader steps over the bytes to here, so image bytes never
# join with page text into one token (fresh reviews, 2026-11-21).
$script:InlineImageEnd = '\sEI(?=[\s/\[<(%]|$)'

$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'
# A scanned PDF's JPEG page images, found by Get-PdfAllText: read with Windows OCR when the PDF has no text.
$script:PdfScanImages = @()
$script:PdfSlowReason = 'This PDF is too large or damaged to read in time: check by hand'
$script:PdfPartReason = 'Part of this PDF could not be read (a damaged or unusual string): check by hand'
$script:PdfEncryptedReason = 'Encrypted PDF (password or label): check by hand'

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
    $walk = $true
    while ($walk) {
    $walk = $false; $inWalk = 0
    foreach ($token in $Tokens.Matches($Content, $next)) {
        if ((++$inWalk) -gt $script:WalkTokens) { $walk = $true; break }   # this token is found again from $next
        if ((++$count % 100) -eq 0) { Assert-PdfTime }
        $t = $token.Value
        if ($token.Index -gt $next -and $Content.IndexOf('(', $next, $token.Index - $next) -ge 0) { Assert-PdfGap $Content $next $token.Index }
        $next = $token.Index + $token.Length
        $first = $t[0]
        if ($first -eq '%') { continue }   # a comment
        if ($t -ceq 'BI') { $inImage = $true; $stack.Clear(); continue }
        if ($inImage) {
            if ($t -ceq 'ID') { $next = Skip-PdfInlineImage $Content $next $ImageEnd; $inImage = $false; $walk = $true; break }   # else: its keys
            continue
        }
        if ($t -eq '[') { $array = New-Object System.Collections.Generic.List[string]; continue }
        if ($t -eq ']') { if ($null -ne $array) { $stack.Add($array) }; $array = $null; continue }
        if ($first -eq '(' -or $first -eq '<' -or $first -eq '/' -or ($first -ge '0' -and $first -le '9') -or $first -eq '-' -or $first -eq '+' -or $first -eq '.') {
            if ($null -eq $array) { $stack.Add($t); continue }
            # No real TJ array is this long: a hostile one would hold gigabytes, so the PDF is reported (r14).
            if ($array.Count -ge $script:MaxArrayItems) { throw (New-Object System.FormatException (Get-PdfStringError)) }
            $array.Add($t); continue
        }
        # Only text operators change the text; drawing operators just drop their operands.
        if ($null -ne $array) { continue }   # an operator inside an open [ ] (damaged): the array goes on, as before
        if ($t -ceq 'Td') { Add-PdfMove $stack $State }   # the most common operator: straight to its handler (speed)
        elseif ($script:TextOperators.Contains($t)) { Invoke-PdfTextOperator $t $stack $State $Fonts }
        $stack.Clear()
        if ($t -ceq 'ET') { return $next }
    }
    }
    if ($next -lt $Content.Length -and $Content.IndexOf('(', $next) -ge 0) { Assert-PdfGap $Content $next $Content.Length }
    return $Content.Length
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
        $sorted = @($letters.ToArray() | Sort-Object); $base = [Math]::Max(0.0, $sorted[[int]($sorted.Count / 2)])
    }
    return @{ Base = $base; Letters = $letters.Count; Gaps = $gaps }
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
    if ($base -eq 0 -and $spacing.Letters -gt 0 -and $spacing.Letters * 2 -ge $spacing.Gaps) { $base = $BlockBase }
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
    # Letter spacing is measured per line (a letter-spaced heading must not join the words of the body below it):
    # within a line, the median gap between two one-glyph pieces, when there are at least $MinSpacingGaps of them and
    # they are at least half of the line's gaps; a gap is a space only when wider than that plus $WordGapEms.
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
    Add-PdfNote 'Part of this PDF is damaged: check by hand'
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
            $State.Shown.Add([double]($move * $State.Size * $scale))
            if ($move -gt 0) { $State.Pieces.Add(@{ Gap = [double]$move; Single = ($State.LastLength -eq 1) }) }
        }
    }
}

function Get-PdfText {
    <#
    .SYNOPSIS
      The text of a PDF: the text of every page and the values of its form fields. A PDF that is encrypted, has
      no text (a scanned image) or is not a PDF gets a reason instead. Fonts without a ToUnicode map and other
      unusual encodings can give partial text: what can be read is checked.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [hashtable] Text, Reason (empty when the text was read), Note (a part not read, for a "check by hand" row; or empty).
    #>
    param([byte[]]$Bytes)
    $text = $script:PdfLatin1.GetString($Bytes)
    if ($text.IndexOf('%PDF') -lt 0 -or $text.IndexOf('%PDF') -gt 1024) {
        return @{ Text = ''; Reason = 'Not a valid PDF (damaged, or another type with a .pdf name): check by hand' }
    }
    if ($text -match '/Encrypt\s*(?:\d+\s+\d+\s+R|<<)') { return @{ Text = ''; Reason = $script:PdfEncryptedReason } }
    Reset-PdfOldObjects   # this PDF's own older versions only
    $script:PdfScanImages = @()   # and its own scanned pages
    Start-PdfBudget
    try { $all = Get-PdfAllText $text }
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.TimeoutException] -or $e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
            if ($e -is [System.FormatException] -and $e.Message -eq (Get-PdfStringError)) { return @{ Text = ''; Reason = $script:PdfPartReason } }
        }
        throw   # any other error is the caller's (Get-TextFromBytes reports the file)
    }
    finally { $notes = @(Stop-PdfBudget) }
    $note = (@($notes) -join '; ')
    # Runs of spaces and tabs become one space: the detectors never need them, and a page of 15 million spaces would
    # otherwise cost the later checks a gigabyte (fresh review, 2026-11-21).
    $all = [regex]::Replace($all, '[ \t]{2,}', ' ')
    if (-not $all.Trim()) {
        if ($script:PdfScanImages.Count -gt 0) {
            $scan = Get-ScanPagesText $script:PdfScanImages $note
            if (-not $scan.Reason -and -not $scan.Text.Trim()) { $scan.Reason = $script:PdfNoTextReason }
            return $scan
        }
        $reason = $script:PdfNoTextReason; if ($note) { $reason = $note }
        return @{ Text = ''; Reason = $reason; Note = '' }
    }
    return @{ Text = $all; Reason = ''; Note = $note }
}

function Get-PdfPageResources {
    <#
    .SYNOPSIS
      The resources dictionary of a page: its own (inline or by reference), else the nearest one up its /Parent
      chain (a Pages node), so a page that inherits its fonts gets the right ones.
    .PARAMETER Page
      The page object.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] the dictionary text to look up fonts in.
    #>
    param($Page, $Objects)
    $node = $Page
    for ($level = 0; $level -lt 32 -and $node; $level++) {   # 32: deeper than any real page tree; stops a loop
        $ref = [regex]::Match($node.Dict, '/Resources\s+(\d+)\s+\d+\s+R')
        if ($ref.Success -and $Objects.Contains($ref.Groups[1].Value)) { return $Objects[$ref.Groups[1].Value].Dict }
        if ($node.Dict -match '/Resources\s*<<') { return $node.Dict }
        $parent = [regex]::Match($node.Dict, '/Parent\s+(\d+)\s+\d+\s+R')
        $node = $null
        if ($parent.Success -and $Objects.Contains($parent.Groups[1].Value)) { $node = $Objects[$parent.Groups[1].Value] }
    }
    return $Page.Dict
}

function Get-PdfLooseText {
    <#
    .SYNOPSIS
      The text of content streams no page or form reads: older page versions kept by an incremental update ("SSN"
      edited out of a page is still in the file) and streams of pages since removed. A stream counts as content when
      it has a text block (BT, Tj or TJ, ET) and is not XML. One that cannot be read gives a note, not a failure.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Fonts
      The fonts of the whole file.
    .PARAMETER Read
      The streams already read.
    .OUTPUTS
      [string] text, one item per stream.
    #>
    param($Objects, [hashtable]$Fonts, $Read)
    $loose = @($Objects.Values) + (Get-PdfOldObjects)
    foreach ($obj in $loose) {
        if (-not $obj.Stream -or $Read.Contains($obj) -or $obj.Dict -match '/Type\s*/(?:ObjStm|XRef)\b') { continue }
        $start = $obj.Stream.TrimStart()
        if ($start.StartsWith('<') -or $obj.Stream -notmatch '\bBT\b' -or $obj.Stream -notmatch '\bT[Jj]\b|\)\s*[''"]') { continue }
        try { Convert-PdfContentToText $obj.Stream $Fonts }
        catch {
            if ($_.Exception -is [System.FormatException] -and $_.Exception.Message -eq (Get-PdfStringError)) {
                Add-PdfNote 'Part of this PDF could not be read: check by hand'
            }
            else { throw }
        }
    }
}

function Get-PdfAllText {
    <#
    .SYNOPSIS
      Page text (each page with its own fonts; the fonts of the whole file when a page names none, as when its
      resources are inherited), form XObjects, dictionaries (Get-PdfDictText), XFA data and content streams no page
      reads (Get-PdfLooseText), joined.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [string]
    #>
    param([string]$Text)
    $objects = Get-PdfObjects $Text
    $script:PdfScanImages = @($objects.Values | Where-Object { $_.ContainsKey('Image') -and $_.Image } | ForEach-Object { $_.Image })
    Assert-PdfTime
    $allFonts = Get-PdfFonts $objects
    Assert-PdfTime
    $parts = New-Object System.Collections.Generic.List[string]
    $read = New-Object System.Collections.Generic.HashSet[object]   # the content streams read as pages or forms
    foreach ($page in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/Page(?![s\w])' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $page $objects) $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $contents = [regex]::Match($page.Dict, '/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)').Groups[1].Value
        foreach ($content in [regex]::Matches($contents, '(\d+)\s+\d+\s+R')) {
            $stream = $objects[$content.Groups[1].Value]
            if (-not $stream) { continue }
            if ($stream.Stream) { [void]$read.Add($stream); $parts.Add((Convert-PdfContentToText $stream.Stream $fonts)); continue }
            # /Contents 5 0 R where object 5 is an array of streams ([6 0 R 7 0 R]): read each of them.
            foreach ($inner in [regex]::Matches($stream.Dict, '(\d+)\s+\d+\s+R')) {
                $part = $objects[$inner.Groups[1].Value]
                if ($part -and $part.Stream) { [void]$read.Add($part); $parts.Add((Convert-PdfContentToText $part.Stream $fonts)) }
            }
        }
    }
    # Form XObjects (/Subtype /Form, drawn with "Do"): reusable page parts that can hold text of their own. Read each
    # once; a form field's appearance can repeat its value, so a hit may show twice (better than a miss).
    foreach ($form in @($objects.Values | Where-Object { $_.Stream -and $_.Dict -match '/Subtype\s*/Form\b' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $form $objects) $objects   # also /Resources 10 0 R
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        [void]$read.Add($form)
        $parts.Add((Convert-PdfContentToText $form.Stream $fonts))
    }
    foreach ($line in @(Get-PdfDictText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfXfaText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfLooseText $objects $allFonts $read)) { $parts.Add($line) }
    return ($parts -join "`n")
}

# ==== Module TextExtract.psm1 ====
<#
TextExtract.psm1 - turns file bytes into plain text, in memory.

  Purpose:     Plain text from Word, Excel and PowerPoint files (.docx/.xlsx/.pptx and macro versions), PDF,
               OpenDocument, Visio, RTF, text and HTML; images (OCR), old .doc/.xls/.ppt and .msg; and each file
               inside a .zip or an .eml (the PAM, 2026-11-20). Other types (scripts and anything else), encrypted
               files and files that cannot be read are reported as NOT SCANNED, so a person checks them by hand.
  Inputs:      file bytes and the file name (for its extension).
  Outputs:     an object with Status ('OK' or 'NOT SCANNED'), Reason, Text, Note and Parts (Get-TextFromBytes).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  PdfText.psm1 (Get-PdfText), OcrText.psm1 (Get-ImageText), OleText.psm1 (Get-OleText) and
               ContainerText.psm1 (Get-ArchiveParts, Get-EmailParts), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Add-Type -AssemblyName System.IO.Compression

# =====================================================================================
# Settings: which files are read, and which parts of an Office file hold text
# =====================================================================================

# Plain text types, read as they are.
$script:PlainTextExtensions = @('.txt', '.csv', '.md', '.log', '.json', '.xml', '.tsv')

# HTML types, read and stripped of tags.
$script:HtmlExtensions = @('.html', '.htm')

# Office Open XML types (a zip of XML parts, ECMA-376) and the parts that hold their text:
# body, headers, footers, comments, notes; worksheets and cell comments; slides and speaker notes.
# A worksheet is read cell by cell (Convert-SheetXmlToText), with its shared strings put back into their cells.
$script:OfficeParts = @{
    '.docx' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.docm' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.xlsx' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.xlsm' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.pptx' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    '.pptm' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    # OpenDocument (ODF: a zip with content.xml) and Visio (pages in visio/pages); the PAM, 2026-11-20: round 3.
    '.odt' = '^content\.xml$'
    '.ods' = '^content\.xml$'
    '.odp' = '^content\.xml$'
    '.vsdx' = '^visio/pages/page\d+\.xml$'
}

# Images, read with Windows OCR (OcrText.psm1); old binary Office and Outlook files (OleText.psm1); files that
# hold other files (ContainerText.psm1): their parts are read one by one, at most $MaxContainerDepth levels deep.
$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif')
$script:OleExtensions = @('.doc', '.xls', '.ppt', '.msg')
$script:ContainerExtensions = @('.zip', '.eml')
$script:MaxContainerDepth = 2

# First bytes of a zip (every .docx/.xlsx/.pptx) and of an OLE compound file. Office saves a file in OLE form when it
# is encrypted: a password, or a sensitivity label (CUI) with encryption; an old .doc/.xls/.ppt renamed is OLE too.
# The text of an encrypted file cannot be read without the Microsoft Information Protection SDK (not on the gov device).
$script:ZipMagic = '504B'
$script:OleMagic = 'D0CF11E0A1B11AE1'
$script:EncryptedReason = 'Encrypted, password-protected or labelled with encryption (or an old Office format with a new name): check by hand'

# =====================================================================================
# Markup to text
# =====================================================================================

# Between two cells of a worksheet row: a mark that no detector reads across, so values of neighbouring cells
# (123 | 45 | 6789) never join into one SSN, card or phone number (the PAM, 2026-11-20: Excel false positives).
$script:CellBreak = ' | '

function Convert-SheetXmlToText {
    <#
    .SYNOPSIS
      A worksheet as text: one line per row, cells split by $CellBreak. Only what a person sees in a cell is kept:
      text (shared or inline), whole numbers and formula results. Formula text is dropped, and so are numbers with
      a decimal part or an exponent (a time or a calculated value: 45231.416666666664 is no SSN or phone number).
    .PARAMETER Xml
      The worksheet XML.
    .PARAMETER Shared
      The shared strings of the workbook, in order (a cell of type "s" holds an index into them).
    .OUTPUTS
      [string]
    #>
    param([string]$Xml, [string[]]$Shared)
    # Tags may carry a namespace prefix (<x:row>, <x:c>): some tools write them so, and a missed tag would be a silent miss.
    $lines = foreach ($row in [regex]::Matches($Xml, '(?s)<(?:\w+:)?row\b[^>]*>(.*?)</(?:\w+:)?row>')) {
        $cells = foreach ($cell in [regex]::Matches($row.Groups[1].Value, '(?s)<(?:\w+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?c>)')) {
            $type = [regex]::Match($cell.Groups[1].Value, '\bt="(\w+)"').Groups[1].Value
            $inner = $cell.Groups[2].Value -replace '(?s)<(?:\w+:)?f\b[^>]*/>|<(?:\w+:)?f\b[^>]*>.*?</(?:\w+:)?f>', ''   # formula text: not data
            $value = [System.Net.WebUtility]::HtmlDecode([regex]::Match($inner, '(?s)<(?:\w+:)?v>(.*?)</(?:\w+:)?v>').Groups[1].Value)
            switch ($type) {
                's' { $index = 0; if ([int]::TryParse($value, [ref]$index) -and $index -lt $Shared.Count) { $Shared[$index] } }
                'inlineStr' { Convert-XmlToText $inner }
                'str' { $value }
                default { if ($type -notin @('b', 'e') -and $value -match '^-?\d+$') { $value } }   # whole numbers only
            }
        }
        (@($cells | Where-Object { $_ }) -join $script:CellBreak)
    }
    return ($lines -join "`n")
}

function Convert-XmlToText {
    <#
    .SYNOPSIS
      Strips Office XML to plain text. Paragraph, row and cell ends become line breaks or spaces;
      every other tag is removed with no space, so text that Word splits into several runs
      ("123-45-" and "6789") joins again.
    .PARAMETER Xml
      The XML text of one part.
    .OUTPUTS
      [string]
    #>
    param([string]$Xml)
    # Word field codes are not shown on the page: PAGEREF _Toc123456789, HYPERLINK addresses, form IDs. They gave
    # false positives (a 9-digit bookmark number read as an SSN), so their text is dropped (the PAM, 2026-11-20).
    $text = $Xml -replace '(?s)<w:instrText\b[^>]*?(?:/>|>.*?</w:instrText>)', ''
    $text = $text -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
    $text = $text -replace '<text:(?:s|tab|line-break)\b[^>]*/>', ' '
    $text = $text -replace '</w:p>|</a:p>|</row>|</text:p>|</text:h>|</table:table-row>|</Text>', "`n"
    $text = $text -replace '</table:table-cell>', $script:CellBreak   # ODF cells: never joined into one number
    $text = $text -replace '</si>|</c>|</w:tc>|</a:tc>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

function Convert-HtmlToText {
    <#
    .SYNOPSIS
      Strips HTML (.html files) to plain text. Script and style blocks are dropped;
      block ends become line breaks.
    .PARAMETER Html
      The HTML text.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $text = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ' '
    $text = $text -replace '(?i)<br\s*/?>|</p>|</div>|</li>|</tr>', "`n"
    $text = $text -replace '(?i)</td>|</th>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

# =====================================================================================
# Reading bytes
# =====================================================================================

function Read-ZipEntry {
    # The text of one zip entry (UTF-8, as Office writes its XML).
    param($Entry)
    $reader = New-Object System.IO.StreamReader($Entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-SharedStrings {
    # The shared strings of a workbook in order (xl/sharedStrings.xml), each as plain text; none for other files.
    param($Zip)
    $entry = $Zip.GetEntry('xl/sharedStrings.xml')
    if (-not $entry) { return @() }
    foreach ($item in [regex]::Matches((Read-ZipEntry $entry), '(?s)<(?:\w+:)?si>(.*?)</(?:\w+:)?si>')) { Convert-XmlToText $item.Groups[1].Value }
}

function Get-OfficeProps {
    <#
    .SYNOPSIS
      The names an Office file stores about itself (docProps/core.xml: creator, last modified by), as Office saved
      them; empty for other files (the accountability columns, the PAM 2026-11-21).
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [hashtable] Creator, LastModifiedBy.
    #>
    param([byte[]]$Bytes)
    $props = @{ Creator = ''; LastModifiedBy = '' }
    if ($Bytes.Length -lt 4 -or $Bytes[0] -ne 0x50 -or $Bytes[1] -ne 0x4B) { return $props }   # not a zip
    try {
        $memory = New-Object System.IO.MemoryStream(, $Bytes)
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $core = $zip.GetEntry('docProps/core.xml')
            if ($core) {
                $xml = Read-ZipEntry $core
                $props.Creator = [System.Net.WebUtility]::HtmlDecode([regex]::Match($xml, '<dc:creator>([^<]*)</dc:creator>').Groups[1].Value)
                $props.LastModifiedBy = [System.Net.WebUtility]::HtmlDecode([regex]::Match($xml, '<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>').Groups[1].Value)
            }
        }
        finally { $zip.Dispose(); $memory.Dispose() }
    }
    catch { return @{ Creator = ''; LastModifiedBy = '' } }   # a damaged zip: its reader reports it
    return $props
}

function Get-ZipPartText {
    <#
    .SYNOPSIS
      Opens an Office file (a zip) in memory and returns the text of the parts whose names match.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER PartPattern
      Regular expression on the part name (from $OfficeParts).
    .OUTPUTS
      [string] the parts' text, joined with line breaks.
    #>
    param([byte[]]$Bytes, [string]$PartPattern)
    $memory = New-Object System.IO.MemoryStream(, $Bytes)
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $shared = @(Get-SharedStrings $zip)
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match $PartPattern } | Sort-Object FullName)) {
            $xml = Read-ZipEntry $entry
            if ($entry.FullName -match '^xl/worksheets/') { $texts.Add((Convert-SheetXmlToText $xml $shared)) }
            else { $texts.Add((Convert-XmlToText $xml)) }
        }
        return ($texts -join "`n")
    }
    finally { $zip.Dispose(); $memory.Dispose() }
}

function Convert-RtfToText {
    <#
    .SYNOPSIS
      Rich text (.rtf) as plain text: groups that hold no page text (font and colour tables, pictures, \*
      destinations) dropped; \par, \line and \tab as breaks; \'hh and \uN characters decoded; other control
      words and braces removed.
    .PARAMETER Rtf
      The RTF text.
    .OUTPUTS
      [string]
    #>
    param([string]$Rtf)
    # Drop whole groups that are not page text, with the groups nested in them (a balancing group counts the braces).
    $skip = '\{\\(?:\*|fonttbl|colortbl|stylesheet|pict|themedata|colorschememapping|datastore|latentstyles)' +
        '(?>[^{}]+|\{(?<d>)|\}(?<-d>))*(?(d)(?!))\}'
    $Rtf = (New-Object System.Text.RegularExpressions.Regex($skip, 'None', [TimeSpan]::FromSeconds(10))).Replace($Rtf, '')
    $text = $Rtf -replace '\\(?:par|line|row|sect|page)\b ?', "`n" -replace '\\(?:tab|cell)\b ?', ' '
    # \uN (with its fallback character, \'hh or ?, dropped) and \'hh in one pass, so a fallback is never decoded twice.
    $out = New-Object System.Text.StringBuilder
    $at = 0
    foreach ($m in [regex]::Matches($text, '\\u(-?\d+) ?(?:\\''[0-9A-Fa-f]{2}|\?)?|\\''([0-9A-Fa-f]{2})')) {
        [void]$out.Append($text, $at, $m.Index - $at)
        if ($m.Groups[1].Success) { [void]$out.Append([char](([int]$m.Groups[1].Value) -band 0xFFFF)) }
        else { [void]$out.Append([char][Convert]::ToInt32($m.Groups[2].Value, 16)) }
        $at = $m.Index + $m.Length
    }
    [void]$out.Append($text, $at, $text.Length - $at)
    $text = $out.ToString()
    $text = $text -replace '\\[A-Za-z]+-?\d* ?', '' -replace '\\([{}\\])', '$1' -replace '[{}]', ''
    return $text
}

function Get-PlainText {
    <#
    .SYNOPSIS
      Decodes a text file. A UTF-8 or UTF-16 byte order mark is honoured; no mark means UTF-8.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    $reader = New-Object System.IO.StreamReader((New-Object System.IO.MemoryStream(, $Bytes)), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-NotZipReason {
    <#
    .SYNOPSIS
      Why an Office file's bytes are not a zip, from their first bytes; empty when they are one. Without this check
      the zip reader says only "End of central directory not found" (the PAM's DoD365 run, 2026-11-20).
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string] the reason a person can act on, or ''.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return 'Empty file (0 bytes)' }
    $head = ($Bytes[0..([Math]::Min(7, $Bytes.Length - 1))] | ForEach-Object { $_.ToString('X2') }) -join ''
    if ($head.StartsWith($script:ZipMagic)) { return '' }
    if ($head.StartsWith($script:OleMagic)) { return $script:EncryptedReason }
    # Decoded with its byte order mark, if any (UTF-8 or UTF-16), so a web page with a mark is still seen.
    $start = (Get-PlainText ([byte[]]$Bytes[0..([Math]::Min(255, $Bytes.Length - 1))])).TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
    if ($start.StartsWith('<')) { return 'The download gave a web page, not the file: open it in SharePoint and check by hand' }
    return 'Not a valid Office file (damaged, or another type with an Office name): check by hand'
}

function Test-SupportedFile {
    # True when Get-TextFromBytes can read this file type (so other types are not downloaded).
    param([string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($script:PlainTextExtensions -contains $extension) -or ($script:HtmlExtensions -contains $extension) -or
        $script:OfficeParts.ContainsKey($extension) -or $extension -in @('.pdf', '.rtf') -or $script:ImageExtensions -contains $extension -or
        $script:OleExtensions -contains $extension -or $script:ContainerExtensions -contains $extension
}

function Get-TextFromBytes {
    <#
    .SYNOPSIS
      Returns the text of one file, or why it was not scanned. A zip or an e-mail gives Parts: the result of each
      file inside it (read the same way, up to $MaxContainerDepth levels deep).
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER FileName
      The file name; only its extension is used.
    .PARAMETER Depth
      How many containers this file sits in (0 for a file in the library).
    .OUTPUTS
      [pscustomobject] Status ('OK' or 'NOT SCANNED'), Reason, Text, Note (a part of a read file that was not read),
      Parts (the results of the files inside, each with a Name; empty for other files).
    #>
    param([byte[]]$Bytes, [string]$FileName, [int]$Depth = 0)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    $result = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = ''; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) }
    try {
        $read = $null   # a reader's hashtable (Text, Reason, Note) for the types that give one
        if ($script:PlainTextExtensions -contains $extension) { $result.Text = Get-PlainText $Bytes }
        elseif ($script:HtmlExtensions -contains $extension) { $result.Text = Convert-HtmlToText (Get-PlainText $Bytes) }
        elseif ($extension -eq '.rtf') { $result.Text = Convert-RtfToText (Get-PlainText $Bytes) }
        elseif ($extension -eq '.pdf') { $read = Get-PdfText $Bytes }
        elseif ($script:ImageExtensions -contains $extension) { $read = Get-ImageText $Bytes }
        elseif ($script:OleExtensions -contains $extension) { $read = Get-OleText $Bytes $extension }
        elseif ($script:ContainerExtensions -contains $extension) { Add-ContainerParts $result $Bytes $extension $Depth }
        elseif ($script:OfficeParts.ContainsKey($extension)) {
            $notZip = Get-NotZipReason $Bytes
            if ($notZip) { $result.Status = 'NOT SCANNED'; $result.Reason = $notZip }
            else { $result.Text = Get-ZipPartText $Bytes $script:OfficeParts[$extension] }
        }
        else {
            $result.Status = 'NOT SCANNED'
            $result.Reason = "File type '$extension' is not read by this script: check by hand"
        }
        if ($read) {
            if ($read.Reason) { $result.Status = 'NOT SCANNED'; $result.Reason = $read.Reason }
            else { $result.Text = $read.Text; if ($read.ContainsKey('Note')) { $result.Note = $read.Note } }
        }
    }
    catch {
        # A damaged or protected file must not stop the scan: report it and go on.
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file (damaged or protected?): ' + $_.Exception.Message
    }
    return $result
}

function Add-ContainerParts {
    <#
    .SYNOPSIS
      Reads each file in a zip or e-mail into $Result.Parts, named "inner/path" (a part's own parts are flattened
      as "inner.zip/file"). A container deeper than $MaxContainerDepth is not opened: NOT SCANNED with a reason.
    .PARAMETER Result
      The container's result (Get-TextFromBytes).
    .PARAMETER Bytes
      The container bytes.
    .PARAMETER Extension
      The container's extension: zip or eml, with its point.
    .PARAMETER Depth
      How many containers the container sits in.
    #>
    param($Result, [byte[]]$Bytes, [string]$Extension, [int]$Depth)
    if ($Depth -ge $script:MaxContainerDepth) {
        $Result.Status = 'NOT SCANNED'; $Result.Reason = 'An archive or e-mail inside another: open it and check by hand'
        return
    }
    $parts = if ($Extension -eq '.zip') { @(Get-ArchiveParts $Bytes) } else { @(Get-EmailParts $Bytes) }
    foreach ($part in $parts) {
        if ($part.Reason) { $one = [pscustomobject]@{ Status = 'NOT SCANNED'; Reason = $part.Reason; Text = ''; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) } }
        elseif ($null -ne $part.Text) {
            $text = $part.Text; if ($part.Name -like '*.html') { $text = Convert-HtmlToText $text }
            $one = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = $text; Note = ''; Parts = (New-Object System.Collections.Generic.List[object]) }
        }
        else { $one = Get-TextFromBytes -Bytes $part.Bytes -FileName $part.Name -Depth ($Depth + 1) }
        $one | Add-Member -NotePropertyName Name -NotePropertyValue ([string]$part.Name) -Force
        $Result.Parts.Add($one)
        $inners = @(); if ($null -ne $one.Parts -and $one.Parts.Count -gt 0) { $inners = $one.Parts.ToArray() }
        foreach ($inner in $inners) {
            $inner.Name = $part.Name + '/' + $inner.Name
            $Result.Parts.Add($inner)
        }
    }
}

# ==== Module ReaderPool.psm1 ====
<#
ReaderPool.psm1 - reads several files at once in PowerShell runspaces, and each identical file only once per run.

  Purpose:     Speed (the PAM, 2026-11-21): files are read and checked for PII by 1 to 8 workers at a time, while the
               main thread lists, downloads (SharePoint) and reports. The number of workers follows the load: it
               grows while the throughput grows and shrinks when it falls. A file whose content (SHA-256) was
               already read in this run is not read again: it gets the same hits. Nothing is kept after the run.
  Inputs:      jobs: a file's bytes (SharePoint) or its local or network path (drives), and its name.
  Outputs:     results: Status, Reason, Note, Hits, Parts, Hash, Duplicate, Seconds, Size.
  How to run:  imported by Find-ChannelPII.ps1; each worker loads the same reader code (the readers above this
               module and this module) once.
  Depends on:  PiiDetectors.psm1 (Find-PiiInText) and TextExtract.psm1 (Get-TextFromBytes), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7). Runspaces are part of PowerShell: no module, no
               compiled code, no other process.
  Network:     none here (a network-drive path is opened read-only with the person's Windows logon).
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>


# Workers: at most this many (and at most one per processor); the first target is 2.
$script:MaxWorkers = 8
$script:FirstWorkers = 2
# The target is reviewed after every this many finished files: up one when the MB per second rose by more than
# $GrowthStep, down one when it fell by more than $DropStep.
$script:ReviewEvery = 6
$script:GrowthStep = 0.05
$script:DropStep = 0.15
# How long the main thread waits for a worker between screen updates (milliseconds).
$script:WaitSliceMs = 250
# Paths longer than this are opened with the \\?\ prefix (Windows' classic limit is 260 characters).
$script:LongPathAt = 240

# The reader modules, in load order, for a development run (the bundle gives the reader block's text instead).
$script:ReaderModules = @('PiiDetectors', 'PiiMasking', 'OcrText', 'OleText', 'ContainerText', 'PdfCore', 'PdfObjects', 'PdfFonts',
    'PdfText', 'TextExtract', 'ReaderPool')

# The script each worker runs per job: it loads the reader code once per runspace, then reads one file.
$script:PoolWorkerScript = @'
param($Init, $Job, $Seen)
if (-not $global:IonReaderReady) {
    if ($Init.ContainsKey('Code')) { New-Module -Name IonReaders -ScriptBlock ([scriptblock]::Create($Init.Code)) | Import-Module -Global }
    else { Import-Module $Init.Paths -Global -Force }
    $global:IonReaderReady = $true
}
Invoke-PooledRead $Job $Seen
'@

function Read-LocalFileBytes {
    <#
    .SYNOPSIS
      The bytes of one local or network file, opened read-only and shared (others may keep it open and write):
      the file is never locked, changed or touched otherwise. Long paths get the \\?\ prefix.
    .PARAMETER Path
      The full path (X:\... or \\server\share\...).
    .OUTPUTS
      [byte[]]
    #>
    param([string]$Path)
    $open = $Path
    if ($Path.Length -gt $script:LongPathAt -and -not $Path.StartsWith('\\?\')) {
        if ($Path.StartsWith('\\')) { $open = '\\?\UNC\' + $Path.Substring(2) } else { $open = '\\?\' + $Path }
    }
    $stream = New-Object System.IO.FileStream($open, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]'ReadWrite, Delete')
    try {
        $memory = New-Object System.IO.MemoryStream
        try { $stream.CopyTo($memory); return , $memory.ToArray() } finally { $memory.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Get-ContentHash {
    <#
    .SYNOPSIS
      The SHA-256 of a file's bytes as hex text, with the FIPS-validated Windows provider (DoD computers enforce
      FIPS mode); empty when it is not available, and then the file is simply read.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    try {
        $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
        try { return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') } finally { $sha.Dispose() }
    }
    catch { return '' }
}

function Invoke-PooledRead {
    <#
    .SYNOPSIS
      Reads one file and checks its text for PII (runs in a worker). A file whose content another job of this run
      already claimed is not read again: Duplicate, with the Id of that job.
    .PARAMETER Job
      @{ Id; Name; Bytes (or $null); Path (when Bytes is $null) }.
    .PARAMETER Seen
      A synchronized hashtable shared by all workers: content hash and extension -> the Id of the job that reads it.
    .OUTPUTS
      [hashtable] Id, Status, Reason, Note, Hits, Parts, Hash, Duplicate, Original, Seconds, Size, Props.
    #>
    param([hashtable]$Job, [hashtable]$Seen)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $result = @{ Id = $Job.Id; Status = 'OK'; Reason = ''; Note = ''; Hits = @(); Parts = @(); Hash = ''; Duplicate = $false
        Original = $null; Seconds = 0.0; Size = 0; Props = $null }
    try {
        $bytes = $Job.Bytes
        if ($null -eq $bytes) { $bytes = Read-LocalFileBytes $Job.Path }
        $result.Size = $bytes.Length
        # Same content and same extension (the extension decides the reader): read once.
        $hash = Get-ContentHash $bytes
        if ($hash) { $result.Hash = $hash + '|' + [System.IO.Path]::GetExtension($Job.Name).ToLowerInvariant() }
        if ($result.Hash -and $null -ne $Seen) {
            $claimed = $false
            [System.Threading.Monitor]::Enter($Seen.SyncRoot)
            try { if ($Seen.ContainsKey($result.Hash)) { $result.Original = $Seen[$result.Hash] } else { $Seen[$result.Hash] = $Job.Id; $claimed = $true } }
            finally { [System.Threading.Monitor]::Exit($Seen.SyncRoot) }
            if (-not $claimed) { $result.Duplicate = $true; return $result }
        }
        $text = Get-TextFromBytes -Bytes $bytes -FileName $Job.Name
        $result.Status = $text.Status; $result.Reason = $text.Reason; $result.Note = $text.Note
        if ($text.Status -eq 'OK') { $result.Hits = @(Find-PiiInText $text.Text) }
        if ($result.Hits.Count -gt 0) { $result.Props = Get-OfficeProps $bytes }   # who made and changed it (Office files)
        # No @() around the list itself: @() around an empty generic list property fails in PowerShell.
        if ($null -ne $text.Parts -and $text.Parts.Count -gt 0) { $result.Parts = $text.Parts.ToArray() }
    }
    catch {
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file (in use, damaged or protected?): ' + $_.Exception.Message
    }
    finally { $result.Seconds = $clock.Elapsed.TotalSeconds }
    return $result
}

function New-ReaderPool {
    <#
    .SYNOPSIS
      Opens the worker pool: up to $MaxWorkers runspaces (at most one per processor), $FirstWorkers busy at first.
    .PARAMETER Init
      How a worker loads the reader code: @{ Code = <text> } (the paste bundle) or @{ Paths = <module files> }; when
      empty, the files of the reader modules loaded now (a development run).
    .OUTPUTS
      [hashtable] the pool state for Add-PoolJob, Receive-PoolResult and Close-ReaderPool.
    #>
    param([hashtable]$Init)
    if (-not $Init) { $Init = @{ Paths = @(Get-Module $script:ReaderModules | ForEach-Object { $_.Path }) } }
    $max = [Math]::Max(1, [Math]::Min($script:MaxWorkers, [Environment]::ProcessorCount))
    $pool = [runspacefactory]::CreateRunspacePool(1, $max)
    $pool.Open()
    return @{ Pool = $pool; Init = $Init; Jobs = (New-Object System.Collections.Generic.List[object]); Max = $max
        Target = [Math]::Min($script:FirstWorkers, $max); Seen = [hashtable]::Synchronized(@{}); Done = 0
        WindowBytes = 0L; WindowClock = [Diagnostics.Stopwatch]::StartNew(); LastRate = 0.0 }
}

function Add-PoolJob {
    <#
    .SYNOPSIS
      Hands one file to a worker; first waits (and returns the finished results) while the target number of
      workers is busy.
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Job
      @{ Id; Name; Bytes or Path } (see Invoke-PooledRead).
    .OUTPUTS
      The results that finished while waiting (Receive-PoolResult).
    #>
    param([hashtable]$State, [hashtable]$Job)
    while ($State.Jobs.Count -ge $State.Target) { Receive-PoolResult $State -Wait }
    $shell = [powershell]::Create()
    $shell.RunspacePool = $State.Pool
    [void]$shell.AddScript($script:PoolWorkerScript).AddArgument($State.Init).AddArgument($Job).AddArgument($State.Seen)
    $State.Jobs.Add(@{ Shell = $shell; Handle = $shell.BeginInvoke(); Id = $Job.Id })
}

function Receive-PoolResult {
    <#
    .SYNOPSIS
      The results of the workers that have finished (with -Wait: waits until at least one has, a slice at a time).
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Wait
      Wait for at least one result when none is ready (and jobs are running).
    .OUTPUTS
      [hashtable] results of Invoke-PooledRead.
    #>
    param([hashtable]$State, [switch]$Wait)
    while ($true) {
        $ready = @($State.Jobs | Where-Object { $_.Handle.IsCompleted })
        if ($ready.Count -gt 0 -or -not $Wait -or $State.Jobs.Count -eq 0) { break }
        [void]$State.Jobs[0].Handle.AsyncWaitHandle.WaitOne($script:WaitSliceMs)
    }
    foreach ($job in $ready) {
        [void]$State.Jobs.Remove($job)
        try {
            $output = @($job.Shell.EndInvoke($job.Handle))
            $result = $output | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
            if (-not $result) { $result = @{ Id = $job.Id; Status = 'NOT SCANNED'; Reason = 'The reader gave no result: check by hand' } }
        }
        catch { $result = @{ Id = $job.Id; Status = 'NOT SCANNED'; Reason = 'The reader failed: ' + $_.Exception.Message } }
        finally { $job.Shell.Dispose() }
        Update-PoolTarget $State $result
        $result
    }
}

function Update-PoolTarget {
    <#
    .SYNOPSIS
      Counts one finished file and, every $ReviewEvery files, moves the number of busy workers by one: up while the
      MB per second grows, down when it falls or a file could not be read.
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Result
      The finished result.
    #>
    param([hashtable]$State, [hashtable]$Result)
    $State.Done++
    if ($Result.ContainsKey('Size')) { $State.WindowBytes += [long]$Result.Size }
    $failed = $Result.Status -ne 'OK' -and $Result.Reason -match '^Could not read the file|^The reader'
    if ($failed -and $State.Target -gt 1) { $State.Target--; return }
    if ($State.Done % $script:ReviewEvery -ne 0) { return }
    $rate = $State.WindowBytes / [Math]::Max(0.001, $State.WindowClock.Elapsed.TotalSeconds)
    if ($rate -gt $State.LastRate * (1 + $script:GrowthStep) -and $State.Target -lt $State.Max) { $State.Target++ }
    elseif ($rate -lt $State.LastRate * (1 - $script:DropStep) -and $State.Target -gt 1) { $State.Target-- }
    $State.LastRate = $rate
    $State.WindowBytes = 0L
    $State.WindowClock.Restart()
}

function Close-ReaderPool {
    <#
    .SYNOPSIS
      Stops the workers that still run (after a cancel or an error) and closes the pool.
    .PARAMETER State
      The pool of New-ReaderPool.
    #>
    param([hashtable]$State)
    foreach ($job in $State.Jobs.ToArray()) {
        try { $job.Shell.Stop() } catch { Write-Verbose 'A worker had already stopped.' }   # nothing left to stop
        $job.Shell.Dispose()
    }
    $State.Jobs.Clear()
    $State.Pool.Close()
    $State.Pool.Dispose()
}

}
$null = New-Module -Name IonReaders -ScriptBlock $IonReaderBlock
$IonWorkerInit = @{ Code = $IonReaderBlock.ToString() }

# ==== Module XlsxWriter.psm1 ====
<#
XlsxWriter.psm1 - writes a small Excel workbook (.xlsx) with .NET only: no Excel, no module.

  Purpose:     The results file of Ion: one workbook with tabs (the PAM, 2026-11-20: a CSV cannot have tabs).
               Each tab has a bold header row; a list tab also has the header frozen and filter buttons on, and its
               link column holds clickable links (the PAM, 2026-11-20).
  Inputs:      the tabs to write (New-XlsxSheet), each with its columns and rows of cells (text or whole numbers).
  Outputs:     one .xlsx file (Office Open XML, ECMA-376: a zip of XML parts).
  Depends on:  .NET System.IO.Compression (loaded here).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the one .xlsx file the caller names; it never overwrites a file that exists.
  Credentials: none.

  Safety:      every text cell is an inline string, never a formula, so text from a scanned file cannot run in Excel.
#>

Add-Type -AssemblyName System.IO.Compression

# Excel limits: 31 characters in a tab name, 32,767 in a cell (ECMA-376 and Excel's specifications and limits).
$script:MaxSheetName = 31
$script:MaxCellText = 32767

# Excel opens a link of at most 2,079 characters (a longer one stays plain text in its cell).
$script:MaxLinkChars = 2079

# XML namespaces of the parts.
$script:MainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Three cell styles: 0 = normal, 1 = bold (the header row and the Summary labels), 2 = a link (blue, underlined).
$script:StylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>'

function New-XlsxSheet {
    <#
    .SYNOPSIS
      One tab of the workbook.
    .PARAMETER Name
      The tab name (cut to 31 characters; characters Excel refuses are replaced).
    .PARAMETER Columns
      The header row.
    .PARAMETER Rows
      The rows: each an array of cells, a cell being text or a whole number.
    .PARAMETER Widths
      Column widths in characters, one per column.
    .PARAMETER List
      Freeze the header row and turn the filter buttons on (a list of hits); off for the Summary tab.
    .PARAMETER LinkColumn
      The column (0-based) whose https:// cells become clickable links; -1 for none.
    .OUTPUTS
      [pscustomobject]
    #>
    param([string]$Name, [string[]]$Columns, [object[]]$Rows, [int[]]$Widths, [switch]$List, [int]$LinkColumn = -1)
    [pscustomobject]@{ Name = $Name; Columns = $Columns; Rows = @($Rows); Widths = $Widths; List = [bool]$List; LinkColumn = $LinkColumn }
}

function ConvertTo-XmlText {
    # Text for an XML part: characters XML does not allow are dropped, the rest is escaped, and it is cut to Excel's cell limit.
    param([AllowEmptyString()][string]$Text)
    $clean = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', ''
    if ($clean.Length -gt $script:MaxCellText) { $clean = $clean.Substring(0, $script:MaxCellText) }
    return [System.Security.SecurityElement]::Escape($clean)
}

function Get-ColumnLetter {
    # The column letters of a 1-based column number: 1 = A, 27 = AA.
    param([int]$Number)
    $letters = ''
    while ($Number -gt 0) { $Number--; $letters = [char](65 + $Number % 26) + $letters; $Number = [Math]::Floor($Number / 26) }
    return $letters
}

function ConvertTo-CellXml {
    # One cell: a whole number as a number, anything else as an inline string (never a formula).
    param([string]$Ref, $Value, [int]$Style)
    $styleAttr = ''
    if ($Style -gt 0) { $styleAttr = ' s="' + $Style + '"' }
    if ($Value -is [int] -or $Value -is [long]) { return '<c r="{0}"{1}><v>{2}</v></c>' -f $Ref, $styleAttr, $Value }
    return '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Ref, $styleAttr, (ConvertTo-XmlText ([string]$Value))
}

function Get-SheetLinks {
    <#
    .SYNOPSIS
      The clickable links of one tab: each https:// cell in its link column, with its cell and relationship ID.
      A link is a relationship of the sheet (ECMA-376 hyperlink), never a HYPERLINK formula (the safety rule above).
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .OUTPUTS
      [pscustomobject] Ref (for example G2), Id (rId1, rId2 ...), Target.
    #>
    param($Sheet)
    if ($Sheet.LinkColumn -lt 0) { return }
    $letter = Get-ColumnLetter ($Sheet.LinkColumn + 1)
    $rowNumber = 1
    $n = 0
    foreach ($row in $Sheet.Rows) {
        $rowNumber++
        $cells = @($row)
        if ($Sheet.LinkColumn -ge $cells.Count) { continue }
        $target = [string]$cells[$Sheet.LinkColumn]
        if ($target -notmatch '^https://\S+$' -or $target.Length -gt $script:MaxLinkChars) { continue }
        $n++
        [pscustomobject]@{ Ref = $letter + $rowNumber; Id = 'rId' + $n; Target = $target }
    }
}

function ConvertTo-SheetRelsXml {
    # The relationships part of one tab: one external target per link.
    param([object[]]$Links)
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    foreach ($link in $Links) {
        $xml += '<Relationship Id="{0}" Type="{1}/hyperlink" Target="{2}" TargetMode="External"/>' -f $link.Id, $script:RelNs, (ConvertTo-XmlText $link.Target)
    }
    return $xml + '</Relationships>'
}

function ConvertTo-SheetXml {
    <#
    .SYNOPSIS
      The worksheet XML of one tab.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .PARAMETER Links
      Output of Get-SheetLinks for this tab.
    .OUTPUTS
      [string]
    #>
    param($Sheet, [object[]]$Links = @())
    $linked = @{}
    foreach ($link in $Links) { $linked[$link.Ref] = $true }
    $xml = New-Object System.Text.StringBuilder
    [void]$xml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + $script:MainNs + '" xmlns:r="' + $script:RelNs + '">')
    if ($Sheet.List) {
        [void]$xml.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    }
    if ($Sheet.Widths) {
        [void]$xml.Append('<cols>')
        for ($i = 0; $i -lt $Sheet.Widths.Count; $i++) {
            [void]$xml.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $Sheet.Widths[$i]))
        }
        [void]$xml.Append('</cols>')
    }
    [void]$xml.Append('<sheetData><row r="1">')
    for ($i = 0; $i -lt $Sheet.Columns.Count; $i++) { [void]$xml.Append((ConvertTo-CellXml ((Get-ColumnLetter ($i + 1)) + '1') $Sheet.Columns[$i] 1)) }
    [void]$xml.Append('</row>')
    $rowNumber = 1
    foreach ($row in $Sheet.Rows) {
        $rowNumber++
        [void]$xml.Append('<row r="' + $rowNumber + '">')
        $cells = @($row)
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $style = 0
            $ref = (Get-ColumnLetter ($i + 1)) + $rowNumber
            if (-not $Sheet.List -and $i -eq 0) { $style = 1 }   # Summary tab: the label column is bold
            if ($linked.ContainsKey($ref)) { $style = 2 }
            [void]$xml.Append((ConvertTo-CellXml $ref $cells[$i] $style))
        }
        [void]$xml.Append('</row>')
    }
    [void]$xml.Append('</sheetData>')
    if ($Sheet.List -and $Sheet.Columns.Count -gt 0) {
        [void]$xml.Append(('<autoFilter ref="A1:{0}{1}"/>' -f (Get-ColumnLetter $Sheet.Columns.Count), [Math]::Max(2, $rowNumber)))
    }
    if ($Links.Count -gt 0) {
        [void]$xml.Append('<hyperlinks>')
        foreach ($link in $Links) { [void]$xml.Append(('<hyperlink ref="{0}" r:id="{1}"/>' -f $link.Ref, $link.Id)) }
        [void]$xml.Append('</hyperlinks>')
    }
    [void]$xml.Append('</worksheet>')
    return $xml.ToString()
}

function Get-SafeSheetNames {
    <#
    .SYNOPSIS
      Tab names Excel accepts: no : \ / ? * [ ], at most 31 characters, none twice (a number is added).
    .PARAMETER Names
      The wanted names, in order.
    .OUTPUTS
      [string[]] the names to use, in the same order.
    #>
    param([string[]]$Names)
    $used = @{}
    foreach ($name in $Names) {
        $clean = ($name -replace '[:\\/?*\[\]]', '-').Trim()
        if (-not $clean) { $clean = 'Sheet' }
        if ($clean.Length -gt $script:MaxSheetName) { $clean = $clean.Substring(0, $script:MaxSheetName) }
        $candidate = $clean
        $n = 2
        while ($used.ContainsKey($candidate.ToLower())) {
            $suffix = ' (' + $n + ')'
            $candidate = $clean.Substring(0, [Math]::Min($clean.Length, $script:MaxSheetName - $suffix.Length)) + $suffix
            $n++
        }
        $used[$candidate.ToLower()] = $true
        $candidate
    }
}

function Add-ZipText {
    # Adds one UTF-8 text part (no byte order mark) to the zip.
    param($Zip, [string]$Name, [string]$Text)
    $writer = New-Object System.IO.StreamWriter($Zip.CreateEntry($Name).Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Export-Xlsx {
    <#
    .SYNOPSIS
      Writes the tabs to a new .xlsx file. Throws if the file exists (the caller picks a new name).
    .PARAMETER Path
      The new file.
    .PARAMETER Sheets
      The tabs, in order (New-XlsxSheet).
    #>
    param([string]$Path, [object[]]$Sheets)
    $names = @(Get-SafeSheetNames @($Sheets | ForEach-Object { $_.Name }))
    $types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $book = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + $script:MainNs + '" xmlns:r="' + $script:RelNs + '"><sheets>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $types += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        $book += '<sheet name="' + (ConvertTo-XmlText $names[$i - 1]) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'
        $rels += '<Relationship Id="rId' + $i + '" Type="' + $script:RelNs + '/worksheet" Target="worksheets/sheet' + $i + '.xml"/>'
    }
    $rels += '<Relationship Id="rId' + ($Sheets.Count + 1) + '" Type="' + $script:RelNs + '/styles" Target="styles.xml"/></Relationships>'
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::CreateNew)   # CreateNew: never overwrite
    $zip = $null
    $written = $false
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        Add-ZipText $zip '[Content_Types].xml' ($types + '</Types>')
        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">' +
            '<Relationship Id="rId1" Type="' + $script:RelNs + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        Add-ZipText $zip 'xl/workbook.xml' ($book + '</sheets></workbook>')
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rels
        Add-ZipText $zip 'xl/styles.xml' $script:StylesXml
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $links = @(Get-SheetLinks $Sheets[$i - 1])
            Add-ZipText $zip ('xl/worksheets/sheet' + $i + '.xml') (ConvertTo-SheetXml $Sheets[$i - 1] $links)
            if ($links.Count -gt 0) { Add-ZipText $zip ('xl/worksheets/_rels/sheet' + $i + '.xml.rels') (ConvertTo-SheetRelsXml $links) }
        }
        $written = $true
    }
    finally {
        if ($zip) { $zip.Dispose() }
        $stream.Dispose()
        # A half-written workbook would not open: remove it, so a rerun can use the name.
        if (-not $written) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
    }
}

# ==== Module Report.psm1 ====
<#
Report.psm1 - the scan results for a person: a summary on screen and one Excel workbook with tabs.

  Purpose:     Holds the report rows, sorts them worst PII type first, shows the summary (totals and hits per type)
               and saves the workbook: Summary tab first, one tab per type, Low confidence, Not scanned
               (the PAM, 2026-11-20: tabs to cut noise; no .txt or .csv).
  Inputs:      report rows (New-ReportRow) from ChannelScan.psm1; values in them are already masked.
  Outputs:     the summary on screen; <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx.
  Depends on:  PiiDetectors.psm1 (Find-PiiInText), PiiMasking.psm1 (Protect-PiiName, Protect-PiiText), XlsxWriter.psm1,
               imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the workbook and its site and day folders under the output folder (default Documents\Ion).
               A new file each run; an existing result is never overwritten.
  Credentials: none.
#>


# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes
# (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# The tool's name on screen and in the workbook (the PAM, 2026-11-20).
$script:ToolTitle = 'Ion - DoD365 PII Reporter'

# PII types from most to least damaging, for sorting and for the order of the tabs. AFI 33-332 gives no ranking
# (A3.1 lists only the factors for judging harm), so this is the PAM's order (2026-11-20): SSN and date of birth
# first, then identity documents and accounts, then contact details.
$script:SeverityOrder = @('SSN', 'Date of birth', 'Passport', 'Payment card', 'DoD ID', 'Home address', 'Name', 'Phone', 'Personal e-mail')

# How many files the screen summary lists (the workbook lists all).
$script:TopItems = 10

# Screen colours (the PAM, 2026-11-20: less grey, more format). Write-Host has only the 16 console colours:
# Cyan = headings and labels, White = values, Green = clean or done, Yellow = check by hand and Medium,
# Red = possible PII, High and errors, Gray = hints and the labels of the folder rows, DarkGray = rules and table
# borders, DarkCyan (teal) = folder names on the folder rows (the PAM, 2026-11-20).
$script:ScreenColors = @{ Title = 'Cyan'; Label = 'Cyan'; Text = 'White'; Good = 'Green'; Accent = 'Yellow'; Alert = 'Red'
    Muted = 'Gray'; Rule = 'DarkGray'; Folder = 'DarkCyan' }

# Not-scanned reasons in groups for the counts on screen and on the Summary tab: a pattern on the reason, its group.
# A file type that is not read is counted per extension (Get-NotScannedGroup).
$script:NotScannedGroups = [ordered]@{
    '^Encrypted' = 'Encrypted or labelled with encryption'
    '^Larger than' = 'Too large'
    '^Empty file' = 'Empty file'
    '^No text in this PDF' = 'PDF without text (scanned image?)'
    'web page' = 'Download gave a web page'
    '^Download failed' = 'Download failed'
    '^Not a valid (Office file|PDF)|^Could not read|too large or damaged to read in time|^Part of this PDF' = 'Damaged or not readable'
    'could not be fully read|^A font in this PDF|^Some characters in this PDF' = 'PDF fonts not fully readable'
    '^Files attached inside this PDF' = 'PDF attachments'
    '^Image: Windows OCR|^Scanned PDF: Image: Windows OCR' = 'Images: OCR not available'
    '^Image:|^Scanned PDF' = 'Images and scans not fully read'
    '^Old Excel file' = 'Old Excel: numbers not read'
    'old Office or Outlook' = 'Old Office or Outlook not readable'
    '^An archive or e-mail inside another|in this archive|Too large to unpack|Could not unpack' = 'Archives: files not unpacked'
}


# What to do with PII found where it should not be: the breach steps of the rule source, AFI 33-332 para 3.1.2 as
# changed by AFI33-332_DAFGM2026-01 (7 July 2026; edition: docs/DECISIONS.md).
# The wording is the README's "If it finds PII" section; scripts/build_site.py --check fails when they differ (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01 (7 July 2026), ' +
                       'page 1, second paragraph).'

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the folder (masked by the caller).
    .PARAMETER Link
      A link to the item, or empty.
    .PARAMETER Modified
      When the item last changed, or empty.
    .PARAMETER Detector
      The PII type, or the note word.
    .PARAMETER Confidence
      High, Medium, Low, or empty for a note.
    .PARAMETER MaskedValue
      The masked value, or empty.
    .PARAMETER Context
      The masked snippet, or the reason for a note.
    .OUTPUTS
      [pscustomobject] Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link, and the
      accountability fields (empty until Set-RowAccountability fills them): ModifiedBy, Created, CreatedBy, AddedOn, AddedBy.
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source = $Source; Channel = $Channel; Item = $Item; Detector = $Detector; Confidence = $Confidence
        MaskedValue = $MaskedValue; Context = $Context; Modified = $Modified; Link = $Link
        ModifiedBy = ''; Created = ''; CreatedBy = ''; AddedOn = ''; AddedBy = ''
    }
}

function Add-IncompleteRow {
    <#
    .SYNOPSIS
      Adds one INCOMPLETE row when the scan did not reach its end and no INCOMPLETE row is there yet.
      Called from the script's 'finally' block, which also runs after Ctrl+C (when 'catch' does not).
    .PARAMETER Rows
      The list of report rows (changed in place).
    .PARAMETER Finished
      True when every step ran.
    .OUTPUTS
      None.
    #>
    param($Rows, [bool]$Finished)
    if ($Finished) { return }
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) { return }
    $Rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
}

function Write-Ui {
    <#
    .SYNOPSIS
      Writes one piece of screen text in a palette colour ($ScreenColors). Every screen line of Ion goes through here.
    .PARAMETER Text
      The text (already masked by the caller).
    .PARAMETER Kind
      Title, Label, Text, Good, Accent, Alert, Muted, Rule or Folder.
    .PARAMETER NoNewline
      Keep writing on the same line (for a line with several colours).
    #>
    param([AllowEmptyString()][string]$Text, [ValidateSet('Title', 'Label', 'Text', 'Good', 'Accent', 'Alert', 'Muted', 'Rule', 'Folder')][string]$Kind = 'Text',
          [switch]$NoNewline)
    Write-Host $Text -ForegroundColor $script:ScreenColors[$Kind] -NoNewline:$NoNewline
}

function Get-SeverityIndex {
    # Where a PII type stands in $SeverityOrder (0 = worst); a type not in the list comes last.
    param([string]$Type)
    $index = [Array]::IndexOf($script:SeverityOrder, $Type)
    if ($index -lt 0) { return $script:SeverityOrder.Count }
    return $index
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: its worst type, highest confidence, number of hits and types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Worst, Highest, Hits, Channel, Source, Item, Detectors; worst type first, then highest confidence.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Channel, Item)) {
        $first = $group.Group[0]
        $worst = ($group.Group | Sort-Object { Get-SeverityIndex $_.Detector }, { -$script:Rank[$_.Confidence] } | Select-Object -First 1)
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Worst = $worst.Detector; Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = ((@($group.Group | Select-Object -ExpandProperty Detector -Unique) | Sort-Object { Get-SeverityIndex $_ }) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { Get-SeverityIndex $_.Worst } }, @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true },
        @{ Expression = 'Hits'; Descending = $true })
}

function Get-TypeCounts {
    <#
    .SYNOPSIS
      The hits per PII type, worst type first, split by confidence.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, High, Medium, Low, Total.
    #>
    param([object[]]$Hits)
    foreach ($group in ($Hits | Group-Object Detector | Sort-Object { Get-SeverityIndex $_.Name })) {
        $count = @{ High = 0; Medium = 0; Low = 0 }
        foreach ($hit in $group.Group) { $count[$hit.Confidence]++ }
        [pscustomobject]@{ Type = $group.Name; High = $count.High; Medium = $count.Medium; Low = $count.Low; Total = $group.Count }
    }
}

function Get-NotScannedGroup {
    <#
    .SYNOPSIS
      The group of one not-scanned row: a type not read by its extension, else by its reason ($NotScannedGroups).
    .PARAMETER Row
      A report row that is not a hit (NOT SCANNED, BLOCKED or INCOMPLETE).
    .OUTPUTS
      [string] the group name.
    #>
    param($Row)
    if ($Row.Detector -ne 'NOT SCANNED') { return $Row.Detector }   # BLOCKED, INCOMPLETE
    if ($Row.Context -match '^File type') {
        $extension = [System.IO.Path]::GetExtension($Row.Item).ToLower()
        if (-not $extension) { $extension = '(no extension)' }
        return 'Type not read: ' + $extension
    }
    foreach ($pattern in $script:NotScannedGroups.Keys) { if ($Row.Context -match $pattern) { return $script:NotScannedGroups[$pattern] } }
    return 'Other'
}

function Get-NotScannedCounts {
    <#
    .SYNOPSIS
      How many files were not scanned, per reason, most first (the PAM, 2026-11-20: see at once what to check by hand).
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Reason, Count.
    #>
    param([object[]]$Rows)
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    foreach ($group in ($notes | Group-Object { Get-NotScannedGroup $_ } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)) {
        [pscustomobject]@{ Reason = $group.Name; Count = $group.Count }
    }
}

function Get-ScanMetrics {
    <#
    .SYNOPSIS
      The scan's totals, as label and value pairs (the screen summary and the Summary tab show the same).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([object[]]$Rows, [int]$ItemsScanned)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $metrics = [ordered]@{}
    $metrics['Files read'] = $ItemsScanned
    $metrics['Files with possible PII'] = @(Get-ReportSummary $Rows).Count
    $metrics['Possible PII hits'] = $hits.Count
    $metrics['High / Medium / Low'] = '{0} / {1} / {2}' -f @($hits | Where-Object { $_.Confidence -eq 'High' }).Count,
        @($hits | Where-Object { $_.Confidence -eq 'Medium' }).Count, @($hits | Where-Object { $_.Confidence -eq 'Low' }).Count
    $metrics['Not scanned / incomplete'] = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) }).Count
    return $metrics
}

function Format-BoxTable {
    <#
    .SYNOPSIS
      A text table with ASCII borders (no other characters: Windows PowerShell 5.1 and ISE show them all alike).
      Text columns are left-aligned, number columns right-aligned.
    .PARAMETER Columns
      The header texts.
    .PARAMETER Rows
      The rows, each an array of cells.
    .OUTPUTS
      [string[]] the border line, the header, the border, one line per row, the border.
    #>
    param([string[]]$Columns, [object[]]$Rows)
    $widths = for ($i = 0; $i -lt $Columns.Count; $i++) {
        [Math]::Max($Columns[$i].Length, (@($Rows | ForEach-Object { ([string]$_[$i]).Length }) + 0 | Measure-Object -Maximum).Maximum)
    }
    $border = '+' + ((@($widths) | ForEach-Object { '-' * ($_ + 2) }) -join '+') + '+'
    # One table line: each cell padded to its column width, numbers to the right.
    $line = {
        param($cells)
        $parts = for ($i = 0; $i -lt $Columns.Count; $i++) {
            $cell = [string]$cells[$i]
            if ($cells[$i] -is [int]) { ' ' + $cell.PadLeft($widths[$i]) + ' ' } else { ' ' + $cell.PadRight($widths[$i]) + ' ' }
        }
        '|' + ($parts -join '|') + '|'
    }
    $border
    & $line $Columns
    $border
    foreach ($row in $Rows) { & $line $row }
    $border
}

function Get-TableLines {
    <#
    .SYNOPSIS
      A boxed table as screen lines: borders grey, the header in the label colour, each row in its own colour.
    .PARAMETER Columns
      The header texts.
    .PARAMETER Rows
      The rows, each an array of cells.
    .PARAMETER Kinds
      One Write-Ui colour per row.
    .OUTPUTS
      [pscustomobject] Text, Kind, LabelLength (0), one per line.
    #>
    param([string[]]$Columns, [object[]]$Rows, [string[]]$Kinds)
    $table = @(Format-BoxTable $Columns $Rows)
    for ($i = 0; $i -lt $table.Count; $i++) {
        $kind = 'Rule'
        if ($i -eq 1) { $kind = 'Label' } elseif ($i -ge 3 -and $i -lt $table.Count - 1) { $kind = $Kinds[$i - 3] }
        [pscustomobject]@{ Text = $table[$i]; Kind = $kind; LabelLength = 0 }
    }
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The summary for the screen, line by line with its palette colour: the scan's details and totals, a table of
      possible PII per type (worst first; red with High hits, yellow with Medium, white otherwise), a table of the
      files with the most serious hits, and the not-scanned files per reason. Every hit is listed in the workbook.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [pscustomobject] Text, Kind (a Write-Ui colour), LabelLength (characters at the start in the label colour), one per line.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $out = New-Object System.Collections.Generic.List[object]
    # Adds one screen line with its colour; the first LabelLength characters are written in the label colour.
    $add = { param($text, $kind, $labelLength = 0) $out.Add([pscustomobject]@{ Text = $text; Kind = $kind; LabelLength = $labelLength }) }
    & $add '' 'Text'; & $add ('==== ' + $script:ToolTitle + ' ====') 'Title'
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) {
        & $add ('{0,-26} {1}' -f ($pair.Key + ':'), $pair.Value) 'Text' 27
    }
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    if ($hits.Count -eq 0) { & $add '' 'Text'; & $add 'No possible PII found in the files that were read.' 'Good' }
    else {
        $types = @(Get-TypeCounts $hits)
        & $add '' 'Text'; & $add 'POSSIBLE PII BY TYPE (worst first)' 'Title'
        $kinds = @($types | ForEach-Object { if ($_.High -gt 0) { 'Alert' } elseif ($_.Medium -gt 0) { 'Accent' } else { 'Text' } })
        Get-TableLines @('Type', 'High', 'Medium', 'Low', 'Total') @($types | ForEach-Object { , @($_.Type, $_.High, $_.Medium, $_.Low, $_.Total) }) $kinds |
            ForEach-Object { $out.Add($_) }
        $top = @(Get-ReportSummary $Rows | Select-Object -First $script:TopItems)
        & $add '' 'Text'; & $add ('FILES WITH THE MOST SERIOUS HITS (first {0}; all are in the workbook)' -f $script:TopItems) 'Title'
        $kinds = @($top | ForEach-Object { if ($_.Highest -eq 'High') { 'Alert' } else { 'Accent' } })
        Get-TableLines @('Worst', 'Highest', 'Hits', $FolderLabel, 'File') @($top | ForEach-Object { , @($_.Worst, $_.Highest, $_.Hits, $_.Channel, $_.Item) }) $kinds |
            ForEach-Object { $out.Add($_) }
    }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        & $add '' 'Text'; & $add 'NOT SCANNED: CHECK THESE BY HAND (the workbook lists each file)' 'Title'
        Get-TableLines @('Reason', 'Files') @($skipped | ForEach-Object { , @($_.Reason, $_.Count) }) @($skipped | ForEach-Object { 'Accent' }) |
            ForEach-Object { $out.Add($_) }
    }
    & $add '' 'Text'
    & $add 'Confidence is a pattern guess, not proof: open each item and check it before you act.' 'Muted'
    & $add $script:BreachAdvice 'Accent'
    return $out.ToArray()
}

# =====================================================================================
# The workbook
# =====================================================================================

function Get-ResultPath {
    <#
    .SYNOPSIS
      A new workbook path: <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx, with (2), (3) ... added
      if the name exists, so a result is never overwritten. Creates the folders. The site name is masked like any
      name the tool shows (found PII only), and characters Windows refuses in a folder name are replaced.
    .PARAMETER OutputFolder
      The base folder.
    .PARAMETER SiteUrl
      The site address; its last part (teams/<name>) names the site folder.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$SiteUrl)
    $siteName = [uri]::UnescapeDataString(($SiteUrl.TrimEnd('/') -split '/')[-1])
    $siteName = ((Protect-PiiName $siteName) -replace '[\\/:*?"<>|]', '-').Trim(' .')
    if (-not $siteName) { $siteName = 'site' }
    if ($siteName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $siteName = '_' + $siteName }   # names Windows keeps for devices
    $now = Get-Date
    $folder = Join-Path (Join-Path $OutputFolder $siteName) ($now.ToString('yyyy-MM-dd'))
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
    $base = 'Ion Results ' + $now.ToString('HH-mm-ss')
    $path = Join-Path $folder ($base + '.xlsx')
    $copy = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $folder ('{0} ({1}).xlsx' -f $base, $copy); $copy++ }
    return $path
}

function Get-HitCells {
    # One row of a hit tab: the hit's cells in the order of $HitColumns.
    param($Hit)
    @($Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context, $Hit.Modified, $Hit.Link, $Hit.Source,
        $Hit.ModifiedBy, $Hit.Created, $Hit.CreatedBy, $Hit.AddedOn, $Hit.AddedBy)
}

function Set-RowAccountability {
    <#
    .SYNOPSIS
      Fills the accountability fields of a file's rows: who last changed it, when it was made and by whom, and who
      put the PII in and when (where the file's history shows it; else the file's maker or "not recorded").
    .PARAMETER Rows
      The file's rows (New-ReportRow).
    .PARAMETER Who
      @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; ByHit } (any may be missing); ByHit: "type|masked value"
      -> @{ AddedOn; AddedBy } for each hit dated from the file's history.
    #>
    param([object[]]$Rows, [hashtable]$Who)
    if (-not $Who) { return }
    foreach ($row in $Rows) {
        foreach ($key in @('ModifiedBy', 'Created', 'CreatedBy', 'AddedOn', 'AddedBy')) {
            if ($Who.ContainsKey($key) -and $Who[$key]) { $row.$key = [string]$Who[$key] }
        }
        $hitKey = $row.Detector + '|' + $row.MaskedValue
        if ($Who.ContainsKey('ByHit') -and $Who.ByHit -and $Who.ByHit.ContainsKey($hitKey)) {
            $row.AddedOn = [string]$Who.ByHit[$hitKey].AddedOn; $row.AddedBy = [string]$Who.ByHit[$hitKey].AddedBy
        }
    }
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first, then one tab per PII type with High or Medium hits (worst type
      first), then Low confidence and Not scanned when they have rows (the PAM, 2026-11-20).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the Summary tab.
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site): the name of the folder column.
    .OUTPUTS
      [object[]] New-XlsxSheet objects.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    # Accountability (the PAM, 2026-11-21): who changed and made the file, and who put the PII in and when (by e-mail).
    $columns = @('Confidence', $FolderLabel, 'File', 'Masked value', 'Context', 'Modified', 'Link', 'Found in',
        'Modified by', 'Created', 'Created by', 'PII added', 'PII added by')
    $widths = @(11, 22, 45, 18, 60, 17, 50, 10, 30, 17, 30, 17, 30)
    $byRank = @{ Expression = { $script:Rank[$_.Confidence] }; Descending = $true }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header) @(28, 70, 9, 9, 9)))
    foreach ($type in @(Get-TypeCounts $hits | Where-Object { $_.High + $_.Medium -gt 0 } | ForEach-Object { $_.Type })) {
        $typeHits = @($hits | Where-Object { $_.Detector -eq $type -and $_.Confidence -ne 'Low' } | Sort-Object $byRank, Channel, Item)
        $sheets.Add((New-XlsxSheet $type $columns @($typeHits | ForEach-Object { , (Get-HitCells $_) }) $widths -List -LinkColumn 6))
    }
    $low = @($hits | Where-Object { $_.Confidence -eq 'Low' } | Sort-Object { Get-SeverityIndex $_.Detector }, Channel, Item)
    if ($low.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Low confidence' (@('Type') + $columns) @($low | ForEach-Object { , (@($_.Detector) + (Get-HitCells $_)) }) (@(16) + $widths) -List -LinkColumn 7))
    }
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    if ($notes.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'File', 'Reason', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, $_.Link) }) @(14, 22, 45, 70, 50) -List -LinkColumn 4))
    }
    return $sheets.ToArray()
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the hits per type (worst first), the not-scanned
      files per reason, then how to read the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) { , @($pair.Key, $pair.Value) }
    , @('')
    , @('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total')
    foreach ($type in @(Get-TypeCounts $hits)) { , @($type.Type, $type.High, $type.Medium, $type.Low, $type.Total) }
    if ($hits.Count -eq 0) { , @('None found in the files that were read.') }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        , @('')
        , @('Not scanned: check by hand', 'Files')
        foreach ($group in $skipped) { , @($group.Reason, $group.Count) }
    }
    , @('')
    , @('Tabs', 'One tab per type with High or Medium hits, worst type first; then Low confidence; then Not scanned.')
    , @('Confidence', 'A pattern guess, not proof: open each item and check it before you act.')
    , @('If you find PII', $script:BreachAdvice)
    , @('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.')
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the workbook (Summary tab first). Returns its path.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the summary.
    .PARAMETER OutputFolder
      The base folder; the workbook goes in <site>\<day> under it.
    .PARAMETER SiteUrl
      The site address (names the site folder).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$OutputFolder, [string]$SiteUrl,
          [string]$FolderLabel = 'Folder')
    foreach ($line in @(Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header -FolderLabel $FolderLabel)) {
        if ($line.LabelLength -gt 0) { Write-Ui $line.Text.Substring(0, $line.LabelLength) 'Label' -NoNewline; Write-Ui $line.Text.Substring($line.LabelLength) $line.Kind }
        else { Write-Ui $line.Text $line.Kind }
    }
    $path = Get-ResultPath $OutputFolder $SiteUrl
    Export-Xlsx -Path $path -Sheets (Get-WorkbookSheets $Rows $ItemsScanned $Header $FolderLabel)
    Write-Ui ''
    Write-Ui "Saved (values masked): $path" 'Good'
    Write-Ui 'Handle this file as CUI: it shows where PII is.' 'Accent'
    return $path
}

function Add-PartRows {
    <#
    .SYNOPSIS
      The rows of one file inside a zip or an e-mail, named "<file>/<inner path>" (PII in the name masked): its
      hits, a hit for PII in its name, and a NOT SCANNED row when it was not read or read only in part.
    .PARAMETER Part
      One result of Get-TextFromBytes Parts (with its Name).
    .PARAMETER Channel
      The channel name as the report shows it.
    .PARAMETER Item
      The container file as the report shows it.
    .PARAMETER Link
      The container file's link (an inner file has no link of its own).
    .PARAMETER Modified
      The container file's change time.
    .PARAMETER Rows
      The list that receives the rows.
    #>
    param($Part, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified, $Rows)
    $name = $Item + '/' + (Protect-PiiName $Part.Name)
    foreach ($hit in @(Find-PiiInText $Part.Name)) {
        $Rows.Add((New-ReportRow 'File name' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the name of a file inside it'))
    }
    if ($Part.Status -ne 'OK') { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Reason))); return }
    if ($Part.Note) { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Note))) }
    foreach ($hit in @(Find-PiiInText $Part.Text)) {
        $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
}

# ==== Module ScanMeter.psm1 ====
<#
ScanMeter.psm1 - the progress bars and the per-folder lines of a scan: counts only, so the screen shows no PII.

  Purpose:     Lets the person follow a scan: files done of the total, read, not scanned, possible PII,
               time used and time left (the PAM, 2026-11-20). No found value is shown; channel and file
               names show only masked PII.
  Inputs:      the report row list and the counts the scan functions of ChannelScan.psm1 give.
  Outputs:     a meter (hashtable), status text, percent, a per-file progress bar and one coloured line per folder,
               counts in fixed widths (the PAM, 2026-11-20).
  Depends on:  PiiMasking.psm1 (Protect-PiiName), Report.psm1 (Write-Ui), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (screen only).
  Credentials: none.
#>


# From this size on, the per-file bar says the step can take a minute (5 MB: a slow download on a gov network
# takes tens of seconds; the author's choice, 2026-11-20).
$script:LargeFileBytes = 5MB

# The per-file bar is shown only from this size on: a small file is done in a moment, and each bar update costs time in
# Windows PowerShell 5.1 (the author's choice, 2026-11-20: faster scans of many small files).
$script:FileBarBytes = 1MB

# A coloured line of live counts is printed under the progress bar at most this often (and after each folder).
$script:LiveLineSeconds = 10

# Width of each count on the live lines: 5 characters (up to 99,999), so the columns do not move (the PAM, 2026-11-20).
$script:CountWidth = 5

# Length of the grey rule after each folder's lines.
$script:RuleWidth = 100

# The name column of the folder rows: as wide as the longest folder name, at most this; the total row says "Total".
$script:MaxNameWidth = 40
$script:TotalName = 'Total'

# Short PII type names for the live line, in the PAM's order (Report.psm1 $SeverityOrder).
$script:ShortTypeNames = [ordered]@{ 'SSN' = 'SSN'; 'Date of birth' = 'DOB'; 'Passport' = 'Passport'; 'Payment card' = 'Card'
    'DoD ID' = 'DoD ID'; 'Home address' = 'Address'; 'Name' = 'Name'; 'Phone' = 'Phone'; 'Personal e-mail' = 'E-mail' }

function New-ScanMeter {
    <#
    .SYNOPSIS
      Starts a meter for one scan step: items done, read, not scanned, possible PII, and the time.
    .PARAMETER Total
      How many items the step has (0 when not known up front).
    .OUTPUTS
      [hashtable] the meter (Update-ScanMeter and Get-ScanMeterText use it).
    #>
    param([int]$Total)
    # Clock: anything with an Elapsed [TimeSpan] (a running stopwatch here; a fixed time in the tests).
    @{ Total = $Total; Done = 0; Read = 0; Skipped = 0; Hits = 0; HitItems = 0; Types = @{}; LastLine = [TimeSpan]::Zero; NameWidth = 0
       Clock = [Diagnostics.Stopwatch]::StartNew() }
}

function Update-ScanMeter {
    <#
    .SYNOPSIS
      Counts one item into the meter from the report rows it added.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Rows
      The report row list.
    .PARAMETER FirstRow
      The row count before the item was scanned.
    .PARAMETER ReadCount
      The items whose text was read (the scan function's return value).
    #>
    param([hashtable]$Meter, $Rows, [int]$FirstRow, [int]$ReadCount)
    $new = @(for ($r = $FirstRow; $r -lt $Rows.Count; $r++) { $Rows[$r] })
    $hits = @($new | Where-Object { $_.Confidence })   # a row with a confidence is a possible PII hit
    $Meter.Done++
    $Meter.Read += $ReadCount
    $Meter.Skipped += @($new | Where-Object { $_.Detector -eq 'NOT SCANNED' -or $_.Detector -eq 'BLOCKED' }).Count
    $Meter.Hits += $hits.Count
    $Meter.HitItems += @($hits | Select-Object -ExpandProperty Item -Unique).Count   # files with a hit
    foreach ($hit in $hits) { $Meter.Types[$hit.Detector] = 1 + [int]$Meter.Types[$hit.Detector] }
}

function Format-Duration {
    # A time as minutes:seconds; the minutes go past 59 (65 minutes = 65:00), so a long scan never wraps to 05:00.
    param([TimeSpan]$Time)
    return '{0:00}:{1:00}' -f [int][Math]::Floor($Time.TotalMinutes), $Time.Seconds
}

function Format-Count {
    # A count right-aligned in $CountWidth characters, so the columns of the live lines stay in place (the PAM, 2026-11-20).
    param([long]$Value)
    return ([string]$Value).PadLeft($script:CountWidth)
}

function Get-TimeLeftText {
    # ", about mm:ss left" from the pace so far; empty when the total is not known or the step is done.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0 -or $Meter.Done -le 0 -or $Meter.Done -ge $Meter.Total) { return '' }
    $elapsed = $Meter.Clock.Elapsed
    return ', about ' + (Format-Duration ([TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done))))) + ' left'
}

function Get-ScanMeterText {
    <#
    .SYNOPSIS
      The meter as one status line with fixed-width counts: counts, time used and, when the total is known, time left.
      The parts that change in width (time, PII per type) come last.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .OUTPUTS
      [string] for example "   12/   40 files | read    10 | not scanned     2 | possible PII     3 in     2 files | 00:42 used, about 01:48 left".
    #>
    param([hashtable]$Meter, [string]$Noun)
    $text = Format-Count $Meter.Done
    if ($Meter.Total -gt 0) { $text += '/' + (Format-Count $Meter.Total) }
    $types = Get-TypeCountText $Meter
    if ($types) { $types = ' | ' + $types }
    $text += ' {0}s | read {1} | not scanned {2} | possible PII {3} in {4} {0}s | {5} used{6}{7}' -f
        $Noun, (Format-Count $Meter.Read), (Format-Count $Meter.Skipped), (Format-Count $Meter.Hits), (Format-Count $Meter.HitItems),
        (Format-Duration $Meter.Clock.Elapsed), (Get-TimeLeftText $Meter), $types
    return $text
}

function Get-TypeCountText {
    # The hits per type so far, worst type first, as "SSN 2, DOB 1"; empty when there is none.
    param([hashtable]$Meter)
    if (-not $Meter.ContainsKey('Types') -or $Meter.Types.Count -eq 0) { return '' }
    $parts = foreach ($type in $script:ShortTypeNames.Keys) {
        if ($Meter.Types[$type]) { '{0} {1}' -f $script:ShortTypeNames[$type], $Meter.Types[$type] }
    }
    return (@($parts) -join ', ')
}

function Write-CountPart {
    # One part of a folder or total row: " | " (dark grey), the label (grey), the count (fixed width) in its colour.
    param([string]$Label, [long]$Value, [string]$Kind)
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Label + ' ') 'Muted' -NoNewline
    Write-Ui (Format-Count $Value) $Kind -NoNewline
}

function Get-CountKind {
    # The colour of a count: White when 0; else Yellow (not scanned: check by hand) or Red (possible PII).
    param([long]$Value, [string]$WhenSome)
    if ($Value -gt 0) { return $WhenSome }
    return 'Text'
}

function Get-NameWidth {
    # The width of the name column: the longest (shown) folder name, at least "Total", at most $MaxNameWidth.
    param([string[]]$Names)
    $width = $script:TotalName.Length
    foreach ($name in $Names) { $width = [Math]::Max($width, $name.Length) }
    return [Math]::Min($width, $script:MaxNameWidth)
}

function Format-Name {
    # A name padded (or cut, ending in "...") to the name column's width, so every row lines up.
    param([string]$Name, [int]$Width)
    if ($Name.Length -gt $Width) { return $Name.Substring(0, $Width - 3) + '...' }
    return $Name.PadRight($Width)
}

function Get-RowWidth {
    # The name column's width of a meter (a meter made in a test may have none).
    param([hashtable]$Meter)
    if ($Meter.ContainsKey('NameWidth') -and $Meter.NameWidth -gt 0) { return [int]$Meter.NameWidth }
    return $script:TotalName.Length
}

function Write-MeterLine {
    <#
    .SYNOPSIS
      Prints the running total as one row in the same columns as the folder rows: "Total", files done of the total
      and percent, read, not scanned, possible PII, time used and left, then the PII per type. At most every
      $LiveLineSeconds, unless -Force. Counts only: no file name and no value.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .PARAMETER Force
      Print now (after a folder).
    #>
    param([hashtable]$Meter, [string]$Noun, [switch]$Force)
    $elapsed = $Meter.Clock.Elapsed
    if (-not $Force -and ($elapsed - $Meter.LastLine).TotalSeconds -lt $script:LiveLineSeconds) { return }
    $Meter.LastLine = $elapsed
    Write-Ui ('  ' + (Format-Name $script:TotalName (Get-RowWidth $Meter))) 'Text' -NoNewline
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Noun + 's ') 'Muted' -NoNewline
    Write-Ui ('{0}/{1} {2,3}%' -f (Format-Count $Meter.Done), (Format-Count $Meter.Total), (Get-MeterPercent $Meter)) 'Text' -NoNewline
    Write-CountPart 'read' $Meter.Read 'Text'
    Write-CountPart 'not scanned' $Meter.Skipped (Get-CountKind $Meter.Skipped 'Accent')
    Write-CountPart 'possible PII' $Meter.Hits (Get-CountKind $Meter.Hits 'Alert')
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ('{0} used{1}' -f (Format-Duration $elapsed), (Get-TimeLeftText $Meter)) 'Muted' -NoNewline
    $types = Get-TypeCountText $Meter
    if ($types) { Write-Ui ' | ' 'Rule' -NoNewline; Write-Ui $types 'Alert' -NoNewline }
    Write-Ui ''
}

function Get-MeterPercent {
    # The meter's percent done (0-100); 0 when the total is not known.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0) { return 0 }
    return [int][Math]::Min(100, 100 * $Meter.Done / $Meter.Total)
}

function Write-FileStep {
    <#
    .SYNOPSIS
      A second progress bar for the file in work: its size and its step, so a large file does not look hung.
      PnP gives no byte count while it downloads, so the bar moves by step (download, read, check), not by byte.
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER Step
      What is being done now.
    .PARAMETER Percent
      How far through the file's steps (0-100).
    #>
    param([long]$Size, [string]$Step, [int]$Percent)
    if ($Size -lt $script:FileBarBytes) { return }
    $status = '{0} - {1:N1} MB' -f $Step, ($Size / 1MB)
    if ($Size -ge $script:LargeFileBytes) { $status += ' (a large file: this step can take a minute)' }
    # No -ParentId: Windows PowerShell ISE indents a child bar; the PAM wants both bars aligned (2026-11-20).
    Write-Progress -Id 2 -Activity 'This file' -Status $status -PercentComplete $Percent
}

function Write-ChannelLine {
    <#
    .SYNOPSIS
      Prints one row when a folder (a channel) is done: its (masked) name first in teal, in a fixed-width column,
      then its counts in fixed widths with grey labels (the PAM, 2026-11-20). Not scanned in yellow and possible
      PII in red when there are some; other numbers white.
    .PARAMETER Channel
      The channel name (masked here).
    .PARAMETER Before
      A copy of the meter before the channel (Done, Read, Skipped, Hits).
    .PARAMETER Meter
      The meter after the channel.
    .PARAMETER Noun
      What an item is, for example 'file'.
    #>
    param([string]$Channel, [hashtable]$Before, [hashtable]$Meter, [string]$Noun)
    $hits = $Meter.Hits - $Before.Hits
    $skipped = $Meter.Skipped - $Before.Skipped
    Write-Ui ('  ' + (Format-Name (Protect-PiiName $Channel) (Get-RowWidth $Meter))) 'Folder' -NoNewline
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Noun + 's ') 'Muted' -NoNewline
    # Padded to the width of the total row's "done/total percent", so the next columns line up.
    Write-Ui ((Format-Count ($Meter.Done - $Before.Done)).PadRight($script:CountWidth * 2 + 6)) 'Text' -NoNewline
    Write-CountPart 'read' ($Meter.Read - $Before.Read) 'Text'
    Write-CountPart 'not scanned' $skipped (Get-CountKind $skipped 'Accent')
    Write-CountPart 'possible PII' $hits (Get-CountKind $hits 'Alert')
    Write-Ui ''
}

function Write-RuleLine {
    # A grey line that ends one folder's block on screen.
    Write-Ui ('  ' + ('-' * $script:RuleWidth)) 'Rule'
}

# ==== Module ScanQueue.psm1 ====
<#
ScanQueue.psm1 - turns the reader pool's results into report rows, in the order files finish.

  Purpose:     The link between the scans (ChannelScan.psm1 for SharePoint, DriveScan.psm1 for drives) and the
               reader pool (ReaderPool.psm1): each file handed in gets its rows and is counted on the meter when its
               worker finishes; a file with the same content as one read before gets that file's hits.
  Inputs:      per file: how the report shows it (folder, item, link, change time) and its bytes or path.
  Outputs:     report rows (Report.psm1 New-ReportRow) in the caller's row list; the meter and the live line.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  ReaderPool.psm1, PiiMasking.psm1, Report.psm1 (New-ReportRow, Add-PartRows) and ScanMeter.psm1,
               imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none here.
  Writes:      none (rows in memory).
  Credentials: none.
#>


function New-ScanQueue {
    <#
    .SYNOPSIS
      A queue for one scan: the reader pool, the report rows, the meter and what each file in work is.
    .PARAMETER Init
      How the workers load the reader code (New-ReaderPool).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .OUTPUTS
      [hashtable] the queue.
    #>
    param([hashtable]$Init, $Rows, [hashtable]$Meter)
    return @{ Pool = (New-ReaderPool $Init); Rows = $Rows; Meter = $Meter; Info = @{}; Results = @{}; Waiting = @{}; NextId = 0
        Seconds = 0.0 }
}

function Submit-ScanFile {
    <#
    .SYNOPSIS
      Hands one file to the readers; adds the rows of the files that finished meanwhile.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Info
      @{ Channel; Item; Link; Modified } as the report shows the file (already masked).
    .PARAMETER Name
      The file name (its extension decides how it is read).
    .PARAMETER Bytes
      The file bytes (SharePoint), or $null.
    .PARAMETER Path
      The file path when Bytes is $null (drives).
    #>
    param([hashtable]$Queue, [hashtable]$Info, [string]$Name, [byte[]]$Bytes, [string]$Path)
    $Queue.NextId++
    $id = $Queue.NextId
    $Queue.Info[$id] = $Info
    $job = @{ Id = $id; Name = $Name; Bytes = $Bytes; Path = $Path }
    foreach ($result in @(Add-PoolJob $Queue.Pool $job)) { Resolve-ScanResult $Queue $result }
}

function Resolve-ScanResult {
    <#
    .SYNOPSIS
      Adds the rows of one finished file (and of the files with the same content that waited for it) and counts
      them on the meter.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Result
      A result of Invoke-PooledRead.
    #>
    param([hashtable]$Queue, [hashtable]$Result)
    if ($Result.ContainsKey('Seconds')) { $Queue.Seconds += [double]$Result.Seconds }
    $source = $Result
    if ($Result.ContainsKey('Duplicate') -and $Result.Duplicate) {
        if (-not $Queue.Results.ContainsKey($Result.Original)) {
            if (-not $Queue.Waiting.ContainsKey($Result.Original)) { $Queue.Waiting[$Result.Original] = New-Object System.Collections.Generic.List[int] }
            $Queue.Waiting[$Result.Original].Add($Result.Id)
            return
        }
        $source = $Queue.Results[$Result.Original]
    }
    else { $Queue.Results[$Result.Id] = $Result }
    Add-QueuedRows $Queue $Result.Id $source
    if ($Queue.Waiting.ContainsKey($Result.Id)) {
        foreach ($waiting in $Queue.Waiting[$Result.Id]) { Add-QueuedRows $Queue $waiting $Result }
        $Queue.Waiting.Remove($Result.Id)
    }
}

function Add-QueuedRows {
    <#
    .SYNOPSIS
      The rows of one file from a result (its own, or the one of a file with the same content), then the meter
      and the live line.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Id
      The file's job number.
    .PARAMETER Result
      The result whose hits the file gets.
    #>
    param([hashtable]$Queue, [int]$Id, [hashtable]$Result)
    $info = $Queue.Info[$Id]
    $Queue.Info.Remove($Id)
    $rows = $Queue.Rows
    $first = $rows.Count
    $read = 0
    if ($Result.Status -ne 'OK') {
        $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Reason)))
    }
    else {
        $read = 1
        # A file read in part (a PDF over a cap, damaged, or with attachments): a "check by hand" row next to its hits.
        if ($Result.Note) { $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Note))) }
        foreach ($hit in $Result.Hits) {
            $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
        }
        foreach ($part in $Result.Parts) { Add-PartRows $part $info.Channel $info.Item $info.Link $info.Modified $rows }
        if ($Result.Hits.Count -gt 0 -or $rows.Count -gt $first) { Add-WhoToRows $info $Result @($rows | Select-Object -Skip $first) }
    }
    Update-ScanMeter $Queue.Meter $rows $first $read
    Write-MeterLine $Queue.Meter 'file'
}

function Invoke-GroupScan {
    <#
    .SYNOPSIS
      Scans the groups (folders) one after the other: each file is started (its name checked, then handed to the
      readers or given a NOT SCANNED row), and the folder's line is written when all its files have their rows.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Groups
      The groups, each with Name and Files.
    .PARAMETER StartFile
      { param($Group, $File) } that starts one file; returns true when the file went to the readers.
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    #>
    param([hashtable]$Queue, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile)
    $meter = $Queue.Meter
    for ($c = 0; $c -lt $Groups.Count; $c++) {
        $group = $Groups[$c]
        $before = $meter.Clone()
        foreach ($file in $group.Files) {
            # The file in work, with PII masked (the PAM, 2026-11-20: a slow file is then visible).
            Write-Progress -Id 1 -Activity ("Scanning files - folder {0} of {1}: '{2}'" -f ($c + 1), $Groups.Count, (Protect-PiiName $group.Name)) `
                -Status (Get-ScanMeterText $meter 'file') -PercentComplete (Get-MeterPercent $meter) -CurrentOperation ('Now: ' + (& $ShownFile $group $file))
            $firstRow = $Queue.Rows.Count
            if (-not (& $StartFile $group $file)) {
                Update-ScanMeter $meter $Queue.Rows $firstRow 0
                Write-MeterLine $meter 'file'   # the running total, at most every $LiveLineSeconds
            }
        }
        Wait-ScanQueue $Queue   # the folder's line counts all its files
        Write-ChannelLine $group.Name $before $meter 'file'
    }
    Write-Progress -Id 1 -Activity 'Scanning files' -Completed
}

function Add-WhoToRows {
    <#
    .SYNOPSIS
      The accountability columns of a file's rows (the PAM, 2026-11-21): from its SharePoint fields and version
      history (Info.Who and Info.Added), or from the drive file (Info.DrivePath), plus the names an Office file stores.
    .PARAMETER Info
      The file's queue info.
    .PARAMETER Result
      Its reader result (Hits, Props).
    .PARAMETER FileRows
      The rows of this file.
    #>
    param([hashtable]$Info, [hashtable]$Result, [object[]]$FileRows)
    $props = $null; if ($Result.ContainsKey('Props')) { $props = $Result.Props }
    if ($Info.ContainsKey('DrivePath')) { $who = Get-DriveWho $Info.DrivePath $props }
    elseif ($Info.ContainsKey('Who')) {
        $who = $Info.Who.Clone()
        if ($Info.ContainsKey('Added') -and $Result.Hits.Count -gt 0) { $who.ByHit = (& $Info.Added $Result.Hits) }
    }
    else { return }
    Set-RowAccountability $FileRows $who
}

function Wait-ScanQueue {
    <#
    .SYNOPSIS
      Waits until every file handed in has its rows (the end of a folder, or of the scan).
    .PARAMETER Queue
      Output of New-ScanQueue.
    #>
    param([hashtable]$Queue)
    while ($Queue.Pool.Jobs.Count -gt 0) {
        foreach ($result in @(Receive-PoolResult $Queue.Pool -Wait)) { Resolve-ScanResult $Queue $result }
    }
}

function Close-ScanQueue {
    <#
    .SYNOPSIS
      Closes the readers (also after a cancel: workers still running are stopped). Returns the reading time of all
      workers together, for the header.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .OUTPUTS
      [TimeSpan]
    #>
    param([hashtable]$Queue)
    Close-ReaderPool $Queue.Pool
    return [TimeSpan]::FromSeconds($Queue.Seconds)
}

function Get-QueueWorkers {
    # The number of files read at once now (for the live line and the header).
    param([hashtable]$Queue)
    return $Queue.Pool.Target
}

# ==== Module Accountability.psm1 ====
<#
Accountability.psm1 - who made and changed a file with PII, and who put the PII in and when.

  Purpose:     The PAM, 2026-11-21: for accountability and to date a possible incident, each file with a hit shows
               who last modified it, when it was made and by whom, and who added the PII and when, by e-mail
               address (the report is CUI and stays in DoD365). SharePoint and Teams: the file's fields and its
               version history (the first version that holds each hit). Drives: the file's owner (a Windows logon
               name; no directory lookup: "No reconnaissance"), its times, and the names an Office file stores.
  Inputs:      the file's list item, path or reader result, and its hits.
  Outputs:     @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; ByHit } for Report.psm1 Set-RowAccountability.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  TextExtract.psm1 (Get-TextFromBytes), PiiDetectors.psm1 (Find-PiiInText) and, for SharePoint,
               PnP.PowerShell 1.12.0 (Get-PnPFileVersion, Get-PnPContext; read-only), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the signed-in site (version history of files with a hit).
  Writes:      none. Old versions are read in memory only.
  Credentials: none (the PnP sign-in of the scan).
#>


# Old versions read per file with a hit, oldest first (a file with more is dated from the oldest ones read).
$script:MaxVersionsRead = 30
$script:NoHistory = 'Not recorded (no version history)'
$script:DriveNoHistory = 'Not recorded (a drive keeps no file history)'

function Get-PersonText {
    <#
    .SYNOPSIS
      A person field as text: the e-mail address of a SharePoint user value, else its name, else the text itself.
    .PARAMETER Value
      A field value (FieldUserValue, User or text).
    .OUTPUTS
      [string]
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    foreach ($name in @('Email', 'LookupValue', 'Title')) {
        $property = $Value.PSObject.Properties[$name]
        if ($property -and $property.Value) { return [string]$property.Value }
    }
    return [string]$Value
}

function Get-DateText {
    # A date as yyyy-MM-dd HH:mm (the report's form); empty when missing.
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Get-SharePointWho {
    <#
    .SYNOPSIS
      Who made and last changed a SharePoint file, from its list item (Author, Created, Editor).
    .PARAMETER File
      The list item.
    .OUTPUTS
      [hashtable] ModifiedBy, Created, CreatedBy.
    #>
    param($File)
    return @{ ModifiedBy = (Get-PersonText $File['Editor']); Created = (Get-DateText $File['Created']); CreatedBy = (Get-PersonText $File['Author']) }
}

function Get-HitKey {
    # The same hit in two versions: its type and masked value.
    param($Hit)
    return ([string]$Hit.Detector + '|' + [string]$Hit.MaskedValue)
}

function Find-PiiAddedVersion {
    <#
    .SYNOPSIS
      For each hit of a SharePoint file, the first version that holds it: who saved that version and when. A hit
      in no older version came with the current one (the last modifier). Read-only; versions stay in memory.
    .PARAMETER Ref
      The file's server-relative path.
    .PARAMETER Name
      The file name (for its reader).
    .PARAMETER Hits
      The current hits (Detector, MaskedValue).
    .PARAMETER Who
      Get-SharePointWho of the file, with Modified (the current version's time).
    .OUTPUTS
      [hashtable] hit key -> @{ AddedOn; AddedBy }.
    #>
    param([string]$Ref, [string]$Name, [object[]]$Hits, [hashtable]$Who)
    $byHit = @{}
    $open = @($Hits | ForEach-Object { Get-HitKey $_ } | Sort-Object -Unique)
    try {
        $versions = @(Get-PnPFileVersion -Url $Ref -ErrorAction Stop | Sort-Object Created | Select-Object -First $script:MaxVersionsRead)
        $context = Get-PnPContext
        foreach ($version in $versions) { $context.Load($version.CreatedBy) }
        if ($versions.Count -gt 0) { $context.ExecuteQuery() }
    }
    catch { $versions = $null }
    if ($null -eq $versions) {
        foreach ($key in $open) { $byHit[$key] = @{ AddedOn = $script:NoHistory; AddedBy = $script:NoHistory } }
        return $byHit
    }
    foreach ($version in $versions) {
        if ($open.Count -eq 0) { break }
        try {
            $stream = $version.OpenBinaryStream(); $context.ExecuteQuery()
            $memory = New-Object System.IO.MemoryStream
            try { $stream.Value.CopyTo($memory); $bytes = $memory.ToArray() } finally { $memory.Dispose() }
            $text = Get-TextFromBytes -Bytes $bytes -FileName $Name
        }
        catch { continue }   # a version that cannot be read: the next one dates the hit
        if ($text.Status -ne 'OK') { continue }
        $found = @(Find-PiiInText $text.Text | ForEach-Object { Get-HitKey $_ })
        foreach ($key in @($open | Where-Object { $found -contains $_ })) {
            $byHit[$key] = @{ AddedOn = (Get-DateText $version.Created); AddedBy = (Get-PersonText $version.CreatedBy) }
        }
        $open = @($open | Where-Object { -not $byHit.ContainsKey($_) })
    }
    # Not in any older version: the current version added it (or, with no older version, the file's maker).
    foreach ($key in $open) {
        if ($versions.Count -eq 0) { $byHit[$key] = @{ AddedOn = $Who.Created; AddedBy = $Who.CreatedBy } }
        else { $byHit[$key] = @{ AddedOn = $Who.Modified; AddedBy = $Who.ModifiedBy } }
    }
    return $byHit
}

function Get-DriveWho {
    <#
    .SYNOPSIS
      Who made and changed a file on a drive: its owner (usually its maker; a Windows logon name), its creation
      time, and the names an Office file stores (creator, last modified by). Windows does not record who last
      changed other files.
    .PARAMETER Path
      The file path.
    .PARAMETER Props
      The Office names of the reader result (Creator, LastModifiedBy), or $null.
    .OUTPUTS
      [hashtable] ModifiedBy, Created, CreatedBy, AddedOn, AddedBy.
    #>
    param([string]$Path, [hashtable]$Props)
    $owner = ''
    try { $owner = (Get-Acl -LiteralPath $Path -ErrorAction Stop).Owner } catch { $owner = '' }   # no owner where the system keeps none
    $created = ''
    try { $created = (Get-Item -LiteralPath $Path -ErrorAction Stop).CreationTime.ToString('yyyy-MM-dd HH:mm') } catch { $created = '' }
    $modifiedBy = 'Not recorded (Windows keeps no last editor)'
    $createdBy = $owner
    if ($Props) {
        if ($Props.LastModifiedBy) { $modifiedBy = $Props.LastModifiedBy + ' (as saved in the file)' }
        if (-not $createdBy -and $Props.Creator) { $createdBy = $Props.Creator + ' (as saved in the file)' }
    }
    return @{ ModifiedBy = $modifiedBy; Created = $created; CreatedBy = $createdBy; AddedOn = $script:DriveNoHistory; AddedBy = $script:DriveNoHistory }
}

# ==== Module ChannelScan.psm1 ====
<#
ChannelScan.psm1 - reads one team's or site's files and turns them into report rows.

  Purpose:     The SharePoint and Teams steps of Find-ChannelPII.ps1, one function per step: read the
               Teams link or site address, sign in to the team's site, find the library, list the channel
               folders, scan each file with a progress bar and live counts. Posts are not read (the PAM,
               2026-11-20: one sign-in only; reading posts needs an admin-approved Graph permission).
  Inputs:      the Teams link or site address the person gives.
  Outputs:     report rows (Report.psm1 New-ReportRow) added to the list the caller passes in.
  Depends on:  PnP.PowerShell 1.12.0; PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1, Report.psm1,
               ScanMeter.psm1 (all imported first by Find-ChannelPII.ps1).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the one SharePoint site the person gives. Nothing lists other sites or teams (no network
               reconnaissance).
  Writes:      none in SharePoint or Teams: read calls only. Rows go to an in-memory list.
  Credentials: the person's own CAC sign-in (PnP web login). Nothing is stored.
#>


# =====================================================================================
# Settings
# =====================================================================================

# A team site's file library: Teams names it "Documents" (URL "Shared Documents"). It is scanned first.
$script:LibraryNames = @('Shared Documents', 'Documents')

# A document library is list template 101 (SharePoint ListTemplateType.DocumentLibrary).
$script:DocumentLibraryTemplate = 101

# Visible libraries that hold SharePoint's own files, not people's files: not scanned (the PAM, 2026-11-20: Site Assets,
# Preservation Hold Library and Site Collection Documents are scanned).
$script:NotScannedLibraries = @('Style Library', 'Form Templates')

# Teams makes one top-level folder per standard channel. "Forms" is SharePoint's own folder, not a channel.
$script:NotChannelFolders = @('Forms')

# Time spent downloading and spent reading and checking, for the summary (it shows where a slow scan loses its time).
$script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }

# How the report and the screen name the files at the top of the library (in no folder).
$script:RootGroupName = '(top of the library)'
# A drive folder: X:\<folder> (a mapped drive) or \\<server>\<share>\<folder> (UNC), with or without quotes.
$script:DrivePathPattern = '^"?(?:[A-Za-z]:\\|\\\\[^\\/\s]+\\[^\\/]+)'

# Office types that SharePoint opens by file ID (Doc.aspx?sourcedoc=): the link for a file whose path holds PII.
$script:IdLinkExtensions = @('.docx', '.docm', '.doc', '.xlsx', '.xlsm', '.xls', '.pptx', '.pptm', '.ppt')

# Items per request when listing the library. SharePoint refuses views over 5,000 items
# (list view threshold); 500 keeps each request small.
$script:ListPageSize = 500

# Accepted site address: https://<host>/sites/<name> or /teams/<name>, also as a file's "Copy link"
# (https://<host>/:w:/r/teams/<name>/...). Anything after the site name is cut off, so the scan can only
# touch that one site. Other sharing links (/:w:/s/, /:w:/t/) are refused: they do not name the site.
$script:SiteUrlPattern = '^(?<host>https://[A-Za-z0-9.-]+)/(?::[a-z]:/r/)?(?<site>(?:sites|teams)/[^/?#\s]+)'

# Accepted Teams link: a team's or a channel's "Copy link" (https://<teams-host>/l/team/... or /l/channel/...).
# Only its groupId (the team's Microsoft 365 group) is used; the host is not, so any Teams cloud's link works.
# The path may hold spaces (a channel name) but not a second address ('://' or '/l/'), so text pasted after the link is not read.
$script:TeamsLinkPattern = '^https://[A-Za-z0-9.-]+/l/(?<kind>team|channel)/(?:(?!://|/l/)[^?#\r\n])+\?(?:[^#\s]*&)?groupId=' +
                           '(?<id>[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})(?![0-9A-Fa-f-])'

# =====================================================================================
# Site, library and channels
# =====================================================================================

function Add-UrlScheme {
    <#
    .SYNOPSIS
      Adds https:// to an address typed without a scheme (usaf.dps.mil/teams/<team-site>); other text stays as it is.
    .PARAMETER Url
      The address the person gave.
    .OUTPUTS
      [string] the trimmed address, with https:// when it had no scheme.
    #>
    param([string]$Url)
    $Url = $Url.Trim()
    if ($Url -and $Url -notmatch '^[A-Za-z][A-Za-z0-9+.-]*://') { return 'https://' + $Url }
    return $Url
}

function Test-TeamSiteUrl {
    <#
    .SYNOPSIS
      Checks that the address is one SharePoint site and returns it cut to the site name.
    .PARAMETER Url
      The address the person gave (it may go on to a library or a file, or be a file's "Copy link"; https:// may be left out).
    .OUTPUTS
      [string] https://<host>/sites/<name> or https://<host>/teams/<name>
    #>
    param([string]$Url)
    $match = [regex]::Match((Add-UrlScheme $Url), $script:SiteUrlPattern, 'IgnoreCase')
    if (-not $match.Success) {
        throw "Not a team site address: '$Url'. Expected https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>."
    }
    return $match.Groups['host'].Value + '/' + $match.Groups['site'].Value
}

function Get-ScanTarget {
    <#
    .SYNOPSIS
      Reads what the person gave: a Teams link (team or channel) or a team site address.
      Throws when it is neither.
    .PARAMETER Address
      The Teams link or the site address.
    .OUTPUTS
      [pscustomobject] GroupId ([guid] from a Teams link, else $null), SiteUrl (the site, else ''), IsChannel
      (true for a channel's link) and DrivePath (a drive or network folder, else '').
    #>
    param([string]$Address)
    # A folder on a mapped drive (X:\...) or a network share (\\server\share\...): scanned by DriveScan.psm1.
    if ($Address -match $script:DrivePathPattern) { return [pscustomobject]@{ GroupId = $null; SiteUrl = ''; IsChannel = $false; DrivePath = $Address.Trim('"') } }
    $link = [regex]::Match((Add-UrlScheme $Address), $script:TeamsLinkPattern, 'IgnoreCase')
    if ($link.Success) {
        return [pscustomobject]@{ GroupId = [guid]$link.Groups['id'].Value; SiteUrl = ''; IsChannel = ($link.Groups['kind'].Value -ieq 'channel'); DrivePath = '' }
    }
    if ($Address -match '/l/(team|channel)/') {
        throw 'This Teams link has no groupId. In Teams: the ... next to the team or channel name > Copy link.'
    }
    return [pscustomobject]@{ GroupId = $null; SiteUrl = (Test-TeamSiteUrl $Address); IsChannel = $false; DrivePath = '' }
}

function Assert-SiteGroup {
    <#
    .SYNOPSIS
      Checks that the connected site belongs to the team in the Teams link; throws when not (Connect-TeamSite
      then signs out), so nothing of another site is read.
    .PARAMETER GroupId
      The group ID from the Teams link.
    #>
    param([guid]$GroupId)
    $actual = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    if ($actual -ne $GroupId) {
        throw 'This site does not belong to the team in the Teams link. Nothing was scanned.'
    }
}

function ConvertTo-FriendlyError {
    <#
    .SYNOPSIS
      A sign-in error in words a person can act on: no access, site not found, sign-in not finished. Used only for
      the sign-in step (Connect-TeamSite); the original message always follows, masked, so no real error is hidden.
    .PARAMETER Message
      The error message.
    .OUTPUTS
      [string]
    #>
    param([string]$Message)
    $details = Protect-PiiText $Message
    if ($Message -match '\((401|403)\)|\bforbidden\b|access (is )?denied|do(es)? not have (access|permission)') {
        return "No access: your account cannot open this site. Ask the site's owner for access, or check the address. ($details)"
    }
    if ($Message -match '\(404\)|could not be resolved|remote name could not') { return "Site not found: check the address. ($details)" }
    if ($Message -match 'user cancel+ed|cancel+ed by the user|prompt cancel+ed') { return "The sign-in did not finish (was the sign-in window closed?). ($details)" }
    return $details
}

function Read-ScanTarget {
    <#
    .SYNOPSIS
      Asks for a Teams link or site address until it is a valid one, or until the person presses Enter (exit).
      A wrong paste is explained and asked again; nothing is signed in until the address is valid.
    .PARAMETER Given
      An address given up front (the -SiteUrl setting); used first, without asking.
    .PARAMETER Prompt
      The question to ask.
    .OUTPUTS
      Output of Get-ScanTarget, or $null when the person pressed Enter.
    #>
    param([string]$Given, [string]$Prompt)
    $answer = $Given
    while ($true) {
        if (-not $answer) { $answer = Read-Host $Prompt }
        if (-not $answer -or -not $answer.Trim()) { return $null }
        try { return Get-ScanTarget $answer.Trim() }
        catch {
            Write-Ui ('That is not a Teams link or a site address: ' + (Protect-PiiText $_.Exception.Message)) 'Alert'
            Write-Ui 'Paste it again (Teams: ... > Copy link; or the SharePoint address), or press Enter to exit.' 'Muted'
            $answer = ''
        }
    }
}

function Read-TeamSiteAddress {
    # For a Teams link: asks for the team's site address until it is valid; Enter gives up (an error the caller reports).
    while ($true) {
        $typed = Read-Host 'Team site address (https://usaf.dps.mil/teams/<team-site>), or Enter to go back'
        if (-not $typed -or -not $typed.Trim()) { throw 'No site address given.' }
        try { return Test-TeamSiteUrl $typed }
        catch { Write-Ui ('Not a site address: ' + (Protect-PiiText $_.Exception.Message)) 'Alert' }
    }
}

function Connect-TeamSite {
    <#
    .SYNOPSIS
      Signs in (own CAC, PnP web login) to the site and returns its address. For a Teams link it asks for the site
      address (a lookup at the tenant root hung on DoD365: the PAM, 2026-11-20), then checks that the site belongs to
      the team in the link. It then checks that the person can open the site. On any error or cancel it signs out
      and throws a message a person can act on (ConvertTo-FriendlyError).
    .PARAMETER Target
      Output of Get-ScanTarget.
    .OUTPUTS
      [string] the site the scan may touch.
    #>
    param($Target)
    $site = $Target.SiteUrl
    $ok = $false   # set only at the end: a cancel (Ctrl+C, Cancel on a prompt) skips 'catch' but not 'finally'
    try {
        if ($null -ne $Target.GroupId) {
            if ($Target.IsChannel) {
                Write-Ui 'Channel link: the scan covers the standard channels of its team. A private or shared channel has its own site: give that site address instead.' 'Accent'
            }
            Write-Ui "Teams link: the script needs this team's site address. In Teams: the team's Shared tab > ... > Open in SharePoint, then copy the address." 'Accent'
            $site = Read-TeamSiteAddress
        }
        Write-Ui "Signing in to $site (a sign-in window opens; if you do not see it, press Alt+Tab; if it says you have no access, close it) ..." 'Muted'
        # PnP prints a warning that suggests -Interactive (and its update notice): not useful to the person, so hidden (the PAM, 2026-11-20).
        Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
        $null = Get-PnPWeb -ErrorAction Stop   # can this person open the site? A refusal here is "no access", before any scan
        if ($null -ne $Target.GroupId) { Assert-SiteGroup $Target.GroupId }
        $ok = $true
        return $site
    }
    catch { throw (ConvertTo-FriendlyError $_.Exception.Message) }
    finally {
        # No scan starts after an error or a cancel here: sign out, so no connection stays open in the ISE session.
        if (-not $ok) {
            try { Disconnect-PnPOnline } catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error
        }
    }
}

function Get-SiteDocumentLibraries {
    <#
    .SYNOPSIS
      The document libraries of the connected site that hold people's files: every visible library (template 101),
      without SharePoint's own ($NotScannedLibraries). The team's "Documents" library comes first. Only this one
      site: subsites are not read (the PAM, 2026-11-20).
    .OUTPUTS
      [pscustomobject] Title, List (the PnP list), ServerRelativeUrl (/sites/<name>/Shared Documents),
      SiteRelativeUrl (Shared Documents); one per library.
    #>
    $lists = @(Get-PnPList -ErrorAction Stop | Where-Object {
        [int]$_.BaseTemplate -eq $script:DocumentLibraryTemplate -and -not $_.Hidden -and $script:NotScannedLibraries -notcontains $_.Title
    })
    $libraries = foreach ($list in $lists) {
        $rootUrl = (Get-PnPProperty -ClientObject $list -Property RootFolder -ErrorAction Stop).ServerRelativeUrl.TrimEnd('/')
        $siteRelative = $rootUrl -replace '^.*?/(sites|teams)/[^/]+/', ''
        [pscustomobject]@{ Title = [string]$list.Title; List = $list; ServerRelativeUrl = $rootUrl; SiteRelativeUrl = $siteRelative }
    }
    $libraries = @($libraries | Sort-Object @{ Expression = { $script:LibraryNames -notcontains $_.SiteRelativeUrl } }, Title)
    if ($libraries.Count -eq 0) { throw 'No document library with files on this site. Check the address, or ask the site owner.' }
    return $libraries
}

function Get-ChannelFolderNames {
    <#
    .SYNOPSIS
      Lists the channel folders: the top-level folders of the library, without SharePoint's own. The scan
      covers all of them (the PAM, 2026-11-20: no channel choice). A plain SharePoint site may have none.
    .PARAMETER Library
      One library of Get-SiteDocumentLibraries.
    .OUTPUTS
      [string[]] folder names, sorted.
    #>
    param($Library)
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $Library.SiteRelativeUrl -ItemType Folder -ErrorAction Stop
    return @($folders | ForEach-Object { $_.Name } | Where-Object { $script:NotChannelFolders -notcontains $_ } | Sort-Object)
}

# =====================================================================================
# Files
# =====================================================================================

function Get-FieldValue {
    # Reads one field of a list item as text (dates as yyyy-MM-dd HH:mm); empty when missing.
    param($Item, [string]$Name)
    $value = $Item[$Name]
    if ($null -eq $value) { return '' }
    if ($value -is [datetime]) { return $value.ToString('yyyy-MM-dd HH:mm') }
    return [string]$value
}

function Get-SafeLink {
    <#
    .SYNOPSIS
      A link for the report that holds no PII and opens the exact file (the PAM, 2026-11-20). It is the file's own
      path when its path below the library holds no PII (Protect-PiiName on the whole path changes nothing). Else, for an Office file,
      a link by the file's ID (<site>/_layouts/15/Doc.aspx?sourcedoc={id}), which names nothing; else the deepest
      folder whose name is clean (at least the library). Each part is escaped, so names with '#', '%' or spaces work.
    .PARAMETER Origin
      https://<host>
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .PARAMETER UniqueId
      The file's ID (the list field UniqueId), or empty.
    .OUTPUTS
      [string]
    #>
    param([string]$Origin, [string]$ServerRelativeUrl, [string]$UniqueId = '')
    $parts = $ServerRelativeUrl.TrimStart('/').Split('/')
    # Joins path parts into a link path, each part escaped.
    $escape = { param($names) ($names | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/' }
    # The first three parts are the site and the library, which the person gave or Teams named.
    $keep = [Math]::Min(3, $parts.Count)
    # Each step checks the whole path below the library, so a label in a parent folder ("CAC/<number>") is seen.
    while ($keep -lt $parts.Count) {
        $below = $parts[3..$keep] -join '/'
        if ((Protect-PiiName $below) -cne $below) { break }
        $keep++
    }
    if ($keep -eq $parts.Count) { return $Origin + '/' + (& $escape $parts) }
    $id = [guid]::Empty
    $extension = [System.IO.Path]::GetExtension($parts[-1]).ToLower()
    if ($parts.Count -ge 2 -and $script:IdLinkExtensions -contains $extension -and [guid]::TryParse($UniqueId, [ref]$id)) {
        return '{0}/{1}/_layouts/15/Doc.aspx?sourcedoc=%7B{2}%7D' -f $Origin, (& $escape $parts[0..1]), $id
    }
    return $Origin + '/' + (& $escape $parts[0..($keep - 1)])
}

function Get-ShownItem {
    <#
    .SYNOPSIS
      A file's path inside its folder, as the report and the progress bar show it: PII masked, checked together with
      the folder name, so a label in the folder ("CAC", "DOB") and a number in the file name are seen as one hit.
    .PARAMETER Channel
      The folder (channel) name, not masked.
    .PARAMETER RawItem
      The path inside the folder, not masked.
    .OUTPUTS
      [string] the masked path inside the folder.
    #>
    param([string]$Channel, [string]$RawItem)
    $masked = Protect-PiiName ($Channel + '/' + $RawItem)
    $cut = $masked.IndexOf('/')
    # A hit across the folder's slash: fall back to the strict mask of the path alone.
    if ($cut -lt 0) { return Protect-PiiText $RawItem }
    return $masked.Substring($cut + 1)
}

function Get-FileBytes {
    <#
    .SYNOPSIS
      Downloads one file into memory (never to disk).
    .DESCRIPTION
      Tries Get-PnPFile -AsMemoryStream first. If this PnP version has no such switch, it uses the
      SharePoint client call OpenBinaryStream on the same connection.
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .OUTPUTS
      [byte[]]
    #>
    param([string]$ServerRelativeUrl)
    try {
        $stream = Get-PnPFile -Url $ServerRelativeUrl -AsMemoryStream -ErrorAction Stop
        return $stream.ToArray()
    }
    catch [System.Management.Automation.ParameterBindingException] {
        # This PnP version has no -AsMemoryStream: use the client call instead (same connection, same rights).
        $context = Get-PnPContext
        $binary = $context.Web.GetFileByServerRelativeUrl($ServerRelativeUrl).OpenBinaryStream()
        $context.ExecuteQuery()
        $memory = New-Object System.IO.MemoryStream
        try { $binary.Value.CopyTo($memory); return $memory.ToArray() } finally { $memory.Dispose() }
    }
}

function Get-ChannelFiles {
    <#
    .SYNOPSIS
      The files (not folders) under one channel folder, at any depth.
    .PARAMETER AllItems
      Every item of the library (one Get-PnPListItem call for all channels).
    .PARAMETER Prefix
      /sites/<name>/<library>/<channel>/
    .OUTPUTS
      List items.
    #>
    param([object[]]$AllItems, [string]$Prefix)
    $AllItems | Where-Object {
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and (Get-FieldValue $_ 'FileRef').StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

function Get-RootFiles {
    # The files at the top of the library, in no folder (a plain SharePoint site keeps files there; a team has none).
    param([object[]]$AllItems, [string]$LibraryUrl)
    $prefix = $LibraryUrl + '/'
    $AllItems | Where-Object {
        $ref = Get-FieldValue $_ 'FileRef'
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and $ref.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            $ref.Substring($prefix.Length).IndexOf('/') -lt 0
    }
}

function Invoke-OneFileScan {
    <#
    .SYNOPSIS
      Starts one file: a hit for PII in its name, then a NOT SCANNED row (unsupported, too large, download failed)
      or its download, handed to the readers (its text rows follow when a worker finishes).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER File
      The list item.
    .PARAMETER Channel
      The channel folder name.
    .PARAMETER Prefix
      The channel folder's server-relative path with a trailing '/'.
    .PARAMETER Origin
      https://<host>
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file went to the readers (else its rows are in and it still needs counting).
    #>
    param([hashtable]$Queue, $File, [string]$Channel, [string]$Prefix, [string]$Origin, [int]$MaxFileMB)
    $rows = $Queue.Rows
    $name = Get-FieldValue $File 'FileLeafRef'
    $ref = Get-FieldValue $File 'FileRef'
    $rawChannel = $Channel
    $Channel = Protect-PiiName $Channel   # from here on, the channel name as the report shows it
    $modified = Get-FieldValue $File 'Modified'
    # The path is shown in the report: mask PII in it (with its folder name), and report a name that holds PII as a hit.
    $rawItem = $ref.Substring($Prefix.Length)
    $item = Get-ShownItem $rawChannel $rawItem
    $link = Get-SafeLink $Origin $ref (Get-FieldValue $File 'UniqueId')
    foreach ($hit in @(Find-PiiInText ($rawChannel + '/' + $rawItem))) {
        $rows.Add((New-ReportRow 'File name' $Channel $item $link $modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $size = 0L
    [void][long]::TryParse((Get-FieldValue $File 'File_x0020_Size'), [ref]$size)
    $skipReason = ''
    if (-not (Test-SupportedFile $name)) { $skipReason = 'File type is not read by this script: check by hand' }
    elseif ($size -gt ($MaxFileMB * 1MB)) { $skipReason = "Larger than $MaxFileMB MB: check by hand" }
    if ($skipReason) { $rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $skipReason)); return $false }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    Write-FileStep $size 'Downloading' 30
    try { $bytes = Get-FileBytes $ref }
    catch {
        # One file that cannot be downloaded must not stop the scan: note it and go on.
        $reason = Protect-PiiText ('Download failed: ' + $_.Exception.Message)
        $rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $reason)); return $false
    }
    finally { $script:Timing.Download += $clock.Elapsed; Write-Progress -Id 2 -Activity 'This file' -Completed }
    # The main thread downloads the next file while workers read this one (speed, the PAM 2026-11-21).
    # Accountability: who made and changed the file; for a file with a hit, who added each hit (version history).
    $who = Get-SharePointWho $File
    $who.Modified = $modified
    $added = { param($hits) Find-PiiAddedVersion $ref $name $hits $who }.GetNewClosure()
    Submit-ScanFile $Queue @{ Channel = $Channel; Item = $item; Link = $link; Modified = $modified; Who = $who; Added = $added } $name $bytes ''
    return $true
}

function Get-ScanTiming {
    # The time of the last file scan spent downloading and spent reading and checking, as minutes:seconds text.
    $format = { param($t) '{0:00}:{1:00}' -f [int][Math]::Floor($t.TotalMinutes), $t.Seconds }
    [ordered]@{ 'Time downloading' = (& $format $script:Timing.Download); 'Time reading and checking (all workers)' = (& $format $script:Timing.Check) }
}

function Get-ScanGroups {
    <#
    .SYNOPSIS
      The scan groups of one library: one per top-level folder, then the files at the top of the library.
    .PARAMETER Library
      One library of Get-SiteDocumentLibraries.
    .PARAMETER Folders
      The top-level folder names.
    .PARAMETER Label
      Put before each group name ("Library / ") when the site has several libraries; else empty.
    .OUTPUTS
      [pscustomobject] Name, Prefix, Files; a group without files is kept only when it is a folder.
    #>
    param($Library, [string[]]$Folders, [string]$Label)
    $fields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified', 'UniqueId', 'Created', 'Author', 'Editor')
    $allItems = @(Get-PnPListItem -List $Library.List -PageSize $script:ListPageSize -Fields $fields -ErrorAction Stop)
    foreach ($folder in $Folders) {
        $prefix = $Library.ServerRelativeUrl + '/' + $folder + '/'
        [pscustomobject]@{ Name = $Label + $folder; Prefix = $prefix; Files = @(Get-ChannelFiles $allItems $prefix) }
    }
    $rootFiles = @(Get-RootFiles $allItems $Library.ServerRelativeUrl)
    if ($rootFiles.Count -gt 0) {
        [pscustomobject]@{ Name = $Label + $script:RootGroupName; Prefix = $Library.ServerRelativeUrl + '/'; Files = $rootFiles }
    }
}

function Invoke-FileScan {
    <#
    .SYNOPSIS
      Scans every file in the given libraries' top-level folders and at their top, and adds report rows.
    .PARAMETER Library
      One or more libraries of Get-SiteDocumentLibraries.
    .PARAMETER Channels
      The channel folder names to scan when there is one library; with several libraries, each library's own
      top-level folders are scanned.
    .PARAMETER SiteUrl
      The site address (for the links in the report).
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .OUTPUTS
      [int] the number of files whose text was scanned.
    #>
    param($Library, [string[]]$Channels, [string]$SiteUrl, [int]$MaxFileMB, $Rows, [hashtable]$Init)
    $origin = ([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }
    $libraries = @($Library)
    $groups = @(foreach ($one in $libraries) {
        if ($libraries.Count -eq 1) { Get-ScanGroups $one $Channels ''; continue }
        # Several libraries: one you cannot open (Preservation Hold Library often) is one BLOCKED row; the rest go on.
        # Collected first, so a library that fails part way adds no groups next to its BLOCKED row.
        try { $libraryGroups = @(Get-ScanGroups $one @(Get-ChannelFolderNames $one) ($one.Title + ' / ')); $libraryGroups }   # "Library / Folder"
        catch {
            $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $one.Title) '(library)' '' '' 'BLOCKED' '' '' ('Could not read this library: ' + (Protect-PiiText $_.Exception.Message))))
        }
    })
    $total = 0
    foreach ($group in $groups) { $total += $group.Files.Count }
    Write-Ui ('==== Scanning {0} files in {1} folders ====' -f $total, $groups.Count) 'Title'
    $meter = New-ScanMeter $total
    $meter.NameWidth = Get-NameWidth @($groups | ForEach-Object { Protect-PiiName $_.Name })
    $queue = New-ScanQueue $Init $Rows $meter
    # Starts one file of a group (Invoke-GroupScan calls it); then the file's masked path for the progress bar.
    $start = { param($group, $file) Invoke-OneFileScan $queue $file $group.Name $group.Prefix $origin $MaxFileMB }
    # The file's masked path inside its folder, for the progress bar.
    $shown = { param($group, $file) Get-ShownItem $group.Name ((Get-FieldValue $file 'FileRef').Substring($group.Prefix.Length)) }
    try { Invoke-GroupScan $queue $groups $start $shown }
    finally { $script:Timing.Check = Close-ScanQueue $queue }
    # The folder rows form one table; the total closes it, under a rule.
    Write-RuleLine
    Write-MeterLine $meter 'file' -Force
    return $meter.Read
}

# =====================================================================================
# Team or plain site
# =====================================================================================

function Test-TeamSite {
    <#
    .SYNOPSIS
      True when the connected site belongs to a team (a Microsoft 365 group): its folders are then channels.
      The report names the folder column "Channel" for a team and "Folder" for a plain site (the PAM, 2026-11-20).
    .OUTPUTS
      [bool]
    #>
    $groupId = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    return [bool]($groupId -and $groupId -ne [guid]::Empty)
}

# ==== Module DriveScan.psm1 ====
<#
DriveScan.psm1 - scans a folder on a mapped or network drive (X:\<folder> or \\<server>\<share>\<folder>).

  Purpose:     The PAM, 2026-11-21: the same check as for a SharePoint site, for one folder and all its subfolders on
               a drive the person can open. Nothing above the folder, no other share or server is touched ("No
               reconnaissance"). Files are opened read-only and shared, with the person's own Windows logon.
  Inputs:      the pasted folder path; the reader workers' set-up (New-ReaderPool).
  Outputs:     report rows; the results workbook (Report.psm1 Publish-ScanReport).
  How to run:  imported by Find-ChannelPII.ps1; Invoke-DriveScan runs when the pasted target is a folder path.
  Depends on:  PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1 (Test-SupportedFile), Report.psm1,
               ScanMeter.psm1, ScanQueue.psm1 and ChannelScan.psm1 (Get-ShownItem), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7, where the tests run on a temporary folder).
  Network:     only the given folder's server, through the person's Windows logon (SMB); nothing else.
  Writes:      only the results workbook under the results folder.
  Credentials: none (the Windows logon).
#>


# Files directly in the pasted folder form this group (each subfolder is its own group).
$script:DriveTopGroup = '(top of the folder)'

function Get-DriveGroups {
    <#
    .SYNOPSIS
      The files under one folder, grouped by its top-level subfolders (then the files at its top), with the folders
      that could not be opened and the folder links that were not followed (a link can lead outside the folder).
    .PARAMETER Root
      The folder path.
    .OUTPUTS
      [hashtable] Groups (each @{ Name; Files }), Blocked (each @{ Folder; Reason }), Links (folder paths).
    #>
    param([string]$Root)
    $Root = $Root.TrimEnd('\', '/')
    $slashes = [char[]]@('\', '/')   # Windows uses \; the tests also run where / is the separator
    $result = @{ Groups = (New-Object System.Collections.Generic.List[object]); Blocked = (New-Object System.Collections.Generic.List[object])
        Links = (New-Object System.Collections.Generic.List[string]) }
    $byName = @{}
    $top = @{ Name = $script:DriveTopGroup; Files = (New-Object System.Collections.Generic.List[object]) }
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $folder = $stack.Pop()
        $relative = $folder.Substring($Root.Length).TrimStart($slashes)
        try { $entries = @((New-Object System.IO.DirectoryInfo($folder)).GetFileSystemInfos()) }
        catch { $result.Blocked.Add(@{ Folder = $relative; Reason = $_.Exception.Message }); continue }
        foreach ($entry in $entries) {
            $rel = ($relative + '\' + $entry.Name).TrimStart($slashes)
            if ($entry -is [System.IO.DirectoryInfo]) {
                if ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { $result.Links.Add($rel); continue }
                $stack.Push($entry.FullName)
                continue
            }
            $groupName = $script:DriveTopGroup
            $inner = $rel
            $cut = $rel.IndexOfAny($slashes)
            if ($cut -gt 0) { $groupName = $rel.Substring(0, $cut); $inner = $rel.Substring($cut + 1) }
            $group = $top
            if ($groupName -ne $script:DriveTopGroup) {
                if (-not $byName.ContainsKey($groupName)) { $byName[$groupName] = @{ Name = $groupName; Files = (New-Object System.Collections.Generic.List[object]) } }
                $group = $byName[$groupName]
            }
            $group.Files.Add(@{ Path = $entry.FullName; Inner = ($inner -replace '[\\/]', '/'); Name = $entry.Name; Size = [long]$entry.Length
                Modified = $entry.LastWriteTime.ToString('yyyy-MM-dd HH:mm') })
        }
    }
    foreach ($name in @($byName.Keys | Sort-Object)) { $result.Groups.Add($byName[$name]) }
    if ($top.Files.Count -gt 0) { $result.Groups.Add($top) }
    return $result
}

function Get-DriveLink {
    <#
    .SYNOPSIS
      A file:/// link to a file, so the workbook opens it; none when the path holds PII (the link would show it).
    .PARAMETER Path
      The file path.
    .OUTPUTS
      [string] the link, or empty.
    #>
    param([string]$Path)
    if (@(Find-PiiInText $Path).Count -gt 0) { return '' }
    $absolute = $Path; if ($Path.StartsWith('/')) { $absolute = 'file://' + $Path }   # / paths: the test machines only
    try { return ([uri]$absolute).AbsoluteUri } catch { return '' }
}

function Invoke-DriveFile {
    <#
    .SYNOPSIS
      Starts one file of a drive scan: a hit for PII in its name, then a NOT SCANNED row (unsupported or too
      large) or the file handed to the readers, which open it themselves.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Group
      The file's group name (a top-level subfolder), not masked.
    .PARAMETER File
      One file of Get-DriveGroups.
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file went to the readers.
    #>
    param([hashtable]$Queue, [string]$Group, [hashtable]$File, [int]$MaxFileMB)
    $rows = $Queue.Rows
    $channel = Protect-PiiName $Group
    $item = Get-ShownItem $Group $File.Inner
    $link = Get-DriveLink $File.Path
    foreach ($hit in @(Find-PiiInText ($Group + '/' + $File.Inner))) {
        $rows.Add((New-ReportRow 'File name' $channel $item $link $File.Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $skipReason = ''
    if (-not (Test-SupportedFile $File.Name)) { $skipReason = 'File type is not read by this script: check by hand' }
    elseif ($File.Size -gt ($MaxFileMB * 1MB)) { $skipReason = "Larger than $MaxFileMB MB: check by hand" }
    if ($skipReason) { $rows.Add((New-ReportRow 'File' $channel $item $link $File.Modified 'NOT SCANNED' '' '' $skipReason)); return $false }
    Submit-ScanFile $Queue @{ Channel = $channel; Item = $item; Link = $link; Modified = $File.Modified; DrivePath = $File.Path } $File.Name $null $File.Path
    return $true
}

function Add-DriveListRows {
    <#
    .SYNOPSIS
      A BLOCKED row for each folder that could not be opened, and a NOT SCANNED row for each folder link not
      followed (a link can lead outside the folder).
    .PARAMETER Found
      Output of Get-DriveGroups.
    .PARAMETER Rows
      The list that receives the rows.
    #>
    param([hashtable]$Found, $Rows)
    foreach ($blocked in $Found.Blocked) {
        $reason = 'Could not open this folder: ' + (Protect-PiiText $blocked.Reason)
        $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $blocked.Folder) '(folder)' '' '' 'BLOCKED' '' '' $reason))
    }
    foreach ($link in $Found.Links) {
        $reason = 'A link to another folder is not followed (it can lead outside the folder): check by hand'
        $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $link) '(folder link)' '' '' 'NOT SCANNED' '' '' $reason))
    }
}

function Get-DriveHeader {
    <#
    .SYNOPSIS
      The Summary tab's header lines of a drive scan (every name masked like any name the tool shows).
    .PARAMETER Path
      The folder.
    .PARAMETER Groups
      The groups scanned.
    .PARAMETER Used
      The whole run's time.
    .PARAMETER Reading
      The time all workers spent reading and checking.
    .PARAMETER Workers
      The number of files read at once at the end.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path, [object[]]$Groups, [TimeSpan]$Used, [TimeSpan]$Reading, [int]$Workers)
    # A time as minutes:seconds.
    $format = { param($t) '{0:00}:{1:00}' -f [int][Math]::Floor($t.TotalMinutes), $t.Seconds }
    $header = [ordered]@{}
    $header['Folder'] = Protect-PiiName $Path
    $header['Folders'] = ($Groups | ForEach-Object { Protect-PiiName $_.Name }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = & $format $Used
    $header['Time reading and checking (all workers)'] = & $format $Reading
    $header['Files read at once (at the end)'] = [string]$Workers
    return $header
}

function Invoke-DriveScan {
    <#
    .SYNOPSIS
      Scans one drive folder and all its subfolders with the reader workers, then saves the workbook (also after an
      error or a cancel: the results so far are saved and say INCOMPLETE). Returns the workbook path.
    .PARAMETER Path
      The folder (X:\<folder> or \\<server>\<share>\<folder>).
    .PARAMETER OutputFolder
      The base folder for the results.
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([string]$Path, [string]$OutputFolder, [int]$MaxFileMB, [hashtable]$Init)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[object]
    $finished = $false
    $groups = @()
    $workers = 0
    $reading = [TimeSpan]::Zero
    $meter = New-ScanMeter 0
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw 'This folder cannot be opened (check the path, and that you can open it in File Explorer).' }
        Write-Ui 'Listing the files ...' 'Muted'
        $found = Get-DriveGroups $Path
        $groups = $found.Groups.ToArray()
        Add-DriveListRows $found $rows
        $total = ($groups | ForEach-Object { $_.Files.Count } | Measure-Object -Sum).Sum
        Write-Ui ('==== Scanning {0} files in {1} folders ====' -f $total, $groups.Count) 'Title'
        $meter = New-ScanMeter ([int]$total)
        $meter.NameWidth = Get-NameWidth @($groups | ForEach-Object { Protect-PiiName $_.Name })
        $queue = New-ScanQueue $Init $rows $meter
        try {
            # Starts one file (Invoke-GroupScan calls it).
            $start = { param($group, $file) Invoke-DriveFile $queue $group.Name $file $MaxFileMB }
            # The file's masked path inside its folder, for the progress bar.
            $shown = { param($group, $file) Get-ShownItem $group.Name $file.Inner }
            Invoke-GroupScan $queue $groups $start $shown
            $workers = Get-QueueWorkers $queue
        }
        finally { $reading = Close-ScanQueue $queue }
        Write-RuleLine
        Write-MeterLine $meter 'file' -Force
        $finished = $true
    }
    catch {
        $reason = Protect-PiiText $_.Exception.Message
        Write-Ui ('The scan stopped early: ' + $reason) 'Alert'
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
    }
    finally {
        Add-IncompleteRow -Rows $rows -Finished $finished; $header = Get-DriveHeader $Path $groups $clock.Elapsed $reading $workers
        $name = $Path.TrimEnd('\') -replace '\\', '/'   # the results folder is named after the folder's last part
        $path = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $meter.Read -Header $header -OutputFolder $OutputFolder -SiteUrl $name -FolderLabel 'Folder'
    }
    return $path
}

# ==== Body ==== (scripts/build_paste.py copies everything from this line into the paste bundle; the lines above
# become the bundle's Settings block and its inline module code)

function Invoke-SiteFiles {
    <#
    .SYNOPSIS
      Scans one signed-in site: its folders and files, then saves the workbook and signs out (also after an error or
      a cancel: the results so far are saved and say INCOMPLETE). Returns the workbook path.
    .PARAMETER SiteUrl
      The signed-in site.
    .PARAMETER OutputFolder
      The base folder for the results.
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .OUTPUTS
      [string] the workbook path.
    #>
    param([string]$SiteUrl, [string]$OutputFolder, [int]$MaxFileMB)
    $path = $null
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[object]
    $scanned = 0
    $selected = @()
    $folderLabel = 'Folder'   # 'Channel' when the site belongs to a team (the PAM, 2026-11-20)
    $listName = ''            # what the header lists: 'Channels', 'Folders' or 'Libraries'
    $finished = $false   # set only after the last step; a cancelled run (Ctrl+C) skips 'catch' but not 'finally'
    try {
        # Only the column name depends on it: if the site's group cannot be read, the scan goes on with 'Folder'.
        try { if (Test-TeamSite) { $folderLabel = 'Channel' } }
        catch { Write-Verbose 'Team check not possible: the folder column is named Folder.' }
        $libraries = @(Get-SiteDocumentLibraries)
        if ($libraries.Count -eq 1) { $selected = @(Get-ChannelFolderNames $libraries[0]) }
        else { $folderLabel = 'Library / Folder'; $selected = @($libraries | ForEach-Object { $_.Title }) }   # every library of the site
        $listName = $folderLabel + 's'; if ($libraries.Count -gt 1) { $listName = 'Libraries' }
        Write-Ui ($listName + ' to scan: ') 'Label' -NoNewline
        Write-Ui ((($selected | ForEach-Object { Protect-PiiName $_ }) -join ', ') + '; also any files at the top of each library') 'Text'
        $scanned += Invoke-FileScan -Library $libraries -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows -Init $IonWorkerInit
        $finished = $true
    }
    catch {
        # Keep what was found so far: the results say the scan stopped early, and why, in words a person can act on.
        $reason = Protect-PiiText $_.Exception.Message
        Write-Ui ('The scan stopped early: ' + $reason) 'Alert'
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
    }
    finally {
        Add-IncompleteRow -Rows $rows -Finished $finished
        $header = [ordered]@{}
        $header['Site'] = $SiteUrl
        if (-not $listName) { $listName = $folderLabel + 's' }
        $header[$listName] = ($selected | ForEach-Object { Protect-PiiName $_ }) -join ', '
        $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
        $header['Time used'] = '{0:00}:{1:00}' -f [int][Math]::Floor($clock.Elapsed.TotalMinutes), $clock.Elapsed.Seconds
        foreach ($pair in (Get-ScanTiming).GetEnumerator()) { $header[$pair.Key] = $pair.Value }
        try { $path = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $scanned -Header $header -OutputFolder $OutputFolder -SiteUrl $SiteUrl -FolderLabel $folderLabel }
        finally {
            # Sign out even when the results could not be saved (for example, a folder that cannot be written).
            try { Disconnect-PnPOnline }
            catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error for the person
        }
    }
    return $path
}

function Open-ScanResult {
    <#
    .SYNOPSIS
      Asks whether to open the saved workbook now (Enter = yes; the PAM, 2026-11-20) and opens it in Excel.
      Only the local file is opened; nothing is sent.
    .PARAMETER Path
      The workbook path, or empty when nothing was saved.
    #>
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return }
    $answer = Read-Host 'Open the workbook now? [Y/n]'
    if ($answer -and $answer.Trim() -match '^(n|no)$') { return }
    try { Invoke-Item -LiteralPath $Path }
    catch { Write-Ui ('The workbook could not be opened: ' + (Protect-PiiText $_.Exception.Message) + '. Open it from the folder above.') 'Accent' }
}

Write-Ui '==== Ion - DoD365 PII Reporter ====' 'Title'
Write-Ui 'Read-only. Every PII value found is masked on screen and in the results.' 'Muted'

# ---- Ask for a site, sign in, scan, save; then offer another site, until the person presses Enter ----
# A wrong paste is explained and asked again (Read-ScanTarget); a site the person cannot open is reported as
# "no access" and nothing is saved for it (Connect-TeamSite); only the one site given is touched each time.
$target = Read-ScanTarget $SiteUrl 'Teams link, SharePoint site address (https://usaf.dps.mil/teams/<team-site>) or folder (X:\<folder>), or Enter to exit'
while ($null -ne $target) {
    $site = $null
    if ($target.DrivePath) {
        # A drive or network folder: no sign-in (the Windows logon); read-only; nothing above the folder.
        try { Open-ScanResult (Invoke-DriveScan -Path $target.DrivePath -OutputFolder $OutputFolder -MaxFileMB $MaxFileMB -Init $IonWorkerInit) }
        catch { Write-Ui ('The results could not be saved: ' + (Protect-PiiText $_.Exception.Message)) 'Alert' }
        Write-Ui ''
        $target = Read-ScanTarget '' 'Scan another? Paste a Teams link, a site address or a folder (X:\<folder>), or press Enter to exit'
        continue
    }
    try { $site = Connect-TeamSite $target }
    catch { Write-Ui $_.Exception.Message 'Alert' }
    if ($site) {
        # A results file that cannot be saved (a folder that cannot be written) is reported; the next site is still offered.
        try { Open-ScanResult (Invoke-SiteFiles -SiteUrl $site -OutputFolder $OutputFolder -MaxFileMB $MaxFileMB) }
        catch { Write-Ui ('The results could not be saved: ' + (Protect-PiiText $_.Exception.Message)) 'Alert' }
    }
    Write-Ui ''
    $target = Read-ScanTarget '' 'Scan another? Paste a Teams link, a site address or a folder (X:\<folder>), or press Enter to exit'
}
Write-Ui 'Ion is done: nothing stays signed in.' 'Good'

# ==== Clean-up ====
} finally {
    Set-StrictMode -Off
    $ErrorActionPreference = $previousErrorAction
}
