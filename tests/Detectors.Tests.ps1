<#
Detectors.Tests.ps1 - tests of PiiDetectors.psm1 and PiiMasking.psm1: what is found, what is not, and masking.

  Purpose:     True and false cases for each PII type; overlap; masked values and snippets never
               show a full value.
  Inputs:      $Sample (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- SSN ----
Assert-Equal 'SSN/High/***-**-6789' (Get-HitSummary "Member SSN: $($Sample.Ssn) on file") 'SSN with label'
Assert-Equal 'SSN/Medium/***-**-6789' (Get-HitSummary "ref $($Sample.Ssn) end") 'SSN with dashes, no label'
Assert-Equal 'SSN/Low/***-**-6789' (Get-HitSummary "value $($Sample.SsnBare) end") 'SSN bare digits'
foreach ($bad in @('SsnArea000', 'SsnArea666', 'SsnArea9xx', 'SsnGroup00', 'SsnSerial0', 'SsnMixed')) {
    Assert-Equal 0 (Get-HitSummary "SSN $($Sample[$bad])").Count "SSN refused: $bad"
}

# ---- DoD ID ----
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary "DoD ID: $($Sample.DodId)") 'DoD ID with label'
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary "EDIPI $($Sample.DodId)") 'EDIPI with label'
Assert-Equal 0 (Get-HitSummary "tracking $($Sample.DodId) shipped").Count '10 digits without label not reported'
# A CAC name in a digital signature block is a DoD ID, never an SSN, also when the digits are drawn with gaps (the PAM, 2026-11-21).
$cacName = 'DOE' + '.' + 'JANE' + '.Q.'
$spacedId = ($Sample.DodId.Substring(0, 3) + ' ' + $Sample.DodId.Substring(3, 2) + ' ' + $Sample.DodId.Substring(5, 4) + ' ' + $Sample.DodId.Substring(9))
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "Digitally signed by $cacName$($Sample.DodId)") -join ',') 'CAC name: the DoD ID, no SSN'
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "Digitally signed by $cacName$spacedId Date: 2026.11.20") -join ',') 'CAC name, digits with gaps: the DoD ID, no SSN'
Assert-Equal 0 (Get-HitSummary ('Digitally signed by ' + $Sample.DodId.Substring(0, 9) + ' ' + $Sample.DodId.Substring(9) + ' Date')).Count 'Signature: 9 digits then one more after a space: no SSN'
Assert-Equal 'SSN/High/***-**-6789' ((Get-HitSummary "SSN: $($Sample.Ssn) 2 copies") -join ',') 'A labelled SSN with a number after it is still found'
$bareSsnX = $Sample.Ssn -replace '-', ''
# Second review of PR 49: a count after an SSN, words that are no CAC name, and a one-letter first name.
foreach ($case in @("assigned $bareSsnX 1 slot", "Unsigned $($Sample.Ssn -replace '-', ' ') 2", "$bareSsnX 3 pages", "$($Sample.Ssn -replace '-', ' ') 2 copies")) {
    Assert-Equal 1 @(Get-HitSummary $case | Where-Object { $_ -like 'SSN/*' }).Count "SSN with a count after it still found: $($case -replace '\d', '#')"
}
foreach ($word in @('Acct.No.', 'Ref.No.', 'REF.NO.', 'backup.config.', 'Ph.No.')) {
    Assert-Equal 0 @(Get-HitSummary ($word + $Sample.DodId) | Where-Object { $_ -like 'DoD ID/*' }).Count "No CAC name: $word"
}
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary ('DOE' + '.J.' + $Sample.DodId)) -join ',') 'CAC name with a one-letter first name'
Assert-Equal 'Low' (@(Find-PiiInText "Name: Jane Doe signed $cacName$($Sample.DodId)" | Where-Object { $_.Detector -eq 'Name' })[0].Confidence) 'An e-signature DoD ID does not raise a name to High'
# Review of PR 49: an unlabelled SSN next to a row number, a date, initials or an abbreviation is still found.
$bareSsn = $Sample.Ssn -replace '-', ''
foreach ($case in @("Row 12 $($Sample.Ssn) DOE JANE", "1 $($Sample.Ssn)", "$($Sample.Ssn) 05/12/1990", "SMITH, J.A. $($Sample.Ssn)",
        "see e.g. $($Sample.Ssn)", "U.S. $bareSsn", "($($Sample.Ssn -replace '-', ' ')) 05/12/1990")) {
    Assert-Equal 1 @(Get-HitSummary $case | Where-Object { $_ -like 'SSN/*' }).Count "Unlabelled SSN still found: $($case -replace '\d', '#')"
}

# ---- Payment card ----
Assert-Equal 'Payment card/High/************1111' (Get-HitSummary 'GTC card 4111 1111 1111 1111 exp') 'Card with label'
Assert-Equal 'Payment card/Medium/************4444' (Get-HitSummary 'number 5555-5555-5555-4444') 'Card without label'
Assert-Equal 0 (Get-HitSummary 'number 4111 1111 1111 1112').Count 'Card failing Luhn not reported'
Assert-Equal 1 (Get-HitSummary 'card 4111111111111111').Count 'Card digits not also reported as SSN (overlap)'

# Many overlapping hits: the fast overlap removal keeps exactly what a check against every kept hit keeps.
$random = New-Object System.Random 332   # fixed seed: the same invented hits every run
$types = @('SSN', 'DoD ID', 'Payment card', 'Phone', 'Name')
$manyHits = @(foreach ($n in 1..1500) {
    [pscustomobject]@{ Index = $random.Next(0, 3000); Length = $random.Next(1, 40)
                       Detector = $types[$random.Next(0, $types.Count)]; Confidence = @('High', 'Medium', 'Low')[$random.Next(0, 3)] }
})
$fastKept = & (Get-Module PiiDetectors) { param($hits) Remove-OverlappingHits $hits } $manyHits
$order = & (Get-Module PiiDetectors) { $script:DetectorOrder }
$slowSorted = $manyHits | Sort-Object @{ Expression = { [array]::IndexOf($order, $_.Detector) } },
    @{ Expression = { @{ High = 3; Medium = 2; Low = 1 }[$_.Confidence] }; Descending = $true }, Index
$slowKept = New-Object System.Collections.Generic.List[object]
foreach ($hit in $slowSorted) {
    $clash = @($slowKept | Where-Object { $hit.Index -lt ($_.Index + $_.Length) -and $_.Index -lt ($hit.Index + $hit.Length) }).Count
    if ($clash -eq 0) { $slowKept.Add($hit) }
}
function ConvertTo-HitListText {
    # The kept hits as one line (start+length:type each), so two lists compare as text.
    param($List)
    return ($List | ForEach-Object { '{0}+{1}:{2}' -f $_.Index, $_.Length, $_.Detector }) -join ' '
}
Assert-Equal (ConvertTo-HitListText $slowKept) (ConvertTo-HitListText $fastKept) 'Many overlapping hits: same hits kept as the full check'
Assert-Equal $true ($fastKept.Count -gt 50) 'Many overlapping hits: a real mix kept'

# ---- Date of birth ----
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'DOB: 03/14/1985') 'DOB numeric'
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'Date of birth 14 Mar 1985') 'DOB day month year'
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'born March 14, 1985') 'DOB month name'
Assert-Equal 0 (Get-HitSummary 'Suspense 03/14/2027').Count 'Date without birth label not reported'

# ---- Passport, phone, e-mail ----
Assert-Equal 'Passport/Medium/*****567' (Get-HitSummary "Passport no. $($Sample.Passport)") 'Passport with label'
Assert-Equal 'Phone/Medium/***-***-0142' (Get-HitSummary "my cell $($Sample.Phone)") 'Phone with personal label'
Assert-Equal 'Phone/Low/***-***-0142' (Get-HitSummary "office $($Sample.PhoneParen)") 'Phone without label'
Assert-Equal 'Personal e-mail/Low/***@example.com' (Get-HitSummary "write to $($Sample.Email)") 'Personal e-mail'
Assert-Equal 0 (Get-HitSummary $Sample.GovEmail).Count '.gov e-mail not reported'
Assert-Equal 'Personal e-mail/Medium/***@gmail.com' (Get-HitSummary ('write to pat.sample' + '@gmail.com')) 'E-mail at a personal provider: Medium'
Assert-Equal 'Personal e-mail/Low/***@mail.example.com' (Get-HitSummary ('write to pat.sample' + '@mail.example.com')) 'A company subdomain named mail is not a personal provider'

# ---- False positives in typical documents (the PAM, 2026-11-20) ----
Assert-Equal 0 (Get-HitSummary ('PAGEREF _Toc' + $Sample.SsnBare)).Count 'A Word bookmark number is no SSN'
Assert-Equal 0 (Get-HitSummary ('Form ABC' + $Sample.SsnBare + ' rev')).Count 'Digits glued to letters are no SSN'
Assert-Equal 0 (Get-HitSummary ('tracking 1234' + '5678' + '9012' + '3452')).Count 'A Luhn-valid number with no card issuer prefix is no card'
Assert-Equal $true (Test-Luhn ('12345678' + '90123452')) 'The tracking sample passes Luhn (so the issuer check is what drops it)'
Assert-Equal 0 (Get-HitSummary ('build v1-555-201' + '-0142x')).Count 'A phone-shaped run glued to letters is no phone'
Assert-Equal 0 (Get-HitSummary $Sample.MilEmail).Count '.mil e-mail not reported'

# ---- Home address ----
Assert-Equal 'Home address/High/[address masked]' (Get-HitSummary "Home address: $($Sample.Street), $($Sample.CityZip)") 'Address with label'
Assert-Equal 'Home address/Medium/[address masked]' (Get-HitSummary "send to $($Sample.Street), $($Sample.CityZip) today") 'Address with ZIP'
Assert-Equal 'Home address/Low/[address masked]' (Get-HitSummary "meet at $($Sample.Street) today") 'Street only'

# ---- Name ----
Assert-Equal 'Name/Low/P*** Q*** S***' (Get-HitSummary 'Name: Pat Q. Sample') 'Labelled name'
Assert-Equal 'Name/High/S*** P*** Q***' ((Get-HitSummary "NOK: SAMPLE, PAT Q. SSN $($Sample.Ssn)")[0]) 'Name next to SSN is High'
Assert-Equal 'Name/Low/P*** S***' (Get-HitSummary 'Spouse: Pat Sample, Phone ext 4') 'Name stops at the next label'
Assert-Equal 0 (Get-HitSummary 'Pat Sample wrote the brief').Count 'Name without label not reported'

# ---- Speed: a 100,000-letter run glued with '@' must not stall the scan (atomic top-level label) ----
$longRun = 'x' + '@' + 'b.' + ('a' * 100000) + '@'
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$null = Find-PiiInText $longRun
$timer.Stop()
Assert-Equal $true ($timer.Elapsed.TotalSeconds -lt 2) "Long glued run scanned in under 2 s (took $([int]$timer.Elapsed.TotalMilliseconds) ms)"

# ---- Speed: a long dotted run glued with '@' must not stall either detection or masking (bounded domain) ----
$dottedRun = 'x' + '@' + ('ab.' * 40000) + 'cc' + '@'
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$null = Find-PiiInText $dottedRun
$null = Protect-PiiText $dottedRun
$timer.Stop()
Assert-Equal $true ($timer.Elapsed.TotalSeconds -lt 2) "Long dotted run found and masked in under 2 s (took $([int]$timer.Elapsed.TotalMilliseconds) ms)"

# ---- Masking: snippets never show digits, e-mail names or found names ----
$hits = @(Find-PiiInText "SSN $($Sample.Ssn) then contact person: $($Sample.Email) later")
Assert-Equal $false ($hits[0].Context -match '\d') 'Snippet has no digits'
Assert-Equal $false ($hits[0].Context -match 'pat|sample') 'Snippet cut through an e-mail shows no part of its name'
$quoted = @(Find-PiiInText "SSN $($Sample.Ssn) mail $($Sample.EmailQuote)")
Assert-Equal $false (($quoted | ForEach-Object { $_.Context }) -join ' ' -match 'connor|jane') "Snippet hides an e-mail name with an apostrophe"
$named = @(Find-PiiInText "Name: Pat Q. Sample SSN $($Sample.Ssn)")
Assert-Equal $false (($named | ForEach-Object { $_.Context }) -join ' ' -match 'Pat|Sample') 'Snippet hides a found name'

# ---- Detection in names: underscores and accented e-mail names ----
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary "CAC_$($Sample.DodId).docx") 'DoD ID after an underscore'
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'Roster DOB_01-02-1990.xlsx') 'DOB after an underscore'
Assert-Equal 'Personal e-mail/Low/***@example.net' (Get-HitSummary "notes $($Sample.EmailAccent) here") 'E-mail with accents'
Assert-Equal 1 (Get-HitSummary "mail $($Sample.EmailIdn) here").Count 'E-mail with a non-ASCII domain'
Assert-Equal 1 (Get-HitSummary ("see " + $Sample.MilEmail + '/' + $Sample.Email)).Count 'E-mail right after a slash'
$afterAt = Get-HitSummary ('note x' + '@' + $Sample.Email)
Assert-Equal 'Personal e-mail/Low/***@example.com' ($afterAt -join ',') 'E-mail right after an @: only the real address, masked'
# An address that ends a sentence and then has a closing quote, a dash or a star is still found. (A letter or digit
# right after the final period reads as more domain, so such an address is not reported: an accepted loss.)
foreach ($tail in @(".$([char]0x201D)", ".$([char]0x2019)", ".'", '.*', '.-')) {
    Assert-Equal 'Personal e-mail/Low/***@example.com' ((Get-HitSummary ("mail me at " + $Sample.Email + $tail + ' soon')) -join ',') "E-mail followed by [$tail]"
}
# A long run of short addresses glued by '@' is read in linear time.
$gluedRun = (('ab' + '.cd' + '@') * 60000)
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$null = Find-PiiInText $gluedRun
$timer.Stop()
Assert-Equal $true ($timer.Elapsed.TotalSeconds -lt 2) "Long glued run of addresses scanned in under 2 s (took $([int]$timer.Elapsed.TotalMilliseconds) ms)"
# An address followed by punctuation is still found: ? ! ' curly quote ... ellipsis -ok * 's .
foreach ($tail in @('?', '!', "'", [string][char]0x201D, '...', [string][char]0x2026, '-ok', '*', "'s", '.')) {
    Assert-Equal 'Personal e-mail/Low/***@example.com' ((Get-HitSummary ("mail me at " + $Sample.Email + $tail + ' soon')) -join ',') "E-mail followed by [$tail]"
}

# ---- Protect-PiiText: names and messages shown as they are; every digit and e-mail name hidden ----
Assert-Equal 'Roster ***-**-####.docx' (Protect-PiiText "Roster $($Sample.Ssn).docx") 'File name with SSN masked'
Assert-Equal 'FY## Roster.docx' (Protect-PiiText 'FY27 Roster.docx') 'Digits in a name masked'
Assert-Equal 'Team Roster.docx' (Protect-PiiText 'Team Roster.docx') 'Name with nothing to mask unchanged'
Assert-Equal '***@example.org.docx' (Protect-PiiText "$($Sample.EmailQuote).docx") 'E-mail name with apostrophe hidden'
Assert-Equal '***@example.net notes.txt' (Protect-PiiText "$($Sample.EmailAccent) notes.txt") 'E-mail name with accents hidden'
Assert-Equal 'CAC_******####.docx' (Protect-PiiText "CAC_$($Sample.DodId).docx") 'DoD ID after an underscore hidden'
Assert-Equal $false (Protect-PiiText "$($Sample.EmailIdn) notes.txt").Contains('john') 'E-mail name with a non-ASCII domain hidden'
Assert-Equal 'Folder/***@example.com.docx' (Protect-PiiText "Folder/$($Sample.Email).docx") 'Folders before an address stay readable'

# ---- Protect-PiiName: site, folder and file names keep their digits; only found PII is masked ----
Assert-Equal 'FY27 Roster v2.docx' (Protect-PiiName 'FY27 Roster v2.docx') 'Name: digits that are not PII kept'
Assert-Equal 'Roster ***-**-6789.docx' (Protect-PiiName "Roster $($Sample.Ssn).docx") 'Name: SSN masked like a found value'
Assert-Equal 'Folder/***@example.com.docx' (Protect-PiiName "Folder/$($Sample.Email).docx") 'Name: e-mail name hidden'
Assert-Equal $false (Protect-PiiName "CAC_$($Sample.DodId).docx").Contains($Sample.DodId) 'Name: DoD ID never shown in full'

# ---- Validators ----
Assert-Equal $true (Test-Luhn '4111111111111111') 'Luhn true'
Assert-Equal $false (Test-Luhn '4111111111111112') 'Luhn false'

# ---- Unreadable text (the PAM, 2026-11-21): a Low hit inside symbol-font noise is dropped; High and Medium stay ----
$noise = -join @(1..40 | ForEach-Object { [char](0xA4 + ($_ % 8)) })   # Latin-1 signs, as a symbol font read as text gives
Assert-Equal 0 @(Find-PiiInText ($noise + ' ' + $Sample.Phone + ' ' + $noise)).Count 'Garbled: a Low phone in symbol noise is dropped'
Assert-Equal 'Phone/Low' ((Get-HitSummary ('Call the front desk at ' + $Sample.Phone + ' for the schedule of the week.')) -replace '/[^/]*$', '') `
    'Garbled: the same phone in plain words stays'
Assert-Equal 'SSN/High' ((Get-HitSummary ($noise + ' SSN ' + $Sample.Ssn + ' ' + $noise)) -replace '/[^/]*$', '') 'Garbled: a High hit in noise stays'
$privateNoise = -join @(1..40 | ForEach-Object { [char](0xE000 + $_) })
Assert-Equal 0 @(Find-PiiInText ($privateNoise + ' ' + $Sample.Phone + ' ' + $privateNoise)).Count 'Garbled: private-use characters count as noise'
Assert-Equal 'Phone/Low' ((Get-HitSummary $Sample.Phone) -replace '/[^/]*$', '') 'Garbled: a hit with little text around it is kept'
# Normal layouts are not noise (review of PR 31): line lists, tab-separated rows, sheet cells, box-drawing tables.
$phone = $Sample.Phone
$layouts = [ordered]@{
    'a CRLF list of short lines' = "Ops`r`nRm 2`r`nBldg 5`r`nDSN`r`n$phone`r`nFax`r`nN/A`r`nRm 3`r`nOK"
    'tab-separated rows'         = "10`t20`t30`t40`r`n50`t$phone`t60`t70`r`n80`t90`t11`t12"
    'sheet cells'                = "SSgt | Doe | Y | N | 1 | $phone | Y | Y | N | 2 | 3 | N"
    'a box-drawing table'        = ([string][char]0x2502 + ' Phone ' + [char]0x2502 + " $phone " + [char]0x2502 + "`n" + ([string][char]0x2500 * 30))
}
foreach ($layout in $layouts.Keys) {
    Assert-Equal 'Phone/Low' ((Get-HitSummary $layouts[$layout]) -replace '/[^/]*$', '') ('Garbled: ' + $layout + ' keeps its Low hit')
}

# ---- The 5 Ion defects of the R4e benchmark (archimedes_lever docs/reviews/2026-11-23-tr-r4e-ion-benchmark.md) ----
# 1. Street types in any case.
Assert-Equal 'Home address/Medium/[address masked]' ((Get-HitSummary ('42 ELM ST, SPRINGFIELD, IL ' + '62704')) -join ',') 'Address in capitals'
Assert-Equal 'Home address/High/[address masked]' ((Get-HitSummary 'Home: 12 OAK RD') -join ',') 'Capital street type with a label'
Assert-Equal 'Home address/Low/[address masked]' ((Get-HitSummary '42 Elm st') -join ',') 'Lower-case street type'
# 2. A 9-digit passport number is a passport, not an SSN; an SSN label still wins; a dashed SSN stays an SSN.
$passport9 = '1234' + '56789'
foreach ($case in @("Passport: $passport9", "PASSPORT NO`n$passport9", "her passport number is $passport9")) {
    Assert-Equal 'Passport/Medium/******789' ((Get-HitSummary $case) -join ',') "9-digit passport: $($case -replace '\d', '#' -replace "`n", ' ')"
}
Assert-Equal 'SSN/High/***-**-6789' ((Get-HitSummary "Passport and SSN: $passport9") -join ',') 'SSN label near a passport word: SSN'
Assert-Equal 1 @(Get-HitSummary "Passport photo; ref $($Sample.Ssn)" | Where-Object { $_ -like 'SSN/*' }).Count 'Dashed SSN after a passport word: SSN'
# 3. An underscore in an e-mail name: the whole address is the hit, so no part of the name stays visible.
$underscoreMail = 'first_last' + '@gmail.com'
$mailSpan = @(Get-PiiSpan "Write to $underscoreMail today")[0]
Assert-Equal $underscoreMail $mailSpan.Value 'E-mail with an underscore: whole address found'
Assert-Equal $false ((Protect-PiiText "Write to $underscoreMail") -match 'first') 'E-mail with an underscore: name hidden'
# 4. A name stops at a line break and never takes the next label.
Assert-Equal 'Kenji D. Abernathy' (@(Get-PiiSpan "Name: Kenji D. Abernathy`nShip to: depot")[0].Value) 'Name stops at a line break'
Assert-Equal 'P*** L*** | S*** R***' ((@(Find-PiiInText "Name: Pat Lee`nName: Sam Roe") | ForEach-Object { $_.MaskedValue }) -join ' | ') 'Two name rows: both names found'
Assert-Equal 'Jane Doe' (@(Get-PiiSpan "Name: Jane`nDoe")[0].Value -replace '\s+', ' ') 'A name wrapped after the first word is still found'
# 5. A middle initial without a dot.
Assert-Equal 'DOE, JANE M' (@(Get-PiiSpan 'Name: DOE, JANE M')[0].Value) 'Middle initial without a dot kept'
Assert-Equal 'J*** M*** D***' ((Find-PiiInText 'Member: Jane M Doe, unit').MaskedValue) 'First name, initial without a dot, last name'
