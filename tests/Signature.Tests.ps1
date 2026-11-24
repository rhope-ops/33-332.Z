<#
Signature.Tests.ps1 - tests of the e-signature DoD ID (PiiDetectors.psm1 Find-SignedIdHits, Add-SignedIdHits): 25 layouts, decoys, masking, speed.

  Purpose:     Each layout of docs/reviews/2026-11-23-ion-esignature-layouts.md (this session's 25 and the source's 25) gives
               exactly one hit, a DoD ID (Low), and no SSN, except source layout OF by design; decoys give no DoD ID;
               values with gaps and line breaks stay masked; no regex stalls.
  Inputs:      $Sample (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- The 25 layouts (invented name and ID; the ID is built from parts in TestKit.ps1) ----
$sigId = $Sample.DodId
$sigName = 'DOE' + '.JANE.Q.'
$sigDate = "Date: 2026.11.20 10:15:22 -06'00'"
$sigNbsp = [string][char]0xA0
$sig3241 = $sigId.Substring(0, 3) + ' ' + $sigId.Substring(3, 2) + ' ' + $sigId.Substring(5, 4) + ' ' + $sigId.Substring(9)
$SignatureLayouts = [ordered]@{
    'S01 base'                    = "Digitally signed by $sigName$sigId`n$sigDate"
    'S02 CN= distinguished name'  = "CN=$sigName$sigId, OU=USAF, OU=PKI, OU=DoD, O=U.S. Government, C=US"
    'S03 no middle name'          = "Digitally signed by DOE.JANE.$sigId"
    'S04 suffix'                  = "DOE.JOHN.Q.JR.$sigId"
    'S05 hyphen and apostrophe'   = "O'BRIEN-LEE.PAT.R.$sigId"
    'S06 digits 3-2-4-1'          = "Digitally signed by $sigName$sig3241`n$sigDate"
    'S07 every digit apart'       = "Digitally signed by $sigName$(($sigId.ToCharArray()) -join ' ')"
    'S08 digits 5-5'              = "$sigName$($sigId.Substring(0, 5)) $($sigId.Substring(5))"
    'S09 LF before the ID'        = "Digitally signed by $sigName`n$sigId`n$sigDate"
    'S10 CRLF before the ID'      = "Digitally signed by $sigName`r`n$sigId`r`n$sigDate"
    'S11 space before the ID'     = "Digitally signed by $sigName $sigId"
    'S12 tab before the ID'       = "$sigName`t$sigId"
    'S13 no-break space before'   = "$sigName$sigNbsp$sigId"
    'S14 LF inside the ID'        = "Digitally signed by $sigName$($sigId.Substring(0, 6))`n$($sigId.Substring(6))`n$sigDate"
    'S15 LF after a name dot'     = "Digitally signed by DOE.JANE.`nQ.$sigId"
    'S16 LF inside a name'        = "Digitally signed by DOE.JA`nNE.Q.$sigId"
    'S17 title case, signed'      = "Digitally signed by Doe.Jane.Q.$sigId"
    'S18 two-word first name'     = "Digitally signed by DOE.MARY ANN.$sigId"
    'S19 no-break-space gaps'     = "Digitally signed by $sigName$($sig3241.Replace(' ', $sigNbsp))"
    'S20 signed, ID 3-2-4-1 only' = "Digitally signed by $sig3241 $sigDate"
    'S21 signed, ID only'         = "Digitally signed by $sigId"
    'S22 file name, underscores'  = "Signed_DOE_JANE_Q_$sigId.pdf"
    'S23 space, then 3-2-4-1'     = "$sigName $sig3241"
    'S24 CRLF, then 3-2-4-1'      = "$sigName`r`n$sig3241"
    'S25 one digit, then nine'    = "$sigName$($sigId.Substring(0, 1)) $($sigId.Substring(1))"
}
# A CAC name, a gap, then 9+1 digits (S23, S24, OE) is a DoD ID Low (the PAM's gap rule, 2026-11-24, question dialog):
# no layout stays an SSN. A roster row without a name, and a count after an SSN, stay SSNs (tests below and Detectors.Tests.ps1).
$sigAsSsn = @{}
foreach ($layout in $SignatureLayouts.Keys) {
    $expected = if ($sigAsSsn.ContainsKey($layout)) { $sigAsSsn[$layout] } else { 'DoD ID/Low/******7890' }
    Assert-Equal $expected ((Get-HitSummary $SignatureLayouts[$layout]) -join ',') "E-signature layout $layout"
}

# ---- The source's 25 layouts (tool retirement session, 2026-11-23; same shapes, this repo's invented name and ID) ----
$srcName = 'DOE' + '.JOHN.A.'
$sig91 = $sigId.Substring(0, 9) + ' ' + $sigId.Substring(9)
$SourceLayouts = [ordered]@{
    'O01' = "Digitally signed by $srcName$sigId"; 'O02' = "DOE.JOHN.$sigId"; 'O03' = "Digitally signed by $srcName $sigId"
    'O04' = "$srcName`n$sigId"; 'O05' = "$srcName$($sigId.Substring(0, 5))`n$($sigId.Substring(5))"; 'O06' = "DOE.JOHN.A.JR.$sigId"
    'O07' = "SMITH-JONES.JANE.B.$sigId"; 'O08' = "$sigId`n$srcName"; 'O09' = "Digitally signed by $srcName$(($sigId.ToCharArray()) -join ' ')"
    'O10' = "$srcName$sigId Date: 2026.11.21 08:15:00 -0600"; 'O11' = "Digitally signed by doe.john.a.$sigId"; 'O12' = "DOE.MARY ANN.C.$sigId"
    'O13' = "CN=$srcName$sigId, OU=USAF, OU=PKI, OU=DoD, O=U.S. Government, C=US"; 'O14' = "O'BRIEN.SEAN.P.$sigId"
    'O15' = "Digitally signed by DOE.JOHN.A $sigId"; 'OB' = "Digitally signed by DOE JOHN A $sig3241"
    'OC' = "$srcName$($sigId.Substring(0, 9))`n$($sigId.Substring(9))"; 'OD' = "Digitally signed by`n$srcName`n$sigId"; 'OE' = "$srcName $sig91"
    'OG' = "Signature: DOE, JOHN A. $sig91"; 'OH' = "Digitally signed by $srcName  $sigId"; 'OI' = "$srcName$sigId`nReason: I am the author"
    'OJ' = "$srcName$($sigId.Substring(0, 3))-$($sigId.Substring(3, 2))-$($sigId.Substring(5, 4))-$($sigId.Substring(9))"
}
foreach ($layout in $SourceLayouts.Keys) {
    $expected = if ($sigAsSsn.ContainsKey($layout)) { $sigAsSsn[$layout] } else { 'DoD ID/Low/******7890' }
    Assert-Equal $expected ((Get-HitSummary $SourceLayouts[$layout]) -join ',') "E-signature source layout $layout"
}
# By design: no dotted name and no signature word. A stays an SSN (still reported); F, 10 spaced digits alone, is not reported.
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "DOE JOHN A $sig3241") -join ',') 'E-signature source layout OA: DoD ID (the PAM, 2026-11-24)'
Assert-Equal 0 (Get-HitSummary (($sigId.ToCharArray()) -join ' ')).Count 'E-signature source layout OF: no hit, by design'

# ---- Decoys: no DoD ID where there is no signature ID; an SSN with a label or dashes stays an SSN ----
$sigSsnBare = $Sample.Ssn -replace '-', ' '
$sigDecoys = [ordered]@{
    'date after signed'          = 'Signed on 2026 11 20 12:34 by the board'
    'signature without an ID'    = "Digitally signed by $sigName`n$sigDate"
    'phone after signed'         = "Signed copy, office $($Sample.Phone)"
    'spaced phone after signed'  = "Signed copy, cell $($Sample.Phone -replace '-', ' ')"
    '11 digits after a name'     = "$sigName$($sigId)4"
    'lower-case words, unsigned' = "see backup.config.$sigId"
    'title-case name, unsigned'  = "Doe.Jane.Q.$sigId"
    'ID, digits on the same line' = "Digitally signed by $sigName$sigId 2"
}
foreach ($decoy in $sigDecoys.Keys) {
    Assert-Equal 0 @(Get-HitSummary $sigDecoys[$decoy] | Where-Object { $_ -like 'DoD ID/*' }).Count "E-signature decoy: $decoy"
}
Assert-Equal 'SSN/High/***-**-6789' ((Get-HitSummary "Digitally signed by $sigName SSN $sigSsnBare 0") -join ',') 'Signature: a labelled SSN stays an SSN'
Assert-Equal 'Phone/Medium/***-***-0142' ((Get-HitSummary "Signed copy, cell $($Sample.Phone -replace '-', ' ')") -join ',') 'Signature: a spaced cell phone stays a phone'
Assert-Equal 'SSN/Medium/***-**-6789' ((Get-HitSummary "Signed by: $($Sample.Ssn)") -join ',') 'Signature: a dashed SSN stays an SSN'
Assert-Equal 'DoD ID/High/******7890' ((Get-HitSummary "EDIPI $sigId signed by $sigName$sigId") -join ',').Split(',')[0] 'A labelled DoD ID stays High'
Assert-Equal 'DoD ID/High/******7890' ((Get-HitSummary "Signed, EDIPI $sig3241") -join ',') 'A labelled DoD ID with gaps in a signature: High'

# ---- Review of PR 60: phones and dashed SSNs near a CAC name or a signature word keep their type and confidence ----
$revPhone = $Sample.Phone -replace '-', ' '
$revCases = [ordered]@{
    'Phone/Low/***-***-0142'    = @("$sigName $revPhone", "Signed by Doe.Jane.Q. $revPhone", "$sigName`n$revPhone")
    'Phone/Medium/***-***-0142' = @("Signed form.`nCall my cell.`n$($Sample.Phone)", "CALL ME.`nCELL PHONE.`n$($Sample.Phone)", "Signature page.`nMy cell.`n$revPhone")
    'SSN/Medium/***-**-6789'    = @("$sigName`t$($Sample.Ssn)`t5", "$sigName`n$($Sample.Ssn)`n2 copies", "Signed by applicant $sigSsnBare`n1", "$sigSsnBare 0`n$sigName")
}
foreach ($expected in $revCases.Keys) {
    foreach ($case in $revCases[$expected]) {
        Assert-Equal $expected ((Get-HitSummary $case) -join ',') "Not an e-signature ID: $($case -replace '\d', '#' -replace "`r?`n", ' / ')"
    }
}
Assert-Equal 0 @(Get-HitSummary 'Order 4455667788 REF.NO. 5' | Where-Object { $_ -like 'DoD ID/*' }).Count 'No CAC name: digits before REF.NO.'
# Second review of PR 60.
$revMail = 'jane.doe.' + $sigId + '@gmail.com'
$revCases2 = [ordered]@{
    'Phone/Low/***-***-0142'               = @("$sigName$($Sample.Phone)", "AB.CD.$($Sample.Phone)", "signed by doe.jane.q.$($Sample.Phone)")
    'SSN/Medium/***-**-6789'               = @("$sigName$($Sample.Ssn) 5 copies", "$sigName$($Sample.Ssn)`t5", "CN=$sigName$($Sample.Ssn) 2",
        "Signature: ____`nMember $sigSsnBare 2 kids", "Signature`t$sigSsnBare`t1")
    'Personal e-mail/Medium/***@gmail.com' = @("Digitally signed by $revMail")
}
foreach ($expected in $revCases2.Keys) {
    foreach ($case in $revCases2[$expected]) {
        Assert-Equal $expected ((Get-HitSummary $case) -join ',') "Not an e-signature ID (2): $($case -replace '\d', '#' -replace "`r?`n", ' / ')"
    }
}
Assert-Equal 0 @(Get-HitSummary 'signed 20261120 10:15:22' | Where-Object { $_ -like 'DoD ID/*' }).Count 'No ID: a yyyymmdd date after signed'
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$null = Find-PiiInText (("$sigName$sig3241" + "`n") * 2000)
$timer.Stop()
Assert-Equal $true ($timer.Elapsed.TotalSeconds -lt 8) "E-signature: 2,000 signature rows in under 8 s (took $([int]$timer.Elapsed.TotalMilliseconds) ms)"

# ---- Masking: a value with a gap, a tab or a line break inside shows no full ID anywhere ----
foreach ($layout in @('S06 digits 3-2-4-1', 'S09 LF before the ID', 'S12 tab before the ID', 'S14 LF inside the ID', 'S19 no-break-space gaps')) {
    $sigHit = @(Find-PiiInText $SignatureLayouts[$layout])[0]
    Assert-Equal $false ($sigHit.MaskedValue + $sigHit.Context -match '\d{5}|\d\D\d\D\d\D\d\D\d') "E-signature masked: value and snippet, $layout"
    Assert-Equal $false ((Protect-PiiText $SignatureLayouts[$layout]) -match '\d') "E-signature masked: Protect-PiiText, $layout"
}
Assert-Equal $false (Protect-PiiName $SignatureLayouts['S22 file name, underscores']).Contains($sigId) 'E-signature masked: file name'

# ---- Speed: long runs that look like CAC names or digit groups must not stall the scan ----
$sigRuns = @(('AB.' * 40000), (('AB' + "`n") * 40000), ('AB.CD EF GH IJ.' * 8000), ('Signed ' + ('1 ' * 50000)), ('signed DOE.' + ('A.' * 40000)))
$timer = [System.Diagnostics.Stopwatch]::StartNew()
foreach ($run in $sigRuns) { $null = Find-PiiInText $run }
$timer.Stop()
Assert-Equal $true ($timer.Elapsed.TotalSeconds -lt 4) "E-signature: long name and digit runs scanned in under 4 s (took $([int]$timer.Elapsed.TotalMilliseconds) ms)"

# ---- No loss (third review of PR 60): the e-signature pass only adds. Every hit found without it is still there, the same,
# except an SSN (not High) wholly inside an added DoD ID. Values next to CAC names and signature words, from parts. ----
$invAt = '@'; $invId = $sigId; $invPhone = $Sample.Phone
$invValues = @($invId, $sig3241, $sig91, $Sample.Ssn, $sigSsnBare, ($sigSsnBare -replace ' ', ''), $invPhone, ($invPhone -replace '-', ' '),
    ($invPhone -replace '-', '.'), '01-02-1990', 'X1234567', '4111 1111 1111 1111', "jdoe$invAt$('gmail.com')", "$invId$invAt$('gmail.com')")
$invLost = 0; $invCount = 0
foreach ($context in @('', 'Digitally signed by ', 'Name: Jane Doe ', 'born ', 'Name: John Smith DoD ID: ')) {
    foreach ($name in @($sigName, 'Doe.Jane.Q.', 'DOE JANE Q ', '')) {
        foreach ($gap in @('', ' ', "`n", "`t")) {
            foreach ($value in $invValues) {
                foreach ($tail in @('', ' 2 copies', "`n2 ROE.JOHN.A.", "`t5")) {
                    $text = $context + $name + $gap + $value + $tail
                    $off = & (Get-Module PiiDetectors) { param($t) Remove-OverlappingHits @(Get-AllHits $t.Replace('_', ' ')) } $text
                    $on = @(Get-PiiSpan $text)
                    $invCount++
                    foreach ($hit in $off) {
                        if (@($on | Where-Object { $_.Detector -eq $hit.Detector -and $_.Index -eq $hit.Index -and $_.Length -eq $hit.Length -and
                                ($_.Confidence -eq $hit.Confidence -or $_.Detector -eq 'Name') }).Count) { continue }
                        $swapped = $hit.Detector -eq 'SSN' -and $hit.Confidence -ne 'High' -and @($on | Where-Object { $_.Detector -eq 'DoD ID' -and
                            $_.Index -le $hit.Index -and $hit.Index + $hit.Length -le $_.Index + $_.Length }).Count
                        if (-not $swapped) { $invLost++ }
                    }
                }
            }
        }
    }
}
Assert-Equal 0 $invLost "E-signature pass only adds: no hit lost or changed in $invCount texts"

# Fourth review of PR 60: an added ID never raises a name, and no hit is lost in unreadable text.
Assert-Equal 'Low' (@(Find-PiiInText "Name: John Smith DoD ID: $sigName$($sigId.Substring(0, 5))`n$($sigId.Substring(5))" |
    Where-Object { $_.Detector -eq 'Name' })[0].Confidence) 'An added e-signature DoD ID does not raise a name'
$sigNoise = -join (1..16 | ForEach-Object { [char](0xE000 + $_) })
Assert-Equal 1 @(Get-HitSummary "${sigNoise}signed by doe.jane.q. $sigSsnBare`n0$sigNoise").Count 'Unreadable text around a signature: the SSN is kept'

# Fifth review of PR 60: a date and time after a CAC name is no ID.
foreach ($case in @("$sigName 2026 11 20 10:15", "DOE.JANE.`t20261120 10")) {
    Assert-Equal 0 @(Get-HitSummary $case | Where-Object { $_ -like 'DoD ID/*' }).Count "No ID: a date after a CAC name ($($case -replace "`t", ' '))"
}

# Sixth review of PR 60: a signature word cut at the window edge is no word; mid dates and 3-7 phones are no ID.
$cutWord = 'Signed form. ' + ('k' * 80) + ' assigned' + (' k' * 16) + '  1 ' + ($sigSsnBare -replace ' ', '')
Assert-Equal 1 @(Get-HitSummary $cutWord | Where-Object { $_ -like 'SSN/*' }).Count '"assigned" at the window edge is no signature word: SSN kept'
Assert-Equal 0 @(Get-HitSummary "Digitally signed by $sigName`n11 20 2026 0930" | Where-Object { $_ -like 'DoD ID/*' }).Count 'No ID: a month-day-year date'
Assert-Equal 0 @(Get-HitSummary ('Signed. Phone 312 555' + '0123') | Where-Object { $_ -like 'DoD ID/*' }).Count 'No ID: a 3-7 phone after signed'

# Seventh review of PR 60: a hit just before the ID must not hide it (the digit scan starts again at the hit's end), and a
# candidate that needs a signature word must not block a sure one at the same digits.
$sigPhone1 = '+1 ' + $Sample.Phone
$sigAddress = '4 Oak' + ' Ln, Springfield, IL ' + '627' + '01'
$sigProse = 'The form is filed for review by the office. ' * 3
$sigAfterHit = [ordered]@{
    'phone, ID, name'         = "Contact $sigPhone1`n$sigId`n$sigName"
    'signed, phone, ID'       = "Digitally signed by`n$sigPhone1`n$sigId"
    'address, ID, name'       = "$sigAddress`n$sigId`n$sigName"
    'signed, address, ID'     = "Signed. $sigAddress`n$sigId"
    'Thank you., Signed., ID' = "Signed form. $sigProse Thank you.`nSigned.`n$sigId"
    'Approved., Signed., ID'  = "Signed form. $sigProse Approved.`nSigned.`n$sigId"
    'Approved., Signed. only' = "Approved.`nSigned.`n$sigId"
    'Reviewed., signature'    = "Signature page. $sigProse Reviewed.`nDigital signature.`n$sigId"
}
foreach ($case in $sigAfterHit.Keys) {
    Assert-Equal 1 @(Get-HitSummary $sigAfterHit[$case] | Where-Object { $_ -eq 'DoD ID/Low/******7890' }).Count "E-signature ID after a hit or a second word ($case)"
}
# A date, time, year or page number and one gap before bare digits after a signature word: the ID is found (the PAM,
# 2026-11-24, "1a": the guard of 013ca82 narrowed). A lone digit and a gap still block it (the ID could be 11 digits).
$sigLeadIns = @("Digitally signed by`n2026.11.20 $sigId", "Digitally signed by`nPage 3 $sig3241", "Digitally signed by 2026-11-20 $sigId",
    "Digitally signed by 11/20/2026 $sigId", "Digitally signed by 10:15:22 $sigId", "Digitally signed by 2026 $sigId")
foreach ($case in $sigLeadIns) {
    Assert-Equal 1 @(Get-HitSummary $case | Where-Object { $_ -eq 'DoD ID/Low/******7890' }).Count "ID after a date, time, year or page ($($case -replace "`n", ' ' -replace '\d', '#'))"
}
Assert-Equal 0 @(Get-HitSummary "Digitally signed by 7 $sigId" | Where-Object { $_ -like 'DoD ID/*' }).Count 'No ID after a lone digit and a gap'
# After a CAC name a 3-7 split is an ID; a 3-3-4 phone there stays a phone (the PAM, 2026-11-24, "1a").
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "$sigName $($sigId.Substring(0, 3)) $($sigId.Substring(3))") -join ',') 'CAC name, a gap, a 3-7 split: an ID'
$sigNamePhone = @(Find-PiiInText "$sigName $($Sample.Phone -replace '-', ' ')" | ForEach-Object { $_.Detector })
Assert-Equal 'True False' ('{0} {1}' -f ($sigNamePhone -contains 'Phone'), ($sigNamePhone -contains 'DoD ID')) 'CAC name, a gap, a 3-3-4 phone: a phone' 

# ---- Review of PR 57 (the PAM, 2026-11-24: close PR 57, fix its cases on main) ----
$sig9 = $sigId.Substring(0, 9)
$sigCaps = 'DOE' + ' JANE Q'
$pr57Ids = [ordered]@{
    'capitals, no dots, 3-2-4-1'      = "$sigCaps $sig3241"
    'Signature:, comma name, 9+1'     = "Signature: DOE, JANE Q. $sig9 $($sigId.Substring(9))"
    'dotted name, a space, 9+1'       = "$sigName $sig9 $($sigId.Substring(9))"
    'lower-case name after signed'    = "Digitally signed by $($sigName.ToLower())$sigId"
    'lower-case name, a space'        = "Digitally signed by $($sigName.ToLower()) $sigId"
}
foreach ($case in $pr57Ids.Keys) {
    Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary $pr57Ids[$case]) -join ',') "Review of PR 57: DoD ID ($case)"
}
$pr57None = [ordered]@{
    'signed. in a sentence'          = "The contract was signed. Tracking $sigId"
    'date m-d-yy after a name'       = "$($sigName)11-23-26 1015"
    'date m-d-yyyy after a name'     = "$($sigName)1-5-2026 1015"
    'date yyyy-m-d after a name'     = "$($sigName)2026-1-5 1015"
    'capitals heading, plain digits' = "THE UNIT A $sigId"
    'capitals roster, next row'      = "1 $sigCaps $sig9`n2 DOE JOHN"
    'small words, no dots'           = "assigned to bob a $sig9 $($sigId.Substring(9))"
    'dotted small words, no signed'  = "report.final.v.$sigId"
    'capitals heading, other groups' = "BLDG NO $($sigId.Substring(0, 4)) $($sigId.Substring(4))"
    'name, 1+9 after a gap'          = "$sigName $($sigId.Substring(0, 1)) $($sigId.Substring(1))"
}
foreach ($case in $pr57None.Keys) {
    Assert-Equal 0 @(Get-HitSummary $pr57None[$case] | Where-Object { $_ -like 'DoD ID/*' }).Count "Review of PR 57: no DoD ID ($case)"
}
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "$($sigName)$($sigId.Substring(0, 3))-$($sigId.Substring(3, 2))-$($sigId.Substring(5, 4))-$($sigId.Substring(9))") -join ',') `
    'Review of PR 57: a dashed ID after a name is no date'
# Review of PR 69: a title before the name is no sentence end; a dashed 4-2-4 ID after a name is no date.
foreach ($case in @("Digitally signed by Lt. Col. Doe $sigId", "Signature: Mr. Doe $sigId")) {
    Assert-Equal 1 @(Get-HitSummary $case | Where-Object { $_ -eq 'DoD ID/Low/******7890' }).Count "A title is no sentence end ($($case -replace '\d', '#'))"
}
Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary "$($sigName)$($sigId.Substring(0, 4))-$($sigId.Substring(4, 2))-$($sigId.Substring(6))") -join ',') 'A dashed 4-2-4 ID after a name'
