<#
Signature.Tests.ps1 - tests of the e-signature DoD ID (PiiPatterns.psm1 Find-SignedIdHits): 25 layouts, decoys, masking, speed.

  Purpose:     Each layout of docs/reviews/2026-11-23-ion-esignature-layouts.md (this session's 25 and the source's 25) gives
               exactly one hit, a DoD ID (Low), and no SSN, except 2 source layouts by design; decoys give no DoD ID;
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
foreach ($layout in $SignatureLayouts.Keys) {
    Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary $SignatureLayouts[$layout]) -join ',') "E-signature layout $layout"
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
    Assert-Equal 'DoD ID/Low/******7890' ((Get-HitSummary $SourceLayouts[$layout]) -join ',') "E-signature source layout $layout"
}
# By design: no dotted name and no signature word. A stays an SSN (still reported); F, 10 spaced digits alone, is not reported.
Assert-Equal 'SSN/Medium/***-**-6789' ((Get-HitSummary "DOE JOHN A $sig3241") -join ',') 'E-signature source layout OA: SSN, by design'
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
