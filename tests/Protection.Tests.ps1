<#
Protection.Tests.ps1 - tests that protected files and PII stay protected: the not-scanned reasons for protected files
(a sensitivity label with encryption told apart from a password, in PDFs and Office files; a zlib header of any window
size), and negative tests that no raw PII value leaks into a hit, a reason, a file name, the screen or the workbook.

  Purpose:     The PAM's gov run of 2026-11-21: labelled and password-protected files in groups of their own, and good
               PDFs not called damaged. AGENTS.md "Masked output": every found value is masked where it is shown or saved.
  Inputs:      $Sample, New-PdfBytes and Get-ResultFiles (TestKit.ps1); invented values built from parts.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      one workbook in a new temporary folder, deleted at the end of this file.
  Credentials: none.
#>

# A sensitivity label with encryption is told apart from a password (the PAM, 2026-11-21).
$labelPdf = New-PdfBytes 'BT (x) Tj ET' -Extra @('<< /Filter /MicrosoftIRMServices /V 2 >>') -Trailer '/Encrypt 9 0 R'
Assert-Equal $true ((Get-TextFromBytes -Bytes $labelPdf -FileName 'l.pdf').Reason -match '^Sensitivity label') 'PDF: Purview encryption named as a sensitivity label'
$wrapperPdf = New-PdfBytes 'BT (This document is protected) Tj ET' -Extra @('<< /Type /EmbeddedFile /Subtype /EncryptedPayload >>')
Assert-Equal $true ((Get-TextFromBytes -Bytes $wrapperPdf -FileName 'w.pdf').Reason -match '^Sensitivity label') 'PDF: a protected-PDF wrapper is not read as clean'
# A zlib header with a smaller window (0x58, not 0x78) is still skipped, so the stream is not called damaged.
$zlibSmall = & (Get-Module PdfCore) {
    $memory = New-Object System.IO.MemoryStream
    $deflate = New-Object System.IO.Compression.DeflateStream($memory, [System.IO.Compression.CompressionMode]::Compress, $true)
    $plain = [System.Text.Encoding]::ASCII.GetBytes('BT (Roster) Tj ET'); $deflate.Write($plain, 0, $plain.Length); $deflate.Dispose()
    $raw = [byte[]](@(0x58, 0x85) + $memory.ToArray())   # 0x5885 is a multiple of 31
    Start-PdfBudget
    $out = Expand-PdfStream ($script:Latin1.GetString($raw))
    $notes = @(Stop-PdfBudget)
    '{0}|{1}' -f $out, $notes.Count
}
Assert-Equal 'BT (Roster) Tj ET|0' $zlibSmall 'PDF: a zlib header of any window size is skipped, no damage note'

# Office files in OLE form: a password, or a sensitivity label (rights management names in the directory).
$oleHead = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1) + (New-Object byte[] 504)
$oleLocked = [byte[]]($oleHead + [System.Text.Encoding]::Unicode.GetBytes('EncryptedPackage') + (New-Object byte[] 64))
Assert-Equal $true ((Get-TextFromBytes -Bytes $oleLocked -FileName 'locked.xlsx').Reason -match '^Encrypted \(password') 'New Office in OLE form: a password named'
$oleLabel = [byte[]]($oleHead + [System.Text.Encoding]::Unicode.GetBytes('DRMEncryptedDataSpace EncryptedPackage') + (New-Object byte[] 64))
Assert-Equal $true ((Get-TextFromBytes -Bytes $oleLabel -FileName 'cui.docx').Reason -match '^Sensitivity label') 'New Office in OLE form: a sensitivity label named'
$oldIrm = [byte[]]($oleHead + [System.Text.Encoding]::Unicode.GetBytes('DRMDataSpace DRMContent') + (New-Object byte[] 64))
Assert-Equal $true ((Get-TextFromBytes -Bytes $oldIrm -FileName 'cui.doc').Reason -match '^Sensitivity label') 'Old Office with rights management: a sensitivity label named'

# ---- No PII leaks (fleet v2 Stage 3B, 2026-11-23): a raw value never reaches a hit, a reason, a name, the workbook
# or the screen (AGENTS.md "Masked output"). Invented values, built from parts. ----
$leakCard = '4111' + ' 1111' + ' 1111' + ' 1111'
$leakDob = '03/14/' + '1985'
$leakRaw = @($Sample.Ssn, $Sample.SsnBare, $Sample.DodId, $Sample.Phone, $Sample.Passport, $Sample.Email, $Sample.Street, $leakCard, $leakDob)
$leakText = "SSN: $($Sample.Ssn). SSN $($Sample.SsnBare). DoD ID: $($Sample.DodId). Cell: $($Sample.Phone). " +
    "Passport: $($Sample.Passport). Mail $($Sample.Email). Home address: $($Sample.Street), $($Sample.CityZip). " +
    "Card: $leakCard. DOB: $leakDob."
# Returns the raw values found in a text (none is the pass).
$leakIn = { param([string]$Text) @($leakRaw | Where-Object { $Text.Contains($_) }) }
Assert-Equal $leakRaw.Count @(& $leakIn $leakText).Count 'No leak: the check itself finds every raw value in the unmasked text'
$leakHits = @(Find-PiiInText $leakText)
Assert-Equal $true ($leakHits.Count -ge 7) 'No leak: the sample text gives hits to check'
Assert-Equal 0 @(& $leakIn (($leakHits | ForEach-Object { $_.MaskedValue + ' ' + $_.Context }) -join ' ')).Count 'No leak: hit values and snippets are masked'
Assert-Equal 0 @(& $leakIn (Protect-PiiText ('Could not read: ' + $leakText))).Count 'No leak: a reason text is masked'
Assert-Equal 0 @(& $leakIn (Protect-PiiName ("Roster $($Sample.Ssn) $($Sample.Email).docx"))).Count 'No leak: a file name is masked'
$leakFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-leak-' + [guid]::NewGuid())
try {
    $leakRows = @($leakHits | ForEach-Object { New-ReportRow 'File' 'General' 'Roster.docx' 'https://x/y' '2026-11-01' $_.Detector $_.Confidence $_.MaskedValue $_.Context })
    $leakRows += New-ReportRow 'File' 'General' 'Locked.pdf' '' '' 'NOT SCANNED' '' '' (Protect-PiiText ('Encrypted: ' + $Sample.Ssn))
    $leakHeader = [ordered]@{ Site = 'https://tenant.example.com/teams/TeamA'; Channels = 'General' }
    $leakScreen = (Publish-ScanReport -Rows $leakRows -ItemsScanned 1 -Header $leakHeader -OutputFolder $leakFolder -SiteUrl $leakHeader.Site -FolderLabel 'Channel' 6>&1 |
        ForEach-Object { [string]$_ }) -join "`n"
    Assert-Equal 0 @(& $leakIn $leakScreen).Count 'No leak: the screen summary shows no raw value'
    $leakBook = @(Get-ResultFiles $leakFolder)
    Assert-Equal 1 $leakBook.Count 'No leak: one workbook saved'
    Add-Type -AssemblyName System.IO.Compression
    # ZipArchive on a stream, as TestKit.ps1 reads workbooks: Windows PowerShell 5.1 has no ZipFile type without an extra assembly.
    $leakStream = [System.IO.File]::OpenRead($leakBook[0].FullName)
    $leakZip = $null
    try {
        $leakZip = New-Object System.IO.Compression.ZipArchive($leakStream, [System.IO.Compression.ZipArchiveMode]::Read)
        $leakXml = ($leakZip.Entries | ForEach-Object { $r = New-Object System.IO.StreamReader($_.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }) -join "`n"
    }
    finally { if ($leakZip) { $leakZip.Dispose() }; $leakStream.Dispose() }
    Assert-Equal 0 @(& $leakIn ([System.Net.WebUtility]::HtmlDecode($leakXml))).Count 'No leak: every part of the workbook is free of raw values'
}
finally { Remove-Item -LiteralPath $leakFolder -Recurse -Force -ErrorAction SilentlyContinue }
