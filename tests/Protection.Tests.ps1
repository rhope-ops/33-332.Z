<#
Protection.Tests.ps1 - tests of the not-scanned reasons for protected files: a sensitivity label with encryption
(Microsoft Purview) told apart from a password, in PDFs and Office files, and a zlib header of any window size.

  Purpose:     The PAM's gov run of 2026-11-21: labelled and password-protected files in groups of their own, and good
               PDFs not called damaged.
  Inputs:      New-PdfBytes (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
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
