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

Set-StrictMode -Version 2.0

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

# Names that only a file protected by rights management has (MS-OFFCRYPTO 2.2.11, IRM: a Microsoft Purview sensitivity
# label with encryption uses it): DRMEncrypted* in a new-style file kept in OLE form, DRMDataSpace and DRMContent in an
# old .doc/.xls/.ppt; a password-protected file has StrongEncryptionDataSpace instead (the PAM, 2026-11-21: tell
# labels from passwords).
$script:LabelStreamNames = @('DRMEncryptedDataSpace', 'DRMEncryptedTransform', 'DRMDataSpace', 'DRMContent')
$script:LabelReason = 'Sensitivity label with encryption (Microsoft Purview): open it and check by hand'
$script:PasswordReason = 'Encrypted (password-protected): check by hand'

# An old .xls keeps numbers as binary values, not text: a number typed into a cell (an SSN without dashes) is not
# read, so every .xls gets this note. Old .doc, .ppt and .msg files keep their text as text.
$script:XlsNote = 'Old Excel file: numbers stored as values are not read: check by hand'
$script:MsgAttachNote = 'This Outlook message has attachments, which are not read: open them and check by hand'
# The stream name of an attachment, as UTF-16 in the file's directory (bytes as Latin-1 characters).
$script:MsgAttachMark = -join ('__attach_version1.0_'.ToCharArray() | ForEach-Object { [string]$_ + [char]0 })

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
    if (Test-OleName $Text $script:EncryptedStreamNames) { return $true }
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

function Test-OleName {
    # True when the OLE directory (bytes as Latin-1 text) holds one of the names, as UTF-16LE.
    param([string]$Text, [string[]]$Names)
    foreach ($name in $Names) {
        $utf16 = -join @($name.ToCharArray() | ForEach-Object { [string]$_ + [char]0 })
        if ($Text.IndexOf($utf16, [StringComparison]::Ordinal) -ge 0) { return $true }
    }
    return $false
}

function Get-OleLockReason {
    <#
    .SYNOPSIS
      Why an OLE file cannot be read: a sensitivity label with encryption, a password, or '' when it is not encrypted.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    $text = $script:OleLatin1.GetString($Bytes)
    if (Test-OleName $text $script:LabelStreamNames) { return $script:LabelReason }
    if (Test-OleEncrypted $text $Bytes) { return $script:PasswordReason }
    return ''
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
    $lock = Get-OleLockReason $Bytes
    if ($lock) { return @{ Text = ''; Reason = $lock; Note = '' } }
    $text = $script:OleLatin1.GetString($Bytes)
    $runs = New-Object System.Collections.Generic.List[string]
    foreach ($run in [regex]::Matches($text, $script:Utf16RunPattern)) {
        $runs.Add([System.Text.Encoding]::Unicode.GetString($script:OleLatin1.GetBytes($run.Value)))
    }
    foreach ($run in [regex]::Matches($text, $script:AnsiRunPattern)) { $runs.Add($run.Value) }
    if ($runs.Count -eq 0) { return @{ Text = ''; Reason = 'No readable text in this old Office or Outlook file: check by hand'; Note = '' } }
    $note = ''; if ($Extension -eq '.xls') { $note = $script:XlsNote }
    # An Outlook message's attachments are stored as their own streams (__attach_version1.0_#n): not read here.
    if ($Extension -eq '.msg' -and $text.Contains($script:MsgAttachMark)) { $note = $script:MsgAttachNote }
    return @{ Text = ($runs -join "`n"); Reason = ''; Note = $note }
}

Export-ModuleMember -Function Get-OleText, Test-OleSignature, Get-OleLockReason
