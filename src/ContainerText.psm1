<#
ContainerText.psm1 - text extraction from images (Windows OCR), old Office/Outlook OLE files, and archives/e-mail.

  Purpose:     Extracts text in memory from images and scanned JPEG PDF pages (Windows.Media.Ocr), binary .doc,
               .xls, .ppt and .msg files, and .zip archives / .eml e-mail parts, with roster header promotion.
  Inputs:      file bytes (or scan images / HTML / XML nodes).
  Outputs:     text strings and archive/e-mail parts lists.
  Depends on:  nothing (TextExtract.psm1 and PdfExtract.psm1 call into it).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# ==== OcrText.psm1 ====
# Seconds one WinRT step (decode, read) may take; frames read from one multi-page TIFF.
$script:OcrStepMilliseconds = 30000
$script:MaxImageFrames = 30
# A scanned PDF (no text layer), or the pages with no text of a PDF with text: their JPEG images are read, at most this
# many per PDF (Get-ScanPagesText).
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
      The text of a scanned PDF (JPEG page images, no text layer), or of the pages with no text of a PDF with text,
      read with Windows OCR, image by image, at most $MaxScanPages; more images give a note, OCR that is not available
      a reason.
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

# ==== OleText.psm1 ====
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
# old .doc/.xls/.ppt; a password-protected file has StrongEncryptionDataSpace instead (the Administrator, 2026-11-21: tell
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

# ==== ContainerText.psm1 ====
# Zip limits: files read from one archive, the size of one unpacked file, and of all of them together.
$script:MaxArchiveEntries = 2000
$script:MaxEntryBytes = 25MB
$script:MaxArchiveBytes = 200MB

# An e-mail's MIME parts nest at most this deep (a hostile message cannot recurse without end).
$script:MaxMimeDepth = 5
# Parts read from one e-mail (a hostile message can hold hundreds of thousands); the count of the one in work.
$script:MaxMimeParts = 500
# Headers that decide how a part is read: one value each (a second, different one is a conflict, Split-MimeEntity).
$script:DecidingHeaders = @('content-type', 'content-transfer-encoding', 'content-disposition')
# Headers that hold no message text (how a part is coded, and mail-server stamps and signatures): not read as text.
$script:NoTextHeaders = @('content-type', 'content-transfer-encoding', 'content-disposition', 'content-id', 'mime-version',
    'received', 'dkim-signature', 'authentication-results', 'received-spf')
$script:MimeParts = 0

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

function Get-ZipEncryptedFlags {
    <#
    .SYNOPSIS
      For each entry of a zip, in directory order (the order of ZipArchive.Entries): whether it is encrypted (bit 0
      of its flags). Read from the central directory that the end record points to, so bytes inside a file are
      never taken for a record; .NET opens a stored encrypted entry without an error (reviews of PR 19).
    .PARAMETER Bytes
      The archive bytes.
    .OUTPUTS
      [bool[]], or $null when the directory cannot be read.
    #>
    param([byte[]]$Bytes)
    $end = -1
    for ($at = $Bytes.Length - 22; $at -ge [Math]::Max(0, $Bytes.Length - 65557); $at--) {   # the end record: PK 05 06
        if ($Bytes[$at] -eq 0x50 -and $Bytes[$at + 1] -eq 0x4B -and $Bytes[$at + 2] -eq 5 -and $Bytes[$at + 3] -eq 6) { $end = $at; break }
    }
    if ($end -lt 0) { return $null }
    [long]$count = [BitConverter]::ToUInt16($Bytes, $end + 10)
    [long]$offset = [BitConverter]::ToUInt32($Bytes, $end + 16)
    # Zip64: the locator before the end record; the full fields compared as decimals (0xFFFFFFFF would read as -1).
    if (($count -eq 65535 -or $offset -eq 4294967295) -and $end -ge 20) {
        $locator = $end - 20
        if ($Bytes[$locator] -ne 0x50 -or $Bytes[$locator + 1] -ne 0x4B -or $Bytes[$locator + 2] -ne 6 -or $Bytes[$locator + 3] -ne 7) { return $null }
        $zip64 = [BitConverter]::ToInt64($Bytes, $locator + 8)
        if ($zip64 -lt 0 -or $zip64 + 56 -gt $Bytes.Length -or [BitConverter]::ToUInt32($Bytes, [int]$zip64) -ne 0x06064B50) { return $null }
        $count = [BitConverter]::ToInt64($Bytes, $zip64 + 32); $offset = [BitConverter]::ToInt64($Bytes, $zip64 + 48)
    }
    $flags = New-Object System.Collections.Generic.List[bool]
    for ($i = 0; $i -lt $count; $i++) {
        if ($offset -lt 0 -or $offset + 46 -gt $Bytes.Length -or [BitConverter]::ToUInt32($Bytes, [int]$offset) -ne 0x02014B50) { return $null }
        $flags.Add((([BitConverter]::ToUInt16($Bytes, [int]$offset + 8)) -band 1) -eq 1)
        $offset += 46 + [BitConverter]::ToUInt16($Bytes, [int]$offset + 28) + [BitConverter]::ToUInt16($Bytes, [int]$offset + 30) + [BitConverter]::ToUInt16($Bytes, [int]$offset + 32)
    }
    return , $flags.ToArray()
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
        $encrypted = Get-ZipEncryptedFlags $Bytes
        if ($null -eq $encrypted -or $encrypted.Count -ne $zip.Entries.Count) {
            New-Part '(archive)' -Reason 'The archive directory could not be checked for encrypted files: check by hand'
            $encrypted = $null
        }
        $index = -1
        foreach ($entry in $zip.Entries) {
            $index++
            if (-not $entry.Name) { continue }   # a folder
            $count++
            if ($count -gt $script:MaxArchiveEntries) {
                New-Part '(more files)' -Reason "More than $script:MaxArchiveEntries files in this archive: the rest is not scanned: check by hand"
                break
            }
            if ($null -ne $encrypted -and $encrypted[$index]) { New-Part $entry.FullName -Reason 'Encrypted (password-protected) in the archive: check by hand'; continue }
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

function Get-HeaderKey {
    # A header value for comparing two copies: without spaces around ; and =, without quotes, in lower case.
    param([string]$Value)
    return (($Value -replace '\s*([;=])\s*', '$1' -replace '"', '').Trim().ToLowerInvariant())
}

function Split-MimeEntity {
    <#
    .SYNOPSIS
      A MIME entity (a message or one part) as its headers (unfolded, names in lower case) and its body text.
    .PARAMETER Text
      The entity's text (Latin-1 characters, one per byte).
    .OUTPUTS
      [hashtable] Headers, Body, Conflict (a deciding header given twice with different values).
    #>
    param([string]$Text)
    # The headers end at a blank line, or at the first line that is neither "Name: value" nor a continuation
    # (a part may start with text and no blank line): such a line and all after it are the body (reviews of PR 19).
    $end = [regex]::Match($Text, '\A(?:[!-9;-~]+:[^\r\n]*(?:\r?\n[ \t][^\r\n]*)*(?:\r?\n|\z))*')
    $head = $end.Value
    $body = $Text.Substring($end.Length) -replace '\A\r?\n', ''
    # Values per name in a list, joined once (a header given thousands of times stays linear). A name given twice
    # keeps both values, except the headers that decide how the part is read: they keep their first, and a second,
    # different value marks the part Conflict (it is read one way, and gets a "check by hand" reason).
    $values = @{}
    $conflict = $false
    foreach ($line in (($head -replace '\r?\n[ \t]+', ' ') -split '\r?\n')) {
        $colon = $line.IndexOf(':')
        if ($colon -le 0) { continue }
        $name = $line.Substring(0, $colon).Trim().ToLower(); $value = $line.Substring($colon + 1).Trim()
        if (-not $values.ContainsKey($name)) { $values[$name] = New-Object System.Collections.Generic.List[string] }
        elseif ($script:DecidingHeaders -contains $name) {
            if ((Get-HeaderKey $values[$name][0]) -ne (Get-HeaderKey $value)) { $conflict = $true }   # spacing, quotes and case do not count
            continue
        }
        $values[$name].Add($value)
    }
    $headers = @{}
    foreach ($name in $values.Keys) { $headers[$name] = $values[$name] -join "`n" }
    return @{ Headers = $headers; Body = $body; Conflict = $conflict }
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

function Get-HeaderLines {
    <#
    .SYNOPSIS
      The headers of a message or part as "Name: value" lines, encoded words decoded; the headers in $NoTextHeaders
      and ARC-* left out, as they hold no text of the message (Content-Description is kept: it is free text).
    .PARAMETER Headers
      The headers of Split-MimeEntity (names in lower case).
    .OUTPUTS
      [string] lines.
    #>
    param([hashtable]$Headers)
    foreach ($name in @($Headers.Keys | Sort-Object)) {
        if ($script:NoTextHeaders -contains $name -or $name -like 'arc-*') { continue }
        (Get-Culture).TextInfo.ToTitleCase($name) + ': ' + (ConvertFrom-EncodedWords $Headers[$name])
    }
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
    if ($entity.Conflict) { New-Part '(message part)' -Reason 'An e-mail part with conflicting Content headers (it may be read another way): check by hand' }
    if ($Depth -gt 0) {   # a part's own headers other than its content ones (a line such as "Note: ..." is kept)
        $own = @(Get-HeaderLines $entity.Headers)
        if ($own.Count -gt 0) { New-Part '(part headers)' -Text ($own -join "`n") }
    }
    $type = [string]$entity.Headers['content-type']
    $encoding = [string]$entity.Headers['content-transfer-encoding']
    if ($type -match '(?i)^\s*multipart/') {
        $boundary = Get-HeaderParameter $type 'boundary'
        if ($Depth -ge $script:MaxMimeDepth) { New-Part '(message part)' -Reason 'An e-mail part nested too deep to read: check by hand'; return }
        if (-not $boundary) { New-Part '(message text)' -Text $entity.Body; return }   # damaged: the body is still read as text
        $pieces = $entity.Body -split ('(?m)^--' + [regex]::Escape($boundary) + '(?:--)?[ \t]*\r?$')
        for ($i = 1; $i -lt $pieces.Count; $i++) {   # piece 0 is the preamble
            $piece = $pieces[$i] -replace '^\r?\n', ''
            if (-not $piece.Trim()) { continue }
            if ((++$script:MimeParts) -gt $script:MaxMimeParts) {
                New-Part '(more parts)' -Reason "More than $script:MaxMimeParts parts in this e-mail: the rest is not scanned: check by hand"; return
            }
            Get-MimeParts $piece ($Depth + 1)
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
    $lines = Get-HeaderLines $entity.Headers
    New-Part '(message headers)' -Text (@($lines) -join "`n")
    $script:MimeParts = 0
    Get-MimeParts $text 0
}

# ---- Roster columns: a PII label only in a column header (the Administrator, 2026-11-24; design: review of PR 96) ----
# The text stays as it is; each header-column value that no hit touches is written again after it (Add-RosterSection).
# A header cell: the whole cell is one label of this list (case ignored; quotes, and a trailing ':', '#' or '.', allowed).
$script:RosterHeaderPattern = [regex]::new('^\s*"?\s*(?:(?<ssn>SSN|SSAN)|(?<dod>DoD\s*ID|DODID|EDIPI)|(?<dob>DOB|Date\s+of\s+birth)|(?<pp>Passport(?:\s+No)?))\s*[:#.]*\s*"?\s*$', 'IgnoreCase')
$script:RosterLabelText = [ordered]@{ ssn = 'SSN'; dod = 'DoD ID'; dob = 'DOB'; pp = 'Passport' }   # also the order of the section's groups
$script:RosterWords = [regex]::new('SSN|SSAN|DoD|EDIPI|DOB|birth|Passport', 'IgnoreCase')   # none in the first rows: no header (speed)
$script:RosterHeaderRows = 5   # the header row is looked for in this many first rows (the author's choice, 2026-11-24)
$script:RosterCellChars = 40   # a header-like row: no cell longer than this, none with a digit (the author's choice, 2026-11-24)
$script:RosterMaxChars = 1MB   # a longer text gets no section: its read time stays as main's (5 MB: 30 s -> 105 s; the author's choice, 2026-11-24)
# Gap: more than the farthest a detector reads (120: $CacNameChars); '(roster)' stops a name or address running on over spaces.
$script:RosterBreak = "`n" + (' ' * 121) + "`n(roster)`n"

function Get-RosterRowKind {
    # A header-like row (no cell over $RosterCellChars, no digit; an empty row too): its label columns (column -> label; maybe none). Else $null.
    param([string]$Line, $Cells)
    $labels = @{}; foreach ($c in $Cells) {
        $text = $Line.Substring($c[1], $c[2]); if ($c[2] -gt $script:RosterCellChars -or $text -match '\d') { return $null }
        $m = $script:RosterHeaderPattern.Match($text)
        if ($m.Success) { foreach ($k in $script:RosterLabelText.Keys) { if ($m.Groups[$k].Success) { $labels[$c[0]] = $script:RosterLabelText[$k] } } }
    }
    return $labels
}

function Add-RosterValues {
    <#
    .SYNOPSIS
      Adds to Found each cell below the header row (the first header-like row with a label cell, in the first
      $RosterHeaderRows) that is in a header column and holds a digit: @(start in the whole text, length, label). The
      header's columns end at a header-like row with text after an empty or missing row (a second table), and at a csv
      line that opens a quoted field over lines (its columns cannot be told).
    .PARAMETER Lines
      The text lines, one per row (joined by line breaks, they are the text from Offset on).
    .PARAMETER Rows
      A sheet's rows, per row @(texts, column numbers, row number); $null for a .csv file.
    .PARAMETER BreakLength
      The length of the text between two cells of a sheet.
    .PARAMETER Field
      The field pattern of a .csv file (group 1: one field); $null for a sheet.
    .PARAMETER Offset
      Where the first line starts in the whole text.
    .PARAMETER Found
      The list the cells are added to.
    #>
    param([string[]]$Lines, $Rows, [int]$BreakLength, [regex]$Field, [int]$Offset, [System.Collections.Generic.List[object]]$Found)
    if ($Lines.Count -eq 0 -or $Offset + [string]::Join("`n", $Lines).Length -gt $script:RosterMaxChars) { return }
    if (-not $script:RosterWords.IsMatch(($Lines[0..([Math]::Min($script:RosterHeaderRows, $Lines.Count) - 1)] -join "`n"))) { return }
    $cells = New-Object System.Collections.Generic.List[object]; $head = -1; $labels = @{}
    for ($r = 0; $r -lt $Lines.Count; $r++) {   # each cell as [int[]] (column, start, length); later rows only after a header row
        $one = [System.Collections.Generic.List[object]]::new(); $at = 0; $cells.Add($one)   # ::new: New-Object per row is slow
        if ($Field) { foreach ($m in $Field.Matches($Lines[$r])) { $at++; $one.Add([int[]]@($at, $m.Groups[1].Index, $m.Groups[1].Length)) } }
        else { for ($i = 0; $i -lt $Rows[$r][0].Count; $i++) { $one.Add([int[]]@($Rows[$r][1][$i], $at, $Rows[$r][0][$i].Length)); $at += $Rows[$r][0][$i].Length + $BreakLength } }
        if ($head -lt 0) { $kind = Get-RosterRowKind $Lines[$r] $cells[$r]; if ($null -ne $kind -and $kind.Count -gt 0) { $head = $r; $labels = $kind } }
        if ($head -lt 0 -and $r -ge $script:RosterHeaderRows - 1) { return }
    }
    if ($head -lt 0) { return }
    $start = $Offset; $emptyBefore = $false
    for ($r = 0; $r -lt $Lines.Count; $r++) {
        $line = $Lines[$r]; $at = $start; $start += $line.Length + 1
        if ($r -le $head) { continue }
        if ($Field -and ($line.Length - $line.Replace('"', '').Length) % 2 -eq 1) { return }
        $empty = $line -notmatch '[^\s",]'
        if (-not $Field -and $Rows[$r][2] -gt $Rows[$r - 1][2] + 1) { $emptyBefore = $true }   # a missing sheet row is empty
        if ($emptyBefore -and -not $empty -and $null -ne (Get-RosterRowKind $line $cells[$r])) { return }
        $emptyBefore = $empty
        foreach ($c in $cells[$r]) {
            if ($labels.Contains($c[0]) -and $line.Substring($c[1], $c[2]) -match '\d') { $Found.Add([object[]]@(($at + $c[1]), $c[2], $labels[$c[0]])) }
        }
    }
}

function Add-RosterSection {
    <#
    .SYNOPSIS
      Text, then (when a found cell has no hit of Text on it) $RosterBreak and the section: "<Label>: <value>" per
      line, grouped by label (a label never reaches a value of another type), with $RosterBreak between the groups.
    .PARAMETER Text
      The whole text of the file, as main reads it.
    .PARAMETER Found
      The cells of Add-RosterValues, in text order.
    .OUTPUTS
      [string]
    #>
    param([string]$Text, [System.Collections.Generic.List[object]]$Found)
    if ($Found.Count -eq 0 -or $Text.Length -gt $script:RosterMaxChars) { return $Text }
    $hits = @(Get-PiiSpan $Text); $starts = [int[]]@($hits | ForEach-Object { $_.Index }); [Array]::Sort($starts, $hits)
    $groups = [ordered]@{}; foreach ($label in $script:RosterLabelText.Values) { $groups[$label] = [System.Collections.Generic.List[string]]::new() }
    $h = 0
    foreach ($cell in $Found) {
        $end = $cell[0] + $cell[1]
        while ($h -lt $hits.Count -and $hits[$h].Index + $hits[$h].Length -le $cell[0]) { $h++ }
        $hit = $false; for ($j = $h; $j -lt $hits.Count -and $hits[$j].Index -lt $end; $j++) { if ($hits[$j].Index + $hits[$j].Length -gt $cell[0]) { $hit = $true; break } }
        if (-not $hit) { $groups[$cell[2]].Add($cell[2] + ': ' + $Text.Substring($cell[0], $cell[1])) }
    }
    $section = @($groups.Values | Where-Object { $_.Count -gt 0 } | ForEach-Object { $_ -join "`n" })
    if ($section.Count -eq 0) { return $Text }
    return ($Text + $script:RosterBreak + ($section -join $script:RosterBreak))
}

function Add-CsvRosterLabels {
    # A .csv or .tsv text with its roster section (Add-RosterSection); a field is quoted ("...", kept) or runs to the next Delimiter.
    param([string]$Text, [string]$Delimiter)
    $d = [regex]::Escape($Delimiter); $found = New-Object System.Collections.Generic.List[object]
    Add-RosterValues $Text.Split("`n") $null 0 ([regex]::new('(?:^|' + $d + ')("(?:[^"]|"")*"|[^"\r' + $d + ']*)')) 0 $found
    return (Add-RosterSection $Text $found)
}

Export-ModuleMember -Function Get-ImageText, Get-ScanPagesText, Get-OleText, Test-OleSignature, Get-OleLockReason, Get-ArchiveParts, Get-EmailParts, Read-LimitedStream, Add-RosterValues,
    Add-CsvRosterLabels, Add-RosterSection
