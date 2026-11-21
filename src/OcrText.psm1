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

Set-StrictMode -Version 2.0

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

Export-ModuleMember -Function Get-ImageText, Get-ScanPagesText
