<#
PdfText.psm1 - the text of a PDF file, in memory, with .NET only (no module; the PAM, 2026-11-20).

  Purpose:     Page text (Tj/TJ text in Flate, ASCII85 or hex streams, with ToUnicode font maps), filled-in form-field
               values, comments and file info, so the detectors can check a PDF (ISO 32000). An encrypted PDF,
               a PDF with no text (a scanned image) or a file that is not a PDF gets a reason instead.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text and Reason (Get-PdfText).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters), PdfObjects.psm1 (objects, strings, dictionary text),
               PdfFonts.psm1 (font maps), PdfContent.psm1 (the text of a content stream; split from this file,
               fleet v2 DEC 6, 2026-11-23) and OcrText.psm1 (scanned pages), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:PdfLatin1 = [System.Text.Encoding]::GetEncoding(28591)


$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'
# A scanned PDF's JPEG page images, found by Get-PdfAllText: read with Windows OCR when the PDF has no text.
$script:PdfScanImages = @()
# Not-scanned reasons, each its own group on the Not scanned tab (the PAM, 2026-11-21: tell labels from passwords, and
# a time limit from damage: many PDFs called damaged open fine). The limits are $PdfSecondsLimit and $PdfRegexTimeout of PdfCore.psm1.
$script:PdfSlowReason = 'This PDF took longer than the time limit to read (2 minutes; the file may be fine): check by hand'
$script:PdfSearchSlowReason = 'One search in this PDF took longer than its time limit (2 seconds; the file may be fine): check by hand'
$script:PdfPartReason = 'Part of this PDF could not be read (an unusual string; the file may still open fine): check by hand'
$script:PdfEncryptedReason = 'Encrypted PDF (a password, or limits set by its author; it may open without one): check by hand'
# Microsoft Purview protection: an encrypted PDF whose security handler is Microsoft's, or a wrapper PDF that holds the
# protected document as an encrypted payload (ISO 32000-2; its own pages only say the file is protected).
$script:PdfLabelReason = 'Sensitivity label with encryption (Microsoft Purview): open it and check by hand'
$script:PdfLabelPattern = '/Filter\s*/MicrosoftIRMServices\b|/EncryptedPayload\b'

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
    if ($text -match $script:PdfLabelPattern) { return @{ Text = ''; Reason = $script:PdfLabelReason } }
    if ($text -match '/Encrypt\s*(?:\d+\s+\d+\s+R|<<)') { return @{ Text = ''; Reason = $script:PdfEncryptedReason } }
    Reset-PdfOldObjects   # this PDF's own older versions only
    $script:PdfScanImages = @()   # and its own scanned pages
    Start-PdfBudget
    try { $all = Get-PdfAllText $text }
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSearchSlowReason } }
            if ($e -is [System.TimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
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

Export-ModuleMember -Function Get-PdfText
