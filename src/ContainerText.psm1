<#
ContainerText.psm1 - unpacks files that hold other files: zip archives and e-mails (.eml), in memory.

  Purpose:     The parts of a .zip (each file inside) and of an .eml (its headers and text, and each attachment),
               so TextExtract.psm1 reads each part with its own reader and the report names it as
               "archive.zip/inner/file" (the PAM, 2026-11-20: round 3 readers). Limits on the number of files and
               the unpacked size guard against zip bombs; a part over a limit gets a reason instead of its bytes.
               It also labels roster columns: a PII label only in a column header of a worksheet or .csv/.tsv file is
               written before each value below it, so the detectors read it (the PAM, 2026-11-24).
  Inputs:      the file bytes; for roster columns, the rows of a worksheet or the text of a .csv/.tsv file.
  Outputs:     parts: Name, and Bytes (a file to read), Text (text already read) or Reason (not read, and why); text.
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Depends on:  PiiDetectors.psm1 (Get-PiiSpan: Add-RosterSection), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

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

# ---- Roster columns: a PII label only in a column header (the PAM, 2026-11-24; design: review of PR 96) ----
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

Export-ModuleMember -Function Get-ArchiveParts, Get-EmailParts, Read-LimitedStream, Add-RosterValues, Add-CsvRosterLabels, Add-RosterSection
