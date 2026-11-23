<#
Report.psm1 - the scan results for a person: the report rows and the summary on screen.

  Purpose:     Holds the report rows, sorts them worst PII type first and shows the summary (totals and hits per type).
               The workbook (Summary tab first, one tab per type, Low confidence, Not scanned; the PAM, 2026-11-20: tabs
               to cut noise; no .txt or .csv) is built by ReportBook.psm1, split from this file (fleet v2 DEC 6, 2026-11-23).
  Inputs:      report rows (New-ReportRow) from ChannelScan.psm1; values in them are already masked.
  Outputs:     the summary on screen; report rows.
  Depends on:  PiiDetectors.psm1 (Find-PiiInText), PiiMasking.psm1 (Protect-PiiName, Protect-PiiText), imported first;
               at run time also ScanMeter.psm1 (the screen tables), imported later by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (ReportBook.psm1 saves the workbook).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# PII types from most to least damaging, for sorting and for the order of the tabs. AFI 33-332 gives no ranking
# (A3.1 lists only the factors for judging harm), so this is the PAM's order (2026-11-20): SSN and date of birth
# first, then identity documents and accounts, then contact details.
$script:SeverityOrder = @('SSN', 'Date of birth', 'Passport', 'Payment card', 'DoD ID', 'Home address', 'Name', 'Phone', 'Personal e-mail')

$script:TopItems = 10   # how many files the screen summary lists (the workbook lists all)
# Under the table of types: what High, Medium and Low mean (the PAM, 2026-11-21: plain words for a unit member).
$script:ConfidenceKey = @('High = labelled ("SSN:") or a sure pattern. Medium = right form, no label.', 'Low = weak match, often not PII.')

# Screen colours (the PAM, 2026-11-20: less grey, more format), of the 16 console colours: Cyan = headings and labels,
# White = values, Green = clean or done, Yellow = check by hand and Medium, Red = possible PII, High and errors,
# Gray = hints and folder-row labels, DarkGray = rules and table borders, DarkCyan (teal) = folder names (the PAM, 2026-11-20).
$script:ScreenColors = @{ Title = 'Cyan'; Label = 'Cyan'; Text = 'White'; Good = 'Green'; Accent = 'Yellow'; Alert = 'Red'
    Muted = 'Gray'; Rule = 'DarkGray'; Folder = 'DarkCyan' }

# Not-scanned reasons in groups for the counts on screen and on the Summary tab: a pattern on the reason, its group.
# A file type that is not read is counted per extension (Get-NotScannedGroup).
$script:NotScannedGroups = [ordered]@{
    '^Sensitivity label' = 'Sensitivity label with encryption'
    '^Encrypted' = 'Encrypted (password or author limits)'
    '^An old Office file' = 'Old Office format with a new name'
    '^Larger than' = 'Too large'
    '^Empty file' = 'Empty file'
    '^No text in this PDF' = 'PDF without text (scanned image?)'
    'web page' = 'Download gave a web page'
    '^Download failed' = 'Download failed'
    'longer than the time limit|longer than its time limit' = 'PDF time limit reached (file may be fine)'
    '^Part of this PDF is not well formed|^Part of this PDF could not be read' = 'PDF parts not read (file may open fine)'
    '^Part of this PDF' = 'PDF parts not unpacked'
    '^Not a valid (Office file|PDF)' = 'Not a valid file (damaged or wrong name)'
    '^Could not read' = 'Could not read (in use, damaged or protected)'
    'could not be fully read|^A font in this PDF|^Some characters in this PDF' = 'PDF fonts not fully readable'
    '^Files attached inside this PDF' = 'PDF attachments'
    '^Image: Windows OCR|^Scanned PDF: Image: Windows OCR' = 'Images: OCR not available'
    '^Image:|^Scanned PDF' = 'Images and scans not fully read'
    '^Old Excel file' = 'Old Excel: numbers not read'
    'old Office or Outlook' = 'Old Office or Outlook not readable'
    '^An archive or e-mail inside another|in this archive|Too large to unpack|Could not unpack' = 'Archives: files not unpacked'
}

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the folder (masked by the caller).
    .PARAMETER Link
      A link to the item, or empty.
    .PARAMETER Modified
      When the item last changed, or empty.
    .PARAMETER Detector
      The PII type, or the note word.
    .PARAMETER Confidence
      High, Medium, Low, or empty for a note.
    .PARAMETER MaskedValue
      The masked value, or empty.
    .PARAMETER Context
      The masked snippet, or the reason for a note.
    .OUTPUTS
      [pscustomobject] Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link, and the
      accountability fields (empty until Set-RowAccountability fills them): ModifiedBy, Created, CreatedBy, AddedOn, AddedBy,
      HistoryNote.
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source = $Source; Channel = $Channel; Item = $Item; Detector = $Detector; Confidence = $Confidence
        MaskedValue = $MaskedValue; Context = $Context; Modified = $Modified; Link = $Link
        ModifiedBy = ''; Created = ''; CreatedBy = ''; AddedOn = ''; AddedBy = ''; HistoryNote = ''
    }
}

function Add-IncompleteRow {
    <#
    .SYNOPSIS
      Adds one INCOMPLETE row when the scan did not reach its end and no INCOMPLETE row is there yet.
      Called from the script's 'finally' block, which also runs after Ctrl+C (when 'catch' does not).
    .PARAMETER Rows
      The list of report rows (changed in place).
    .PARAMETER Finished
      True when every step ran.
    .OUTPUTS
      None.
    #>
    param($Rows, [bool]$Finished)
    if ($Finished) { return }
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) { return }
    $Rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
}

function Write-Ui {
    <#
    .SYNOPSIS
      Writes one piece of screen text in a palette colour ($ScreenColors). Every screen line of Ion goes through here.
    .PARAMETER Text
      The text (already masked by the caller).
    .PARAMETER Kind
      Title, Label, Text, Good, Accent, Alert, Muted, Rule or Folder.
    .PARAMETER NoNewline
      Keep writing on the same line (for a line with several colours).
    #>
    param([AllowEmptyString()][string]$Text, [ValidateSet('Title', 'Label', 'Text', 'Good', 'Accent', 'Alert', 'Muted', 'Rule', 'Folder')][string]$Kind = 'Text',
          [switch]$NoNewline)
    Write-Host $Text -ForegroundColor $script:ScreenColors[$Kind] -NoNewline:$NoNewline
}

function Get-IonMark {
    # The app's mark as text: the Greek letter gamma (Logo Studies, the PAM's choice of 2026-11-20), written as a code so the script stays ASCII.
    return [string][char]0x3B3
}

function Write-Banner {
    # A heading line "==== <text> ====" in the title colour. With -Mark (the start line next to "Ion" only; the PAM,
    # 2026-11-21): "==== <mark> <text> ====", the mark in the mark's eye colour (yellow).
    param([string]$Text, [switch]$Mark)
    if (-not $Mark) { Write-Ui ('==== ' + $Text + ' ====') 'Title'; return }
    Write-Ui '==== ' 'Title' -NoNewline; Write-Ui (Get-IonMark) 'Accent' -NoNewline; Write-Ui (' ' + $Text + ' ====') 'Title'
}

function Get-SeverityIndex {
    # Where a PII type stands in $SeverityOrder (0 = worst); a type not in the list comes last.
    param([string]$Type)
    $index = [Array]::IndexOf($script:SeverityOrder, $Type)
    if ($index -lt 0) { return $script:SeverityOrder.Count }
    return $index
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: its worst type, highest confidence, number of hits and types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Worst, Highest, Hits, Channel, Source, Item, Detectors; worst type first, then highest confidence.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Channel, Item)) {
        $first = $group.Group[0]
        $worst = ($group.Group | Sort-Object { Get-SeverityIndex $_.Detector }, { -$script:Rank[$_.Confidence] } | Select-Object -First 1)
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Worst = $worst.Detector; Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = ((@($group.Group | Select-Object -ExpandProperty Detector -Unique) | Sort-Object { Get-SeverityIndex $_ }) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { Get-SeverityIndex $_.Worst } }, @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true },
        @{ Expression = 'Hits'; Descending = $true })
}

function Get-TypeCounts {
    <#
    .SYNOPSIS
      The hits per PII type, worst type first, split by confidence.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, High, Medium, Low, Total.
    #>
    param([object[]]$Hits)
    foreach ($group in ($Hits | Group-Object Detector | Sort-Object { Get-SeverityIndex $_.Name })) {
        $count = @{ High = 0; Medium = 0; Low = 0 }
        foreach ($hit in $group.Group) { $count[$hit.Confidence]++ }
        [pscustomobject]@{ Type = $group.Name; High = $count.High; Medium = $count.Medium; Low = $count.Low; Total = $group.Count }
    }
}

function Get-NotScannedGroup {
    <#
    .SYNOPSIS
      The group of one not-scanned row: a type not read by its extension, else by its reason ($NotScannedGroups).
    .PARAMETER Row
      A report row that is not a hit (NOT SCANNED, BLOCKED or INCOMPLETE).
    .OUTPUTS
      [string] the group name.
    #>
    param($Row)
    if ($Row.Detector -ne 'NOT SCANNED') { return $Row.Detector }   # BLOCKED, INCOMPLETE
    if ($Row.Context -match '^File type') {
        $extension = [System.IO.Path]::GetExtension($Row.Item).ToLower()
        if (-not $extension) { $extension = '(no extension)' }
        return 'Type not read: ' + $extension
    }
    foreach ($pattern in $script:NotScannedGroups.Keys) { if ($Row.Context -match $pattern) { return $script:NotScannedGroups[$pattern] } }
    return 'Other'
}

function Get-NotScannedCounts {
    <#
    .SYNOPSIS
      How many files were not scanned, per reason, most first (the PAM, 2026-11-20: see at once what to check by hand).
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Reason, Count.
    #>
    param([object[]]$Rows)
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    foreach ($group in ($notes | Group-Object { Get-NotScannedGroup $_ } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)) {
        [pscustomobject]@{ Reason = $group.Name; Count = $group.Count }
    }
}

function Get-ScanMetrics {
    <#
    .SYNOPSIS
      The scan's totals, as label and value pairs (the screen summary and the Summary tab show the same).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([object[]]$Rows, [int]$ItemsScanned)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $metrics = [ordered]@{}
    $metrics['Files read'] = $ItemsScanned
    $metrics['Files with possible PII'] = @(Get-ReportSummary $Rows).Count
    $metrics['Possible PII hits'] = $hits.Count
    $metrics['High / Medium / Low'] = '{0} / {1} / {2}' -f @($hits | Where-Object { $_.Confidence -eq 'High' }).Count,
        @($hits | Where-Object { $_.Confidence -eq 'Medium' }).Count, @($hits | Where-Object { $_.Confidence -eq 'Low' }).Count
    $metrics['Not scanned / incomplete'] = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) }).Count
    return $metrics
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The summary for the screen, line by line with its palette colour: the scan's details and totals, a table of
      possible PII per type (worst first; red with High hits, yellow with Medium, white otherwise), a table of the
      files with the most serious hits, and the not-scanned files per reason. Every hit is listed in the workbook.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [pscustomobject] Text, Kind (a Write-Ui colour), LabelLength (characters at the start in the label colour), one per line.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $out = New-Object System.Collections.Generic.List[object]
    # Adds one screen line with its colour; the first LabelLength characters are written in the label colour.
    $add = { param($text, $kind, $labelLength = 0) $out.Add([pscustomobject]@{ Text = $text; Kind = $kind; LabelLength = $labelLength }) }
    & $add '' 'Text'; & $add '==== Results ====' 'Title'
    # A scan that stopped early says so first, in red: its results cover only the files read so far.
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) {
        & $add ('INCOMPLETE: the scan stopped early. Only the {0} {1} read are in these results; scan again.' -f $ItemsScanned, @('files', 'file')[[int]($ItemsScanned -eq 1)]) 'Alert'
    }
    # The values start in one column, after the longest label (labels differ: site, drive, timing lines).
    $pairs = @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())
    $labelWidth = ($pairs | ForEach-Object { ([string]$_.Key).Length + 1 } | Measure-Object -Maximum).Maximum
    foreach ($pair in $pairs) {
        & $add (($pair.Key + ':').PadRight($labelWidth) + ' ' + $pair.Value) 'Text' ($labelWidth + 1)
    }
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    if ($hits.Count -eq 0) { & $add '' 'Text'; & $add 'No possible PII found in the files that were read.' 'Good' }
    else {
        $types = @(Get-TypeCounts $hits)
        & $add '' 'Text'; & $add 'POSSIBLE PII BY TYPE (worst first)' 'Title'
        $kinds = @($types | ForEach-Object { if ($_.High -gt 0) { 'Alert' } elseif ($_.Medium -gt 0) { 'Accent' } else { 'Text' } })
        Get-TableLines @('Type', 'High', 'Medium', 'Low', 'Total') @($types | ForEach-Object { , @($_.Type, $_.High, $_.Medium, $_.Low, $_.Total) }) $kinds |
            ForEach-Object { $out.Add($_) }
        foreach ($keyLine in $script:ConfidenceKey) { & $add $keyLine 'Muted' }
        $top = @(Get-ReportSummary $Rows | Select-Object -First $script:TopItems)
        & $add '' 'Text'; & $add ('FILES WITH THE MOST SERIOUS HITS (first {0}; all are in the workbook)' -f $script:TopItems) 'Title'
        $kinds = @($top | ForEach-Object { if ($_.Highest -eq 'High') { 'Alert' } else { 'Accent' } })
        Get-TableLines @('Worst', 'Highest', 'Hits', $FolderLabel, 'File') @($top | ForEach-Object { , @($_.Worst, $_.Highest, $_.Hits, $_.Channel, $_.Item) }) $kinds |
            ForEach-Object { $out.Add($_) }
    }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        & $add '' 'Text'; & $add 'NOT SCANNED: CHECK THESE BY HAND (the workbook lists each file)' 'Title'
        Get-TableLines @('Reason', 'Files') @($skipped | ForEach-Object { , @($_.Reason, $_.Count) }) @($skipped | ForEach-Object { 'Accent' }) |
            ForEach-Object { $out.Add($_) }
    }
    return $out.ToArray()
}

function Set-RowAccountability {
    <#
    .SYNOPSIS
      Fills the accountability fields of a file's rows: who last changed it, when it was made and by whom, and who
      put the PII in and when (where the file's history shows it; else the file's maker or "not recorded").
    .PARAMETER Rows
      The file's rows (New-ReportRow).
    .PARAMETER Who
      @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; HistoryNote; ByHit } (any may be missing); ByHit:
      "type|masked value" -> @{ AddedOn; AddedBy; Note } for each hit dated from the file's history.
    #>
    param([object[]]$Rows, [hashtable]$Who)
    if (-not $Who) { return }
    foreach ($row in $Rows) {
        foreach ($key in @('ModifiedBy', 'Created', 'CreatedBy', 'AddedOn', 'AddedBy', 'HistoryNote')) {
            if ($Who.ContainsKey($key) -and $Who[$key]) { $row.$key = [string]$Who[$key] }
        }
        $hitKey = $row.Detector + '|' + $row.MaskedValue
        if ($Who.ContainsKey('ByHit') -and $Who.ByHit -and $Who.ByHit.ContainsKey($hitKey)) {
            $entry = $Who.ByHit[$hitKey]
            $row.AddedOn = [string]$entry.AddedOn; $row.AddedBy = [string]$entry.AddedBy
            if ($entry.ContainsKey('Note')) { $row.HistoryNote = [string]$entry.Note }
        }
    }
}

function Add-PartRows {
    <#
    .SYNOPSIS
      The rows of one file inside a zip or an e-mail, named "<file>/<inner path>" (PII in the name masked): its
      hits, a hit for PII in its name, and a NOT SCANNED row when it was not read or read only in part.
    .PARAMETER Part
      One result of Get-TextFromBytes Parts (with its Name).
    .PARAMETER Channel
      The channel name as the report shows it.
    .PARAMETER Item
      The container file as the report shows it.
    .PARAMETER Link
      The container file's link (an inner file has no link of its own).
    .PARAMETER Modified
      The container file's change time.
    .PARAMETER Rows
      The list that receives the rows.
    #>
    param($Part, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified, $Rows)
    $name = $Item + '/' + (Protect-PiiName $Part.Name)
    foreach ($hit in @(Find-PiiInText $Part.Name)) {
        $Rows.Add((New-ReportRow 'File name' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the name of a file inside it'))
    }
    if ($Part.Status -ne 'OK') { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Reason))); return }
    if ($Part.Note) { $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified 'NOT SCANNED' '' '' (Protect-PiiText $Part.Note))) }
    foreach ($hit in @(Find-PiiInText $Part.Text)) {
        $Rows.Add((New-ReportRow 'File' $Channel $name $Link $Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
}

Export-ModuleMember -Function Write-Ui, Get-IonMark, Write-Banner, Get-SeverityIndex, New-ReportRow, Add-IncompleteRow,
    Get-ReportSummary, Get-TypeCounts, Get-NotScannedGroup, Get-NotScannedCounts, Get-ScanMetrics, Get-ScanSummaryText,
    Add-PartRows, Set-RowAccountability
