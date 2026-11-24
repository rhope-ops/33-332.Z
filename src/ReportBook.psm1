<#
ReportBook.psm1 - the results workbook: its tabs, its Summary rows and where it is saved.

  Purpose:     Builds the workbook tabs (Summary first, Hits and Files as Excel Tables, Not scanned, By account, Method)
               with the review columns, and saves it as a new file under the output folder; the breach advice is on the Summary
               tab and on screen at the end. Split from Report.psm1 (fleet v2 DEC 6, 2026-11-23: 450 lines for
               hand-written code).
  Inputs:      report rows (Report.psm1 New-ReportRow); values in them are already masked.
  Outputs:     <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx, and its path.
  Depends on:  PiiDetectors.psm1 (Get-ConfidenceRank), PiiMasking.psm1 (Protect-PiiName), XlsxWriter.psm1 (New-XlsxSheet, New-XlsxFormula, Export-Xlsx),
               Report.psm1 (Write-Ui, Get-SeverityIndex,
               Get-TypeCounts, Get-NotScannedGroup, Get-NotScannedCounts, Get-ScanMetrics, Get-ScanSummaryText),
               imported first; at run time also ScanMeter.psm1 and Accountability.psm1 (the By account and Method tabs).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the workbook and its site and day folders under the output folder (default Documents\Ion).
               A new file each run; an existing result is never overwritten.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# The tool's name on screen and in the workbook (the PAM, 2026-11-20).
$script:ToolTitle = 'Ion - DoD365 PII Reporter'

$script:MaxNameInPath = 40   # the site name is in the path twice: at most 40 characters keeps it under 260 (review of PR 41)
# C6, the review columns the person fills in (the PAM, 2026-11-21), each with its dropdown (empty = free text); Ion never fills them.
# Order: the answers first, then who and when (the PAM, 2026-11-24, the approved preview).
$script:ReviewColumns = [ordered]@{ 'Is it PII?' = @('Yes', 'No', 'Not sure')
    'Action taken' = @('Removed', 'Moved', 'Access limited', 'Left (not PII)', 'Other')
    'Reported to Privacy Manager' = @('Yes', 'No', 'Not needed'); 'Reviewed by' = @(); 'Review date' = @() }
# What to do with PII found where it should not be (rule source: AFI 33-332 para 3.1.2 as changed by AFI33-332_DAFGM2026-01,
# 7 July 2026; edition: docs/DECISIONS.md). The wording is the README's "If it finds PII" section; scripts/build_site.py --check fails when they differ (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01 (7 July 2026), ' +
                       'page 1, second paragraph).'
# "Start here" on the Summary tab: 4 steps for the reviewer (approved 2026-11-21; STE: one action a step, checked by hand,
# as scripts/check_instructions.py cannot read text in code).
$script:StartHere = @(@('Start here: 1', 'Open the tab Hits. In the column Confidence, filter to High. Do the High hits first.'),
    @('Start here: 2', 'For each hit, click the file name. Find the masked value in the file, then check it.'),
    @('Start here: 3', 'In the same row, select Is it PII?, Action taken and Reported to Privacy Manager.'),
    @('Start here: 4', 'If it is PII where it should not be, report it to your servicing Privacy Manager/Monitor without delay.'))
# Shown as two lines, the advice and its rule source (the PAM, 2026-11-21); one line if no "(AFI ...)." part is found.
$script:BreachLines = @($script:BreachAdvice)
if ($script:BreachAdvice -match '^(.*) \((AFI .*)\)\.$') { $script:BreachLines = @(($Matches[1] + '.'), ('Source: ' + $Matches[2] + '.')) }

# =====================================================================================
# The workbook
# =====================================================================================

function Get-ResultPath {
    <#
    .SYNOPSIS
      A new workbook path: <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx (the PAM, 2026-11-21), with (2), (3) ...
      if the name exists, so a result is never overwritten. Creates the folder. The site name is masked like any name the
      tool shows (found PII only); characters Windows refuses in a folder name are replaced.
    .PARAMETER OutputFolder
      The base folder.
    .PARAMETER SiteUrl
      The site address; its last part (teams/<name>) names the site folder.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$SiteUrl)
    $siteName = [uri]::UnescapeDataString(($SiteUrl.TrimEnd('/') -split '/')[-1])
    $siteName = ((Protect-PiiName $siteName) -replace '[\\/:*?"<>|]', '-').Trim(' .')
    if (-not $siteName) { $siteName = 'site' }
    if ($siteName.Length -gt $script:MaxNameInPath) { $siteName = $siteName.Substring(0, $script:MaxNameInPath).Trim(' .') }   # MAX_PATH 260
    if ($siteName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $siteName = '_' + $siteName }   # names Windows keeps for devices
    $now = Get-Date
    $folder = Join-Path $OutputFolder $siteName
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
    $base = 'Ion {0} {1}' -f $siteName, $now.ToString('yyyy-MM-dd HHmm')
    $path = Join-Path $folder ($base + '.xlsx')
    $copy = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $folder ('{0} ({1}).xlsx' -f $base, $copy); $copy++ }
    return $path
}

function Get-HitCells {
    # One row of the Hits tab, in the order of its columns: ID, what was found, the review columns, then who and when.
    param($Hit, [int]$Number)
    @(('H{0:000}' -f $Number), $Hit.Detector, $Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context) +
        @('') * $script:ReviewColumns.Count +
        @($Hit.Modified, $Hit.ModifiedBy, $Hit.AddedOn, $Hit.AddedBy, $Hit.Created, $Hit.CreatedBy, $Hit.HistoryNote, $Hit.Source, $Hit.Link)
}

function Get-FileRows {
    <#
    .SYNOPSIS
      The rows of the Files tab: one per file with hits, worst first: its counts, its types, how many hits a reviewer
      has answered "Is it PII?" for (a formula on the Hits table), and who saved it.
    .PARAMETER Hits
      The hit rows, in the Hits tab's order.
    .PARAMETER FolderLabel
      The name of the folder column (Channel or Folder), as in the Hits table.
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Hits, [string]$FolderLabel)
    $groups = @($Hits | Group-Object { $_.Channel + "`n" + $_.Item } | ForEach-Object {
        $g = @($_.Group)
        [pscustomobject]@{ First = $g[0]; Count = $g.Count; High = @($g | Where-Object { $_.Confidence -eq 'High' }).Count
            Medium = @($g | Where-Object { $_.Confidence -eq 'Medium' }).Count; Low = @($g | Where-Object { $_.Confidence -eq 'Low' }).Count
            Types = (@($g | ForEach-Object { $_.Detector } | Select-Object -Unique) -join ', ')
            AddedBy = (@($g | ForEach-Object { $_.AddedBy } | Where-Object { $_ } | Select-Object -Unique) -join ', ') }
    } | Sort-Object @{ Expression = 'High'; Descending = $true }, @{ Expression = 'Medium'; Descending = $true },
        @{ Expression = 'Low'; Descending = $true }, @{ Expression = { $_.First.Channel } }, @{ Expression = { $_.First.Item } })
    $row = 1
    foreach ($f in $groups) {
        $row++
        $reviewed = New-XlsxFormula ('SUMPRODUCT((tblHits[{0}]=A{1})*(tblHits[File]=B{1})*(tblHits[Is it PII?]<>""))' -f $FolderLabel, $row) 0
        , @($f.First.Channel, $f.First.Item, [long]$f.Count, [long]$f.High, [long]$f.Medium, [long]$f.Low, $f.Types, $reviewed,
            $f.First.Modified, $f.First.ModifiedBy, $f.AddedBy, $f.First.Link)
    }
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first; Hits (every hit, one Excel Table: filter, sort, pivot) and Files (one row
      per file) when there are hits; then Not scanned, By account and Method (the PAM, 2026-11-24: as previewed).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the Summary tab.
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site): the name of the folder column.
    .OUTPUTS
      [object[]] New-XlsxSheet objects.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $hits = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -gt 0 } |
        Sort-Object @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true }, @{ Expression = { Get-SeverityIndex $_.Detector } }, Channel, Item)
    # The review columns right after the hit, then who and when (accountability names: the PAM, 2026-11-21; the Method tab explains them).
    $columns = @('ID', 'Type', 'Confidence', $FolderLabel, 'File', 'Masked value', 'Context') + @($script:ReviewColumns.Keys) +
        @('Last saved', 'Last saved by', 'PII first saved', 'PII first saved by', 'Created', 'Created by', 'History note', 'Found in', 'Link')
    $widths = @(7, 15, 11, 16, 40, 16, 50) + @(16) * $script:ReviewColumns.Count + @(16, 22, 16, 22, 16, 22, 40, 10, 50)
    $choices = @{}; $offset = 7
    foreach ($name in @($script:ReviewColumns.Keys)) { if ($script:ReviewColumns[$name]) { $choices[$offset] = $script:ReviewColumns[$name] }; $offset++ }
    $after = 7 + $script:ReviewColumns.Count   # the first column after the review columns
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header $hits) @(34, 70, 9, 9, 9, 10)))
    if ($hits.Count -gt 0) {
        $n = 0
        $sheets.Add((New-XlsxSheet 'Hits' $columns @($hits | ForEach-Object { $n++; , (Get-HitCells $_ $n) }) $widths -List -Table 'tblHits' `
            -LinkColumn 4 -LinkTarget ($columns.Count - 1) -Choices $choices -WrapColumns @(6) -DateColumns @($after, ($after + 2), ($after + 4)) `
            -DateChecks @(7 + @($script:ReviewColumns.Keys).IndexOf('Review date')) `
            -FreezeColumns 5 -FillColumn 2))
        $sheets.Add((New-XlsxSheet 'Files' @($FolderLabel, 'File', 'Hits', 'High', 'Medium', 'Low', 'Types', 'Reviewed', 'Last saved',
            'Last saved by', 'PII first saved by', 'Link') @(Get-FileRows $hits $FolderLabel) @(16, 40, 7, 7, 9, 7, 40, 10, 16, 22, 30, 50) `
            -List -Table 'tblFiles' -LinkColumn 1 -LinkTarget 11 -DateColumns @(8) -FreezeColumns 2))
    }
    $notes = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -eq 0 })
    if ($notes.Count -gt 0) {
        # A Group column, so the tab filters by the groups the Summary counts (the PAM, 2026-11-21).
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'File', 'Reason', 'Group', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, (Get-NotScannedGroup $_), $_.Link) }) `
            @(14, 22, 45, 70, 30, 50) -List -Table 'tblNotScanned' -LinkColumn 2 -LinkTarget 5))
    }
    $accounts = @(Get-AccountRows $hits)   # C6 (the PAM, 2026-11-21): whom to contact first
    if ($accounts.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'By account' @('PII first saved by', 'Files', 'High', 'Medium', 'Low') $accounts @(40, 9, 9, 9, 9) -List -Table 'tblAccounts'))
    }
    $method = @(Get-MethodRows | Where-Object { $accounts.Count -gt 0 -or $_[0] -ne 'By account tab' })   # no row for a tab that is not there
    $sheets.Add((New-XlsxSheet 'Method' @('Topic', 'How Ion works it out') $method @(34, 110) -List -Table 'tblMethod' -WrapColumns @(1)))   # the PAM, 2026-11-21
    return $sheets.ToArray()
}

function Get-TypeTableRows {
    <#
    .SYNOPSIS
      The Summary tab's hits per type (worst first): High, Medium, Low, Total and Reviewed, as formulas on the Hits table
      with Ion's values, and a Total row.
    .PARAMETER Hits
      The hit rows (the Hits tab).
    .PARAMETER StartRow
      The sheet row of the table's header row.
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Hits, [int]$StartRow)
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add(@('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total', 'Reviewed'))
    $first = $StartRow + $list.Count   # the sheet row of the first type (row 1 is the title)
    foreach ($type in @(Get-TypeCounts $Hits)) {
        $r = $StartRow + $list.Count
        $cells = @($type.Type)
        foreach ($level in @(@('High', $type.High), @('Medium', $type.Medium), @('Low', $type.Low))) {
            $cells += New-XlsxFormula ('COUNTIFS(tblHits[Type],$A{0},tblHits[Confidence],"{1}")' -f $r, $level[0]) $level[1]
        }
        $list.Add($cells + @((New-XlsxFormula ('SUM(B{0}:D{0})' -f $r) $type.Total), (New-XlsxFormula ('COUNTIFS(tblHits[Type],$A{0},tblHits[Is it PII?],"<>")' -f $r) 0)))
    }
    if ($Hits.Count -eq 0) { $list.Add(@('None found in the files that were read.')) }
    else {
        $last = $StartRow + $list.Count - 1
        $total = @('Total')
        foreach ($c in @('B', 'C', 'D', 'E', 'F')) {
            $value = switch ($c) { 'B' { @($Hits | Where-Object { $_.Confidence -eq 'High' }).Count } 'C' { @($Hits | Where-Object { $_.Confidence -eq 'Medium' }).Count }
                'D' { @($Hits | Where-Object { $_.Confidence -eq 'Low' }).Count } 'E' { $Hits.Count } default { 0 } }
            $total += New-XlsxFormula ('SUM({0}{1}:{0}{2})' -f $c, $first, $last) $value
        }
        $list.Add($total)
    }
    return $list.ToArray()   # one row per item (each row stays an array)
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the reviewer's progress, the hits per type (worst
      first; counts as formulas on the Hits table, so they follow a reviewer's filter-free edits), the not-scanned files
      per reason, then how to read the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER Hits
      The hit rows (the Hits tab).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [object[]]$Hits)
    $out = New-Object System.Collections.Generic.List[object]
    if ($Hits.Count -gt 0) { foreach ($step in $script:StartHere) { $out.Add(@($step[0], $step[1])) }; $out.Add(@('')) }
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) {
        if ($pair.Key -ne 'High / Medium / Low') { $out.Add(@($pair.Key, $pair.Value)) }   # the type table below has these as numbers
        if ($pair.Key -eq 'Run') { $out.Add(@('Finished', (Get-Date -Format 'yyyy-MM-dd HH:mm'))) }
    }
    if ($Hits.Count -gt 0) {
        $out.Add(@('Hits reviewed ("Is it PII?" filled in)', (New-XlsxFormula 'COUNTIF(tblHits[Is it PII?],"<>")' 0)))
        $out.Add(@('Confirmed PII ("Is it PII?" = Yes)', (New-XlsxFormula 'COUNTIF(tblHits[Is it PII?],"Yes")' 0)))
    }
    $out.Add(@(''))
    foreach ($row in @(Get-TypeTableRows $Hits ($out.Count + 2))) { $out.Add($row) }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        $out.Add(@(''))
        $out.Add(@('Not scanned: check by hand', 'Files'))
        foreach ($group in $skipped) { $out.Add(@($group.Reason, $group.Count)) }
    }
    $out.Add(@(''))
    $out.Add(@('Tabs', ('Hits: one row per hit, one Excel Table (filter, sort, pivot); fill in the review columns there. ' +
        'Files: one row per file. Then Not scanned, By account and Method (how each column is worked out).')))
    $out.Add(@('Confidence', 'A pattern guess, not proof: open each item and check it before you act.'))
    $out.Add(@('If you find PII', $script:BreachLines[0]))
    if ($script:BreachLines.Count -gt 1) { $out.Add(@('Rule source', ($script:BreachLines[1] -replace '^Source: ', ''))) }
    $out.Add(@('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.'))
    return , $out.ToArray()
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the workbook (Summary tab first). Returns its path.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the summary.
    .PARAMETER OutputFolder
      The base folder; the workbook goes in <site> under it.
    .PARAMETER SiteUrl
      The site address (names the site folder).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$OutputFolder, [string]$SiteUrl,
          [string]$FolderLabel = 'Folder')
    foreach ($line in @(Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header -FolderLabel $FolderLabel)) {
        if ($line.LabelLength -gt 0) { Write-Ui $line.Text.Substring(0, $line.LabelLength) 'Label' -NoNewline; Write-Ui $line.Text.Substring($line.LabelLength) $line.Kind }
        else { Write-Ui $line.Text $line.Kind }
    }
    # The advice comes after the saved path (the PAM, 2026-11-21), and also when the save fails ('finally').
    try {
        $path = Get-ResultPath $OutputFolder $SiteUrl
        Export-Xlsx -Path $path -Sheets (Get-WorkbookSheets $Rows $ItemsScanned $Header $FolderLabel)
        Write-Ui ''; Write-Ui "Saved (values masked): $path" 'Good'; Write-Ui 'Handle this file as CUI: it shows where PII is.' 'Accent'
    }
    finally {
        Write-Ui ''; Write-Ui 'Confidence is a pattern guess, not proof: open each item and check it before you act.' 'Muted'
        Write-Ui $script:BreachLines[0] 'Accent'; if ($script:BreachLines.Count -gt 1) { Write-Ui $script:BreachLines[1] 'Muted' }
    }
    return $path
}

Export-ModuleMember -Function Get-WorkbookSheets, Get-ResultPath, Publish-ScanReport
