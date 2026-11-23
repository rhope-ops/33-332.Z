<#
ReportBook.psm1 - the results workbook: its tabs, its Summary rows and where it is saved.

  Purpose:     Builds the workbook tabs (Summary first, one tab per PII type, Low confidence, Not scanned) with the
               review columns, and saves it as a new file under the output folder; the breach advice is on the Summary
               tab and on screen at the end. Split from Report.psm1 (fleet v2 DEC 6, 2026-11-23: 450 lines for
               hand-written code).
  Inputs:      report rows (Report.psm1 New-ReportRow); values in them are already masked.
  Outputs:     <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx, and its path.
  Depends on:  PiiDetectors.psm1 (Get-ConfidenceRank), PiiMasking.psm1 (Protect-PiiName), XlsxWriter.psm1,
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
$script:ReviewColumns = [ordered]@{ 'Reviewed by' = @(); 'Review date' = @(); 'Is it PII?' = @('Yes', 'No', 'Not sure')
    'Action taken' = @('Removed', 'Moved', 'Access limited', 'Left (not PII)', 'Other')
    'Reported to Privacy Manager' = @('Yes', 'No', 'Not needed') }
# What to do with PII found where it should not be (rule source: AFI 33-332 para 3.1.2 as changed by AFI33-332_DAFGM2026-01,
# 7 July 2026; edition: docs/DECISIONS.md). The wording is the README's "If it finds PII" section; scripts/build_site.py --check fails when they differ (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01 (7 July 2026), ' +
                       'page 1, second paragraph).'
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
    # One row of a hit tab: the hit's cells in the order of $HitColumns.
    param($Hit)
    @($Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context, $Hit.Modified, $Hit.Link, $Hit.Source,
        $Hit.Created, $Hit.CreatedBy, $Hit.ModifiedBy, $Hit.AddedOn, $Hit.AddedBy, $Hit.HistoryNote) + @('') * $script:ReviewColumns.Count
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first, then one tab per PII type with High or Medium hits (worst type
      first), then Low confidence and Not scanned when they have rows (the PAM, 2026-11-20).
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
    $hits = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -gt 0 })
    # Accountability columns, names and order: the PAM, 2026-11-21 (docs/DECISIONS.md); the Method tab explains them.
    $columns = @('Confidence', $FolderLabel, 'File', 'Masked value', 'Context', 'Last saved', 'Link', 'Found in',
        'Created', 'Created by', 'Last saved by', 'PII first saved', 'PII first saved by', 'History note') + @($script:ReviewColumns.Keys)
    $widths = @(11, 22, 45, 18, 60, 17, 50, 10, 17, 30, 30, 17, 30, 50) + @(20) * $script:ReviewColumns.Count
    $choices = @{}; $offset = $columns.Count - $script:ReviewColumns.Count; $lowChoices = @{}   # the review columns' dropdowns
    foreach ($name in @($script:ReviewColumns.Keys)) {
        if ($script:ReviewColumns[$name]) { $choices[$offset] = $script:ReviewColumns[$name]; $lowChoices[$offset + 1] = $script:ReviewColumns[$name] }
        $offset++
    }
    $byRank = @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header) @(28, 70, 9, 9, 9)))
    foreach ($type in @(Get-TypeCounts $hits | Where-Object { $_.High + $_.Medium -gt 0 } | ForEach-Object { $_.Type })) {
        $typeHits = @($hits | Where-Object { $_.Detector -eq $type -and $_.Confidence -ne 'Low' } | Sort-Object $byRank, Channel, Item)
        $sheets.Add((New-XlsxSheet $type $columns @($typeHits | ForEach-Object { , (Get-HitCells $_) }) $widths -List -LinkColumn 6 -Choices $choices))
    }
    $low = @($hits | Where-Object { $_.Confidence -eq 'Low' } | Sort-Object { Get-SeverityIndex $_.Detector }, Channel, Item)
    if ($low.Count -gt 0) {
        $lowCells = @($low | ForEach-Object { , (@($_.Detector) + (Get-HitCells $_)) })
        $sheets.Add((New-XlsxSheet 'Low confidence' (@('Type') + $columns) $lowCells (@(16) + $widths) -List -LinkColumn 7 -Choices $lowChoices))
    }
    $notes = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -eq 0 })
    if ($notes.Count -gt 0) {
        # A Group column, so the tab filters by the groups the Summary counts (the PAM, 2026-11-21).
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'File', 'Reason', 'Group', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, (Get-NotScannedGroup $_), $_.Link) }) `
            @(14, 22, 45, 70, 30, 50) -List -LinkColumn 5))
    }
    $accounts = @(Get-AccountRows $hits)   # C6 (the PAM, 2026-11-21): whom to contact first
    if ($accounts.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'By account' @('PII first saved by', 'Files', 'High', 'Medium', 'Low') $accounts @(40, 9, 9, 9, 9) -List))
    }
    $method = @(Get-MethodRows | Where-Object { $accounts.Count -gt 0 -or $_[0] -ne 'By account tab' })   # no row for a tab that is not there
    $sheets.Add((New-XlsxSheet 'Method' @('Topic', 'How Ion works it out') $method @(34, 110) -List -WrapColumns @(1)))   # the PAM, 2026-11-21
    return $sheets.ToArray()
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the hits per type (worst first), the not-scanned
      files per reason, then how to read the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header)
    $hits = @($Rows | Where-Object { (Get-ConfidenceRank $_.Confidence) -gt 0 })
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) { , @($pair.Key, $pair.Value) }
    , @('')
    , @('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total')
    foreach ($type in @(Get-TypeCounts $hits)) { , @($type.Type, $type.High, $type.Medium, $type.Low, $type.Total) }
    if ($hits.Count -eq 0) { , @('None found in the files that were read.') }
    $skipped = @(Get-NotScannedCounts $Rows)
    if ($skipped.Count -gt 0) {
        , @('')
        , @('Not scanned: check by hand', 'Files')
        foreach ($group in $skipped) { , @($group.Reason, $group.Count) }
    }
    , @('')
    , @('Tabs', 'One tab per type with High or Medium hits, worst type first; then Low confidence; then Not scanned; then Method (how each column is worked out).')
    , @('Confidence', 'A pattern guess, not proof: open each item and check it before you act.')
    , @('If you find PII', $script:BreachLines[0])
    if ($script:BreachLines.Count -gt 1) { , @('Rule source', ($script:BreachLines[1] -replace '^Source: ', '')) }
    , @('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.')
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
