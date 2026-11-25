<#
ScanQueue.psm1 - turns the reader pool's results into report rows, in the order files finish.

  Purpose:     The link between the scans (ChannelScan.psm1 for SharePoint, DriveScan.psm1 for drives) and the
               reader pool (ReaderPool.psm1): each file handed in gets its rows and is counted on the meter when its
               worker finishes; a file with the same content as one read before gets that file's hits.
  Inputs:      per file: how the report shows it (folder, item, link, change time) and its bytes or path.
  Outputs:     report rows (Report.psm1 New-ReportRow) in the caller's row list; the meter and the live line.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  ReaderPool.psm1, PiiMasking.psm1, TextExtract.psm1 (Test-SupportedFile), Report.psm1 (New-ReportRow,
               Add-PartRows, Write-Ui) and ScanMeter.psm1, imported first by Find-ChannelPII.ps1. The steps both scans
               share (a file's start rows, the scan run around the groups) live here, once.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none here.
  Writes:      none (rows in memory).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# The "PII first saved" cells of a hit whose version history is not read (an e-signature DoD ID): "Not recorded", the
# word the By account tab leaves out, so no account is made up (review of PR 50); the History note says why.
$script:NotLookedUp = 'Not recorded'
$script:NotLookedUpNote = 'Not looked up: a DoD ID in an e-signature; its version history is not read'

function New-ScanQueue {
    <#
    .SYNOPSIS
      A queue for one scan: the reader pool, the report rows, the meter and what each file in work is.
    .PARAMETER Init
      How the workers load the reader code (New-ReaderPool).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .OUTPUTS
      [hashtable] the queue.
    #>
    param([hashtable]$Init, $Rows, [hashtable]$Meter)
    return @{ Pool = (New-ReaderPool $Init); Rows = $Rows; Meter = $Meter; Info = @{}; Results = @{}; Waiting = @{}; NextId = 0
        Seconds = 0.0; Activity = '' }
}

function Submit-ScanFile {
    <#
    .SYNOPSIS
      Hands one file to the readers; adds the rows of the files that finished meanwhile.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Info
      @{ Channel; Item; Link; Modified } as the report shows the file (already masked).
    .PARAMETER Name
      The file name (its extension decides how it is read).
    .PARAMETER Bytes
      The file bytes (SharePoint), or $null.
    .PARAMETER Path
      The file path when Bytes is $null (drives).
    #>
    param([hashtable]$Queue, [hashtable]$Info, [string]$Name, [byte[]]$Bytes, [string]$Path)
    $Queue.NextId++
    $id = $Queue.NextId
    $Queue.Info[$id] = $Info
    $job = @{ Id = $id; Name = $Name; Bytes = $Bytes; Path = $Path }
    # Keeps the scan bar moving while every worker is busy.
    $onWait = { param($busy) Write-WaitBar $Queue $busy }
    foreach ($result in @(Add-PoolJob $Queue.Pool $job $onWait)) { Resolve-ScanResult $Queue $result }
}

function Resolve-ScanResult {
    <#
    .SYNOPSIS
      Adds the rows of one finished file (and of the files with the same content that waited for it) and counts
      them on the meter.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Result
      A result of Invoke-PooledRead.
    #>
    param([hashtable]$Queue, [hashtable]$Result)
    if ($Result.ContainsKey('Seconds')) { $Queue.Seconds += [double]$Result.Seconds }
    $source = $Result
    if ($Result.ContainsKey('Duplicate') -and $Result.Duplicate) {
        if (-not $Queue.Results.ContainsKey($Result.Original)) {
            if (-not $Queue.Waiting.ContainsKey($Result.Original)) { $Queue.Waiting[$Result.Original] = New-Object System.Collections.Generic.List[int] }
            $Queue.Waiting[$Result.Original].Add($Result.Id)
            return
        }
        $source = $Queue.Results[$Result.Original]
    }
    else { $Queue.Results[$Result.Id] = $Result }
    Add-QueuedRows $Queue $Result.Id $source
    if ($Queue.Waiting.ContainsKey($Result.Id)) {
        foreach ($waiting in $Queue.Waiting[$Result.Id]) { Add-QueuedRows $Queue $waiting $Result }
        $Queue.Waiting.Remove($Result.Id)
    }
}

function Add-QueuedRows {
    <#
    .SYNOPSIS
      The rows of one file from a result (its own, or the one of a file with the same content), then the meter
      and the live line.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Id
      The file's job number.
    .PARAMETER Result
      The result whose hits the file gets.
    #>
    param([hashtable]$Queue, [int]$Id, [hashtable]$Result)
    $info = $Queue.Info[$Id]
    $Queue.Info.Remove($Id)
    $rows = $Queue.Rows
    $first = $rows.Count
    $read = 0
    if ($Result.Status -ne 'OK') {
        $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Reason)))
    }
    else {
        $read = 1
        # A file read in part (a PDF over a cap, damaged, or with attachments): a "check by hand" row next to its hits.
        if ($Result.Note) { $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Note))) }
        foreach ($hit in $Result.Hits) {
            $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
        }
        foreach ($part in $Result.Parts) { Add-PartRows $part $info.Channel $info.Item $info.Link $info.Modified $rows }
        if ($Result.Hits.Count -gt 0 -or $rows.Count -gt $first) { Add-WhoToRows $info $Result @($rows | Select-Object -Skip $first) }
    }
    Update-ScanMeter $Queue.Meter $rows $first $read
    Write-MeterLine $Queue.Meter 'file'
}

function Invoke-GroupScan {
    <#
    .SYNOPSIS
      Scans the groups (folders) one after the other: each file is started (its name checked, then handed to the
      readers or given a NOT SCANNED row), and the folder's line is written when all its files have their rows.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Groups
      The groups, each with Name and Files.
    .PARAMETER StartFile
      { param($Queue, $Group, $File) } that starts one file; returns true when the file went to the readers.
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    #>
    param([hashtable]$Queue, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile)
    $meter = $Queue.Meter
    for ($c = 0; $c -lt $Groups.Count; $c++) {
        $group = $Groups[$c]
        $before = $meter.Clone()
        $Queue.Activity = 'Folder {0} of {1}: {2}' -f ($c + 1), $Groups.Count, (Protect-PiiName $group.Name)
        # A scan with its own way to start a folder's files (ChannelScan.psm1: parallel downloads in a planned order).
        $files = $group.Files
        if ($Queue.ContainsKey('GroupScan') -and $Queue.GroupScan) { & $Queue.GroupScan $Queue $group; $files = @() }
        foreach ($file in $files) {
            # The file in work, with PII masked (the PAM, 2026-11-20: a slow file is then visible).
            Write-ScanBar $Queue ('File: ' + (& $ShownFile $group $file))
            $firstRow = $Queue.Rows.Count
            if (-not (& $StartFile $Queue $group $file)) {
                Update-ScanMeter $meter $Queue.Rows $firstRow 0
                Write-MeterLine $meter 'file'   # the running total, at most every $LiveLineSeconds
            }
        }
        Wait-ScanQueue $Queue   # the folder's line counts all its files
        Write-ChannelLine $group.Name $before $meter 'file'
    }
    Write-Progress -Id 1 -Activity 'Scan' -Completed
}

function Add-WhoToRows {
    <#
    .SYNOPSIS
      The accountability columns of a file's rows (the PAM, 2026-11-21): from its SharePoint fields and version
      history (Info.Who and Info.Added), or from the drive file (Info.DrivePath), plus the names an Office file stores.
    .PARAMETER Info
      The file's queue info.
    .PARAMETER Result
      Its reader result (Hits, Props).
    .PARAMETER FileRows
      The rows of this file.
    #>
    param([hashtable]$Info, [hashtable]$Result, [object[]]$FileRows)
    $props = $null; if ($Result.ContainsKey('Props')) { $props = $Result.Props }
    if ($Info.ContainsKey('DrivePath')) {
        $who = Get-DriveWho $Info.DrivePath $props
        if ($Info.ContainsKey('DriveNote')) { $who.HistoryNote = $Info.DriveNote }   # a OneDrive file not synced yet (DriveScan.psm1)
    }
    elseif ($Info.ContainsKey('Who')) {
        $who = $Info.Who.Clone()
        if ($Info.ContainsKey('Added') -and $Result.Hits.Count -gt 0) { $who.ByHit = Get-AddedByHit $Info $Result.Hits }
    }
    else { return }
    Set-RowAccountability $FileRows $who
}

function Get-AddedByHit {
    <#
    .SYNOPSIS
      Who first saved each hit, from the version history, except for a DoD ID in an e-signature (the only Low DoD ID):
      its history is not read, so a file with only such hits costs no version reads (the PAM, 2026-11-21).
    .PARAMETER Info
      The file's queue info (Added: the version-history lookup).
    .PARAMETER Hits
      The file's hits.
    .OUTPUTS
      [hashtable] "type|masked value" -> @{ AddedOn; AddedBy; Note }.
    #>
    param([hashtable]$Info, [object[]]$Hits)
    # True for an e-signature DoD ID (the only Low DoD ID).
    $signature = { param($hit) $hit.Detector -eq 'DoD ID' -and $hit.Confidence -eq 'Low' }
    $looked = @($Hits | Where-Object { -not (& $signature $_) })
    $byHit = @{}
    if ($looked.Count -gt 0) { $byHit = & $Info.Added $looked }
    foreach ($hit in @($Hits | Where-Object { & $signature $_ })) {
        $key = [string]$hit.Detector + '|' + [string]$hit.MaskedValue
        # The same masked value as a labelled DoD ID of this file keeps that one's history (review of PR 50).
        if ($byHit.ContainsKey($key)) { continue }
        $byHit[$key] = @{ AddedOn = $script:NotLookedUp; AddedBy = $script:NotLookedUp; Note = $script:NotLookedUpNote }
    }
    return $byHit
}

function Wait-ScanQueue {
    <#
    .SYNOPSIS
      Waits until every file handed in has its rows (the end of a folder, or of the scan).
    .PARAMETER Queue
      Output of New-ScanQueue.
    #>
    param([hashtable]$Queue)
    # Keeps the scan bar moving while the last files are read.
    $onWait = { param($busy) Write-WaitBar $Queue $busy }
    while ($Queue.Pool.Jobs.Count -gt 0) {
        foreach ($result in @(Receive-PoolResult $Queue.Pool -Wait -OnWait $onWait)) { Resolve-ScanResult $Queue $result }
    }
}

function Write-WaitBar {
    <#
    .SYNOPSIS
      The scan bar while every worker is busy or a folder waits for its last files: the time keeps moving and the bar
      says what it waits for, so the scan does not look hung (the PAM, 2026-11-21).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Busy
      The number of files being read now.
    #>
    param([hashtable]$Queue, [int]$Busy)
    if (-not $Queue.Activity) { return }
    $noun = 'files'; if ($Busy -eq 1) { $noun = 'file' }
    Write-ScanBar $Queue ('Reading {0} {1} (a large or long file takes longer)...' -f $Busy, $noun)
}

function Write-ScanBar {
    <#
    .SYNOPSIS
      The scan bar: the folder, the run so far, then the file in work (CurrentOperation). ISE shows CurrentOperation
      after the status, and a line break in the status does not show (the PAM's layout test, 2026-11-21); layout
      test 2 (docs/reviews/2026-11-21-ise-progress-test.md) decides how the file gets a line of its own.
    .PARAMETER Queue
      Output of New-ScanQueue (Activity: the folder line; Meter).
    .PARAMETER Third
      The third line: the file in work, or what the scan waits for.
    #>
    param([hashtable]$Queue, [string]$Third)
    Write-Progress -Id 1 -Activity $Queue.Activity -Status (Get-ScanMeterText $Queue.Meter 'file') -PercentComplete (Get-MeterPercent $Queue.Meter) `
        -CurrentOperation $Third
}

function Close-ScanQueue {
    <#
    .SYNOPSIS
      Closes the readers (also after a cancel: workers still running are stopped). Returns the reading time of all
      workers together, for the header.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .OUTPUTS
      [TimeSpan]
    #>
    param([hashtable]$Queue)
    Close-ReaderPool $Queue.Pool
    return [TimeSpan]::FromSeconds($Queue.Seconds)
}

function Get-QueueWorkers {
    # The number of files read at the same time now (for the drive scan's header).
    param([hashtable]$Queue)
    return $Queue.Pool.Target
}

function Add-FileStartRows {
    <#
    .SYNOPSIS
      The rows every file gets before it is read, for both scans: a hit row for PII in its name or folder, then a NOT
      SCANNED row when its type is not read or it is over the size limit.
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Info
      @{ Channel; Item; Link; Modified } as the report shows the file (already masked).
    .PARAMETER RawPath
      The folder and file path as written (checked for PII; never shown).
    .PARAMETER Name
      The file name (its extension decides whether it is read).
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file is to be read; false when it got its NOT SCANNED row.
    #>
    param($Rows, [hashtable]$Info, [string]$RawPath, [string]$Name, [long]$Size, [int]$MaxFileMB)
    foreach ($hit in @(Find-PiiInText $RawPath)) {
        $Rows.Add((New-ReportRow 'File name' $Info.Channel $Info.Item $Info.Link $Info.Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $reason = ''
    if (-not (Test-SupportedFile $Name)) { $reason = 'File type is not read by this script: check by hand' }
    elseif ($Size -gt ($MaxFileMB * 1MB)) { $reason = "Larger than $MaxFileMB MB: check by hand" }
    if ($reason) { $Rows.Add((New-ReportRow 'File' $Info.Channel $Info.Item $Info.Link $Info.Modified 'NOT SCANNED' '' '' $reason)) }
    return (-not $reason)
}

function Confirm-ScanStart {
    <#
    .SYNOPSIS
      Before a scan: how many files and MB were found, and how many will be listed but not read (type or size); then
      "Scan now? [Y/n]" (Enter = yes; the PAM, 2026-11-21: a guard against a scope larger than meant).
    .PARAMETER Groups
      The groups, each with Files.
    .PARAMETER Facts
      { param($File) } that gives @(name, size in bytes) of one file.
    .PARAMETER MaxFileMB
      The size limit.
    .PARAMETER NoAsk
      Show the counts but do not ask: the person already said yes to the whole list (several targets in one run).
    .OUTPUTS
      [bool] true to scan.
    #>
    param([object[]]$Groups, [scriptblock]$Facts, [int]$MaxFileMB, [switch]$NoAsk)
    $count = 0; $bytes = 0L; $skipped = 0
    foreach ($group in $Groups) {
        foreach ($file in $group.Files) {
            $name, $size = & $Facts $file
            $count++; $bytes += [long]$size
            if (-not (Test-SupportedFile $name) -or [long]$size -gt $MaxFileMB * 1MB) { $skipped++ }
        }
    }
    if ($count -eq 0) { return $true }   # nothing to read: no question
    # A count with its noun: "1 file", "2 files".
    $plural = { param($n, $noun) if ($n -eq 1) { "1 $noun" } else { '{0:N0} {1}s' -f $n, $noun } }
    Write-Ui ('Found {0} ({1:N1} MB) in {2}; {3:N0} will be listed but not read (file type, or over {4} MB).' -f
        (& $plural $count 'file'), ($bytes / 1MB), (& $plural $Groups.Count 'folder'), $skipped, $MaxFileMB) 'Text'
    if ($NoAsk) { return $true }
    $answer = Read-Host 'Scan now? [Y/n]'
    return -not ($answer -and $answer.Trim() -match '^(n|no)$')
}

function Invoke-ScanRun {
    <#
    .SYNOPSIS
      One scan of the groups, for both scans: the "Scanning N files" banner, the meter, the readers, every group, then
      the rule and the total line. The run's meter, reading time and worker count go in Run, also when the scan
      stops part way (the caller's 'finally' reads them).
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Groups
      The groups, each with Name and Files.
    .PARAMETER StartFile
      { param($Queue, $Group, $File) } that starts one file (Invoke-GroupScan).
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    .PARAMETER Run
      Receives Meter, Reading ([TimeSpan]) and Workers; Declined = $true when the person answered no to "Scan now?".
      AutoStart = $true (set by the caller): the counts show, but "Scan now?" is not asked. GroupScan (optional, set by
      the caller): { param($Queue, $Group) } that starts all files of one folder in place of the per-file loop.
    .PARAMETER Facts
      { param($File) } that gives @(name, size in bytes) of one file, for the question before the scan.
    .PARAMETER MaxFileMB
      The size limit (files over it are listed, not read).
    #>
    param([hashtable]$Init, $Rows, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile, [hashtable]$Run,
          [scriptblock]$Facts, [int]$MaxFileMB)
    $total = 0
    foreach ($group in $Groups) { $total += $group.Files.Count }
    $Run.Meter = New-ScanMeter 0; $Run.Reading = [TimeSpan]::Zero; $Run.Workers = 0; $Run.Declined = $false
    # Declined until a yes: a cancel (Ctrl+C) at the question saves nothing either, as nothing was read (review of PR 37).
    if ($Facts) { $Run.Declined = $true; if (-not (Confirm-ScanStart $Groups $Facts $MaxFileMB -NoAsk:($Run.ContainsKey('AutoStart') -and $Run.AutoStart))) { return }; $Run.Declined = $false }
    Write-Banner ('Scanning {0} files in {1} folders' -f $total, $Groups.Count)
    $meter = New-ScanMeter $total
    $meter.NameWidth = Get-NameWidth @($Groups | ForEach-Object { Protect-PiiName $_.Name })
    $Run.Meter = $meter; $Run.Reading = [TimeSpan]::Zero; $Run.Workers = 0
    $queue = New-ScanQueue $Init $Rows $meter
    if ($Run.ContainsKey('GroupScan')) { $queue.GroupScan = $Run.GroupScan }   # { param($Queue, $Group) }: see Invoke-GroupScan
    try { Invoke-GroupScan $queue $Groups $StartFile $ShownFile; $Run.Workers = Get-QueueWorkers $queue }
    finally { $Run.Reading = Close-ScanQueue $queue }
    # The folder rows form one table; the total closes it, under a rule.
    Write-RuleLine
    Write-MeterLine $meter 'file' -Force
}

Export-ModuleMember -Function New-ScanQueue, Submit-ScanFile, Invoke-GroupScan, Wait-ScanQueue, Close-ScanQueue, Get-QueueWorkers, Add-FileStartRows,
    Invoke-ScanRun, Write-ScanBar
