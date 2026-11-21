<#
ScanQueue.psm1 - turns the reader pool's results into report rows, in the order files finish.

  Purpose:     The link between the scans (ChannelScan.psm1 for SharePoint, DriveScan.psm1 for drives) and the
               reader pool (ReaderPool.psm1): each file handed in gets its rows and is counted on the meter when its
               worker finishes; a file with the same content as one read before gets that file's hits.
  Inputs:      per file: how the report shows it (folder, item, link, change time) and its bytes or path.
  Outputs:     report rows (Report.psm1 New-ReportRow) in the caller's row list; the meter and the live line.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  ReaderPool.psm1, PiiMasking.psm1, Report.psm1 (New-ReportRow, Add-PartRows) and ScanMeter.psm1,
               imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none here.
  Writes:      none (rows in memory).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

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
        Seconds = 0.0 }
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
    foreach ($result in @(Add-PoolJob $Queue.Pool $job)) { Resolve-ScanResult $Queue $result }
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
      { param($Group, $File) } that starts one file; returns true when the file went to the readers.
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    #>
    param([hashtable]$Queue, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile)
    $meter = $Queue.Meter
    for ($c = 0; $c -lt $Groups.Count; $c++) {
        $group = $Groups[$c]
        $before = $meter.Clone()
        foreach ($file in $group.Files) {
            # The file in work, with PII masked (the PAM, 2026-11-20: a slow file is then visible).
            Write-Progress -Id 1 -Activity ("Scanning files - folder {0} of {1}: '{2}'" -f ($c + 1), $Groups.Count, (Protect-PiiName $group.Name)) `
                -Status (Get-ScanMeterText $meter 'file') -PercentComplete (Get-MeterPercent $meter) -CurrentOperation ('Now: ' + (& $ShownFile $group $file))
            $firstRow = $Queue.Rows.Count
            if (-not (& $StartFile $group $file)) {
                Update-ScanMeter $meter $Queue.Rows $firstRow 0
                Write-MeterLine $meter 'file'   # the running total, at most every $LiveLineSeconds
            }
        }
        Wait-ScanQueue $Queue   # the folder's line counts all its files
        Write-ChannelLine $group.Name $before $meter 'file'
    }
    Write-Progress -Id 1 -Activity 'Scanning files' -Completed
}

function Wait-ScanQueue {
    <#
    .SYNOPSIS
      Waits until every file handed in has its rows (the end of a folder, or of the scan).
    .PARAMETER Queue
      Output of New-ScanQueue.
    #>
    param([hashtable]$Queue)
    while ($Queue.Pool.Jobs.Count -gt 0) {
        foreach ($result in @(Receive-PoolResult $Queue.Pool -Wait)) { Resolve-ScanResult $Queue $result }
    }
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
    # The number of files read at once now (for the live line and the header).
    param([hashtable]$Queue)
    return $Queue.Pool.Target
}

Export-ModuleMember -Function New-ScanQueue, Submit-ScanFile, Invoke-GroupScan, Wait-ScanQueue, Close-ScanQueue, Get-QueueWorkers
