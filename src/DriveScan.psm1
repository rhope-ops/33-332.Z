<#
DriveScan.psm1 - scans a folder on a mapped or network drive (X:\<folder> or \\<server>\<share>\<folder>).

  Purpose:     The PAM, 2026-11-21: the same check as for a SharePoint site, for one folder and all its subfolders on
               a drive the person can open. Nothing above the folder, no other share or server is touched ("No
               reconnaissance"). Files are opened read-only and shared, with the person's own Windows logon.
  Inputs:      the pasted folder path; the reader workers' set-up (New-ReaderPool).
  Outputs:     report rows; the results workbook (Report.psm1 Publish-ScanReport).
  How to run:  imported by Find-ChannelPII.ps1; Invoke-DriveScan runs when the pasted target is a folder path.
  Depends on:  PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1 (Test-SupportedFile), Report.psm1,
               ScanMeter.psm1, ScanQueue.psm1 and ChannelScan.psm1 (Get-ShownItem), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7, where the tests run on a temporary folder).
  Network:     only the given folder's server, through the person's Windows logon (SMB); nothing else.
  Writes:      only the results workbook under the results folder.
  Credentials: none (the Windows logon).
#>

Set-StrictMode -Version 2.0

# Files directly in the pasted folder form this group (each subfolder is its own group).
$script:DriveTopGroup = '(top of the folder)'

function Get-DriveGroups {
    <#
    .SYNOPSIS
      The files under one folder, grouped by its top-level subfolders (then the files at its top), with the folders
      that could not be opened and the folder links that were not followed (a link can lead outside the folder).
    .PARAMETER Root
      The folder path.
    .OUTPUTS
      [hashtable] Groups (each @{ Name; Files }), Blocked (each @{ Folder; Reason }), Links (folder paths).
    #>
    param([string]$Root)
    $Root = $Root.TrimEnd('\', '/')
    $slashes = [char[]]@('\', '/')   # Windows uses \; the tests also run where / is the separator
    $result = @{ Groups = (New-Object System.Collections.Generic.List[object]); Blocked = (New-Object System.Collections.Generic.List[object])
        Links = (New-Object System.Collections.Generic.List[string]) }
    $byName = @{}
    $top = @{ Name = $script:DriveTopGroup; Files = (New-Object System.Collections.Generic.List[object]) }
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $folder = $stack.Pop()
        $relative = $folder.Substring($Root.Length).TrimStart($slashes)
        try { $entries = @((New-Object System.IO.DirectoryInfo($folder)).GetFileSystemInfos()) }
        catch { $result.Blocked.Add(@{ Folder = $relative; Reason = $_.Exception.Message }); continue }
        foreach ($entry in $entries) {
            $rel = ($relative + '\' + $entry.Name).TrimStart($slashes)
            if ($entry -is [System.IO.DirectoryInfo]) {
                if ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { $result.Links.Add($rel); continue }
                $stack.Push($entry.FullName)
                continue
            }
            $groupName = $script:DriveTopGroup
            $inner = $rel
            $cut = $rel.IndexOfAny($slashes)
            if ($cut -gt 0) { $groupName = $rel.Substring(0, $cut); $inner = $rel.Substring($cut + 1) }
            $group = $top
            if ($groupName -ne $script:DriveTopGroup) {
                if (-not $byName.ContainsKey($groupName)) { $byName[$groupName] = @{ Name = $groupName; Files = (New-Object System.Collections.Generic.List[object]) } }
                $group = $byName[$groupName]
            }
            $group.Files.Add(@{ Path = $entry.FullName; Inner = ($inner -replace '[\\/]', '/'); Name = $entry.Name; Size = [long]$entry.Length
                Modified = $entry.LastWriteTime.ToString('yyyy-MM-dd HH:mm') })
        }
    }
    foreach ($name in @($byName.Keys | Sort-Object)) { $result.Groups.Add($byName[$name]) }
    if ($top.Files.Count -gt 0) { $result.Groups.Add($top) }
    return $result
}

function Get-DriveLink {
    <#
    .SYNOPSIS
      A file:/// link to a file, so the workbook opens it; none when the path holds PII (the link would show it).
    .PARAMETER Path
      The file path.
    .OUTPUTS
      [string] the link, or empty.
    #>
    param([string]$Path)
    if (@(Find-PiiInText $Path).Count -gt 0) { return '' }
    $absolute = $Path; if ($Path.StartsWith('/')) { $absolute = 'file://' + $Path }   # / paths: the test machines only
    try { return ([uri]$absolute).AbsoluteUri } catch { return '' }
}

function Invoke-DriveFile {
    <#
    .SYNOPSIS
      Starts one file of a drive scan: a hit for PII in its name, then a NOT SCANNED row (unsupported or too
      large) or the file handed to the readers, which open it themselves.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Group
      The file's group name (a top-level subfolder), not masked.
    .PARAMETER File
      One file of Get-DriveGroups.
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file went to the readers.
    #>
    param([hashtable]$Queue, [string]$Group, [hashtable]$File, [int]$MaxFileMB)
    $rows = $Queue.Rows
    $channel = Protect-PiiName $Group
    $item = Get-ShownItem $Group $File.Inner
    $link = Get-DriveLink $File.Path
    foreach ($hit in @(Find-PiiInText ($Group + '/' + $File.Inner))) {
        $rows.Add((New-ReportRow 'File name' $channel $item $link $File.Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $skipReason = ''
    if (-not (Test-SupportedFile $File.Name)) { $skipReason = 'File type is not read by this script: check by hand' }
    elseif ($File.Size -gt ($MaxFileMB * 1MB)) { $skipReason = "Larger than $MaxFileMB MB: check by hand" }
    if ($skipReason) { $rows.Add((New-ReportRow 'File' $channel $item $link $File.Modified 'NOT SCANNED' '' '' $skipReason)); return $false }
    Submit-ScanFile $Queue @{ Channel = $channel; Item = $item; Link = $link; Modified = $File.Modified; DrivePath = $File.Path } $File.Name $null $File.Path
    return $true
}

function Add-DriveListRows {
    <#
    .SYNOPSIS
      A BLOCKED row for each folder that could not be opened, and a NOT SCANNED row for each folder link not
      followed (a link can lead outside the folder).
    .PARAMETER Found
      Output of Get-DriveGroups.
    .PARAMETER Rows
      The list that receives the rows.
    #>
    param([hashtable]$Found, $Rows)
    foreach ($blocked in $Found.Blocked) {
        $reason = 'Could not open this folder: ' + (Protect-PiiText $blocked.Reason)
        $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $blocked.Folder) '(folder)' '' '' 'BLOCKED' '' '' $reason))
    }
    foreach ($link in $Found.Links) {
        $reason = 'A link to another folder is not followed (it can lead outside the folder): check by hand'
        $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $link) '(folder link)' '' '' 'NOT SCANNED' '' '' $reason))
    }
}

function Get-DriveHeader {
    <#
    .SYNOPSIS
      The Summary tab's header lines of a drive scan (every name masked like any name the tool shows).
    .PARAMETER Path
      The folder.
    .PARAMETER Groups
      The groups scanned.
    .PARAMETER Used
      The whole run's time.
    .PARAMETER Reading
      The time all workers spent reading and checking.
    .PARAMETER Workers
      The number of files read at once at the end.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path, [object[]]$Groups, [TimeSpan]$Used, [TimeSpan]$Reading, [int]$Workers)
    # A time as minutes:seconds.
    $format = { param($t) '{0:00}:{1:00}' -f [int][Math]::Floor($t.TotalMinutes), $t.Seconds }
    $header = [ordered]@{}
    $header['Folder'] = Protect-PiiName $Path
    $header['Folders'] = ($Groups | ForEach-Object { Protect-PiiName $_.Name }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = & $format $Used
    $header['Time reading and checking (all workers)'] = & $format $Reading
    $header['Files read at once (at the end)'] = [string]$Workers
    return $header
}

function Invoke-DriveScan {
    <#
    .SYNOPSIS
      Scans one drive folder and all its subfolders with the reader workers, then saves the workbook (also after an
      error or a cancel: the results so far are saved and say INCOMPLETE). Returns the workbook path.
    .PARAMETER Path
      The folder (X:\<folder> or \\<server>\<share>\<folder>).
    .PARAMETER OutputFolder
      The base folder for the results.
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([string]$Path, [string]$OutputFolder, [int]$MaxFileMB, [hashtable]$Init)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[object]
    $finished = $false
    $groups = @()
    $workers = 0
    $reading = [TimeSpan]::Zero
    $meter = New-ScanMeter 0
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw 'This folder cannot be opened (check the path, and that you can open it in File Explorer).' }
        Write-Ui 'Listing the files ...' 'Muted'
        $found = Get-DriveGroups $Path
        $groups = $found.Groups.ToArray()
        Add-DriveListRows $found $rows
        $total = ($groups | ForEach-Object { $_.Files.Count } | Measure-Object -Sum).Sum
        Write-Ui ('==== Scanning {0} files in {1} folders ====' -f $total, $groups.Count) 'Title'
        $meter = New-ScanMeter ([int]$total)
        $meter.NameWidth = Get-NameWidth @($groups | ForEach-Object { Protect-PiiName $_.Name })
        $queue = New-ScanQueue $Init $rows $meter
        try {
            # Starts one file (Invoke-GroupScan calls it).
            $start = { param($group, $file) Invoke-DriveFile $queue $group.Name $file $MaxFileMB }
            # The file's masked path inside its folder, for the progress bar.
            $shown = { param($group, $file) Get-ShownItem $group.Name $file.Inner }
            Invoke-GroupScan $queue $groups $start $shown
            $workers = Get-QueueWorkers $queue
        }
        finally { $reading = Close-ScanQueue $queue }
        Write-RuleLine
        Write-MeterLine $meter 'file' -Force
        $finished = $true
    }
    catch {
        $reason = Protect-PiiText $_.Exception.Message
        Write-Ui ('The scan stopped early: ' + $reason) 'Alert'
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
    }
    finally {
        Add-IncompleteRow -Rows $rows -Finished $finished; $header = Get-DriveHeader $Path $groups $clock.Elapsed $reading $workers
        $name = $Path.TrimEnd('\') -replace '\\', '/'   # the results folder is named after the folder's last part
        $path = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $meter.Read -Header $header -OutputFolder $OutputFolder -SiteUrl $name -FolderLabel 'Folder'
    }
    return $path
}

Export-ModuleMember -Function Get-DriveGroups, Invoke-DriveScan
