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
# A file link in the folder is never read (the PAM, 2026-11-21: "No reconnaissance").
$script:FileLinkReason = 'A link to another file is not followed (it can lead outside the folder): check by hand'

function Get-DriveGroups {
    <#
    .SYNOPSIS
      The files under one folder, grouped by its top-level subfolders (then the files at its top), with the folders
      that could not be opened and the folder links that were not followed (a link can lead outside the folder).
    .PARAMETER Root
      The folder path.
    .OUTPUTS
      [hashtable] Groups (each @{ Name; Files }), Blocked (each @{ Folder; Reason }), Links (folder link paths),
      FileLinks (file link paths: never read, they can lead outside the folder; the PAM, 2026-11-21).
    #>
    param([string]$Root)
    $Root = $Root.TrimEnd('\', '/')
    $slashes = [char[]]@('\', '/')   # Windows uses \; the tests also run where / is the separator
    $result = @{ Groups = (New-Object System.Collections.Generic.List[object]); Blocked = (New-Object System.Collections.Generic.List[object])
        Links = (New-Object System.Collections.Generic.List[string]); FileLinks = (New-Object System.Collections.Generic.List[string]) }
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
            # A file link (any reparse point) is not followed: it can lead outside the folder ("No reconnaissance").
            if ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { $result.FileLinks.Add($rel); continue }
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
    if ((Protect-PiiName $Path) -ne $Path -or @(Find-PiiInText $Path).Count -gt 0) { return '' }   # as strict as Get-SafeLink
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
    $info = @{ Channel = $channel; Item = $item; Link = $link; Modified = $File.Modified }
    if (-not (Add-FileStartRows $rows $info ($Group + '/' + $File.Inner) $File.Name $File.Size $MaxFileMB)) { return $false }
    $info.DrivePath = $File.Path
    Submit-ScanFile $Queue $info $File.Name $null $File.Path
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
    foreach ($link in $Found.FileLinks) {
        $shown = Protect-PiiName $link
        foreach ($hit in @(Find-PiiInText $link)) {   # PII in the link's own name, as for any file
            $Rows.Add((New-ReportRow 'File name' $shown '(file link)' '' '' $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
        }
        $Rows.Add((New-ReportRow 'Scan' $shown '(file link)' '' '' 'NOT SCANNED' '' '' $script:FileLinkReason))
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
      The time all workers spent reading.
    .PARAMETER Workers
      The number of files read at once at the end.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path, [object[]]$Groups, [TimeSpan]$Used, [TimeSpan]$Reading, [int]$Workers)
    $header = [ordered]@{}
    $header['Folder'] = Protect-PiiName $Path
    $header['Folders'] = ($Groups | ForEach-Object { Protect-PiiName $_.Name }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = Format-Duration $Used
    $header['Time reading files (all readers)'] = Format-Duration $Reading
    $header['Files read at the same time (at the end)'] = [string]$Workers
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
    $workbook = $null
    $run = @{ Meter = (New-ScanMeter 0); Reading = [TimeSpan]::Zero; Workers = 0 }   # filled by Invoke-ScanRun
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw 'This folder cannot be opened (check the path, and that you can open it in File Explorer).' }
        Write-Ui 'Listing the files ...' 'Muted'
        $found = Get-DriveGroups $Path
        $groups = $found.Groups.ToArray()
        Add-DriveListRows $found $rows
        # Starts one file (Invoke-GroupScan calls it).
        $start = { param($queue, $group, $file) Invoke-DriveFile $queue $group.Name $file $MaxFileMB }
        # The file's masked path inside its folder, for the progress bar.
        $shown = { param($group, $file) Get-ShownItem $group.Name $file.Inner }
        # A file's name and size, for the question before the scan.
        $facts = { param($file) $file.Name, $file.Size }
        Invoke-ScanRun $Init $rows $groups $start $shown $run $facts $MaxFileMB
        $finished = $true
    }
    catch {
        $reason = Protect-PiiText $_.Exception.Message
        Write-Ui ('The scan stopped early: ' + $reason) 'Alert'
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
    }
    finally {
        # "No" to "Scan now?": nothing was read, so nothing is saved (the PAM, 2026-11-21).
        if ($run.ContainsKey('Declined') -and $run.Declined) { Write-Ui 'Not scanned: nothing was saved.' 'Accent' }
        else {
            Add-IncompleteRow -Rows $rows -Finished $finished
            $header = Get-DriveHeader $Path $groups $clock.Elapsed $run.Reading $run.Workers
            $name = $Path.TrimEnd('\') -replace '\\', '/'   # the results folder is named after the folder's last part
            $workbook = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $run.Meter.Read -Header $header -OutputFolder $OutputFolder -SiteUrl $name -FolderLabel 'Folder'
        }
    }
    return $workbook
}

Export-ModuleMember -Function Get-DriveGroups, Invoke-DriveScan
