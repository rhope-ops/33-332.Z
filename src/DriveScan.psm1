<#
DriveScan.psm1 - scans a folder on a mapped or network drive (X:\<folder> or \\<server>\<share>\<folder>).

  Purpose:     The PAM, 2026-11-21: the same check as for a SharePoint site, for one folder and all its subfolders on
               a drive the person can open. Nothing above the folder, no other share or server is touched ("No
               reconnaissance"). Files are opened read-only and shared, with the person's own Windows logon.
  Inputs:      the pasted folder path; the reader workers' set-up (New-ReaderPool).
  Outputs:     report rows; the results workbook (Report.psm1 Publish-ScanReport).
  How to run:  imported by Find-ChannelPII.ps1; Invoke-DriveScan runs when the pasted target is a folder path.
  Depends on:  PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1 (Test-SupportedFile), Report.psm1,
               ScanMeter.psm1, ScanQueue.psm1, ChannelScan.psm1 (Get-ShownItem, Get-FileBytes, Get-FieldValue,
               Get-SafeLink), Accountability.psm1 and SiteTarget.psm1 (Connect-TeamSite), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7, where the tests run on a temporary folder).
  Network:     only the given folder's server, through the person's Windows logon (SMB). In a folder of the own
               work OneDrive, or of a synced team library or a shortcut in it (Connect-OneDriveWeb): the site of that folder (from
               the OneDrive app), one library and only the pasted folder in it, for the who/when history and to read online-only and web-only files
               into memory; when the web is not used, OneDrive downloads an online-only file only after the person
               answers y (Select-OnlineFiles).
  Writes:      only the results workbook under the results folder (a downloaded OneDrive file is written by OneDrive).
  Credentials: none stored (the Windows logon; in a OneDrive folder, the person's own CAC sign-in).
#>

Set-StrictMode -Version 2.0

# Files directly in the pasted folder form this group (each subfolder is its own group).
$script:DriveTopGroup = '(top of the folder)'
# A file link in the folder is never read (the PAM, 2026-11-21: "No reconnaissance").
$script:FileLinkReason = 'A link to another file is not followed (it can lead outside the folder): check by hand'
# OneDrive (and a SharePoint library synced by OneDrive) keeps its files and folders as reparse points too: they are the
# item itself, not a link, so they are read (the PAM, 2026-11-23: scan a synced OneDrive folder when SharePoint sign-in
# is blocked). A file whose data is only in the cloud has one of these attributes: RecallOnDataAccess (0x400000),
# RecallOnOpen (0x40000) or Offline (0x1000). Ion asks before it reads such files (the PAM, 2026-11-23: "Ask each run").
$script:OnlineOnlyAttributes = 0x441000
$script:OnlineReason = 'Online only (OneDrive), not downloaded: check by hand, or run Ion again and answer y to download'

function Get-OneDriveHomes {
    # The person's own OneDrive folders ("My files"), from Windows' OneDrive environment variables: only these can hold
    # shortcuts to other teams' libraries (Confirm-OneDriveRoot). A synced library folder is not one (review of PR 59).
    return @(foreach ($name in @('OneDrive', 'OneDriveCommercial', 'OneDriveConsumer')) { [Environment]::GetEnvironmentVariable($name) }) |
        Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\', '/') }
}

function Get-OneDriveRoots {
    <#
    .SYNOPSIS
      The folders this person's OneDrive app syncs: its OneDrive folders (environment) and each synced SharePoint
      library folder (the OneDrive app's own list in the registry, read-only). Only these can hold OneDrive items.
    .OUTPUTS
      [string[]] full folder paths.
    #>
    $roots = @(Get-OneDriveHomes)
    try {
        foreach ($tenant in @(Get-ChildItem -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\*\Tenants' -ErrorAction Stop)) {
            $roots += @($tenant.GetValueNames())   # each value name is a synced folder
        }
    }
    catch { Write-Verbose 'No OneDrive list in the registry (not Windows, or no OneDrive account).' }
    return @($roots | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') })
}

function Confirm-OneDriveRoot {
    <#
    .SYNOPSIS
      When the pasted folder is the person's own OneDrive folder (Get-OneDriveHomes), its top folders can be shortcuts to
      other teams' libraries ("Add shortcut to My files"), which look like normal folders. Ion says so, lists the top
      folders (names masked) and asks before it scans them all (the PAM, 2026-11-23: "Warn, then ask"; the rule
      "No reconnaissance"). A folder that holds the OneDrive folder (the user folder) is asked about too. Any other
      folder: true without a question.
    .PARAMETER Path
      The pasted folder.
    .OUTPUTS
      [bool] true to scan.
    #>
    param([string]$Path)
    $folder = [System.IO.Path]::GetFullPath($Path).TrimEnd('\', '/')   # "<folder>\." and "<sub>\.." count as the folder
    $held = @()   # the OneDrive folders that are the pasted folder or are inside it (review of PR 59)
    foreach ($oneDriveHome in @(Get-OneDriveHomes)) {
        try { $full = [System.IO.Path]::GetFullPath($oneDriveHome).TrimEnd('\', '/') } catch { continue }   # a bad variable: not a OneDrive folder
        if (-not [System.IO.Directory]::Exists($full) -or $held -icontains $full) { continue }   # a folder that is gone, or listed twice
        if ($full -ieq $folder -or $full.StartsWith($folder + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) { $held += $full }
    }
    if ($held.Count -eq 0) { return $true }
    # The top folders of the OneDrive folder itself: the ones that can be shortcuts (also when a parent folder was pasted).
    $tops = @(foreach ($oneDrive in $held) { (New-Object System.IO.DirectoryInfo($oneDrive + [System.IO.Path]::DirectorySeparatorChar)).GetDirectories() |
                ForEach-Object { Protect-PiiName $_.Name } })
    if ($held -icontains $folder) { Write-Ui 'This is your whole OneDrive folder. Folders in it can be shortcuts to the libraries of other teams.' 'Accent' }
    else { Write-Ui 'This folder holds your whole OneDrive folder. Folders in OneDrive can be shortcuts to the libraries of other teams.' 'Accent' }
    Write-Ui ('  The top folders in OneDrive: ' + ($tops -join ', ')) 'Muted'
    Write-Ui '  To scan one team only, press Enter, then paste the folder of that team.' 'Muted'
    return ((Read-Host 'Scan all of them? [y/N]') -match '^\s*y')
}

function Get-ReparseKind {
    <#
    .SYNOPSIS
      What a reparse point in the folder is: 'Cloud' only for an item inside a OneDrive sync folder that PowerShell
      gives no LinkType and no Target (the file or folder itself: read); else 'Link' (a symbolic link, a junction,
      or any other kind such as a deduplicated or Azure File Sync file: not read, as before; review of PR 59).
    .PARAMETER Entry
      A FileInfo or DirectoryInfo (or, in tests, an object with FullName, LinkType and Target).
    .PARAMETER Roots
      The OneDrive sync folders (Get-OneDriveRoots).
    .OUTPUTS
      [string] 'Link' or 'Cloud'.
    #>
    param($Entry, [string[]]$Roots)
    $inside = $false
    try { foreach ($root in $Roots) { if ($Entry.FullName.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { $inside = $true } } }
    catch { return 'Link' }
    if (-not $inside) { return 'Link' }
    try {
        $item = [psobject]$Entry
        $type = [string]$item.LinkType
        $target = (@($item.Target) | Where-Object { $_ }) -join ''
    }
    catch { return 'Link' }
    if ($type -or $target) { return 'Link' }
    return 'Cloud'
}

function Select-OnlineFiles {
    <#
    .SYNOPSIS
      Online-only OneDrive files: when some would be read (a supported type, under the size limit), Ion shows how many
      and their MB, and asks whether to download and read them (default no). On "no" they leave the groups and get a
      NOT SCANNED row each (with a hit row for PII in the name). Ion never frees them again (that is a write).
    .PARAMETER Groups
      The groups of Get-DriveGroups (changed in place).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER MaxFileMB
      The size limit.
    #>
    param([object[]]$Groups, $Rows, [int]$MaxFileMB)
    # A file that OneDrive on the web gives (Add-OneDriveWebFiles) is read from there, into memory: not asked about.
    $online = @(foreach ($group in $Groups) { foreach ($file in $group.Files) {
                if ($file.Online -and -not $file.ContainsKey('Web') -and (Test-SupportedFile $file.Name) -and $file.Size -le $MaxFileMB * 1MB) { $file } } })
    if ($online.Count -eq 0) { return }
    $megabytes = ($online | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum / 1MB   # 5.1 reads no hashtable key here
    Write-Ui ('{0} files ({1:N1} MB) here are online only (OneDrive). To read them, Ion must download them to this computer.' -f $online.Count, $megabytes) 'Accent'
    Write-Ui '  They then stay on this computer. To remove them again: right-click the folder, then click Free up space.' 'Muted'
    if ((Read-Host 'Download and read them? [y/N]') -match '^\s*y') { return }
    foreach ($group in $Groups) {
        $keep = New-Object System.Collections.Generic.List[object]   # one pass per group, not one search per file
        $declined = New-Object System.Collections.Generic.List[object]
        foreach ($file in $group.Files) {
            if ($file.Online -and -not $file.ContainsKey('Web') -and (Test-SupportedFile $file.Name) -and $file.Size -le $MaxFileMB * 1MB) { $declined.Add($file) }
            else { $keep.Add($file) }
        }
        $group.Files.Clear(); $group.Files.AddRange($keep)
        foreach ($file in $declined) {
            $info = @{ Channel = (Protect-PiiName $group.Name); Item = (Get-ShownItem $group.Name $file.Inner); Link = (Get-DriveLink $file.Path); Modified = $file.Modified }
            if (Add-FileStartRows $Rows $info ($group.Name + '/' + $file.Inner) $file.Name $file.Size $MaxFileMB) {
                $Rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' $script:OnlineReason))
            }
        }
    }
}

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
    $roots = Get-OneDriveRoots
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
                if (($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and (Get-ReparseKind $entry $roots) -eq 'Link') { $result.Links.Add($rel); continue }
                $stack.Push($entry.FullName)
                continue
            }
            # A file link (a symbolic link) is not followed: it can lead outside the folder ("No reconnaissance").
            if (($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and (Get-ReparseKind $entry $roots) -eq 'Link') { $result.FileLinks.Add($rel); continue }
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
                Modified = $entry.LastWriteTime.ToString('yyyy-MM-dd HH:mm'); Online = (([int]$entry.Attributes -band $script:OnlineOnlyAttributes) -ne 0) })
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
    $web = $null; if ($File.ContainsKey('Web')) { $web = $File.Web }
    if ($web) { $link = Get-SafeLink $File.Origin (Get-FieldValue $web 'FileRef') (Get-FieldValue $web 'UniqueId') } else { $link = Get-DriveLink $File.Path }
    $info = @{ Channel = $channel; Item = $item; Link = $link; Modified = $File.Modified }
    if (-not (Add-FileStartRows $rows $info ($Group + '/' + $File.Inner) $File.Name $File.Size $MaxFileMB)) { return $false }
    if (-not $web) {   # no web copy: the disk only, and the Windows owner and times (as before)
        $info.DrivePath = $File.Path
        if ($File.ContainsKey('NoWeb')) { $info.DriveNote = $script:NotSyncedNote }
        Submit-ScanFile $Queue $info $File.Name $null $File.Path
        return $true
    }
    # The who/when history from the web copy, as for a team site (Find-PiiAddedVersion, Accountability.psm1).
    $ref = Get-FieldValue $web 'FileRef'; $name = $File.Name
    $who = Get-SharePointWho $web
    $who.Modified = Get-FieldValue $web 'Modified'
    $info.Who = $who
    $info.Added = { param($hits) Find-PiiAddedVersion $ref $name $hits $who }.GetNewClosure()
    $info.Access = { Get-SharePointAccess $web }.GetNewClosure()
    if (-not $File.Online -and $File.Path) { Submit-ScanFile $Queue $info $name $null $File.Path; return $true }   # on the disk: read from the disk
    # Online only, or on the web only: read from the web into memory (the PAM, 2026-11-24: nothing goes to the disk).
    try { $bytes = Get-FileBytes $ref $File.Size }
    catch {
        $rows.Add((New-ReportRow 'File' $channel $item $link $File.Modified 'NOT SCANNED' '' '' (Protect-PiiText ('Download failed: ' + $_.Exception.Message))))
        return $false
    }
    finally { Write-Progress -Id 2 -Activity 'Downloading' -Completed }
    Submit-ScanFile $Queue $info $name $bytes ''
    return $true
}

function Add-OneDriveWebFiles {
    <#
    .SYNOPSIS
      Joins the web files of the pasted folder (Connect-OneDriveWeb) to its local files: a local file with a web copy
      gets it (Web, Origin: history, and the read of an online-only file); a web file that is not on the disk joins its
      group (the PAM, 2026-11-24: read through PnP). The web side never adds a file outside the pasted folder.
    .PARAMETER Groups
      The groups of Get-DriveGroups (changed in place).
    .PARAMETER Web
      Output of Connect-OneDriveWeb with Used = $true.
    .OUTPUTS
      [int] the number of web-only files added.
    #>
    param([System.Collections.Generic.List[object]]$Groups, [hashtable]$Web)
    $seen = @{}
    foreach ($group in $Groups) {
        foreach ($file in $group.Files) {
            $rel = $file.Inner; if ($group.Name -ne $script:DriveTopGroup) { $rel = $group.Name + '/' + $file.Inner }
            $key = $rel.ToLowerInvariant(); $seen[$key] = $true
            if ($Web.Files.ContainsKey($key)) { $file.Web = $Web.Files[$key]; $file.Origin = $Web.Origin }
            else { $file.NoWeb = $true }   # on this computer only (not synced yet): no history (design, step 5)
        }
    }
    $added = 0
    foreach ($key in @($Web.Files.Keys | Sort-Object)) {
        if ($seen.ContainsKey($key)) { continue }
        $item = $Web.Files[$key]
        $rel = (Get-FieldValue $item 'FileRef').Substring($Web.Map.WebFolder.Length + 1)
        $groupName = $script:DriveTopGroup; $inner = $rel
        $cut = $rel.IndexOf('/'); if ($cut -gt 0) { $groupName = $rel.Substring(0, $cut); $inner = $rel.Substring($cut + 1) }
        $group = $null; foreach ($g in $Groups) { if ($g.Name -ieq $groupName) { $group = $g; break } }
        if (-not $group) { $group = @{ Name = $groupName; Files = (New-Object System.Collections.Generic.List[object]) }; $Groups.Add($group) }
        $size = 0L; [void][long]::TryParse((Get-FieldValue $item 'File_x0020_Size'), [ref]$size)
        $group.Files.Add(@{ Path = ''; Inner = $inner; Name = (Get-FieldValue $item 'FileLeafRef'); Size = $size; Modified = (Get-FieldValue $item 'Modified')
            Online = $true; Web = $item; Origin = $Web.Origin })
        $added++
    }
    return $added
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
    .PARAMETER WebNote
      Connect-OneDriveWeb's note (empty for a folder that OneDrive does not sync).
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path, [object[]]$Groups, [TimeSpan]$Used, [TimeSpan]$Reading, [int]$Workers, [string]$WebNote = '')
    $header = [ordered]@{}
    $header['Folder'] = Protect-PiiName $Path
    if ($WebNote) { $header['OneDrive on the web'] = $WebNote }
    $header['Folders'] = ($Groups | ForEach-Object { Protect-PiiName $_.Name }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = Format-Duration $Used
    $header['Time reading files (all readers)'] = Format-Duration $Reading
    $header['Files read at the same time (at the end)'] = [string]$Workers
    return $header
}

function Get-DriveScanList {
    <#
    .SYNOPSIS
      The files to scan: the folder's files, joined with their web copies in a OneDrive folder (plan step 4: one
      sign-in; history, online-only and web-only files from the web; local only if that fails), the rows for blocked
      folders and links, and the download question for online-only files with no web copy.
    .PARAMETER Path
      The folder.
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER MaxFileMB
      The size limit.
    .PARAMETER Run
      Receives WebNote (Connect-OneDriveWeb's note for the Summary).
    .PARAMETER SyncEntries
      The OneDrive sync entries, or $null (the registry).
    .PARAMETER OneDriveHomes
      The own OneDrive folders, or $null (Windows).
    .OUTPUTS
      [object[]] the groups.
    #>
    param([string]$Path, $Rows, [int]$MaxFileMB, [hashtable]$Run, [object[]]$SyncEntries, [string[]]$OneDriveHomes)
    Write-Ui 'Listing the files ...' 'Muted'
    $found = Get-DriveGroups $Path
    $web = Connect-OneDriveWeb $Path $SyncEntries $OneDriveHomes
    $Run.WebNote = $web.Note
    if ($web.Used) { $null = Add-OneDriveWebFiles $found.Groups $web }
    $groups = $found.Groups.ToArray()
    Add-DriveListRows $found $Rows
    Select-OnlineFiles $groups $Rows $MaxFileMB
    return , $groups
}

function Invoke-DriveScan {
    <#
    .SYNOPSIS
      Scans one drive folder and all its subfolders with the reader workers, then saves the workbook (also after an
      error or a cancel: the results so far are saved and say INCOMPLETE). Returns the workbook path; $null when the
      person answered no (or cancelled) at "Scan now?".
    .PARAMETER Path
      The folder (X:\<folder> or \\<server>\<share>\<folder>).
    .PARAMETER OutputFolder
      The base folder for the results.
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .PARAMETER AutoStart
      Do not ask "Scan now?" (the person said yes to a list of targets); the OneDrive question still asks.
    .PARAMETER SyncEntries
      The OneDrive sync entries; $null (the default) reads them from the registry. Tests give their own.
    .PARAMETER OneDriveHomes
      The person's own OneDrive folders; $null (the default) reads them from Windows. Tests give their own.
    .OUTPUTS
      [string] the workbook path.
    #>
    param([string]$Path, [string]$OutputFolder, [int]$MaxFileMB, [hashtable]$Init, [switch]$AutoStart, [object[]]$SyncEntries = $null, [string[]]$OneDriveHomes = $null)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[object]
    $finished = $false
    $groups = @(); $workbook = $null
    $run = @{ Meter = (New-ScanMeter 0); Reading = [TimeSpan]::Zero; Workers = 0; AutoStart = [bool]$AutoStart; WebNote = '' }   # filled by Invoke-ScanRun
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw 'This folder cannot be opened (check the path, and that you can open it in File Explorer).' }
        # Own OneDrive folder: shortcuts to other teams' libraries ("No reconnaissance"); declined until a yes (a cancel too).
        $run.Declined = $true; if (-not (Confirm-OneDriveRoot $Path)) { return $null }; $run.Declined = $false
        $groups = Get-DriveScanList $Path $rows $MaxFileMB $run $SyncEntries $OneDriveHomes
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
            $header = Get-DriveHeader $Path $groups $clock.Elapsed $run.Reading $run.Workers $run.WebNote
            $name = $Path.TrimEnd('\') -replace '\\', '/'   # the results folder is named after the folder's last part
            $workbook = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $run.Meter.Read -Header $header -OutputFolder $OutputFolder -SiteUrl $name -FolderLabel 'Folder'
        }
    }
    return $workbook
}

# ---- OneDrive on the web (plan step 4, the OneDrive hybrid; the PAM, 2026-11-24: docs/reviews/2026-11-24-onedrive-hybrid-design.md) ----

# The OneDrive app's list of synced folders (one key per folder: MountPoint, UrlNamespace, FullRemotePath).
$script:SyncEnginesKey = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
# All 3 folder kinds use the web (the PAM, 2026-11-24, with a probe first): probe 4 tested 'Home' (runs 1 and 2); probe 7
# (2026-11-25, runs A and B) tested 'Library' and 'Shortcut' with the web folder from FullRemotePath (7 of 7 and 8 of 8
# names on both sides): docs/reviews/2026-11-25-device-probe-7.md.
# The list fields read for each web file (as ChannelScan.psm1 reads them for a team site).
$script:WebFields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified', 'UniqueId', 'Created', 'Author', 'Editor')
# True after a sign-in to a OneDrive site was tried: Find-ChannelPII.ps1 then signs out at the end of the run.
$script:SignInTried = $false
$script:NoHistoryNote = 'The who/when history is not available for this scan'
$script:NotSyncedNote = 'Not on the web yet (not synced): no version history'

function Get-OneDriveSyncEntries {
    <#
    .SYNOPSIS
      The folders the OneDrive app syncs, from its registry list (read-only): each @{ MountPoint; UrlNamespace;
      FullRemotePath } (FullRemotePath: '' when the key has none).
      None when the list is not there (not Windows, or no OneDrive account).
    .OUTPUTS
      [hashtable[]]
    #>
    $entries = @()
    foreach ($key in @(Get-ChildItem -Path $script:SyncEnginesKey -ErrorAction SilentlyContinue)) {
        $mount = [string]$key.GetValue('MountPoint'); $space = [string]$key.GetValue('UrlNamespace')
        $remote = $key.GetValue('FullRemotePath'); if ($remote -isnot [string]) { $remote = '' }
        if ($mount -and $space) { $entries += @{ MountPoint = $mount; UrlNamespace = $space; FullRemotePath = $remote } }
    }
    return $entries
}

function Get-RemoteSubFolder {
    <#
    .SYNOPSIS
      The synced folder's place below the library root, from FullRemotePath (probe 7): '' for the root itself, for
      example '/Plans' for a folder below it; $null when the value is missing, not https, on another host, outside the
      library, or has a '.' or '..' part (then no web folder).
    .PARAMETER Remote
      The FullRemotePath value ('' when missing).
    .PARAMETER LibraryHost
      The host of the library address (UrlNamespace).
    .PARAMETER Library
      The library root, server-relative, decoded.
    .OUTPUTS
      [string] or $null.
    #>
    param([string]$Remote, [string]$LibraryHost, [string]$Library)
    $uri = $null
    if (-not [uri]::TryCreate($Remote.Trim().TrimEnd('/'), [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.Host -ine $LibraryHost) { return $null }
    $remotePath = [uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
    if ($remotePath -ieq $Library) { return '' }
    if (-not $remotePath.StartsWith($Library + '/', [StringComparison]::OrdinalIgnoreCase)) { return $null }
    $sub = $remotePath.Substring($Library.Length)
    if ($sub -match '(^|/)\.\.?(/|$)') { return $null }
    return $sub
}

function Get-OneDriveWebMap {
    <#
    .SYNOPSIS
      The web place of a local folder: the sync entry whose MountPoint holds it (the longest one: probe 4 found 2
      entries with the same folder and address), its site, its web folder, and its kind. $null when no entry holds it.
      UrlNamespace is the library root. For a shortcut or a synced library, the synced folder can be below that root:
      its place is the part of FullRemotePath below the root (probe 7). Without such a part, WebFolder is '' (local only).
    .PARAMETER Folder
      The pasted folder (full path).
    .PARAMETER Entries
      The sync entries (Get-OneDriveSyncEntries).
    .PARAMETER Homes
      The person's own OneDrive folders (Get-OneDriveHomes).
    .OUTPUTS
      [hashtable] Mount, LibraryPath (the library root, server-relative), Site, Origin, WebFolder, Kind ('Home',
      'Shortcut' or 'Library'); or $null.
    #>
    param([string]$Folder, [object[]]$Entries, [string[]]$Homes)
    $slashes = [char[]]@('\', '/')
    $path = $Folder.TrimEnd($slashes)
    $best = $null
    foreach ($entry in $Entries) {
        $mount = ([string]$entry.MountPoint).TrimEnd($slashes)
        if (-not $mount -or -not $entry.UrlNamespace) { continue }
        $inside = $path -ieq $mount -or ($path.Length -gt $mount.Length -and
            $path.StartsWith($mount, [StringComparison]::OrdinalIgnoreCase) -and $slashes -contains $path[$mount.Length])
        if ($inside -and (-not $best -or $mount.Length -gt $best.Mount.Length)) {
            $remote = ''   # an entry without FullRemotePath (an older OneDrive app, or a test) has no web place below the root
            if ($entry -is [System.Collections.IDictionary]) { $remote = [string]$entry['FullRemotePath'] }
            elseif ($entry.PSObject.Properties['FullRemotePath']) { $remote = [string]$entry.FullRemotePath }
            $best = @{ Mount = $mount; Space = [string]$entry.UrlNamespace; Remote = $remote }
        }
    }
    if (-not $best) { return $null }
    $address = $null
    if (-not [uri]::TryCreate($best.Space.TrimEnd('/'), [UriKind]::Absolute, [ref]$address)) { return $null }   # not a web address: no map (local only)
    $library = [uri]::UnescapeDataString($address.AbsolutePath).TrimEnd('/')
    $sitePath = $library; if ($library -match '^/(personal|sites|teams)/[^/]+') { $sitePath = $Matches[0] }
    $origin = $address.GetLeftPart([UriPartial]::Authority)
    $kind = 'Library'
    foreach ($oneDriveHome in @($Homes | Where-Object { $_ } | ForEach-Object { $_.TrimEnd($slashes) })) {
        if ($best.Mount -ieq $oneDriveHome) { $kind = 'Home'; break }
        if ($best.Mount.StartsWith($oneDriveHome, [StringComparison]::OrdinalIgnoreCase) -and $slashes -contains $best.Mount[$oneDriveHome.Length]) { $kind = 'Shortcut' }
    }
    $rest = $path.Substring($best.Mount.Length) -replace '[\\/]+', '/'
    # A '.' or '..' part in the pasted path can point above the synced folder ("No reconnaissance", review of the map fix):
    # no web folder then, for every kind (a local scan only).
    if ($rest -match '(^|/)\.\.?(/|$)') { return @{ Mount = $best.Mount; LibraryPath = $library; Site = $origin + $sitePath; Origin = $origin; WebFolder = ''; Kind = $kind } }
    $sub = ''
    if ($kind -ne 'Home') {
        $sub = Get-RemoteSubFolder $best.Remote $address.Host $library
        if ($null -eq $sub) { return @{ Mount = $best.Mount; LibraryPath = $library; Site = $origin + $sitePath; Origin = $origin; WebFolder = ''; Kind = $kind } }
    }
    return @{ Mount = $best.Mount; LibraryPath = $library; Site = $origin + $sitePath; Origin = $origin; WebFolder = ($library + $sub + $rest).TrimEnd('/'); Kind = $kind }
}

function Get-OneDriveWebFiles {
    <#
    .SYNOPSIS
      The files under the mapped web folder (any depth), from the one library whose root is the mapped library: the
      relative path in small letters ('/' between folders) -> the list item. The library is found by its own address
      (no list of the site's lists), and SharePoint gives only the items in the mapped folder ("No reconnaissance").
    .PARAMETER Map
      Output of Get-OneDriveWebMap.
    .OUTPUTS
      [hashtable]; throws when no library has that root.
    #>
    param([hashtable]$Map)
    $siteRelative = $Map.LibraryPath.Substring(([uri]$Map.Site).AbsolutePath.TrimEnd('/').Length).TrimStart('/')
    $library = Get-PnPList -Identity $siteRelative -ErrorAction Stop
    $root = ''; if ($library) { $root = (Get-PnPProperty -ClientObject $library -Property RootFolder -ErrorAction Stop).ServerRelativeUrl.TrimEnd('/') }
    if (-not $library -or $root -ine $Map.LibraryPath) { throw 'The OneDrive folder was not found on the web (no library with its address).' }
    $prefix = $Map.WebFolder + '/'
    $files = @{}
    $items = @(Get-PnPListItem -List $library -FolderServerRelativeUrl $Map.WebFolder -PageSize 500 -Fields $script:WebFields -ErrorAction Stop)
    foreach ($item in $items) {   # the prefix check stays: only files under the pasted folder are kept, whatever the query gives
        if ([string]$item['FSObjType'] -ne '0') { continue }
        $ref = [string]$item['FileRef']
        if ($ref.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $files[$ref.Substring($prefix.Length).ToLowerInvariant()] = $item }
    }
    return $files
}

function Connect-OneDriveWeb {
    <#
    .SYNOPSIS
      For a folder that OneDrive syncs: the map, one sign-in to its site (Connect-TeamSite: the full-address site
      check), and its web files. Never throws: a folder that is not synced, a folder with no web place, or any failure
      gives Used = $false and a Note for the Summary (the scan is then local only, as before).
    .PARAMETER Folder
      The pasted folder.
    .PARAMETER Entries
      The sync entries; $null reads them from the registry (tests give their own).
    .PARAMETER Homes
      The person's own OneDrive folders; $null reads them from Windows (tests give their own).
    .OUTPUTS
      [hashtable] Used, Note, Files (Get-OneDriveWebFiles), Map, Origin.
    #>
    param([string]$Folder, [object[]]$Entries = $null, [string[]]$Homes = $null)
    $web = @{ Used = $false; Note = ''; Files = @{}; Map = $null; Origin = '' }
    if ($null -eq $Entries) { $Entries = @(Get-OneDriveSyncEntries) }
    # Only the work OneDrive is the own OneDrive here: a personal (consumer) OneDrive is never signed in to (review of PR 109).
    if ($null -eq $Homes) { $Homes = @([Environment]::GetEnvironmentVariable('OneDriveCommercial') | Where-Object { $_ }) }
    $map = Get-OneDriveWebMap $Folder $Entries $Homes
    if (-not $map) { return $web }   # not a OneDrive folder: a plain drive scan, no note
    $web.Map = $map; $web.Origin = $map.Origin
    if (-not $map.WebFolder) {
        $web.Note = 'Not used: the OneDrive app gives no web place for this synced folder. ' + $script:NoHistoryNote + '.'
        return $web
    }
    try {
        Write-Ui 'This folder is in OneDrive: Ion signs in once to read the who/when history and the online-only files.' 'Muted'
        $script:SignInTried = $true   # before the sign-in: a Ctrl+C during it still gets the run's one sign-out
        $null = Connect-TeamSite ([pscustomobject]@{ SiteUrl = $map.Site; GroupId = $null; IsChannel = $false })
        $web.Files = Get-OneDriveWebFiles $map
        $web.Used = $true
        $web.Note = 'Used: {0} files in this folder on the web; history from the web.' -f $web.Files.Count
    }
    catch {
        # The first sentence only: a sign-in error goes on with advice for a site scan ("Nothing was scanned ...").
        $first = @(([string]$_.Exception.Message) -split "`n" | Where-Object { $_.Trim() })
        $why = ''; if ($first.Count) { $why = ' (' + (Protect-PiiText (($first[0] -split '(?<=\.)\s')[0]).TrimEnd('.')) + ')' }
        Write-Ui ('OneDrive on the web is not used' + $why + ': Ion reads the files on this computer only.') 'Accent'
        $web.Note = 'Not used' + $why + '. ' + $script:NoHistoryNote + '.'
    }
    return $web
}

function Test-OneDriveSignIn {
    # True when a sign-in to a OneDrive site was tried in this run (Find-ChannelPII.ps1 signs out once at the end).
    return $script:SignInTried
}

Export-ModuleMember -Function Get-DriveGroups, Invoke-DriveScan, Get-ReparseKind, Select-OnlineFiles, Get-OneDriveRoots,
    Get-OneDriveHomes, Confirm-OneDriveRoot, Add-OneDriveWebFiles, Get-OneDriveSyncEntries, Get-OneDriveWebMap, Get-OneDriveWebFiles,
    Connect-OneDriveWeb, Test-OneDriveSignIn
