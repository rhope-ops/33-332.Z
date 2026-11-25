<#
OneDriveWeb.psm1 - the web side of a OneDrive folder scan (plan step 4, the OneDrive hybrid).

  Purpose:     When the pasted folder is in a folder that OneDrive syncs, Ion maps it to its web folder (probe 4, runs 1
               and 2: web path = UrlNamespace path + the local path after MountPoint), signs in once to that site, and
               lists the files of that one web folder. DriveScan.psm1 then reads files on the disk from the disk, reads
               online-only and web-only files through PnP into memory (nothing is saved to the disk), and takes the
               who/when history of every file from the web (the PAM, 2026-11-24: docs/reviews/2026-11-24-onedrive-hybrid-design.md).
               A failed sign-in or mapping: the scan is local only, and the Summary says so.
  Inputs:      the pasted folder; the OneDrive app's sync list in the registry (read-only).
  Outputs:     @{ Used; Note; Files; Map; Origin } for Invoke-DriveScan.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  SiteTarget.psm1 (Connect-TeamSite), Report.psm1 (Write-Ui), PiiMasking.psm1 (Protect-PiiText) and
               PnP.PowerShell 1.12.0 (Get-PnPList, Get-PnPProperty, Get-PnPListItem; read-only), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the site of the pasted folder's own OneDrive or library ("No reconnaissance": one library, and
               only the files under the pasted folder are kept).
  Writes:      none.
  Credentials: none stored (the person's own CAC sign-in in the PnP web login).
#>

Set-StrictMode -Version 2.0

# The OneDrive app's list of synced folders (one key per folder: MountPoint, UrlNamespace).
$script:SyncEnginesKey = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
# The folder kinds that use the web (the PAM, 2026-11-24: all three, with a probe first for the 2 not tested). Probe 4
# tested 'Home' (runs 1 and 2). 'Library' and 'Shortcut' join when probe 4 runs A and B pass on the gov device.
$script:WebKinds = @('Home')
# The list fields read for each web file (as ChannelScan.psm1 reads them for a team site).
$script:WebFields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified', 'UniqueId', 'Created', 'Author', 'Editor')
$script:WebLibraryTemplate = 101
$script:WebPersonalTemplate = 700
# True after a sign-in to a OneDrive site was tried: Find-ChannelPII.ps1 then signs out at the end of the run.
$script:SignInTried = $false
$script:NoHistoryNote = 'The who/when history is not available for this scan'

function Get-OneDriveSyncEntries {
    <#
    .SYNOPSIS
      The folders the OneDrive app syncs, from its registry list (read-only): each @{ MountPoint; UrlNamespace }.
      None when the list is not there (not Windows, or no OneDrive account).
    .OUTPUTS
      [hashtable[]]
    #>
    $entries = @()
    foreach ($key in @(Get-ChildItem -Path $script:SyncEnginesKey -ErrorAction SilentlyContinue)) {
        $mount = [string]$key.GetValue('MountPoint'); $space = [string]$key.GetValue('UrlNamespace')
        if ($mount -and $space) { $entries += @{ MountPoint = $mount; UrlNamespace = $space } }
    }
    return $entries
}

function Get-OneDriveWebMap {
    <#
    .SYNOPSIS
      The web place of a local folder: the sync entry whose MountPoint holds it (the longest one: probe 4 found 2
      entries with the same folder and address), its site, its web folder, and its kind. $null when no entry holds it.
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
        if ($inside -and (-not $best -or $mount.Length -gt $best.Mount.Length)) { $best = @{ Mount = $mount; Space = [string]$entry.UrlNamespace } }
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
    return @{ Mount = $best.Mount; LibraryPath = $library; Site = $origin + $sitePath; Origin = $origin; WebFolder = ($library + $rest).TrimEnd('/'); Kind = $kind }
}

function Get-OneDriveWebFiles {
    <#
    .SYNOPSIS
      The files under the mapped web folder (any depth), from the one library whose root is the mapped library: the
      relative path in small letters ('/' between folders) -> the list item. Other items are not kept.
    .PARAMETER Map
      Output of Get-OneDriveWebMap.
    .OUTPUTS
      [hashtable]; throws when no library has that root.
    #>
    param([hashtable]$Map)
    $lists = @(Get-PnPList -ErrorAction Stop | Where-Object { @($script:WebLibraryTemplate, $script:WebPersonalTemplate) -contains [int]$_.BaseTemplate })
    $library = $null
    foreach ($list in $lists) {
        $root = (Get-PnPProperty -ClientObject $list -Property RootFolder -ErrorAction Stop).ServerRelativeUrl.TrimEnd('/')
        if ($root -ieq $Map.LibraryPath) { $library = $list; break }
    }
    if (-not $library) { throw 'The OneDrive folder was not found on the web (no library with its address).' }
    $prefix = $Map.WebFolder + '/'
    $files = @{}
    foreach ($item in @(Get-PnPListItem -List $library -PageSize 500 -Fields $script:WebFields -ErrorAction Stop)) {
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
      check), and its web files. Never throws: a folder that is not synced, a kind not tested yet, or any failure
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
    if ($null -eq $Homes) { $Homes = @(Get-OneDriveHomes) }
    $map = Get-OneDriveWebMap $Folder $Entries $Homes
    if (-not $map) { return $web }   # not a OneDrive folder: a plain drive scan, no note
    $web.Map = $map; $web.Origin = $map.Origin
    if ($script:WebKinds -notcontains $map.Kind) {
        $web.Note = 'Not used: this kind of OneDrive folder (a synced library or a shortcut) is not tested yet. ' + $script:NoHistoryNote + '.'
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
        $first = @(([string]$_.Exception.Message) -split "`n" | Where-Object { $_.Trim() })
        $why = ''; if ($first.Count) { $why = ' (' + (Protect-PiiText $first[0]) + ')' }
        Write-Ui ('OneDrive on the web is not used' + $why + ': Ion reads the files on this computer only.') 'Accent'
        $web.Note = 'Not used' + $why + '. ' + $script:NoHistoryNote + '.'
    }
    return $web
}

function Test-OneDriveSignIn {
    # True when a sign-in to a OneDrive site was tried in this run (Find-ChannelPII.ps1 signs out once at the end).
    return $script:SignInTried
}

Export-ModuleMember -Function Get-OneDriveSyncEntries, Get-OneDriveWebMap, Get-OneDriveWebFiles, Connect-OneDriveWeb, Test-OneDriveSignIn
