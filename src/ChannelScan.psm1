<#
ChannelScan.psm1 - reads one team's or site's files and turns them into report rows.

  Purpose:     The SharePoint and Teams steps of Find-ChannelPII.ps1, one function per step: read the
               Teams link or site address, sign in to the team's site, find the library, list the channel
               folders, scan each file with a progress bar and live counts. Posts are not read (the PAM,
               2026-11-20: one sign-in only; reading posts needs an admin-approved Graph permission).
  Inputs:      the Teams link or site address the person gives.
  Outputs:     report rows (Report.psm1 New-ReportRow) added to the list the caller passes in.
  Depends on:  PnP.PowerShell 1.12.0; PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1, Report.psm1,
               ScanMeter.psm1 (all imported first by Find-ChannelPII.ps1).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the one SharePoint site the person gives. Nothing lists other sites or teams (no network
               reconnaissance).
  Writes:      none in SharePoint or Teams: read calls only. Rows go to an in-memory list.
  Credentials: the person's own CAC sign-in (PnP web login). Nothing is stored.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# A team site's file library: Teams names it "Documents" (URL "Shared Documents"). It is scanned first.
$script:LibraryNames = @('Shared Documents', 'Documents')

# A document library is list template 101 (SharePoint ListTemplateType.DocumentLibrary).
$script:DocumentLibraryTemplate = 101

# Visible libraries that hold SharePoint's own files, not people's files: not scanned (the PAM, 2026-11-20: Site Assets,
# Preservation Hold Library and Site Collection Documents are scanned).
$script:NotScannedLibraries = @('Style Library', 'Form Templates')

# Teams makes one top-level folder per standard channel. "Forms" is SharePoint's own folder, not a channel.
$script:NotChannelFolders = @('Forms')

# Time spent downloading and spent reading and checking, for the summary (it shows where a slow scan loses its time).
$script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }

# How the report and the screen name the files at the top of the library (in no folder).
$script:RootGroupName = '(top of the library)'
# A drive folder: X:\<folder> (a mapped drive) or \\<server>\<share>\<folder> (UNC), with or without quotes.
$script:DrivePathPattern = '^"?(?:[A-Za-z]:\\|\\\\[^\\/\s]+\\[^\\/]+)'

# Office types that SharePoint opens by file ID (Doc.aspx?sourcedoc=): the link for a file whose path holds PII.
$script:IdLinkExtensions = @('.docx', '.docm', '.doc', '.xlsx', '.xlsm', '.xls', '.pptx', '.pptm', '.ppt')

# Items per request when listing the library. SharePoint refuses views over 5,000 items
# (list view threshold); 500 keeps each request small.
$script:ListPageSize = 500

# Accepted site address: https://<host>/sites/<name> or /teams/<name>, also as a file's "Copy link"
# (https://<host>/:w:/r/teams/<name>/...). Anything after the site name is cut off, so the scan can only
# touch that one site. Other sharing links (/:w:/s/, /:w:/t/) are refused: they do not name the site.
$script:SiteUrlPattern = '^(?<host>https://[A-Za-z0-9.-]+)/(?::[a-z]:/r/)?(?<site>(?:sites|teams)/[^/?#\s]+)'

# Accepted Teams link: a team's or a channel's "Copy link" (https://<teams-host>/l/team/... or /l/channel/...).
# Only its groupId (the team's Microsoft 365 group) is used; the host is not, so any Teams cloud's link works.
# The path may hold spaces (a channel name) but not a second address ('://' or '/l/'), so text pasted after the link is not read.
$script:TeamsLinkPattern = '^https://[A-Za-z0-9.-]+/l/(?<kind>team|channel)/(?:(?!://|/l/)[^?#\r\n])+\?(?:[^#\s]*&)?groupId=' +
                           '(?<id>[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})(?![0-9A-Fa-f-])'

# =====================================================================================
# Site, library and channels
# =====================================================================================

function Add-UrlScheme {
    <#
    .SYNOPSIS
      Adds https:// to an address typed without a scheme (usaf.dps.mil/teams/<team-site>); other text stays as it is.
    .PARAMETER Url
      The address the person gave.
    .OUTPUTS
      [string] the trimmed address, with https:// when it had no scheme.
    #>
    param([string]$Url)
    $Url = $Url.Trim()
    if ($Url -and $Url -notmatch '^[A-Za-z][A-Za-z0-9+.-]*://') { return 'https://' + $Url }
    return $Url
}

function Test-TeamSiteUrl {
    <#
    .SYNOPSIS
      Checks that the address is one SharePoint site and returns it cut to the site name.
    .PARAMETER Url
      The address the person gave (it may go on to a library or a file, or be a file's "Copy link"; https:// may be left out).
    .OUTPUTS
      [string] https://<host>/sites/<name> or https://<host>/teams/<name>
    #>
    param([string]$Url)
    $match = [regex]::Match((Add-UrlScheme $Url), $script:SiteUrlPattern, 'IgnoreCase')
    if (-not $match.Success) {
        throw "Not a team site address: '$Url'. Expected https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>."
    }
    return $match.Groups['host'].Value + '/' + $match.Groups['site'].Value
}

function Get-ScanTarget {
    <#
    .SYNOPSIS
      Reads what the person gave: a Teams link (team or channel) or a team site address.
      Throws when it is neither.
    .PARAMETER Address
      The Teams link or the site address.
    .OUTPUTS
      [pscustomobject] GroupId ([guid] from a Teams link, else $null), SiteUrl (the site, else ''), IsChannel
      (true for a channel's link) and DrivePath (a drive or network folder, else '').
    #>
    param([string]$Address)
    # A folder on a mapped drive (X:\...) or a network share (\\server\share\...): scanned by DriveScan.psm1.
    if ($Address -match $script:DrivePathPattern) { return [pscustomobject]@{ GroupId = $null; SiteUrl = ''; IsChannel = $false; DrivePath = $Address.Trim('"') } }
    $link = [regex]::Match((Add-UrlScheme $Address), $script:TeamsLinkPattern, 'IgnoreCase')
    if ($link.Success) {
        return [pscustomobject]@{ GroupId = [guid]$link.Groups['id'].Value; SiteUrl = ''; IsChannel = ($link.Groups['kind'].Value -ieq 'channel'); DrivePath = '' }
    }
    if ($Address -match '/l/(team|channel)/') {
        throw 'This Teams link has no groupId. In Teams: the ... next to the team or channel name > Copy link.'
    }
    return [pscustomobject]@{ GroupId = $null; SiteUrl = (Test-TeamSiteUrl $Address); IsChannel = $false; DrivePath = '' }
}

function Assert-SiteGroup {
    <#
    .SYNOPSIS
      Checks that the connected site belongs to the team in the Teams link; throws when not (Connect-TeamSite
      then signs out), so nothing of another site is read.
    .PARAMETER GroupId
      The group ID from the Teams link.
    #>
    param([guid]$GroupId)
    $actual = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    if ($actual -ne $GroupId) {
        throw 'This site does not belong to the team in the Teams link. Nothing was scanned.'
    }
}

function ConvertTo-FriendlyError {
    <#
    .SYNOPSIS
      A sign-in error in words a person can act on: no access, site not found, sign-in not finished. Used only for
      the sign-in step (Connect-TeamSite); the original message always follows, masked, so no real error is hidden.
    .PARAMETER Message
      The error message.
    .OUTPUTS
      [string]
    #>
    param([string]$Message)
    $details = Protect-PiiText $Message
    if ($Message -match '\((401|403)\)|\bforbidden\b|access (is )?denied|do(es)? not have (access|permission)') {
        return "No access: your account cannot open this site. Ask the site's owner for access, or check the address. ($details)"
    }
    if ($Message -match '\(404\)|could not be resolved|remote name could not') { return "Site not found: check the address. ($details)" }
    if ($Message -match 'user cancel+ed|cancel+ed by the user|prompt cancel+ed') { return "The sign-in did not finish (was the sign-in window closed?). ($details)" }
    return $details
}

function Read-ScanTarget {
    <#
    .SYNOPSIS
      Asks for a Teams link or site address until it is a valid one, or until the person presses Enter (exit).
      A wrong paste is explained and asked again; nothing is signed in until the address is valid.
    .PARAMETER Given
      An address given up front (the -SiteUrl setting); used first, without asking.
    .PARAMETER Prompt
      The question to ask.
    .OUTPUTS
      Output of Get-ScanTarget, or $null when the person pressed Enter.
    #>
    param([string]$Given, [string]$Prompt)
    $answer = $Given
    while ($true) {
        if (-not $answer) { $answer = Read-Host $Prompt }
        if (-not $answer -or -not $answer.Trim()) { return $null }
        try { return Get-ScanTarget $answer.Trim() }
        catch {
            Write-Ui ('That is not a Teams link or a site address: ' + (Protect-PiiText $_.Exception.Message)) 'Alert'
            Write-Ui 'Paste it again (Teams: ... > Copy link; or the SharePoint address), or press Enter to exit.' 'Muted'
            $answer = ''
        }
    }
}

function Read-TeamSiteAddress {
    # For a Teams link: asks for the team's site address until it is valid; Enter gives up (an error the caller reports).
    while ($true) {
        $typed = Read-Host 'Team site address (https://usaf.dps.mil/teams/<team-site>), or Enter to go back'
        if (-not $typed -or -not $typed.Trim()) { throw 'No site address given.' }
        try { return Test-TeamSiteUrl $typed }
        catch { Write-Ui ('Not a site address: ' + (Protect-PiiText $_.Exception.Message)) 'Alert' }
    }
}

function Connect-TeamSite {
    <#
    .SYNOPSIS
      Signs in (own CAC, PnP web login) to the site and returns its address. For a Teams link it asks for the site
      address (a lookup at the tenant root hung on DoD365: the PAM, 2026-11-20), then checks that the site belongs to
      the team in the link. It then checks that the person can open the site. On any error or cancel it signs out
      and throws a message a person can act on (ConvertTo-FriendlyError).
    .PARAMETER Target
      Output of Get-ScanTarget.
    .OUTPUTS
      [string] the site the scan may touch.
    #>
    param($Target)
    $site = $Target.SiteUrl
    $ok = $false   # set only at the end: a cancel (Ctrl+C, Cancel on a prompt) skips 'catch' but not 'finally'
    try {
        if ($null -ne $Target.GroupId) {
            if ($Target.IsChannel) {
                Write-Ui 'Channel link: the scan covers the standard channels of its team. A private or shared channel has its own site: give that site address instead.' 'Accent'
            }
            Write-Ui "Teams link: the script needs this team's site address. In Teams: the team's Shared tab > ... > Open in SharePoint, then copy the address." 'Accent'
            $site = Read-TeamSiteAddress
        }
        Write-Ui "Signing in to $site (a sign-in window opens; if you do not see it, press Alt+Tab; if it says you have no access, close it) ..." 'Muted'
        # PnP prints a warning that suggests -Interactive (and its update notice): not useful to the person, so hidden (the PAM, 2026-11-20).
        Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
        $null = Get-PnPWeb -ErrorAction Stop   # can this person open the site? A refusal here is "no access", before any scan
        if ($null -ne $Target.GroupId) { Assert-SiteGroup $Target.GroupId }
        $ok = $true
        return $site
    }
    catch { throw (ConvertTo-FriendlyError $_.Exception.Message) }
    finally {
        # No scan starts after an error or a cancel here: sign out, so no connection stays open in the ISE session.
        if (-not $ok) {
            try { Disconnect-PnPOnline } catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error
        }
    }
}

function Get-SiteDocumentLibraries {
    <#
    .SYNOPSIS
      The document libraries of the connected site that hold people's files: every visible library (template 101),
      without SharePoint's own ($NotScannedLibraries). The team's "Documents" library comes first. Only this one
      site: subsites are not read (the PAM, 2026-11-20).
    .OUTPUTS
      [pscustomobject] Title, List (the PnP list), ServerRelativeUrl (/sites/<name>/Shared Documents),
      SiteRelativeUrl (Shared Documents); one per library.
    #>
    $lists = @(Get-PnPList -ErrorAction Stop | Where-Object {
        [int]$_.BaseTemplate -eq $script:DocumentLibraryTemplate -and -not $_.Hidden -and $script:NotScannedLibraries -notcontains $_.Title
    })
    $libraries = foreach ($list in $lists) {
        $rootUrl = (Get-PnPProperty -ClientObject $list -Property RootFolder -ErrorAction Stop).ServerRelativeUrl.TrimEnd('/')
        $siteRelative = $rootUrl -replace '^.*?/(sites|teams)/[^/]+/', ''
        [pscustomobject]@{ Title = [string]$list.Title; List = $list; ServerRelativeUrl = $rootUrl; SiteRelativeUrl = $siteRelative }
    }
    $libraries = @($libraries | Sort-Object @{ Expression = { $script:LibraryNames -notcontains $_.SiteRelativeUrl } }, Title)
    if ($libraries.Count -eq 0) { throw 'No document library with files on this site. Check the address, or ask the site owner.' }
    return $libraries
}

function Get-ChannelFolderNames {
    <#
    .SYNOPSIS
      Lists the channel folders: the top-level folders of the library, without SharePoint's own. The scan
      covers all of them (the PAM, 2026-11-20: no channel choice). A plain SharePoint site may have none.
    .PARAMETER Library
      One library of Get-SiteDocumentLibraries.
    .OUTPUTS
      [string[]] folder names, sorted.
    #>
    param($Library)
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $Library.SiteRelativeUrl -ItemType Folder -ErrorAction Stop
    return @($folders | ForEach-Object { $_.Name } | Where-Object { $script:NotChannelFolders -notcontains $_ } | Sort-Object)
}

# =====================================================================================
# Files
# =====================================================================================

function Get-FieldValue {
    # Reads one field of a list item as text (dates as yyyy-MM-dd HH:mm); empty when missing.
    param($Item, [string]$Name)
    $value = $Item[$Name]
    if ($null -eq $value) { return '' }
    if ($value -is [datetime]) { return $value.ToString('yyyy-MM-dd HH:mm') }
    return [string]$value
}

function Get-SafeLink {
    <#
    .SYNOPSIS
      A link for the report that holds no PII and opens the exact file (the PAM, 2026-11-20). It is the file's own
      path when its path below the library holds no PII (Protect-PiiName on the whole path changes nothing). Else, for an Office file,
      a link by the file's ID (<site>/_layouts/15/Doc.aspx?sourcedoc={id}), which names nothing; else the deepest
      folder whose name is clean (at least the library). Each part is escaped, so names with '#', '%' or spaces work.
    .PARAMETER Origin
      https://<host>
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .PARAMETER UniqueId
      The file's ID (the list field UniqueId), or empty.
    .OUTPUTS
      [string]
    #>
    param([string]$Origin, [string]$ServerRelativeUrl, [string]$UniqueId = '')
    $parts = $ServerRelativeUrl.TrimStart('/').Split('/')
    # Joins path parts into a link path, each part escaped.
    $escape = { param($names) ($names | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/' }
    # The first three parts are the site and the library, which the person gave or Teams named.
    $keep = [Math]::Min(3, $parts.Count)
    # Each step checks the whole path below the library, so a label in a parent folder ("CAC/<number>") is seen.
    while ($keep -lt $parts.Count) {
        $below = $parts[3..$keep] -join '/'
        if ((Protect-PiiName $below) -cne $below) { break }
        $keep++
    }
    if ($keep -eq $parts.Count) { return $Origin + '/' + (& $escape $parts) }
    $id = [guid]::Empty
    $extension = [System.IO.Path]::GetExtension($parts[-1]).ToLower()
    if ($parts.Count -ge 2 -and $script:IdLinkExtensions -contains $extension -and [guid]::TryParse($UniqueId, [ref]$id)) {
        return '{0}/{1}/_layouts/15/Doc.aspx?sourcedoc=%7B{2}%7D' -f $Origin, (& $escape $parts[0..1]), $id
    }
    return $Origin + '/' + (& $escape $parts[0..($keep - 1)])
}

function Get-ShownItem {
    <#
    .SYNOPSIS
      A file's path inside its folder, as the report and the progress bar show it: PII masked, checked together with
      the folder name, so a label in the folder ("CAC", "DOB") and a number in the file name are seen as one hit.
    .PARAMETER Channel
      The folder (channel) name, not masked.
    .PARAMETER RawItem
      The path inside the folder, not masked.
    .OUTPUTS
      [string] the masked path inside the folder.
    #>
    param([string]$Channel, [string]$RawItem)
    $masked = Protect-PiiName ($Channel + '/' + $RawItem)
    $cut = $masked.IndexOf('/')
    # A hit across the folder's slash: fall back to the strict mask of the path alone.
    if ($cut -lt 0) { return Protect-PiiText $RawItem }
    return $masked.Substring($cut + 1)
}

function Get-FileBytes {
    <#
    .SYNOPSIS
      Downloads one file into memory (never to disk).
    .DESCRIPTION
      Tries Get-PnPFile -AsMemoryStream first. If this PnP version has no such switch, it uses the
      SharePoint client call OpenBinaryStream on the same connection.
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .OUTPUTS
      [byte[]]
    #>
    param([string]$ServerRelativeUrl)
    try {
        $stream = Get-PnPFile -Url $ServerRelativeUrl -AsMemoryStream -ErrorAction Stop
        return $stream.ToArray()
    }
    catch [System.Management.Automation.ParameterBindingException] {
        # This PnP version has no -AsMemoryStream: use the client call instead (same connection, same rights).
        $context = Get-PnPContext
        $binary = $context.Web.GetFileByServerRelativeUrl($ServerRelativeUrl).OpenBinaryStream()
        $context.ExecuteQuery()
        $memory = New-Object System.IO.MemoryStream
        try { $binary.Value.CopyTo($memory); return $memory.ToArray() } finally { $memory.Dispose() }
    }
}

function Get-ChannelFiles {
    <#
    .SYNOPSIS
      The files (not folders) under one channel folder, at any depth.
    .PARAMETER AllItems
      Every item of the library (one Get-PnPListItem call for all channels).
    .PARAMETER Prefix
      /sites/<name>/<library>/<channel>/
    .OUTPUTS
      List items.
    #>
    param([object[]]$AllItems, [string]$Prefix)
    $AllItems | Where-Object {
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and (Get-FieldValue $_ 'FileRef').StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

function Get-RootFiles {
    # The files at the top of the library, in no folder (a plain SharePoint site keeps files there; a team has none).
    param([object[]]$AllItems, [string]$LibraryUrl)
    $prefix = $LibraryUrl + '/'
    $AllItems | Where-Object {
        $ref = Get-FieldValue $_ 'FileRef'
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and $ref.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            $ref.Substring($prefix.Length).IndexOf('/') -lt 0
    }
}

function Invoke-OneFileScan {
    <#
    .SYNOPSIS
      Starts one file: a hit for PII in its name, then a NOT SCANNED row (unsupported, too large, download failed)
      or its download, handed to the readers (its text rows follow when a worker finishes).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER File
      The list item.
    .PARAMETER Channel
      The channel folder name.
    .PARAMETER Prefix
      The channel folder's server-relative path with a trailing '/'.
    .PARAMETER Origin
      https://<host>
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file went to the readers (else its rows are in and it still needs counting).
    #>
    param([hashtable]$Queue, $File, [string]$Channel, [string]$Prefix, [string]$Origin, [int]$MaxFileMB)
    $rows = $Queue.Rows
    $name = Get-FieldValue $File 'FileLeafRef'
    $ref = Get-FieldValue $File 'FileRef'
    $rawChannel = $Channel
    $Channel = Protect-PiiName $Channel   # from here on, the channel name as the report shows it
    $modified = Get-FieldValue $File 'Modified'
    # The path is shown in the report: mask PII in it (with its folder name), and report a name that holds PII as a hit.
    $rawItem = $ref.Substring($Prefix.Length)
    $item = Get-ShownItem $rawChannel $rawItem
    $link = Get-SafeLink $Origin $ref (Get-FieldValue $File 'UniqueId')
    foreach ($hit in @(Find-PiiInText ($rawChannel + '/' + $rawItem))) {
        $rows.Add((New-ReportRow 'File name' $Channel $item $link $modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $size = 0L
    [void][long]::TryParse((Get-FieldValue $File 'File_x0020_Size'), [ref]$size)
    $skipReason = ''
    if (-not (Test-SupportedFile $name)) { $skipReason = 'File type is not read by this script: check by hand' }
    elseif ($size -gt ($MaxFileMB * 1MB)) { $skipReason = "Larger than $MaxFileMB MB: check by hand" }
    if ($skipReason) { $rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $skipReason)); return $false }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    Write-FileStep $size 'Downloading' 30
    try { $bytes = Get-FileBytes $ref }
    catch {
        # One file that cannot be downloaded must not stop the scan: note it and go on.
        $reason = Protect-PiiText ('Download failed: ' + $_.Exception.Message)
        $rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $reason)); return $false
    }
    finally { $script:Timing.Download += $clock.Elapsed; Write-Progress -Id 2 -Activity 'This file' -Completed }
    # The main thread downloads the next file while workers read this one (speed, the PAM 2026-11-21).
    Submit-ScanFile $Queue @{ Channel = $Channel; Item = $item; Link = $link; Modified = $modified } $name $bytes ''
    return $true
}

function Get-ScanTiming {
    # The time of the last file scan spent downloading and spent reading and checking, as minutes:seconds text.
    $format = { param($t) '{0:00}:{1:00}' -f [int][Math]::Floor($t.TotalMinutes), $t.Seconds }
    [ordered]@{ 'Time downloading' = (& $format $script:Timing.Download); 'Time reading and checking (all workers)' = (& $format $script:Timing.Check) }
}

function Get-ScanGroups {
    <#
    .SYNOPSIS
      The scan groups of one library: one per top-level folder, then the files at the top of the library.
    .PARAMETER Library
      One library of Get-SiteDocumentLibraries.
    .PARAMETER Folders
      The top-level folder names.
    .PARAMETER Label
      Put before each group name ("Library / ") when the site has several libraries; else empty.
    .OUTPUTS
      [pscustomobject] Name, Prefix, Files; a group without files is kept only when it is a folder.
    #>
    param($Library, [string[]]$Folders, [string]$Label)
    $fields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified', 'UniqueId')
    $allItems = @(Get-PnPListItem -List $Library.List -PageSize $script:ListPageSize -Fields $fields -ErrorAction Stop)
    foreach ($folder in $Folders) {
        $prefix = $Library.ServerRelativeUrl + '/' + $folder + '/'
        [pscustomobject]@{ Name = $Label + $folder; Prefix = $prefix; Files = @(Get-ChannelFiles $allItems $prefix) }
    }
    $rootFiles = @(Get-RootFiles $allItems $Library.ServerRelativeUrl)
    if ($rootFiles.Count -gt 0) {
        [pscustomobject]@{ Name = $Label + $script:RootGroupName; Prefix = $Library.ServerRelativeUrl + '/'; Files = $rootFiles }
    }
}

function Invoke-FileScan {
    <#
    .SYNOPSIS
      Scans every file in the given libraries' top-level folders and at their top, and adds report rows.
    .PARAMETER Library
      One or more libraries of Get-SiteDocumentLibraries.
    .PARAMETER Channels
      The channel folder names to scan when there is one library; with several libraries, each library's own
      top-level folders are scanned.
    .PARAMETER SiteUrl
      The site address (for the links in the report).
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .OUTPUTS
      [int] the number of files whose text was scanned.
    #>
    param($Library, [string[]]$Channels, [string]$SiteUrl, [int]$MaxFileMB, $Rows, [hashtable]$Init)
    $origin = ([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }
    $libraries = @($Library)
    $groups = @(foreach ($one in $libraries) {
        if ($libraries.Count -eq 1) { Get-ScanGroups $one $Channels ''; continue }
        # Several libraries: one you cannot open (Preservation Hold Library often) is one BLOCKED row; the rest go on.
        # Collected first, so a library that fails part way adds no groups next to its BLOCKED row.
        try { $libraryGroups = @(Get-ScanGroups $one @(Get-ChannelFolderNames $one) ($one.Title + ' / ')); $libraryGroups }   # "Library / Folder"
        catch {
            $Rows.Add((New-ReportRow 'Scan' (Protect-PiiName $one.Title) '(library)' '' '' 'BLOCKED' '' '' ('Could not read this library: ' + (Protect-PiiText $_.Exception.Message))))
        }
    })
    $total = 0
    foreach ($group in $groups) { $total += $group.Files.Count }
    Write-Ui ('==== Scanning {0} files in {1} folders ====' -f $total, $groups.Count) 'Title'
    $meter = New-ScanMeter $total
    $meter.NameWidth = Get-NameWidth @($groups | ForEach-Object { Protect-PiiName $_.Name })
    $queue = New-ScanQueue $Init $Rows $meter
    # Starts one file of a group (Invoke-GroupScan calls it); then the file's masked path for the progress bar.
    $start = { param($group, $file) Invoke-OneFileScan $queue $file $group.Name $group.Prefix $origin $MaxFileMB }
    # The file's masked path inside its folder, for the progress bar.
    $shown = { param($group, $file) Get-ShownItem $group.Name ((Get-FieldValue $file 'FileRef').Substring($group.Prefix.Length)) }
    try { Invoke-GroupScan $queue $groups $start $shown }
    finally { $script:Timing.Check = Close-ScanQueue $queue }
    # The folder rows form one table; the total closes it, under a rule.
    Write-RuleLine
    Write-MeterLine $meter 'file' -Force
    return $meter.Read
}

# =====================================================================================
# Team or plain site
# =====================================================================================

function Test-TeamSite {
    <#
    .SYNOPSIS
      True when the connected site belongs to a team (a Microsoft 365 group): its folders are then channels.
      The report names the folder column "Channel" for a team and "Folder" for a plain site (the PAM, 2026-11-20).
    .OUTPUTS
      [bool]
    #>
    $groupId = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    return [bool]($groupId -and $groupId -ne [guid]::Empty)
}

Export-ModuleMember -Function Test-TeamSiteUrl, Get-ScanTarget, Read-ScanTarget, ConvertTo-FriendlyError, Connect-TeamSite, Get-SiteDocumentLibraries, Get-ChannelFolderNames,
Invoke-FileScan, Get-ScanTiming, Test-TeamSite, Get-ShownItem
