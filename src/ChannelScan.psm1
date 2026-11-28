<#
ChannelScan.psm1 - target parsing (Teams link, site URL, drive folder), PnP CAC sign-in, libraries, and parallel downloads.

  Purpose:     Validates the user-supplied target, resolves a Teams link to its one site, connects via PnP
               interactive CAC sign-in, lists document libraries and channel folders, and downloads files into memory.
  Inputs:      Teams link, SharePoint site address, or drive path; ReaderPool workers.
  Outputs:     validated target descriptors, SharePoint scan report rows and metrics.
  Depends on:  PnP.PowerShell 1.12.0; PiiEngine.psm1, ReaderPool.psm1, ReportExport.psm1, ScanQueue.psm1, Accountability.psm1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7) with PnP.PowerShell 1.12.0.
  Network:     SharePoint Online (https://usaf.dps.mil/sites/<team-site>) and Graph group lookup on the one signed-in host.
  Writes:      none on SharePoint (read-only).
  Credentials: none stored (interactive CAC sign-in only).
#>

Set-StrictMode -Version 2.0

# ==== SiteTarget.psm1 ====
# What Ion can scan, for the first question and after a wrong paste (the Administrator, 2026-11-21: one short block, not one long line).
$script:TargetHelp = "Paste one of these, or press Enter to exit:`n  - a Teams link: in Teams, ... next to the team or channel > Copy link`n" +
    "  - a SharePoint site address: https://usaf.dps.mil/teams/<team-site> (or /sites/)`n  - a drive folder: X:\<folder> or \\<server>\<share>\<folder>`n" +
    "  Several? Paste one, press Enter, paste the next; press Enter on an empty line to start (one workbook each)."

# A drive folder: X:\<folder> (a mapped drive) or \\<server>\<share>\<folder> (UNC), with or without quotes.
$script:DrivePathPattern = '^"?(?:[A-Za-z]:\\|\\\\[^\\/\s]+\\[^\\/]+)'

# Accepted site address: https://<host>/sites/<name> or /teams/<name>, also as a file's "Copy link"
# (https://<host>/:w:/r/teams/<name>/...). Anything after the site name is cut off, so the scan can only
# touch that one site. Other sharing links (/:w:/s/, /:w:/t/) are refused: they do not name the site.
$script:SiteUrlPattern = '^(?<host>https://[A-Za-z0-9.-]+)/(?::[a-z]:/r/)?(?<site>(?:sites|teams)/[^/?#\s]+)'

# The tenant root of DoD365: a Teams link alone is turned into its team's site here (GroupSiteManager GetSiteStatus, a
# read about the one group in the link; the Administrator's test 2 of 2026-11-21 works). Only a site on this host is used.
# GetSiteStatus reports whether the group's site exists (0 not found, 1 being made, 2 ready); it is not the call that
# creates one (CreateGroupEx). Microsoft documents no write for it; not proven on DoD365 (review of PR 49).
$script:TenantRoot = 'https://usaf.dps.mil'

# When the SharePoint sign-in cannot work on this computer (for example Connect-PnPOnline blocked outside the US: the
# Administrator, 2026-11-23), Ion does not stop: it shows how to scan a synced copy of the same files as a folder.
# Steps in ASD-STE100: one action each, with what the person sees next (review of PR 59).
$script:SyncSteps = "`nSharePoint sign-in did not work on this computer. You can scan a synced copy of the files as a folder." +
    "`nIf the library is open in your browser already, start at the step with Sync." +
    "`n  1. In Teams, click the channel. You see the channel's posts." +
    "`n  2. Click Files. You see the channel's files." +
    "`n  3. Click Open in SharePoint. You see the files in your browser." +
    "`n  4. Click Sync. You see a question to open OneDrive, or the sync starts." +
    "`n  5. If you see the question, click Open. Wait until File Explorer shows the synced folder." +
    "`n  6. In File Explorer, in the synced folder, click the address bar. You see the folder path selected." +
    "`n  7. Press Ctrl+C. The path is on the clipboard." +
    "`n  8. In Ion, at the question 'Scan more', press Ctrl+V. You see the path." +
    "`n  9. Press Enter. Ion scans the folder." +
    "`nFor your own OneDrive files, paste the path of your OneDrive folder at the question 'Scan more'."

# The longest one SharePoint call may wait for its answer (milliseconds), set on the signed-in connection when it has
# no limit or a longer one (the PnP 1.12.0 default is not verified here): a download, the file list, a version read.
# Without it a stalled call can wait without end, and Ctrl+C cannot stop a call in progress (the Administrator's gov run,
# 2026-11-23: a scan hung after 50 minutes). 5 minutes: a 25 MB file on a slow link still downloads.
$script:SharePointCallMs = 300000

# Accepted Teams link: a team's or a channel's "Copy link" (https://<teams-host>/l/team/... or /l/channel/...).
# Only its groupId (the team's Microsoft 365 group) is used; the host is not, so any Teams cloud's link works.
# The path may hold spaces (a channel name) but not a second address ('://' or '/l/'), so text pasted after the link is not read.
$script:TeamsLinkPattern = '^https://[A-Za-z0-9.-]+/l/(?<kind>team|channel)/(?:(?!://|/l/)[^?#\r\n])+\?(?:[^#\s]*&)?groupId=' +
                           '(?<id>[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})(?![0-9A-Fa-f-])'

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
        throw "'$Url' is not a SharePoint site address (https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>)."
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
      Checks that the connected site belongs to the team in the Teams link; throws when not, so nothing of another
      site is read (Connect-TeamSite returns no site; the run signs out at its end).
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
      Asks for a Teams link, site address or drive folder until it is a valid one, or until the person presses Enter
      (exit); what can be pasted is listed first ($TargetHelp). A wrong paste is explained and asked again; nothing is
      signed in until the address is valid.
    .PARAMETER Given
      An address given up front (the -SiteUrl setting); used first, without asking.
    .PARAMETER Prompt
      The question to ask.
    .PARAMETER NoHelp
      Do not list what can be pasted first (after a scan; a wrong paste still lists it).
    .OUTPUTS
      Output of Get-ScanTarget, or $null when the person pressed Enter.
    #>
    param([string]$Given, [string]$Prompt, [switch]$NoHelp)
    $answer = $Given
    if (-not $answer -and -not $NoHelp) { Write-Ui $script:TargetHelp 'Muted' }
    while ($true) {
        if (-not $answer) { $answer = Read-Host $Prompt }
        if (-not $answer -or -not $answer.Trim()) { return $null }
        try { return Get-ScanTarget $answer.Trim() }
        catch {
            Write-Ui ('Not read: ' + (Protect-PiiText $_.Exception.Message)) 'Alert'
            Write-Ui $script:TargetHelp 'Muted'
            $answer = ''
        }
    }
}

function Read-TeamSiteAddress {
    <#
    .SYNOPSIS
      For a Teams link: asks for the team's site address until it is a valid one; Enter gives up (an error the caller
      reports). Nothing is signed in here.
    .OUTPUTS
      [string] the site address, cut to the site.
    #>
    while ($true) {
        $typed = Read-Host 'Team site address (https://usaf.dps.mil/teams/<team-site>), or Enter to go back'
        if (-not $typed -or -not $typed.Trim()) { throw 'No site address given.' }
        try { return Test-TeamSiteUrl $typed }
        catch { Write-Ui (Protect-PiiText $_.Exception.Message) 'Alert' }
    }
}

function Find-TeamSiteByGroup {
    <#
    .SYNOPSIS
      The site of the one team in a Teams link: signs in at the tenant root (own CAC, web login) and asks SharePoint
      for that group's site (GroupSiteManager GetSiteStatus, a read; nothing is listed: "No reconnaissance"). Returns
      '' when the sign-in or the lookup fails, gives no site, or gives a site on another host (Data boundary): the
      caller then asks for the site address, as before.
    .PARAMETER GroupId
      The group ID from the Teams link.
    .OUTPUTS
      [string] https://<tenant root host>/sites/<name> or /teams/<name>, or ''.
    #>
    param([guid]$GroupId)
    Write-Ui "Teams link: looking up this team's site ..." 'Muted'
    Write-Ui '  If a sign-in window opens, sign in with your CAC. If you do not see the window, press Alt+Tab.' 'Muted'
    try {
        Connect-PnPOnline -Url $script:TenantRoot -UseWebLogin -WarningAction SilentlyContinue
        Set-SharePointTimeout
        $answer = Invoke-PnPSPRestMethod -Method Post -Url '/_api/GroupSiteManager/GetSiteStatus' -Content @{ groupId = [string]$GroupId } -ErrorAction Stop
    }
    catch { Write-Verbose ('Team site lookup not possible: ' + $_.Exception.Message); return '' }
    $url = ''
    if ($answer -and $answer.PSObject.Properties['SiteUrl']) { $url = [string]$answer.SiteUrl }
    if (-not $url) { return '' }
    try { $site = Test-TeamSiteUrl $url } catch { return '' }   # not a site address: ask instead
    if (-not $site.StartsWith($script:TenantRoot + '/', [StringComparison]::OrdinalIgnoreCase)) { return '' }
    return $site
}

function Set-SharePointTimeout {
    <#
    .SYNOPSIS
      Sets the time limit of each SharePoint call on the signed-in connection ($SharePointCallMs). A call over it fails
      with a time-out: a download then gets a NOT SCANNED row, and the scan goes on.
    #>
    try { $context = Get-PnPContext } catch { Write-Verbose 'No client context: no time limit set.'; return }   # never a sign-in error
    if (-not ($context -and $context.PSObject.Properties['RequestTimeout'])) { return }
    if ($context.RequestTimeout -le 0 -or $context.RequestTimeout -gt $script:SharePointCallMs) { $context.RequestTimeout = $script:SharePointCallMs }
}

function Assert-ConnectedSite {
    <#
    .SYNOPSIS
      Checks that the signed-in connection reaches the given site: its web's full address (host and path) equals the
      site's address. A sign-in without a sign-out keeps the last site's connection when it fails; this check makes
      sure that no read goes to that other site (also on another host). Throws when the web is another site, or cannot
      be read (no access).
    .PARAMETER Site
      The site address that was signed in to.
    #>
    param([string]$Site)
    $want = [uri]::UnescapeDataString(([uri]$Site).GetLeftPart([UriPartial]::Path)).TrimEnd('/')
    $web = Get-PnPWeb -Includes Url -ErrorAction Stop
    $actual = [uri]::UnescapeDataString([string]$web.Url).TrimEnd('/')
    if (-not [string]::Equals($actual, $want, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The sign-in did not reach this site. Nothing was scanned. Close Windows PowerShell ISE. Open it again. Then run Ion again.'
    }
}

function Connect-TeamSite {
    <#
    .SYNOPSIS
      Signs in (own CAC, PnP web login) to the site and returns its address. For a Teams link (a team's or a
      channel's) it looks the site up at the tenant root (Find-TeamSiteByGroup), and asks for the site address only
      when that fails; it then checks that the site belongs to the team in the link. It then checks that the connection
      reaches that site (Assert-ConnectedSite) and that the person can open it. On any error it throws a message a
      person can act on (ConvertTo-FriendlyError). It does not sign out: the next site connects without a sign-out,
      and Find-ChannelPII.ps1 signs out once at the end of the run, after an error or Ctrl+C too (the Administrator, 2026-11-24).
    .PARAMETER Target
      Output of Get-ScanTarget.
    .OUTPUTS
      [string] the site the scan may touch.
    #>
    param($Target)
    $site = $Target.SiteUrl
    $signingIn = $false   # true only while the sign-in itself runs: only its errors get the sync steps (review of PR 59)
    try {
        if ($null -ne $Target.GroupId) {
            if ($Target.IsChannel) {
                Write-Ui 'Channel link: the scan covers the standard channels of its team. A private or shared channel has its own site: give that site address instead.' 'Accent'
            }
            $site = Find-TeamSiteByGroup $Target.GroupId
            if (-not $site) {
                Write-Ui "The team's site could not be looked up: the script needs its address. In Teams: the team's Shared tab > ... > Open in SharePoint, then copy the address." 'Accent'
                $site = Read-TeamSiteAddress
            }
        }
        Write-Ui ('Signing in to ' + (Protect-PiiName $site) + ' ...') 'Muted'
        Write-Ui '  If a sign-in window opens, sign in with your CAC. If you do not see the window, press Alt+Tab. If it says that you have no access, close it.' 'Muted'
        # PnP prints a warning that suggests -Interactive (and its update notice): not useful to the person, so hidden (the Administrator, 2026-11-20).
        $signingIn = $true
        Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
        $signingIn = $false
        Set-SharePointTimeout
        Assert-ConnectedSite $site   # also: can this person open the site? A refusal here is "no access", before any scan
        if ($null -ne $Target.GroupId) { Assert-SiteGroup $Target.GroupId }
        return $site
    }
    catch {
        $friendly = ConvertTo-FriendlyError $_.Exception.Message
        # Not "no access", "not found" or "sign-in not finished": the sign-in itself does not work here (blocked, missing).
        if ($signingIn -and $friendly -eq (Protect-PiiText $_.Exception.Message)) { $friendly += $script:SyncSteps }
        throw $friendly
    }
}

function Read-ScanTargetList {
    <#
    .SYNOPSIS
      Asks for what to scan, then "Add another, or press Enter to start" until an empty line: one or more targets
      (the Administrator, 2026-11-24: several sites per run, one workbook each). Each target is checked as it is pasted
      (Read-ScanTarget); only the targets the person pastes are touched ("No reconnaissance").
    .PARAMETER Given
      A target given up front (the -SiteUrl setting): the list is that one target, without asking.
    .PARAMETER Prompt
      The first question.
    .PARAMETER NoHelp
      Do not list what can be pasted first.
    .OUTPUTS
      [object[]] outputs of Get-ScanTarget; empty when the person pressed Enter at the first question.
    #>
    param([string]$Given, [string]$Prompt, [switch]$NoHelp)
    $first = Read-ScanTarget $Given $Prompt -NoHelp:$NoHelp
    if ($null -eq $first) { return @() }
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add($first)
    if ($Given) { return $list.ToArray() }
    while ($true) {
        $next = Read-ScanTarget '' ('Add another ({0} so far), or press Enter to start' -f $list.Count) -NoHelp
        if ($null -eq $next) { break }
        $list.Add($next)
    }
    return $list.ToArray()
}

function Get-TargetLabel {
    <#
    .SYNOPSIS
      A target in a few words for the list and the closing table: the site address or the drive folder (both masked
      like any name Ion shows) or "Teams link (team)" / "Teams link (channel)".
    .PARAMETER Target
      Output of Get-ScanTarget.
    .OUTPUTS
      [string]
    #>
    param($Target)
    if ($Target.DrivePath) { return Protect-PiiName $Target.DrivePath }
    if ($Target.SiteUrl) { return Protect-PiiName $Target.SiteUrl }
    if ($Target.IsChannel) { return 'Teams link (channel)' }
    return 'Teams link (team)'
}

# ==== ChannelScan.psm1 ====
# =====================================================================================
# Settings
# =====================================================================================

# A team site's file library: Teams names it "Documents" (URL "Shared Documents"). It is scanned first.
$script:LibraryNames = @('Shared Documents', 'Documents')

# A document library is list template 101 (SharePoint ListTemplateType.DocumentLibrary).
$script:DocumentLibraryTemplate = 101

# Visible libraries that hold SharePoint's own files, not people's files: not scanned (the Administrator, 2026-11-20: Site Assets,
# Preservation Hold Library and Site Collection Documents are scanned).
$script:NotScannedLibraries = @('Style Library', 'Form Templates')

# Teams makes one top-level folder per standard channel. "Forms" is SharePoint's own folder, not a channel.
$script:NotChannelFolders = @('Forms')

# Time spent downloading and spent reading, for the summary (it shows where a slow scan loses its time).
# The per-file bar's heading while a file comes down from SharePoint into memory (never to disk). ISE adds a "." after a
# bar's heading, so two here show as "..." (the Administrator's layout test E, 2026-11-21).
$script:DownloadStep = 'Downloading..'

$script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }

# Streamed download (Get-FileBytes, the Administrator 2026-11-21): File.OpenBinaryDirect on the signed-in connection, so the bar
# counts bytes; $false for the rest of a scan after one failure. Tests replace the opener.
$script:DirectDownload = $true
# Opens the file as a stream on the signed-in PnP connection.
$script:OpenDirectStream = {
    param([string]$ServerRelativeUrl)
    if (-not ($fileType = 'Microsoft.SharePoint.Client.File' -as [type])) { throw 'The SharePoint client library is not loaded.' }
    return $fileType::OpenBinaryDirect((Get-PnPContext), $ServerRelativeUrl).Stream
}

# Parallel downloads (the Administrator, 2026-11-24): 3 at once in a Fast run, 2 in a Normal run (the Administrator, 2026-11-25), each in a
# worker with its own copy of the signed-in connection.
$script:ParallelDownloads = 3
$script:NormalDownloads = 2
# The longest one download may take in a worker (seconds; as the reader limit, ReaderPool.psm1). Over it, the file is
# tried on the main thread.
$script:DownloadSecondsLimit = 180
# How long the main thread waits between bar updates while downloads run (milliseconds).
$script:DownloadWaitMs = 250
$script:DownloadBarMs = 250
# The header line "Parallel downloads" of the last site scan: "2" or "3" (the speed), or "off (<error type>)"; empty before a scan.
$script:ParallelState = ''
# Makes a copy of the signed-in connection (PnP.Framework; test 3 of docs/reviews/2026-11-21-ion-round2-tests.md).
# Tests replace it.
$script:CloneContext = {
    param($Context, [string]$SiteUrl)
    return [Microsoft.SharePoint.Client.ClientContextExtensions]::Clone($Context, $SiteUrl)
}
# What a download worker runs: its copy of the connection reads one file into memory (never to disk), and puts its
# bytes done in the shared table. It gives back the bytes, or the error's type name only (a message can hold a path).
$script:DownloadWorkerScript = @'
param($Copy, $Path, $Progress, $Key)
[System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal
$answer = @{ Bytes = $null; Error = '' }
try {
    $binary = $Copy.Web.GetFileByServerRelativeUrl($Path).OpenBinaryStream()
    $Copy.ExecuteQuery()
    $source = $binary.Value
    $memory = New-Object System.IO.MemoryStream
    try {
        $buffer = New-Object byte[] 65536
        while (($read = $source.Read($buffer, 0, $buffer.Length)) -gt 0) { $memory.Write($buffer, 0, $read); $Progress[$Key] = $memory.Length }
        $answer.Bytes = $memory.ToArray()
    }
    finally { $memory.Dispose(); $source.Dispose() }
}
catch {
    $inner = $_.Exception
    while ($inner -is [System.Management.Automation.MethodInvocationException] -and $inner.InnerException) { $inner = $inner.InnerException }
    $answer.Error = $inner.GetType().Name
}
$answer
'@

# How the report and the screen name the files at the top of the library (in no folder).
$script:RootGroupName = '(top of the library)'

# Office types that SharePoint opens by file ID (Doc.aspx?sourcedoc=): the link for a file whose path holds PII.
$script:IdLinkExtensions = @('.docx', '.docm', '.doc', '.xlsx', '.xlsm', '.xls', '.pptx', '.pptm', '.ppt')

# Items per request when listing the library. SharePoint refuses views over 5,000 items
# (list view threshold); 500 keeps each request small.
$script:ListPageSize = 500

# =====================================================================================
# Site, library and channels
# =====================================================================================

function Get-SiteDocumentLibraries {
    <#
    .SYNOPSIS
      The document libraries of the connected site that hold people's files: every visible library (template 101),
      without SharePoint's own ($NotScannedLibraries). The team's "Documents" library comes first. Only this one
      site: subsites are not read (the Administrator, 2026-11-20).
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
      covers all of them (the Administrator, 2026-11-20: no channel choice). A plain SharePoint site may have none.
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
      A link for the report that holds no PII and opens the exact file (the Administrator, 2026-11-20). It is the file's own
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
      A file large enough for the per-file bar is read from SharePoint as a stream (a read: HTTP GET of the file on the
      signed-in connection, same site, same rights), a chunk at a time, so the bar shows the MB done. If that fails,
      or for a small file, Get-PnPFile -AsMemoryStream gives the whole file in one call; if this PnP version has no
      such switch, the SharePoint client call OpenBinaryStream on the same connection does.
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .PARAMETER Size
      The file size from the list (0 when not known).
    .OUTPUTS
      [byte[]]
    #>
    param([string]$ServerRelativeUrl, [long]$Size = 0)
    # '#' and '%' in a path break the stream call's address: such a file takes the normal download.
    if ($script:DirectDownload -and (Test-FileBarSize $Size) -and $ServerRelativeUrl.IndexOfAny([char[]]'#%') -lt 0) {
        try {
            $stream = & $script:OpenDirectStream $ServerRelativeUrl
            try { $bytes = Copy-StreamWithProgress $stream $Size $script:DownloadStep } finally { $stream.Dispose() }
            if ($bytes.LongLength -eq $Size) { return , $bytes }
            # Not the size the list gives (a sign-in page, for example): the normal download decides.
            throw ('the stream gave {0} bytes, the list says {1}' -f $bytes.LongLength, $Size)
        }
        catch {
            # The streamed download does not work with this sign-in or PnP version: the normal one for the rest of the scan.
            $script:DirectDownload = $false
            Write-Verbose ('Streamed download not possible, normal download used: ' + $_.Exception.Message)
        }
    }
    Write-FileStep $Size $script:DownloadStep   # no byte count: the bar says so
    # The leading commas keep the byte array whole: PowerShell would otherwise unroll it, byte by byte (slow).
    try {
        $stream = Get-PnPFile -Url $ServerRelativeUrl -AsMemoryStream -ErrorAction Stop
        return , $stream.ToArray()
    }
    catch [System.Management.Automation.ParameterBindingException] {
        # This PnP version has no -AsMemoryStream: use the client call instead (same connection, same rights).
        $context = Get-PnPContext
        $binary = $context.Web.GetFileByServerRelativeUrl($ServerRelativeUrl).OpenBinaryStream()
        $context.ExecuteQuery()
        $memory = New-Object System.IO.MemoryStream
        try { $binary.Value.CopyTo($memory); return , $memory.ToArray() } finally { $memory.Dispose() }
    }
}

function Get-FilesByFolder {
    <#
    .SYNOPSIS
      The files (not folders) of a library, sorted once by their top-level folder (any depth below it); the files at
      the top of the library, in no folder, under the key ''. One pass over the items, not one per folder (speed).
    .PARAMETER AllItems
      Every item of the library (one Get-PnPListItem call for all channels).
    .PARAMETER LibraryUrl
      /sites/<name>/<library>
    .OUTPUTS
      [hashtable] top-level folder name (any letter case) -> list of items.
    #>
    param([object[]]$AllItems, [string]$LibraryUrl)
    $prefix = $LibraryUrl + '/'
    $byFolder = @{}
    foreach ($item in $AllItems) {
        if ((Get-FieldValue $item 'FSObjType') -ne '0') { continue }
        $ref = Get-FieldValue $item 'FileRef'
        if (-not $ref.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $slash = $ref.IndexOf('/', $prefix.Length)
        $folder = ''; if ($slash -ge 0) { $folder = $ref.Substring($prefix.Length, $slash - $prefix.Length) }
        if (-not $byFolder.ContainsKey($folder)) { $byFolder[$folder] = New-Object System.Collections.Generic.List[object] }
        $byFolder[$folder].Add($item)
    }
    return $byFolder
}

function Start-FileScan {
    <#
    .SYNOPSIS
      The first part of one file: a hit for PII in its name, then a NOT SCANNED row when it is not read (type, size).
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
      [hashtable] the file to download (Info, Ref, Name, Size, File, NameRows: the rows added here); $null when it got
      its NOT SCANNED row.
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
    $size = 0L
    [void][long]::TryParse((Get-FieldValue $File 'File_x0020_Size'), [ref]$size)
    $info = @{ Channel = $Channel; Item = $item; Link = $link; Modified = $modified }
    $first = $rows.Count
    if (-not (Add-FileStartRows $rows $info ($rawChannel + '/' + $rawItem) $name $size $MaxFileMB)) { return $null }
    $nameRows = New-Object System.Collections.Generic.List[object]
    for ($r = $first; $r -lt $rows.Count; $r++) { $nameRows.Add($rows[$r]) }
    return @{ Info = $info; Ref = $ref; Name = $name; Size = $size; File = $File; NameRows = $nameRows }
}

function Complete-FileScan {
    <#
    .SYNOPSIS
      The last part of one downloaded file: who made and changed it, then its bytes to the readers (its text rows follow
      when a worker finishes).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Prepared
      Output of Start-FileScan.
    .PARAMETER Bytes
      The file bytes.
    #>
    param([hashtable]$Queue, [hashtable]$Prepared, [byte[]]$Bytes)
    $info = $Prepared.Info
    $ref = $Prepared.Ref
    $name = $Prepared.Name
    # Accountability: who made and changed the file; for a file with a hit, who added each hit (version history).
    $who = Get-SharePointWho $Prepared.File
    $who.Modified = $info.Modified
    $info.Who = $who
    $info.Added = { param($hits) Find-PiiAddedVersion $ref $name $hits $who }.GetNewClosure()
    $file = $Prepared.File
    $info.Access = { Get-SharePointAccess $file }.GetNewClosure()
    Submit-ScanFile $Queue $info $name $Bytes ''
}

function Add-DownloadFailedRow {
    <#
    .SYNOPSIS
      The NOT SCANNED row of a file that could not be downloaded, counted on the meter with its file-name hits.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Prepared
      Output of Start-FileScan.
    .PARAMETER Message
      The error message (masked here).
    .PARAMETER Count
      Count the file on the meter (the parallel scan); the one-by-one scan counts it in Invoke-GroupScan.
    #>
    param([hashtable]$Queue, [hashtable]$Prepared, [string]$Message, [switch]$Count)
    $info = $Prepared.Info
    $row = New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText ('Download failed: ' + $Message))
    $Queue.Rows.Add($row)
    if (-not $Count) { return }
    $fileRows = New-Object System.Collections.Generic.List[object]
    foreach ($nameRow in $Prepared.NameRows) { $fileRows.Add($nameRow) }
    $fileRows.Add($row)
    Update-ScanMeter $Queue.Meter $fileRows 0 0
    Write-MeterLine $Queue.Meter 'file'
}

function Invoke-OneFileScan {
    <#
    .SYNOPSIS
      Starts one file on the main thread: Start-FileScan, its download, then Complete-FileScan.
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
    .PARAMETER Pool
      Optional: the download pool (New-DownloadPool) while worker downloads run. Its Busy clock pauses for this
      download only (counted here), and runs again while the file goes to the readers (the workers still download).
    .OUTPUTS
      [bool] true when the file went to the readers (else its rows are in and it still needs counting).
    #>
    param([hashtable]$Queue, $File, [string]$Channel, [string]$Prefix, [string]$Origin, [int]$MaxFileMB, [hashtable]$Pool)
    $prepared = Start-FileScan $Queue $File $Channel $Prefix $Origin $MaxFileMB
    if (-not $prepared) { return $false }
    $paused = $Pool -and $Pool.Busy.IsRunning
    if ($paused) { $Pool.Busy.Stop() }   # no double count: this download's time is counted on its own
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try { $bytes = Get-FileBytes $prepared.Ref $prepared.Size }
    catch {
        # One file that cannot be downloaded must not stop the scan: note it and go on.
        Add-DownloadFailedRow $Queue $prepared $_.Exception.Message; return $false
    }
    finally {
        $script:Timing.Download += $clock.Elapsed; Write-Progress -Id 2 -Activity $script:DownloadStep -Completed
        if ($paused -and $Pool.Jobs.Count -gt 0) { $Pool.Busy.Start() }
    }
    # The main thread downloads the next file while workers read this one (speed, the Administrator 2026-11-21).
    Complete-FileScan $Queue $prepared $bytes
    return $true
}

# =====================================================================================
# Parallel downloads (plan step 6, the Administrator 2026-11-24: docs/reviews/2026-11-24-parallel-downloads-design.md)
# =====================================================================================

function Get-DownloadPlan {
    <#
    .SYNOPSIS
      The order of one folder's files (the Administrator, 2026-11-24): the files to download first, largest first (equal sizes
      keep the list order); then the files that are not downloaded (file type not read, or over the size limit).
    .PARAMETER Files
      The folder's list items.
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [hashtable] Download (list items, in order), Other (list items, in list order).
    #>
    param([object[]]$Files, [int]$MaxFileMB)
    $download = New-Object System.Collections.Generic.List[object]
    $other = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($file in $Files) {
        $size = 0L
        [void][long]::TryParse((Get-FieldValue $file 'File_x0020_Size'), [ref]$size)
        if ((Test-SupportedFile (Get-FieldValue $file 'FileLeafRef')) -and $size -le $MaxFileMB * 1MB) {
            $download.Add([pscustomobject]@{ File = $file; Size = $size; Index = $index })
        }
        else { $other.Add($file) }
        $index++
    }
    # Sort-Object is not stable in Windows PowerShell 5.1: the list index keeps equal sizes in list order.
    $sorted = @($download | Sort-Object @{ Expression = 'Size'; Descending = $true }, @{ Expression = 'Index'; Descending = $false } |
        ForEach-Object { $_.File })
    return @{ Download = $sorted; Other = $other.ToArray() }
}

function Get-ErrorTypeName {
    # The type name of an error, under PowerShell's wrapper of a .NET call error; never the message (it can hold a path).
    param([System.Exception]$Exception)
    $inner = $Exception
    while ($inner -is [System.Management.Automation.MethodInvocationException] -and $inner.InnerException) { $inner = $inner.InnerException }
    return $inner.GetType().Name
}

function New-DownloadPool {
    <#
    .SYNOPSIS
      The download workers of one site scan: one copy of the signed-in connection per worker (no new sign-in), and a
      runspace pool. When a copy cannot be made, the pool is off and the scan downloads one file at a time.
    .PARAMETER SiteUrl
      The signed-in site.
    .OUTPUTS
      [hashtable] the pool state (On, State: the header text "2" or "3", or "off (<reason>)"). The reason is the type
      name of the error, or one of Ion's own names: WrongSize (not the size the list gives), NoResult (the worker gave
      nothing back), NoWorker (no copy of the connection left).
    #>
    param([string]$SiteUrl)
    $count = $script:NormalDownloads
    if ((Get-ScanSpeed) -eq 'Fast') { $count = $script:ParallelDownloads }
    $state = @{ On = $false; State = ''; Count = $count; Copies = @{}; Pool = $null; Stuck = 0; SiteUrl = $SiteUrl
        Jobs = (New-Object System.Collections.Generic.List[object]); Free = (New-Object System.Collections.Generic.Stack[int])
        Progress = [hashtable]::Synchronized(@{}); Busy = (New-Object Diagnostics.Stopwatch)
        BurstClock = (New-Object Diagnostics.Stopwatch); BurstBytes = 0L; BurstDone = 0L; LastBar = 0L; NextKey = 0 }
    try {
        $context = Get-PnPContext
        for ($slot = 0; $slot -lt $count; $slot++) { $state.Copies[$slot] = & $script:CloneContext $context $SiteUrl }
        # .NET Framework allows 2 connections per host to a desktop program: N downloads need N, and the main thread 1.
        # Only this one host; in memory, it ends with the PowerShell window.
        $point = [System.Net.ServicePointManager]::FindServicePoint([uri]$SiteUrl)
        if ($point.ConnectionLimit -lt $count + 1) { $point.ConnectionLimit = $count + 1 }
        $state.Pool = [runspacefactory]::CreateRunspacePool(1, $count)
        $state.Pool.Open()
        for ($slot = $count - 1; $slot -ge 0; $slot--) { $state.Free.Push($slot) }
        $state.On = $true
        $state.State = [string]$count
    }
    catch {
        $state.State = 'off ({0})' -f (Get-ErrorTypeName $_.Exception)
        Close-DownloadPool $state
    }
    return $state
}

function Start-DownloadJob {
    <#
    .SYNOPSIS
      Hands one file to a free download worker (the caller checks that one is free).
    .PARAMETER State
      Output of New-DownloadPool.
    .PARAMETER Prepared
      Output of Start-FileScan.
    #>
    param([hashtable]$State, [hashtable]$Prepared)
    if ($State.Jobs.Count -eq 0) {
        # A new burst: the bar counts from here.
        # LastBar too, so the bar shows at once in each new burst (review of PR 115).
        $State.BurstBytes = 0L; $State.BurstDone = 0L; $State.LastBar = 0L; $State.BurstClock.Restart(); $State.Busy.Start()
    }
    $slot = $State.Free.Pop()
    # The bytes done go under a key of this job, not of the worker: a stopped worker cannot change a later job's count.
    $State.NextKey++
    $key = $State.NextKey
    $State.Progress[$key] = 0L
    $State.BurstBytes += $Prepared.Size
    $shell = [powershell]::Create()
    $shell.RunspacePool = $State.Pool
    [void]$shell.AddScript($script:DownloadWorkerScript).AddArgument($State.Copies[$slot]).AddArgument($Prepared.Ref).AddArgument($State.Progress).AddArgument($key)
    $State.Jobs.Add(@{ Shell = $shell; Handle = $shell.BeginInvoke(); Slot = $slot; Key = $key; Prepared = $Prepared; Clock = [Diagnostics.Stopwatch]::StartNew() })
}

function Write-DownloadBar {
    # The one download bar for all files in flight (Write-DownloadsStep), at most every $DownloadBarMs.
    param([hashtable]$State)
    if ($State.BurstClock.ElapsedMilliseconds - $State.LastBar -lt $script:DownloadBarMs -and $State.LastBar -gt 0) { return }
    $State.LastBar = [Math]::Max(1L, $State.BurstClock.ElapsedMilliseconds)
    $done = $State.BurstDone
    foreach ($job in $State.Jobs) { $done += [long]$State.Progress[$job.Key] }
    Write-DownloadsStep $State.Jobs.Count $State.BurstBytes $done $State.BurstClock.Elapsed $script:DownloadStep
}

function Stop-DownloadJob {
    <#
    .SYNOPSIS
      Stops one download over the time limit without waiting (a stuck network call does not stop). Its copy of the
      connection is not used again: the worker gets a new copy, or is dropped when no copy can be made.
    .PARAMETER State
      Output of New-DownloadPool.
    .PARAMETER Job
      The job (from State.Jobs).
    #>
    param([hashtable]$State, [hashtable]$Job)
    [void]$State.Jobs.Remove($Job)
    try { [void]$Job.Shell.BeginStop($null, $null) } catch { Write-Verbose 'The download had already stopped.' }
    $State.Stuck++
    [void]$State.Pool.SetMaxRunspaces($State.Count + $State.Stuck)
    try { $State.Copies[$Job.Slot] = & $script:CloneContext (Get-PnPContext) $State.SiteUrl; $State.Free.Push($Job.Slot) }
    catch { Write-Verbose 'No new copy of the connection: one download worker less.' }
}

function Receive-DownloadResult {
    <#
    .SYNOPSIS
      Waits (a slice at a time, the bar moving) until at least one download finished or went over the time limit, and
      gives back each: Prepared, Bytes, Error ('' when the bytes are right).
    .PARAMETER State
      Output of New-DownloadPool.
    .OUTPUTS
      [hashtable] one per finished download.
    #>
    param([hashtable]$State)
    while ($true) {
        $ready = @($State.Jobs | Where-Object { $_.Handle.IsCompleted })
        $slow = @($State.Jobs | Where-Object { -not $_.Handle.IsCompleted -and $_.Clock.Elapsed.TotalSeconds -gt $script:DownloadSecondsLimit })
        foreach ($job in $slow) {
            Stop-DownloadJob $State $job
            $State.Progress.Remove($job.Key)
            $State.BurstDone += [long]$job.Prepared.Size
            @{ Prepared = $job.Prepared; Bytes = $null; Error = 'TimeoutException' }
        }
        if ($ready.Count -gt 0 -or $slow.Count -gt 0 -or $State.Jobs.Count -eq 0) { break }
        Write-DownloadBar $State
        [void]$State.Jobs[0].Handle.AsyncWaitHandle.WaitOne($script:DownloadWaitMs)
    }
    foreach ($job in $ready) {
        [void]$State.Jobs.Remove($job)
        $answer = @{ Prepared = $job.Prepared; Bytes = $null; Error = '' }
        try {
            $output = @($job.Shell.EndInvoke($job.Handle)) | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
            if (-not $output) { $answer.Error = 'NoResult' }
            elseif ($output.Error) { $answer.Error = [string]$output.Error }
            else { $answer.Bytes = [byte[]]$output.Bytes }
        }
        catch { $answer.Error = Get-ErrorTypeName $_.Exception }
        finally { $job.Shell.Dispose(); $State.Free.Push($job.Slot); $State.Progress.Remove($job.Key) }
        # Not the size the list gives (a sign-in page, for example): a failure. Size 0 in the list: not checked.
        $size = [long]$job.Prepared.Size
        if (-not $answer.Error -and ($null -eq $answer.Bytes -or ($size -gt 0 -and $answer.Bytes.LongLength -ne $size))) { $answer.Error = 'WrongSize'; $answer.Bytes = $null }
        $State.BurstDone += $size
        $answer
    }
    if ($State.Jobs.Count -eq 0) { $State.Busy.Stop() }
}

function Complete-Download {
    <#
    .SYNOPSIS
      A finished worker download: its bytes to the readers. A failed one is tried once more on the main thread: when
      that works, the copy of the connection is at fault and parallel downloads stop for the rest of the scan; when it
      also fails, the file is at fault: its NOT SCANNED row, and parallel downloads go on. A WrongSize that the main
      thread gets too (the same other size) is a fault of the list size, not of the copy: parallel downloads go on.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER State
      Output of New-DownloadPool.
    .PARAMETER Done
      One result of Receive-DownloadResult.
    #>
    param([hashtable]$Queue, [hashtable]$State, [hashtable]$Done)
    $prepared = $Done.Prepared
    if (-not $Done.Error) { Complete-FileScan $Queue $prepared $Done.Bytes; return }
    $running = $State.Busy.IsRunning
    $State.Busy.Stop()   # the main thread's download time is counted on its own (no double count)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try { $bytes = Get-FileBytes $prepared.Ref $prepared.Size }
    catch { Add-DownloadFailedRow $Queue $prepared $_.Exception.Message -Count; return }
    finally {
        $script:Timing.Download += $clock.Elapsed
        Write-Progress -Id 2 -Activity $script:DownloadStep -Completed
        if ($running -and $State.Jobs.Count -gt 0) { $State.Busy.Start() }
    }
    # The main thread got the same other size: the list size is wrong, not the copy of the connection (review of PR 115).
    $fileFault = $Done.Error -eq 'WrongSize' -and $prepared.Size -gt 0 -and $bytes.LongLength -ne $prepared.Size
    if ($State.On -and -not $fileFault) { $State.On = $false; $State.State = 'off ({0})' -f $Done.Error }
    Complete-FileScan $Queue $prepared $bytes
}

function Invoke-ParallelGroup {
    <#
    .SYNOPSIS
      Starts the files of one folder in the planned order (Get-DownloadPlan), up to $ParallelDownloads downloads at a
      time, then the files that are not downloaded (Invoke-GroupScan waits for the readers and writes the folder line).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Group
      One group of Get-ScanGroups.
    .PARAMETER State
      Output of New-DownloadPool.
    .PARAMETER Origin
      https://<host>
    .PARAMETER MaxFileMB
      The size limit.
    #>
    param([hashtable]$Queue, $Group, [hashtable]$State, [string]$Origin, [int]$MaxFileMB)
    $meter = $Queue.Meter
    $plan = Get-DownloadPlan $Group.Files $MaxFileMB
    # Counts a file that did not go to the readers from the rows it added (as Invoke-GroupScan does).
    $count = { param($first) Update-ScanMeter $meter $Queue.Rows $first 0; Write-MeterLine $meter 'file' }
    $next = 0
    $started = $false
    try {
        while ($next -lt $plan.Download.Count -or $State.Jobs.Count -gt 0) {
            # No worker left (each new copy of the connection failed): one at a time, not an endless loop (review of PR 115).
            if ($State.On -and $State.Free.Count -eq 0 -and $State.Jobs.Count -eq 0) { $State.On = $false; $State.State = 'off (NoWorker)' }
            while ($next -lt $plan.Download.Count -and (-not $State.On -or $State.Free.Count -gt 0)) {
                $file = $plan.Download[$next]; $next++
                $ref = Get-FieldValue $file 'FileRef'
                Write-ScanBar $Queue ('File: ' + (Get-ShownItem $Group.Name $ref.Substring($Group.Prefix.Length)))
                # '#' and '%' break the client call's address (as for the stream): such a file takes the main thread's download.
                if (-not $State.On -or $ref.IndexOfAny([char[]]'#%') -ge 0) {
                    $first = $Queue.Rows.Count
                    # The Busy clock pauses for its download only, not while it goes to the readers (review of PR 115).
                    if (-not (Invoke-OneFileScan $Queue $file $Group.Name $Group.Prefix $Origin $MaxFileMB -Pool $State)) { & $count $first }
                    continue
                }
                $first = $Queue.Rows.Count
                $prepared = Start-FileScan $Queue $file $Group.Name $Group.Prefix $Origin $MaxFileMB
                if ($prepared) { Start-DownloadJob $State $prepared; $started = $true } else { & $count $first }
            }
            if ($State.Jobs.Count -gt 0) { foreach ($done in @(Receive-DownloadResult $State)) { Complete-Download $Queue $State $done } }
        }
    }
    finally { if ($started) { Write-Progress -Id 2 -Activity $script:DownloadStep -Completed } }
    foreach ($file in $plan.Other) {
        Write-ScanBar $Queue ('File: ' + (Get-ShownItem $Group.Name (Get-FieldValue $file 'FileRef').Substring($Group.Prefix.Length)))
        $first = $Queue.Rows.Count
        if (-not (Invoke-OneFileScan $Queue $file $Group.Name $Group.Prefix $Origin $MaxFileMB)) { & $count $first }
    }
}

function Close-DownloadPool {
    <#
    .SYNOPSIS
      Stops the downloads that still run (after a cancel or an error) without waiting, closes the pool and the copies
      of the connection. The one sign-out of the run comes after (Find-ChannelPII.ps1).
    .PARAMETER State
      Output of New-DownloadPool.
    #>
    param([hashtable]$State)
    $running = 0
    foreach ($job in $State.Jobs.ToArray()) {
        if ($job.Handle.IsCompleted) { $job.Shell.Dispose() }
        else { $running++; try { [void]$job.Shell.BeginStop($null, $null) } catch { Write-Verbose 'The download had already stopped.' } }
    }
    $State.Jobs.Clear()
    $State.Busy.Stop()
    if ($State.Pool) {
        if ($running + $State.Stuck -gt 0) { [void]$State.Pool.BeginClose($null, $null) }   # it closes when the stuck workers end
        else { $State.Pool.Close(); $State.Pool.Dispose() }
        $State.Pool = $null
    }
    foreach ($copy in @($State.Copies.Values)) {
        if ($copy -is [System.IDisposable]) { try { $copy.Dispose() } catch { Write-Verbose 'A copy of the connection was already closed.' } }
    }
    $State.Copies = @{}
    $State.On = $false
}

function Get-ScanTiming {
    # The time of the last file scan spent downloading (wall time with a download in flight) and spent reading, as
    # minutes:seconds text, and its parallel downloads ("2" or "3", or "off (<reason>)": New-DownloadPool names the reasons).
    $timing = [ordered]@{ 'Time downloading' = (Format-Duration $script:Timing.Download); 'Time reading files (all readers)' = (Format-Duration $script:Timing.Check) }
    if ($script:ParallelState) { $timing['Parallel downloads'] = $script:ParallelState }
    return $timing
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
    $fields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified', 'UniqueId', 'Created', 'Author', 'Editor')
    $allItems = @(Get-PnPListItem -List $Library.List -PageSize $script:ListPageSize -Fields $fields -ErrorAction Stop)
    $byFolder = Get-FilesByFolder $allItems $Library.ServerRelativeUrl
    foreach ($folder in $Folders) {
        $files = @(); if ($byFolder.ContainsKey($folder)) { $files = $byFolder[$folder].ToArray() }
        [pscustomobject]@{ Name = $Label + $folder; Prefix = ($Library.ServerRelativeUrl + '/' + $folder + '/'); Files = $files }
    }
    $rootFiles = @(); if ($byFolder.ContainsKey('')) { $rootFiles = $byFolder[''].ToArray() }
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
    .PARAMETER Progress
      Optional; gets Meter, Reading, Workers and Declined (Invoke-ScanRun), so a scan stopped part way can still say how
      many files were read, and a "no" to "Scan now?" saves nothing.
    .OUTPUTS
      [int] the number of files whose text was scanned.
    #>
    param($Library, [string[]]$Channels, [string]$SiteUrl, [int]$MaxFileMB, $Rows, [hashtable]$Init, [hashtable]$Progress)
    $origin = ([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $script:Timing = @{ Download = [TimeSpan]::Zero; Check = [TimeSpan]::Zero }
    $script:DirectDownload = $true   # tried again on each site
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
    # Starts one file of a group (Invoke-GroupScan calls it).
    $start = { param($queue, $group, $file) Invoke-OneFileScan $queue $file $group.Name $group.Prefix $origin $MaxFileMB }
    # The file's masked path inside its folder, for the progress bar.
    $shown = { param($group, $file) Get-ShownItem $group.Name ((Get-FieldValue $file 'FileRef').Substring($group.Prefix.Length)) }
    $run = $Progress; if (-not $run) { $run = @{} }
    # A file's name and size, for the question before the scan.
    $facts = { param($file) $size = 0L; [void][long]::TryParse((Get-FieldValue $file 'File_x0020_Size'), [ref]$size); (Get-FieldValue $file 'FileLeafRef'), $size }
    # Parallel downloads in a planned order (the Administrator, 2026-11-24); one file at a time when the pool is off.
    $downloads = New-DownloadPool $SiteUrl
    $script:ParallelState = $downloads.State
    # Starts the files of one folder (Invoke-GroupScan calls it).
    $run.GroupScan = { param($queue, $group) Invoke-ParallelGroup $queue $group $downloads $origin $MaxFileMB }
    try { Invoke-ScanRun $Init $Rows $groups $start $shown $run $facts $MaxFileMB }
    finally {
        if ($run.ContainsKey('Reading')) { $script:Timing.Check = $run.Reading }
        $script:Timing.Download += $downloads.Busy.Elapsed
        $script:ParallelState = $downloads.State
        Close-DownloadPool $downloads
    }
    return $run.Meter.Read
}

# =====================================================================================
# Team or plain site
# =====================================================================================

function Test-TeamSite {
    <#
    .SYNOPSIS
      True when the connected site belongs to a team (a Microsoft 365 group): its folders are then channels.
      The report names the folder column "Channel" for a team and "Folder" for a plain site (the Administrator, 2026-11-20).
    .OUTPUTS
      [bool]
    #>
    $groupId = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    return [bool]($groupId -and $groupId -ne [guid]::Empty)
}

Export-ModuleMember -Function Test-TeamSiteUrl, Get-ScanTarget, Read-ScanTarget, ConvertTo-FriendlyError, Connect-TeamSite, Read-ScanTargetList, Get-TargetLabel, Get-SiteDocumentLibraries,
    Get-ChannelFolderNames, Invoke-FileScan, Get-ScanTiming, Test-TeamSite, Get-ShownItem, Get-FileBytes, Get-FieldValue, Get-SafeLink
