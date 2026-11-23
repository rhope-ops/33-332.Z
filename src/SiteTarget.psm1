<#
SiteTarget.psm1 - what the person gives Ion to scan, and the sign-in to that one site.

  Purpose:     Reads a Teams link (a team's or a channel's), a SharePoint site address or a drive folder; turns a
               Teams link alone into its team's site (a lookup at the tenant root, about that one group); signs in
               to the site with the person's own CAC and checks that it belongs to the team in the link.
  Inputs:      the text the person pastes.
  Outputs:     a scan target (Get-ScanTarget) and the signed-in site address (Connect-TeamSite).
  Depends on:  PnP.PowerShell 1.12.0; PiiMasking.psm1 (Protect-PiiText), Report.psm1 (Write-Ui), imported first.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     the tenant root (one group-site lookup for a Teams link) and the one site; nothing lists other
               sites, teams or users ("No reconnaissance").
  Writes:      none.
  Credentials: the person's own CAC sign-in (PnP web login). Nothing is stored.
#>

Set-StrictMode -Version 2.0

# What Ion can scan, for the first question and after a wrong paste (the PAM, 2026-11-21: one short block, not one long line).
$script:TargetHelp = "Paste one of these, or press Enter to exit:`n  - a Teams link: in Teams, ... next to the team or channel > Copy link`n" +
    "  - a SharePoint site address: https://usaf.dps.mil/teams/<team-site> (or /sites/)`n  - a drive folder: X:\<folder> or \\<server>\<share>\<folder>"

# A drive folder: X:\<folder> (a mapped drive) or \\<server>\<share>\<folder> (UNC), with or without quotes.
$script:DrivePathPattern = '^"?(?:[A-Za-z]:\\|\\\\[^\\/\s]+\\[^\\/]+)'

# Accepted site address: https://<host>/sites/<name> or /teams/<name>, also as a file's "Copy link"
# (https://<host>/:w:/r/teams/<name>/...). Anything after the site name is cut off, so the scan can only
# touch that one site. Other sharing links (/:w:/s/, /:w:/t/) are refused: they do not name the site.
$script:SiteUrlPattern = '^(?<host>https://[A-Za-z0-9.-]+)/(?::[a-z]:/r/)?(?<site>(?:sites|teams)/[^/?#\s]+)'

# The tenant root of DoD365: a Teams link alone is turned into its team's site here (GroupSiteManager GetSiteStatus, a
# read about the one group in the link; the PAM's test 2 of 2026-11-21 works). Only a site on this host is used.
# GetSiteStatus reports whether the group's site exists (0 not found, 1 being made, 2 ready); it is not the call that
# creates one (CreateGroupEx). Microsoft documents no write for it; not proven on DoD365 (review of PR 49).
$script:TenantRoot = 'https://usaf.dps.mil'

# The longest one SharePoint call may wait for its answer (milliseconds), set on the signed-in connection when it has
# no limit or a longer one (the PnP 1.12.0 default is not verified here): a download, the file list, a version read.
# Without it a stalled call can wait without end, and Ctrl+C cannot stop a call in progress (the PAM's gov run,
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
    Write-Ui '  A sign-in window opens now, and a second one for the team site after it (if you do not see one, press Alt+Tab).' 'Muted'
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

function Connect-TeamSite {
    <#
    .SYNOPSIS
      Signs in (own CAC, PnP web login) to the site and returns its address. For a Teams link (a team's or a
      channel's) it looks the site up at the tenant root (Find-TeamSiteByGroup), and asks for the site address only
      when that fails; it then checks that the site belongs to the team in the link. It then checks that the person can open the site. On any error or cancel it signs out
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
            $site = Find-TeamSiteByGroup $Target.GroupId
            if (-not $site) {
                Write-Ui "The team's site could not be looked up: the script needs its address. In Teams: the team's Shared tab > ... > Open in SharePoint, then copy the address." 'Accent'
                $site = Read-TeamSiteAddress
            }
        }
        Write-Ui "Signing in to $site ..." 'Muted'
        Write-Ui '  A sign-in window opens (if you do not see it, press Alt+Tab). If it says you have no access, close it.' 'Muted'
        # PnP prints a warning that suggests -Interactive (and its update notice): not useful to the person, so hidden (the PAM, 2026-11-20).
        Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
        Set-SharePointTimeout
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

Export-ModuleMember -Function Test-TeamSiteUrl, Get-ScanTarget, Read-ScanTarget, ConvertTo-FriendlyError, Connect-TeamSite
