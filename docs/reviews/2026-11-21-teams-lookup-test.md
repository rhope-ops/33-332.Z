# Teams-link lookup test (2026-11-21)

The PAM approved this test in the question dialog of 2026-11-21. Goal: find out whether Ion can turn a Teams link into its team's site address on DoD365 without the tenant-root sign-in that hung on 2026-11-20 (docs/DECISIONS.md). If one method works, Ion can accept a Teams link alone, and the person no longer types the site address.

## What it does

- Read-only. It signs in to your own team's site (the same CAC sign-in as Ion), asks that site's SharePoint for the site of the one group in the link (GroupSiteManager GetSiteStatus, a read), then signs out.
- It accepts only a `https://usaf.dps.mil/teams/...` or `/sites/...` address, touches that one site and asks about one group: nothing is listed ("No reconnaissance"). It sends nothing outside DoD365: there is no Microsoft Graph call (commercial Graph is outside the data boundary; review of PR 38).
- It prints only one result line: "works", "works, but another address", "no answer", or "failed" with an error type name. It never prints an address, a name or an error message, so the line is safe to tell the agent.
- If it hangs for more than a minute, press the red Stop button, close PowerShell ISE (this ends the sign-in), and report "hung".

## How to run it

1. Open Windows PowerShell ISE on the gov device.
2. Copy the whole block below, paste it into the script pane, and press F5.
3. Paste your team's site address when asked, then the same team's Teams link (in Teams: ... next to the team name > Copy link).
4. Tell the agent the result line.

```powershell
# Ion Teams-link lookup test: read-only, one site, one group. Prints one result line, never an address or a message.
$site = (Read-Host 'Your team site address (https://usaf.dps.mil/teams/<team-site>)').Trim().TrimEnd('/')
$link = Read-Host "The same team's Teams link (in Teams: ... > Copy link)"
$groupId = [regex]::Match($link, 'groupId=([0-9a-fA-F-]{36})').Groups[1].Value
if (-not $groupId) { Write-Host 'No groupId in that link: copy the team link again.'; return }
# Only a DoD365 team or site address is used: nothing goes to another host (Data boundary rule).
if ($site -notmatch '^https://usaf\.dps\.mil/(sites|teams)/[^/]+$') { Write-Host 'Not a DoD365 site address: paste it again.'; return }
try {
    Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop
    $answer = Invoke-PnPSPRestMethod -Method Post -Url '/_api/GroupSiteManager/GetSiteStatus' -Content @{ groupId = $groupId } -ErrorAction Stop
    $url = ''
    if ($answer -and $answer.PSObject.Properties['SiteUrl']) { $url = [string]$answer.SiteUrl }
    if (-not $url) { Write-Host 'Result: no answer' }
    elseif ($url.TrimEnd('/') -ieq $site) { Write-Host "Result: works (it gives this team's site)" }
    else { Write-Host 'Result: works, but it gives another address' }
}
catch { Write-Host ('Result: failed ({0})' -f $_.Exception.GetType().Name) }
finally {
    try { Disconnect-PnPOnline -ErrorAction Stop } catch { Write-Host 'Signed out already.' }
}
```

## What the result means

| Result | Meaning | Next step |
|---|---|---|
| works (this team's site) | Once signed in to this team's site, SharePoint gives the site of the group in the link. Whether it does so from any other site is not shown by this test. | Ion still needs one site to sign in to first (the tenant root hung on 2026-11-20). The agent proposes options for the PAM to choose from. |
| works, but another address | The lookup answers, but with an address that differs from the one typed (for example a /sites/ form, or another site). | The agent checks with the PAM before using it. |
| no answer | SharePoint answered without a site address. | Ion keeps asking for the site address. |
| failed (type) | The sign-in or the lookup was refused, or the cmdlet is not there (for example CommandNotFoundException: not in PnP.PowerShell 1.12.0). The test cannot tell these apart. | If Ion's own sign-in works on the same site, the lookup failed: Ion keeps asking for the site address. |
| hung | No answer within a minute. | Ion keeps asking for the site address. |
