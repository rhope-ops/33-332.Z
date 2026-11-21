# Teams-link lookup test (2026-11-21)

The PAM approved this test in the question dialog of 2026-11-21. Goal: find out whether Ion can turn a Teams link into its team's site address on DoD365 without the tenant-root sign-in that hung on 2026-11-20 (docs/DECISIONS.md). If one method works, Ion can accept a Teams link alone, and the person no longer types the site address.

## What it does

- Read-only. It signs in to your own team's site (the same CAC sign-in as Ion), asks SharePoint or Microsoft Graph for the one group in the link, then signs out.
- It touches one site and asks about one group: nothing is listed ("No reconnaissance").
- It prints only "works", "no answer" or the kind of error (an error type name and an HTTP status number). It never prints an address, a name or an error message, so the result is safe to tell the agent.
- If a step hangs for more than a minute, press the red Stop button and report "hung".

## How to run it

1. Open Windows PowerShell ISE on the gov device.
2. Copy the whole block below, paste it into the script pane, and press F5.
3. Paste your team's site address when asked, then the same team's Teams link (in Teams: ... next to the team name > Copy link).
4. Tell the agent the two result lines, A and B.

```powershell
# Ion Teams-link lookup test: read-only. Prints only "works", "no answer" or the kind of error.
$site = (Read-Host 'Your team site address (https://usaf.dps.mil/teams/<team-site>)').Trim().TrimEnd('/')
$link = Read-Host "The same team's Teams link (in Teams: ... > Copy link)"
$groupId = [regex]::Match($link, 'groupId=([0-9a-fA-F-]{36})').Groups[1].Value
if (-not $groupId) { Write-Host 'No groupId in that link: copy the team link again.'; return }
Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
function Show-Result([string]$Name, [scriptblock]$Test) {
    try {
        $url = [string](& $Test)
        if (-not $url) { Write-Host "$Name : no answer" }
        elseif ($url.TrimEnd('/') -ieq $site) { Write-Host "$Name : works (it gives this team's site)" }
        else { Write-Host "$Name : works, but it gives another address" }
    }
    catch {
        $code = ''
        if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        Write-Host ("$Name : failed ({0} {1})" -f $_.Exception.GetType().Name, $code)
    }
}
Show-Result 'A (SharePoint, from your site)' {
    (Invoke-PnPSPRestMethod -Method Post -Url '/_api/GroupSiteManager/GetSiteStatus' -Content @{ groupId = $groupId }).SiteUrl
}
Show-Result 'B (Microsoft Graph)' { (Invoke-PnPGraphMethod -Url ('groups/' + $groupId + '/sites/root?$select=webUrl')).webUrl }
Disconnect-PnPOnline
```

## What the results mean

| Result | Meaning | Next step |
|---|---|---|
| A works (this team's site) | SharePoint can look up a team's site once Ion is signed in to any site the person can open. | Ion still needs one site to sign in to first (the tenant root hung). The agent proposes options for the PAM to choose from, for example the site of the person's last scan. |
| A fails, B works | Graph can do it with the web sign-in. | Same as above, through Graph. |
| Both fail or hang | Neither method works on DoD365 with a CAC web sign-in. | Ion keeps asking for the site address. B is expected to fail: the web sign-in gives SharePoint cookies, not a Graph token (Confidence: medium). |

Not verified here: whether `Invoke-PnPGraphMethod` exists in PnP.PowerShell 1.12.0. If it does not, B prints "failed (CommandNotFoundException)", which also means "Graph is not available this way".
