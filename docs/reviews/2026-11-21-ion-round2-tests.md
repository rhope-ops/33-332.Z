# Ion round 2 tests: parallel downloads and view counts (2026-11-21)

The PAM approved this plan on 2026-11-21 (docs/DECISIONS.md, newest row). Step 1 asks two things of DoD365 before Ion is changed:

- **Test 3, parallel downloads:** can a copy of your sign-in download a file in a worker (a PowerShell runspace), and do 3 workers at once go faster than one after another? If yes, Ion downloads 2 to 4 files at once (step 3 of the plan).
- **Test 4, view counts:** does SharePoint give the view count of one file on usaf.dps.mil? If yes, Ion can count the views after the "PII first saved" date (SharePoint keeps about the last 90 days).

## What they do

- Read-only. Each test signs in to your own team's site with your CAC, as Ion does, reads one file you name, and signs out.
- Each test touches that one site and that one file: nothing is listed ("No reconnaissance"). Test 4 makes no Microsoft Graph call (commercial Graph is outside the data boundary): its addresses are on usaf.dps.mil only.
- Each prints only result lines: "works" with sizes, times or a count, "no answer", or "failed" with an error type name. It never prints an address, a file name or an error message, so the lines are safe to tell the agent.
- If a test hangs for more than a minute, press the red Stop button, close PowerShell ISE (this ends the sign-in), and report "hung".

## How to run them

1. Open Windows PowerShell ISE on the gov device.
2. Copy one whole block below, paste it into the script pane, and press F5.
3. Answer 1: your team site address. Answer 2: the path of one file of 1 to 20 MB on that site, as it shows in the browser's address bar after the site name, for example `Shared Documents/General/Report.pdf`.
4. Tell the agent the "Result 3" or "Result 4" lines.

### Test 3: parallel downloads

```powershell
# Ion test 3: a copy of the sign-in in workers. Read-only, one site, one file. Prints only sizes, times and error type names.
$typed = (Read-Host 'Your team site address (usaf.dps.mil/sites/<name> or /teams/<name>)').Trim().Trim('"')
if ($typed -notmatch '^https?://') { $typed = 'https://' + $typed }
$siteMatch = [regex]::Match($typed, '^https://usaf\.dps\.mil/((?:sites|teams)/[^/?#\s]+)')
if (-not $siteMatch.Success) { Write-Host 'Answer 1 is not a DoD365 site address: run the test again.'; return }
$site = 'https://usaf.dps.mil/' + $siteMatch.Groups[1].Value
$relative = [uri]::UnescapeDataString((Read-Host 'The file path after the site (for example Shared Documents/General/Report.pdf)').Trim().Trim('"', '/'))
$path = '/' + $siteMatch.Groups[1].Value + '/' + $relative
$pool = $null
try {
    Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop
    $context = Get-PnPContext
    # One download on the main connection, as Ion does now: the size and time to compare with.
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $one = (Get-PnPFile -Url $path -AsMemoryStream -ErrorAction Stop).Length
    $mainSeconds = $clock.Elapsed.TotalSeconds
    Write-Host ('Result 3a: main download works ({0:N1} MB in {1:N1} s)' -f ($one / 1MB), $mainSeconds)
    # A copy of the signed-in connection (PnP.Framework ClientContext.Clone keeps the sign-in), used in 3 workers at once.
    $clone = { param($ctx, $url) [Microsoft.SharePoint.Client.ClientContextExtensions]::Clone($ctx, $url) }
    $copies = @(1..3 | ForEach-Object { & $clone $context $site })
    $worker = {
        param($copy, $filePath)
        $watch = [Diagnostics.Stopwatch]::StartNew()
        try {
            $stream = $copy.Web.GetFileByServerRelativeUrl($filePath).OpenBinaryStream()
            $copy.ExecuteQuery()
            $memory = New-Object System.IO.MemoryStream
            $stream.Value.CopyTo($memory)
            '{0}|{1:N1}' -f $memory.Length, $watch.Elapsed.TotalSeconds
        }
        catch { 'failed|' + $_.Exception.GetType().Name }
    }
    $pool = [runspacefactory]::CreateRunspacePool(1, 3); $pool.Open()
    $clock.Restart()
    $jobs = foreach ($copy in $copies) {
        $shell = [powershell]::Create(); $shell.RunspacePool = $pool
        [void]$shell.AddScript($worker).AddArgument($copy).AddArgument($path)
        @{ Shell = $shell; Handle = $shell.BeginInvoke() }
    }
    $answers = foreach ($job in $jobs) { [string]($job.Shell.EndInvoke($job.Handle) | Select-Object -First 1); $job.Shell.Dispose() }
    $allSeconds = $clock.Elapsed.TotalSeconds
    $same = @($answers | Where-Object { $_ -like "$one|*" }).Count
    if ($same -eq 3) { Write-Host ('Result 3b: 3 workers at once work (3 copies in {0:N1} s; one alone took {1:N1} s)' -f $allSeconds, $mainSeconds) }
    else { Write-Host ('Result 3b: workers failed ({0})' -f (($answers | ForEach-Object { ($_ -split '\|')[-1] }) -join ', ')) }
}
catch { Write-Host ('Result 3: failed ({0})' -f $_.Exception.GetType().Name) }
finally {
    if ($pool) { $pool.Close(); $pool.Dispose() }
    try { Disconnect-PnPOnline -ErrorAction Stop } catch { Write-Host 'Signed out already.' }
}
```

### Test 4: view counts

```powershell
# Ion test 4: the view count of one file, on usaf.dps.mil only. Read-only, one site, one file. Prints only counts and error type names.
$typed = (Read-Host 'Your team site address (usaf.dps.mil/sites/<name> or /teams/<name>)').Trim().Trim('"')
if ($typed -notmatch '^https?://') { $typed = 'https://' + $typed }
$siteMatch = [regex]::Match($typed, '^https://usaf\.dps\.mil/((?:sites|teams)/[^/?#\s]+)')
if (-not $siteMatch.Success) { Write-Host 'Answer 1 is not a DoD365 site address: run the test again.'; return }
$site = 'https://usaf.dps.mil/' + $siteMatch.Groups[1].Value
$relative = [uri]::UnescapeDataString((Read-Host 'The file path after the site (for example Shared Documents/General/Report.pdf)').Trim().Trim('"', '/'))
$path = '/' + $siteMatch.Groups[1].Value + '/' + $relative
try {
    Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop
    $item = Get-PnPFile -Url $path -AsListItem -ErrorAction Stop
    $list = Get-PnPProperty -ClientObject $item -Property ParentList
    $siteId = (Get-PnPSite -Includes Id).Id; $webId = (Get-PnPWeb -Includes Id).Id
    $address = '/_api/v2.1/sites/usaf.dps.mil,{0},{1}/lists/{2}/items/{3}' -f $siteId, $webId, $list.Id, $item.Id
    # Method 1: all-time and last-7-days counts (the SharePoint item analytics, as Graph's itemAnalytics).
    try {
        $all = Invoke-PnPSPRestMethod -Method Get -Url ($address + '/analytics/allTime') -ErrorAction Stop
        $count = ''; if ($all -and $all.PSObject.Properties['access']) { $count = [string]$all.access.actionCount }
        if ($count -ne '') { Write-Host ('Result 4a: works ({0} views, all time)' -f $count) } else { Write-Host 'Result 4a: no answer' }
    }
    catch { Write-Host ('Result 4a: failed ({0})' -f $_.Exception.GetType().Name) }
    # Method 2: views per month from a start date (what Ion needs: the views after the "PII first saved" date).
    try {
        $start = (Get-Date).AddDays(-89).ToString('yyyy-MM-dd'); $end = (Get-Date).ToString('yyyy-MM-dd')
        $url = "$address/getActivitiesByInterval(startDateTime='$start',endDateTime='$end',interval='month')"
        $months = Invoke-PnPSPRestMethod -Method Get -Url $url -ErrorAction Stop
        $total = 0; $rows = 0
        if ($months -and $months.PSObject.Properties['value']) {
            foreach ($month in @($months.value)) { $rows++; if ($month.access) { $total += [int]$month.access.actionCount } }
        }
        if ($rows -gt 0) { Write-Host ('Result 4b: works ({0} views in the last 90 days, {1} months)' -f $total, $rows) } else { Write-Host 'Result 4b: no answer' }
    }
    catch { Write-Host ('Result 4b: failed ({0})' -f $_.Exception.GetType().Name) }
}
catch { Write-Host ('Result 4: failed ({0})' -f $_.Exception.GetType().Name) }
finally {
    try { Disconnect-PnPOnline -ErrorAction Stop } catch { Write-Host 'Signed out already.' }
}
```

## What the results mean

| Result | Meaning | Next step |
|---|---|---|
| 3a and 3b work | A copy of the sign-in downloads in workers. The two times show whether 3 at once beat one by one. | Step 3: 2 to 4 parallel downloads, largest first. |
| 3a works, 3b failed | The copy of the sign-in does not work in a worker. | Step 3 needs another way (the agent proposes one); downloads stay one by one. |
| 4a or 4b works | SharePoint on DoD365 gives view counts. | Ion adds a "Views since PII first saved" column (last 90 days at most). |
| 4a and 4b failed or no answer | No view counts on DoD365 through usaf.dps.mil. | No views column. |

## Results

- Waiting on the PAM.
