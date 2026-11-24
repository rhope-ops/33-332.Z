<#
.SYNOPSIS
  Find-ChannelPII.ps1 - Ion, DoD365 PII Reporter: read-only check of one Teams team, SharePoint site or drive folder for possible PII.

.DESCRIPTION
  Scans the files of a team's or site's document libraries: every folder (in a team, one per standard channel) and
  the files at the top of each library; or every file in one drive folder and its subfolders. Two progress bars (the whole scan, and the file in work)
  and one line per folder show the counts as it goes. At the end it shows the hits per PII type on screen and
  saves one Excel workbook: a Summary tab first, then Hits (one Excel Table, every hit) and Files, then
  Not scanned. Every value found is MASKED on screen and in the workbook.

  Purpose:     Help a team or site member (owner rights are not needed) find files that may hold PII, so they can
               review them. A spot check, not a replacement for Microsoft Purview DLP or Content Search.
  Inputs:      the Teams link, SharePoint site address or drive folder (X:\<folder>, \\<server>\<share>\<folder>), asked when not given.
  Outputs:     progress and the summary on screen; <OutputFolder>\<site>\Ion <site> <yyyy-MM-dd HHmm>.xlsx.
  How to run:  .\Find-ChannelPII.ps1   (README.md "Run it" gives every step)
  Run target:  Windows PowerShell 5.1 (also PowerShell 7) with PnP.PowerShell 1.12.0. No other modules.
  Network:     only the one SharePoint site, or the one drive folder's server, given; for a Teams link also the tenant
               root, for one lookup of that team's site (SiteTarget.psm1). Channel posts are not read (the PAM, 2026-11-20).
  Writes:      only the workbook in -OutputFolder (default Documents\Ion); nothing in Teams or SharePoint.
  Credentials: the person's own CAC sign-in in the PnP sign-in window (a drive folder: the Windows logon); nothing is stored.
               One sign-in per host: no sign-out between sites; one sign-out at the end of the run, also after an
               error or Ctrl+C (the PAM, 2026-11-24).

  Policy:      Read-only; one site only (no network reconnaissance); file content stays in memory; values
               masked (the PAM, 2026-11-20). No script file is collected or sent: script types are not
               read and only their (masked) names appear in the report.
  Not scanned: encrypted or password-protected files, script and other unread file types, files over -MaxFileMB, and
               images where Windows OCR is not there (each a NOT SCANNED row with its reason); channel posts; private
               and shared channels (each has its own site: run again with that site's address); 1:1, group and meeting chats.

.PARAMETER SiteUrl
  The Teams link or the SharePoint site address.
  Teams link: in Teams, open the team or one of its channels, click the ... next to its name > Copy link
  (https://<teams-host>/l/team/...?groupId=...). The script asks for the team's site address and checks the site belongs to that team.
  Site address: in Teams, the Shared tab > ... > Open in SharePoint (https://usaf.dps.mil/teams/<team-site> or /sites/);
  a file's Copy link works too. Anything after the site name is cut off.
.PARAMETER OutputFolder
  Where to save the results: a folder per site, then a folder per day. Default: Documents\Ion in your profile.
.PARAMETER MaxFileMB
  Files larger than this are reported as NOT SCANNED. Default 25.

.EXAMPLE
  .\Find-ChannelPII.ps1
  Asks for the Teams link or site address, then scans every folder's files.
.EXAMPLE
  .\Find-ChannelPII.ps1 -SiteUrl https://usaf.dps.mil/teams/<team-site>
#>
[CmdletBinding()]
param(
    [string]$SiteUrl,
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Ion'),
    # 25 MB: larger Office files are rare in a channel and slow to read in memory. The author's choice, 2026-11-20.
    [int]$MaxFileMB = 25
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$IonBuild = 'src'   # run from src/; the paste bundle sets its build ID here (scripts/build_paste.py)

# Modules, in dependency order (each module's header names what it needs).
# The readers (the reader workers load the same files: ReaderPool.psm1), then the scan and report modules.
$readerModules = @('PiiPatterns.psm1', 'PiiSignedIds.psm1', 'PiiDetectors.psm1', 'PiiMasking.psm1', 'OcrText.psm1', 'OleText.psm1', 'ContainerText.psm1',
    'PdfCore.psm1', 'PdfObjects.psm1', 'PdfFonts.psm1', 'PdfShown.psm1', 'PdfLayout.psm1', 'PdfContent.psm1', 'PdfText.psm1', 'TextExtract.psm1', 'ReaderPool.psm1')
$otherModules = @('XlsxWriter.psm1', 'Report.psm1', 'ReportBook.psm1',
    'ScanMeter.psm1', 'ScanQueue.psm1', 'Accountability.psm1', 'SiteTarget.psm1', 'ChannelScan.psm1', 'DriveScan.psm1')
foreach ($module in $readerModules + $otherModules) {
    Import-Module (Join-Path $PSScriptRoot $module) -Force
}
$IonWorkerInit = @{ Paths = @($readerModules | ForEach-Object { Join-Path $PSScriptRoot $_ }) }

# ==== Body ==== (scripts/build_paste.py copies everything from this line into the paste bundle; the lines above
# become the bundle's Settings block and its inline module code)

function Invoke-SiteFiles {
    <#
    .SYNOPSIS
      Scans one signed-in site and saves the workbook (after an error or cancel too: the results so far say
      INCOMPLETE). It does not sign out: the run signs out once, at its end. Returns the workbook path; $null after no (or a cancel) at "Scan now?": nothing read or saved.
    .PARAMETER SiteUrl
      The signed-in site.
    .PARAMETER OutputFolder
      The base folder for the results.
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER AutoStart
      Do not ask "Scan now?": the person said yes to the whole list of targets.
    .OUTPUTS
      [string] the workbook path.
    #>
    param([string]$SiteUrl, [string]$OutputFolder, [int]$MaxFileMB, [switch]$AutoStart)
    $path = $null; $clock = [Diagnostics.Stopwatch]::StartNew(); $rows = New-Object System.Collections.Generic.List[object]
    $scanned = 0; $selected = @()
    $folderLabel = 'Folder'   # 'Channel' when the site belongs to a team (the PAM, 2026-11-20)
    $listName = ''            # what the header lists: 'Channels', 'Folders' or 'Libraries'
    $progress = @{ AutoStart = [bool]$AutoStart }   # Invoke-FileScan puts its live meter here
    $finished = $false   # set only after the last step; a cancelled run (Ctrl+C) skips 'catch' but not 'finally'
    try {
        # Only the column name depends on it: if the site's group cannot be read, the scan goes on with 'Folder'.
        try { if (Test-TeamSite) { $folderLabel = 'Channel' } }
        catch { Write-Verbose 'Team check not possible: the folder column is named Folder.' }
        $libraries = @(Get-SiteDocumentLibraries)
        if ($libraries.Count -eq 1) { $selected = @(Get-ChannelFolderNames $libraries[0]) }
        else { $folderLabel = 'Library / Folder'; $selected = @($libraries | ForEach-Object { $_.Title }) }   # every library of the site
        $listName = $folderLabel + 's'; if ($libraries.Count -gt 1) { $listName = 'Libraries' }
        Write-Ui ($listName + ' to scan: ') 'Label' -NoNewline
        Write-Ui ((($selected | ForEach-Object { Protect-PiiName $_ }) -join ', ') + '; also any files at the top of each library') 'Text'
        $scanned += Invoke-FileScan -Library $libraries -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows -Init $IonWorkerInit -Progress $progress
        $finished = $true
    }
    catch {
        # Keep what was found so far: the results say the scan stopped early, and why, in words a person can act on.
        $reason = Protect-PiiText $_.Exception.Message
        Write-Ui ('The scan stopped early: ' + $reason) 'Alert'
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
    }
    finally {
        if (-not $finished -and $progress.ContainsKey('Meter')) { $scanned = $progress.Meter.Read }   # stopped part way: the live count
        Add-IncompleteRow -Rows $rows -Finished $finished
        if ($progress.ContainsKey('Declined') -and $progress.Declined) { Write-Ui 'Not scanned: nothing was saved.' 'Accent' }   # "No" to "Scan now?"
        else {
            $header = Get-SiteHeader $SiteUrl $listName $folderLabel $selected $clock.Elapsed
            $path = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $scanned -Header $header -OutputFolder $OutputFolder -SiteUrl $SiteUrl -FolderLabel $folderLabel
        }
    }
    return $path
}

function Get-SiteHeader {
    <#
    .SYNOPSIS
      The top of a site scan's summary: the site, what was scanned (as Channels, Folders or Libraries), when, how long.
    .PARAMETER SiteUrl
      The site.
    .PARAMETER ListName
      The label of the scanned list ('Channels', 'Folders' or 'Libraries'); empty when the scan stopped before it.
    .PARAMETER FolderLabel
      'Channel', 'Folder' or 'Library / Folder' (the label when ListName is empty).
    .PARAMETER Selected
      The folders or libraries scanned.
    .PARAMETER Used
      The time the scan took.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$SiteUrl, [string]$ListName, [string]$FolderLabel, [string[]]$Selected, [TimeSpan]$Used)
    if (-not $ListName) { $ListName = $FolderLabel + 's' }   # stopped before the libraries were read
    $header = [ordered]@{ Site = (Protect-PiiName $SiteUrl) }   # masked like any name Ion shows (review of PR 80)
    $header[$ListName] = ($Selected | ForEach-Object { Protect-PiiName $_ }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = Format-Duration $Used
    foreach ($pair in (Get-ScanTiming).GetEnumerator()) { $header[$pair.Key] = $pair.Value }
    return $header
}

function Open-ScanResult {
    <#
    .SYNOPSIS
      Asks whether to open the saved workbook now (Enter = yes; the PAM, 2026-11-20) and opens it in Excel.
      Only the local file is opened; nothing is sent.
    .PARAMETER Path
      The workbook path, or empty when nothing was saved.
    #>
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return }
    $answer = Read-Host 'Open the workbook now? [Y/n]'
    if ($answer -and $answer.Trim() -match '^(n|no)$') { return }
    try { Invoke-Item -LiteralPath $Path }
    catch { Write-Ui ('The workbook could not be opened: ' + (Protect-PiiText $_.Exception.Message) + '. Open it from the folder above.') 'Accent' }
}

function Invoke-ScanTarget {
    <#
    .SYNOPSIS
      Scans one target (a drive folder, or a site after its sign-in) and saves its workbook. With List: no "Scan now?"
      and no "Open the workbook now?" (asked once for the whole list). Never throws: a failure is the result's Note.
    .PARAMETER Target
      Output of Get-ScanTarget.
    .PARAMETER List
      Part of a list of several targets.
    .OUTPUTS
      [pscustomobject] Label, Path (the workbook, or empty), Note (why there is none), Time.
    #>
    param($Target, [switch]$List)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $result = [pscustomobject]@{ Label = (Get-TargetLabel $Target); Path = ''; Note = ''; Time = '' }
    $saving = $true   # false only during the sign-in, whose errors are already in words a person can act on
    try {
        if ($Target.DrivePath) {
            # A drive or network folder: no sign-in (the Windows logon); read-only; nothing above the folder.
            $result.Path = Invoke-DriveScan -Path $Target.DrivePath -OutputFolder $OutputFolder -MaxFileMB $MaxFileMB -Init $IonWorkerInit -AutoStart:$List
        }
        else {
            $script:IonSignInTried = $true   # set before the sign-in, so a Ctrl+C during it still gets the run's one sign-out
            $saving = $false; $site = Connect-TeamSite $Target; $saving = $true
            $result.Label = Protect-PiiName $site
            $result.Path = Invoke-SiteFiles -SiteUrl $site -OutputFolder $OutputFolder -MaxFileMB $MaxFileMB -AutoStart:$List
        }
        if (-not $result.Path) { $result.Note = 'Not scanned: nothing was saved.' }
    }
    catch {
        # A site that cannot be opened, or results that cannot be saved: reported; the next target still runs.
        # A sign-in error is already in words a person can act on, and masked (Connect-TeamSite): shown as it is, so its
        # numbered steps keep their numbers. Any other error is masked here. The closing table shows the first line.
        $shown = $_.Exception.Message
        if ($saving) { $shown = 'The results could not be saved: ' + (Protect-PiiText $shown) }
        Write-Ui $shown 'Alert'
        $result.Note = @($shown -split "`n" | Where-Object { $_.Trim() })[0]
    }
    $result.Time = Format-Duration $clock.Elapsed
    if (-not $List) { Open-ScanResult $result.Path }
    return $result
}

function Show-ListResult {
    <#
    .SYNOPSIS
      After a list of targets: one line per target (its workbook, or why there is none, and the time), then offers to
      open the results folder (Enter = yes). Only the local folder is opened; nothing is sent.
    .PARAMETER Results
      Outputs of Invoke-ScanTarget, in the order scanned.
    #>
    param([object[]]$Results)
    Write-Banner ('Done: {0} of {1} saved' -f @($Results | Where-Object { $_.Path }).Count, $Results.Count)
    $n = 0
    foreach ($r in $Results) {
        $n++
        if ($r.Path) { Write-Ui ('{0,2}. {1} ({2}): {3}' -f $n, $r.Label, $r.Time, (Split-Path $r.Path -Leaf)) 'Good' }
        else { Write-Ui ('{0,2}. {1} ({2}): {3}' -f $n, $r.Label, $r.Time, $r.Note) 'Alert' }
    }
    $answer = Read-Host 'Open the results folder now? [Y/n]'
    if ($answer -and $answer.Trim() -match '^(n|no)$') { return }
    try { Invoke-Item -LiteralPath $OutputFolder } catch { Write-Ui ('The folder could not be opened: ' + (Protect-PiiText $_.Exception.Message)) 'Accent' }
}

Write-Banner 'Ion - DoD365 PII Reporter' -Mark
Write-Ui ('Build ' + $IonBuild) 'Muted'   # which Ion ran: the PAM copies this line with the results (2026-11-24)
Write-Ui 'Read-only. Every PII value found is masked on screen and in the results.' 'Muted'
# What it scans and how to stop early (the PAM, 2026-11-21): a cancel still saves the results so far ('finally').
Write-Ui ('Scans the files of Teams teams, SharePoint sites or drive folders you paste. Files over {0} MB are listed, not read.' -f $MaxFileMB) 'Muted'
Write-Ui 'To stop early: Ctrl+C (in ISE: the red Stop button). The results so far are still saved.' 'Muted'
$script:IonSignInTried = $false   # true after the first site sign-in is tried: the run then signs out once, at its end

# ---- Ask for one or more targets, then scan each and save its workbook; then offer more, until the person presses Enter ----
# A wrong paste is explained and asked again (Read-ScanTarget); a site the person cannot open is reported as "no access"
# and nothing is saved for it (Connect-TeamSite); only the targets the person pastes are touched ("No reconnaissance").
# One sign-in per host (the PAM, 2026-11-24): each site connects without a sign-out between sites (probe 3, second run);
# the run signs out once here, at its end, also after an error or Ctrl+C ('finally' runs on both).
try {
    $targets = @(Read-ScanTargetList $SiteUrl 'What to scan')
    while ($targets.Count -gt 0) {
        $list = $targets.Count -gt 1
        $go = $true
        if ($list) {   # several: listed and asked once (the PAM, 2026-11-24: one workbook per site)
            Write-Ui ('{0} to scan, one after the other (sign-in windows open only when needed):' -f $targets.Count) 'Label'
            $n = 0; foreach ($t in $targets) { $n++; Write-Ui ('{0,2}. {1}' -f $n, (Get-TargetLabel $t)) 'Text' }
            $answer = Read-Host ('Scan these {0} now? [Y/n]' -f $targets.Count)
            $go = -not ($answer -and $answer.Trim() -match '^(n|no)$')
        }
        if ($go) {
            $results = @(foreach ($t in $targets) { Invoke-ScanTarget $t -List:$list; Write-Ui '' })
            if ($list) { Show-ListResult $results }
        }
        Write-Ui ''
        $targets = @(Read-ScanTargetList '' 'Scan more: paste it, or press Enter to exit' -NoHelp)
    }
}
finally {
    if ($script:IonSignInTried) {
        try { Disconnect-PnPOnline }
        catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error for the person
    }
}
if ($script:IonSignInTried) { Write-Ui 'Ion is done: nothing stays signed in.' 'Good' } else { Write-Ui 'Ion is done.' 'Good' }
