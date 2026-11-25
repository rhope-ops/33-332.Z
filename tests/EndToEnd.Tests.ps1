<#
EndToEnd.Tests.ps1 - runs Find-ChannelPII.ps1 from start to end against the fake site.

  Purpose:     The whole script: every folder is scanned (one "Scan now?" question, Enter = yes) with one sign-in, one workbook is
               written, every planted type is found, the progress bars show the counts, a plain site gets a Folder
               column, and a failure part way still saves results that say INCOMPLETE.
  Inputs:      $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      report files in a new temporary folder, deleted at the end of this file.
  Credentials: none.
#>

$script = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'src') 'Find-ChannelPII.ps1'
$runFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-e2e-' + [guid]::NewGuid())
$library = '/sites/TeamA/Shared Documents'

# ---- A full run: one workbook, every planted type, "Scan now?" answered with Enter, the progress bars ----
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
& $script -SiteUrl $Sample.Site -OutputFolder $runFolder 6>$null
$results = @(Get-ResultFiles $runFolder)
Assert-Equal 1 $results.Count 'Script saved one workbook'
# -Include is ignored with -LiteralPath in Windows PowerShell 5.1, so the extension is checked after the listing.
Assert-Equal 0 @(Get-ChildItem -LiteralPath $runFolder -Recurse | Where-Object { $_.Extension -in @('.txt', '.csv') }).Count 'No .txt or .csv any more'
Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $results[0].FullName) 'Script found every planted type'
Assert-Equal 'Channel' (Get-XlsxSheets $results[0].FullName)['Hits'][0][3] 'A team: the folder column is named Channel'
Assert-Equal 1 (Get-FakeCallCount 'Read-Host') 'With the site given, one address question only (another site or exit); "Scan now?" and "Open" are their own'
Assert-Equal $true ((Test-FakePnPCall 'Open question') -and (Test-FakePnPCall ('Open ' + $results[0].FullName))) 'Enter at the open question opens the workbook'
Assert-Equal 1 @(Get-FakePnPConnection).Count 'One sign-in only: no posts sign-in'
Assert-Equal $true ((Get-FakeProgressText) -match '2 of 3 files \(\d+%\) \| read \d+ \| not scanned 0 \| possible PII \d+ in \d+ files \| \d\d:\d\d elapsed') 'Progress bar shows the file counts'
Assert-Equal $true ((Get-FakeProgressText) -match 'Folder 2 of 2: Ops \d+ of 3 files .* File: scan\.pdf') 'Progress bar: the folder and its number, then the file'
Assert-Equal $false ((Get-FakeProgressText).Contains([string][char]0x3B3)) 'The mark is on the start line only, not on the bars (the PAM, 2026-11-21)'
Assert-Equal $false ((Get-FakeProgressText) -match 'Downloading [\d.]+ (of|MB)') 'Small files: no per-file bar (it would only slow the scan)'

Assert-Equal $true (Test-FakePnPCall 'Scan question') 'Before the scan: "Scan now?" is asked'
Assert-Equal 1 (Get-FakeCallCount 'Speed question') 'Speed: asked one time, after the targets'

# ---- "No" to "Scan now?": nothing read, nothing saved, signed out (the PAM, 2026-11-21) ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt; "$library/Ops/scan.exe" = $Txt }
Set-FakeSiteAnswer -ScanAnswer 'n'
$noFolder = Join-Path $runFolder 'declined'
$noShown = (& $script -SiteUrl $Sample.Site -OutputFolder $noFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($noShown -match 'Found 2 files \([\d.]+ MB\) in 2 folders; 1 will be listed but not read') 'Before the scan: files, MB and the ones not read'
Assert-Equal $true ($noShown -match 'Not scanned: nothing was saved') 'Declined: the person is told'
Assert-Equal 0 @(Get-ResultFiles $noFolder).Count 'Declined: nothing saved'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Declined: signed out'
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -ScanAnswer 'cancel'
$cancelFolder = Join-Path $runFolder 'cancelled'
$cancelShown = (& $script -SiteUrl $Sample.Site -OutputFolder $cancelFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal 0 @(Get-ResultFiles $cancelFolder).Count 'Cancel at "Scan now?": nothing read, nothing saved'
Assert-Equal $true ($cancelShown -match 'Not scanned: nothing was saved' -and (Test-FakePnPCall 'Disconnect-PnPOnline')) 'Cancel at "Scan now?": told, and signed out'
# No files at all: no "Scan now?" question.
Reset-FakePnP @{}
& $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $runFolder 'empty') 6>$null
Assert-Equal $false (Test-FakePnPCall 'Scan question') 'No files: no "Scan now?" question'

# ---- A plain SharePoint site (no team): Folder column, top-of-library files ----
Reset-FakePnP @{ "$library/top.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite '' -SiteGroupId ([guid]::Empty)
$siteFolder = Join-Path $runFolder 'site'
& $script -SiteUrl $Sample.Site -OutputFolder $siteFolder 6>$null
$siteBook = Get-XlsxSheets (Get-ResultFiles $siteFolder)[0].FullName
$sitePhone = @($siteBook['Hits'] | Where-Object { $_[1] -eq 'Phone' })[0]
Assert-Equal 'Folder|(top of the library)' (($siteBook['Hits'][0][3], $sitePhone[3]) -join '|') 'A plain site: Folder column; top-of-library file reported'

# ---- The team check fails: the scan goes on, and the column is named Folder ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakePnPOption SiteFails $true
$checkFolder = Join-Path $runFolder 'teamcheck'
& $script -SiteUrl $Sample.Site -OutputFolder $checkFolder 6>$null
$checkBook = Get-XlsxSheets (Get-ResultFiles $checkFolder)[0].FullName
Assert-Equal 'Folder' $checkBook['Hits'][0][3] 'Team check failed: scan went on with the Folder label'
Assert-Equal $false (@($checkBook.Keys) -contains 'Not scanned') 'Team check failed: the scan is not INCOMPLETE'

# ---- No access to the site: a clear message, nothing saved, signed out, another site offered ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakePnPOption NoAccess $true
$noAccessFolder = Join-Path $runFolder 'noaccess'
$shown = (& $script -SiteUrl $Sample.Site -OutputFolder $noAccessFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($shown -match 'No access: your account cannot open this site') 'No access: the person is told'
Assert-Equal 0 @(Get-ResultFiles $noAccessFolder).Count 'No access: nothing saved'
Assert-Equal $true ((Test-FakePnPCall 'Disconnect-PnPOnline') -and (Get-FakeCallCount 'Read-Host') -eq 1) 'No access: signed out, another site offered'

# ---- A wrong paste is explained and asked again; Enter exits ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite @('not a link', $Sample.Site)
$retryFolder = Join-Path $runFolder 'retry'
$shown = (& $script -OutputFolder $retryFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($shown -match "Not read: 'not a link' is not a SharePoint site address") 'Wrong paste: explained'
Assert-Equal $true ($shown -match 'a drive folder: X:') 'Wrong paste: what can be pasted is listed'
$endLine = 'Ion is done: signed out\. Close Windows PowerShell ISE to end the CAC session'
Assert-Equal $true ($shown -match 'To stop early: Ctrl\+C' -and $shown -match $endLine) 'Start: how to stop; end: signed out after a site'
Assert-Equal 1 @(Get-ResultFiles $retryFolder).Count 'Wrong paste: asked again, then the right one scanned'
# Wrong paste, right paste, Enter at "Add another" (starts the one target), Enter at "Scan more" (exits).
Assert-Equal 4 (Get-FakeCallCount 'Read-Host') 'Wrong paste, right paste, then Enter to start and Enter to exit'

# ---- A run that fails part way still saves a report ----
Reset-FakePnP @{}
Set-FakePnPOption ListFails $true
$failFolder = Join-Path $runFolder 'fail'
& $script -SiteUrl $Sample.Site -OutputFolder $failFolder 6>$null
$failBook = @(Get-ResultFiles $failFolder)
Assert-Equal 1 $failBook.Count 'Results saved after a failure'
Assert-Equal 'INCOMPLETE' (Get-XlsxSheets $failBook[0].FullName)['Not scanned'][1][0] 'Results say INCOMPLETE'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out after a failure'

# Stopped part way, after folder 1 was read: the red line counts the files read so far (review of PR 36).
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt; "$library/Ops/a.txt" = $Txt }
Set-FakePnPOption StopInFolder2 $true
$stopShown = (& $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $runFolder 'stop') 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($stopShown -match 'INCOMPLETE: the scan stopped early\. Only the 1 file read') 'Stopped part way: the red line counts the files read'

# ---- Several targets in one run (the PAM, 2026-11-24): asked once, one workbook each, a closing table ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
$siteB = $Sample.Site + 'B'
Set-FakeSiteAnswer -TypedSite @($Sample.Site, $siteB)
$listFolder = Join-Path $runFolder 'list'
$listShown = (& $script -OutputFolder $listFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal 2 @(Get-ResultFiles $listFolder).Count 'Several targets: one workbook each'
Assert-Equal ($Sample.Site + '|' + $siteB) ((Get-FakePnPConnection) -join '|') 'Several targets: each site signed in, in order, and no other'
$questions = '{0} {1} {2} {3}' -f (Get-FakeCallCount 'List question'), (Get-FakeCallCount 'Scan question'), (Get-FakeCallCount 'Open question'), (Get-FakeCallCount 'Folder question')
Assert-Equal '1 0 0 1' $questions 'Several targets: asked once; no "Scan now?" or "Open the workbook?" per site; the folder offered once'
Assert-Equal $true ($listShown -match 'Done: 2 of 2 saved' -and $listShown -match '2 to scan, one after the other') 'Several targets: the list and the closing table'
# One sign-in per host (the PAM, 2026-11-24): no sign-out between the sites; one sign-out, after the last site.
$lastConnect = Get-FakeCallIndex ('Connect ' + $siteB) -Last
Assert-Equal '1 True' ('{0} {1}' -f (Get-FakeCallCount 'Disconnect-PnPOnline'), ((Get-FakeCallIndex 'Disconnect-PnPOnline') -gt $lastConnect)) 'Several targets: one sign-out, after the last site'
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite @($Sample.Site, $siteB) -ScanAnswer 'n'
$noFolder = Join-Path $runFolder 'list-no'
& $script -OutputFolder $noFolder 6>$null
Assert-Equal 0 @(Get-ResultFiles $noFolder).Count 'Several targets: "n" to the list scans nothing'

# Review of PR 80: a site that fails does not stop the list; site labels are masked; "n" signs in nowhere.
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakePnPOption NoAccess $true
Set-FakeSiteAnswer -TypedSite @($Sample.Site, $siteB)
$failFolder2 = Join-Path $runFolder 'list-fail'
$failShown = (& $script -OutputFolder $failFolder2 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($failShown -match 'Done: 0 of 2 saved' -and @(Get-FakePnPConnection).Count -eq 2) 'Several targets: a site with no access does not stop the next one'
Assert-Equal 1 (Get-FakeCallCount 'Disconnect-PnPOnline') 'Several targets, no access: one sign-out, at the end of the run'
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
$ssnSite = $Sample.Site + '_' + $Sample.Ssn
Set-FakeSiteAnswer -TypedSite @($ssnSite, $siteB) -ScanAnswer 'n'
$maskShown = (& $script -OutputFolder (Join-Path $runFolder 'list-mask') 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $false $maskShown.Contains($Sample.Ssn) 'Several targets: site labels are masked'
Assert-Equal 0 @(Get-FakePnPConnection).Count 'Several targets: "n" to the list signs in nowhere'

# Second review of PR 80: a blocked sign-in keeps its numbered steps once per site, the table shows one line each;
# a site name with an SSN shape is masked in "Signing in to", the results header, the workbook Summary and the table.
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakePnPOption ConnectFails $true
Set-FakeSiteAnswer -TypedSite @($Sample.Site, $siteB)
$blockShown = (& $script -OutputFolder (Join-Path $runFolder 'list-block') 6>&1 | ForEach-Object { [string]$_ }) -join "`n"
Assert-Equal 2 ([regex]::Matches($blockShown, '(?m)^\s+1\. In Teams, click the channel')).Count 'Blocked sign-in: the numbered steps show once per site, with their numbers'
Assert-Equal 0 ([regex]::Matches($blockShown, '#\. In Teams')).Count 'Blocked sign-in: no step number is masked'
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite @($ssnSite, $siteB)
$ssnFolder = Join-Path $runFolder 'list-ssn'
$ssnShown = (& $script -OutputFolder $ssnFolder 6>&1 | ForEach-Object { [string]$_ }) -join ' '
$ssnBooks = @(Get-ResultFiles $ssnFolder)
$ssnCells = (@($ssnBooks | ForEach-Object { (Get-XlsxSheets $_.FullName)['Summary'] | ForEach-Object { $_ -join ' ' } })) -join ' '
Assert-Equal '2 False False' ('{0} {1} {2}' -f $ssnBooks.Count, $ssnShown.Contains($Sample.Ssn), $ssnCells.Contains($Sample.Ssn)) 'SSN-shaped site name: masked on screen and in the workbooks'

# ---- One sign-in per host across "Scan more" rounds, and the sign-out at a cancel (the PAM, 2026-11-24) ----
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite @($siteB)
$roundFolder = Join-Path $runFolder 'rounds'
& $script -SiteUrl $Sample.Site -OutputFolder $roundFolder 6>$null
$signOutLast = (Get-FakeCallIndex 'Disconnect-PnPOnline') -gt (Get-FakeCallIndex ('Connect ' + $siteB))
$roundOrder = '{0}|{1}|{2}' -f ((Get-FakePnPConnection) -join ' '), (Get-FakeCallCount 'Disconnect-PnPOnline'), $signOutLast
Assert-Equal ($Sample.Site + ' ' + $siteB + '|1|True') $roundOrder 'Scan more: no sign-out between rounds; one sign-out at the end'
Assert-Equal 2 @(Get-ResultFiles $roundFolder).Count 'Scan more: one workbook per round'
Assert-Equal 1 (Get-FakeCallCount 'Speed question') 'Scan more: the speed is asked one time per run, not per round'
# The answer f gives Fast: the reader pool may use one reader per processor.
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSpeedAnswer ' F '
& $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $runFolder 'fast') 6>$null
Assert-Equal 'Fast' (Get-ScanSpeed) 'Speed: the answer f (any case, spaces) gives Fast'
$fastPool = New-ReaderPool @{}
try { Assert-Equal (Get-ReaderLimit 'Fast' ([Environment]::ProcessorCount)) $fastPool.Max 'Speed: the reader pool follows Fast' } finally { Close-ReaderPool $fastPool }
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
& $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $runFolder 'normal') 6>$null
Assert-Equal 'Normal' (Get-ScanSpeed) 'Speed: Enter gives Normal again in the next run'
# A cancel at "Scan more" (as Ctrl+C) leaves the script with an error: the run still signs out.
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite @('cancel')
$cancelError = ''
try { & $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $runFolder 'cancel-more') 6>$null } catch { $cancelError = $_.Exception.Message }  # expected
Assert-Equal $true ($cancelError -match 'prompt cancelled' -and (Get-FakeCallCount 'Disconnect-PnPOnline') -eq 1) 'Cancel after a site scan: the run still signs out, one time'
# A drive folder only (a mapped-drive form; it need not exist here): no sign-in, so no sign-out.
Reset-FakePnP @{}
& $script -SiteUrl 'X:\Unit\Shared' -OutputFolder (Join-Path $runFolder 'drive-out') 6>$null
Assert-Equal '0 0' ('{0} {1}' -f @(Get-FakePnPConnection).Count, (Get-FakeCallCount 'Disconnect-PnPOnline')) 'Drive folder only: no sign-in and no sign-out'

# ---- A report that cannot be saved: the script still signs out ----
Reset-FakePnP @{}
$blocker = Join-Path $runFolder 'not-a-folder.txt'
Set-Content -LiteralPath $blocker -Value 'x'
$shown = (& $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $blocker 'sub') 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $true ($shown -match 'The results could not be saved') 'Save error reaches the person'
Assert-Equal 1 (Get-FakeCallCount 'Read-Host') 'After a save error another site is still offered'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out when the report could not be saved'

Remove-Item -LiteralPath $runFolder -Recurse -Force
