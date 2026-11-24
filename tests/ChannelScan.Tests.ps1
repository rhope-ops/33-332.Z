<#
ChannelScan.Tests.ps1 - tests of ChannelScan.psm1 against the fake PnP cmdlets (FakePnP.ps1).

  Purpose:     Site address check; team or plain site; the progress meter text; file scan (hits, NOT SCANNED, masked file names, links);
               the download fallback without -AsMemoryStream.
  Inputs:      $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- Site address and channel pick ----
Assert-Equal $Sample.Site (Test-TeamSiteUrl "$($Sample.Site)/Shared%20Documents/Forms/AllItems.aspx") 'Address cut to the site'
$refused = $false
try { Test-TeamSiteUrl 'http://example.com/x' | Out-Null } catch { $refused = $true }  # the refusal is the expected result
Assert-Equal $true $refused 'Address that is not a site refused'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Test-TeamSiteUrl 'https://tenant.example.com/:w:/r/teams/TeamA/Shared%20Documents/a.docx?d=w1') `
    'A file Copy link is cut to the site'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Test-TeamSiteUrl ' tenant.example.com/teams/TeamA/Shared%20Documents') `
    'Address without https:// gets it'
$refused = $false
try { Test-TeamSiteUrl 'https://tenant.example.com/:w:/s/TeamA/EAbc' | Out-Null } catch { $refused = $true }  # expected refusal
Assert-Equal $true $refused 'A sharing link that does not name the site is refused'

# ---- Teams link: the group ID is read; the site address is asked for, checked, and only then scanned ----
$groupId = '11111111-2222-3333-4444-555555555555'
$teamLink = 'https://teams.example.com/l/team/19%3aabc%40thread.tacv2/conversations?groupId=' + $groupId + '&tenantId=x'
$channelLink = 'https://teams.example.com/l/channel/19%3aabc%40thread.tacv2/General?groupId=' + $groupId + '&tenantId=x'
Assert-Equal $groupId (Get-ScanTarget $teamLink).GroupId 'Team link gives the group ID'
Assert-Equal $groupId (Get-ScanTarget $channelLink).GroupId 'Channel link gives the group ID'
Assert-Equal $groupId (Get-ScanTarget ($channelLink -replace '^https://', '')).GroupId 'Teams link without https:// read'
Assert-Equal $null (Get-ScanTarget $Sample.Site).GroupId 'Site address has no group ID'
Assert-Equal $groupId (Get-ScanTarget ('https://teams.example.com/l/channel/link?groupId=' + $groupId)).GroupId 'The page form of a link is read'
$refused = $false
try { Get-ScanTarget 'https://teams.example.com/l/team/19%3aabc/conversations?tenantId=x' | Out-Null } catch { $refused = $true }  # expected
Assert-Equal $true $refused 'Teams link without groupId refused'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA/Shared%20Documents'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Teams link: the typed address, cut to the site'
Assert-Equal 300000 (Get-PnPContext).RequestTimeout 'Sign-in: each SharePoint call has a 5-minute limit (a stalled call cannot hang the scan)'
Assert-Equal $true (Test-FakePnPCall 'Read-Host') 'Teams link, lookup gives no site: the site address is asked for'
Assert-Equal 'https://usaf.dps.mil https://tenant.example.com/teams/TeamA' ((Get-FakePnPConnection) -join ' ') 'Teams link: the tenant root for the lookup, then the team site'
Assert-Equal $true (Test-FakePnPCall ('Lookup /_api/GroupSiteManager/GetSiteStatus ' + $groupId)) 'Teams link: one lookup, of the group in the link'
Assert-Equal $true (Test-FakePnPCall 'Quiet sign-in https://tenant.example.com/teams/TeamA') 'Sign-in warnings hidden'
# The lookup answers (the PAM's test 2, 2026-11-21): a team or channel link alone is enough, no question.
$siteTargetModule = Get-Module SiteTarget
& $siteTargetModule { $script:TenantRoot = 'https://tenant.example.com' }
foreach ($link in @($teamLink, $channelLink)) {
    Reset-FakePnP @{}; Set-FakeLookupSite 'https://tenant.example.com/teams/TeamA'
    Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $link) 6>$null) 'Teams link alone: the looked-up site'
    Assert-Equal $false (Test-FakePnPCall 'Read-Host') 'Teams link alone: no address question'
}
# A looked-up site on another host is never used (Data boundary): the address is asked for instead.
Reset-FakePnP @{}; Set-FakeLookupSite 'https://other.example.org/teams/TeamA'
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Lookup on another host: the typed address'
Assert-Equal $false ((Get-FakePnPConnection) -contains 'https://other.example.org/teams/TeamA') 'Lookup on another host: never signed in there'
# A refused lookup: the address is asked for.
Reset-FakePnP @{}; Set-FakeLookupSite 'fail'
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Lookup refused: the typed address'
# The looked-up site of another team is refused by the group check, as a typed one is.
Reset-FakePnP @{}; Set-FakeLookupSite 'https://tenant.example.com/teams/TeamB'
Set-FakeSiteAnswer -SiteGroupId '99999999-2222-3333-4444-555555555555'
$refused = $false
$wrongTeam = ''
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true; $wrongTeam = $_.Exception.Message }  # expected: another team's site
Assert-Equal $true $refused 'Looked-up site of another team refused'
Assert-Equal $false $wrongTeam.Contains('Click Sync') 'Another team: no sync steps (the sign-in worked)'
& $siteTargetModule { $script:TenantRoot = 'https://usaf.dps.mil' }
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamB' -SiteGroupId '99999999-2222-3333-4444-555555555555'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: another team's site
Assert-Equal $true $refused 'Site of another team refused'
# One sign-in per host (the PAM, 2026-11-24): Connect-TeamSite does not sign out; the run signs out once at its end.
Assert-Equal $false (Test-FakePnPCall 'Disconnect-PnPOnline') 'Site of another team: no sign-out here (the run signs out at its end)'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'not an address'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: bad typed address
Assert-Equal $true $refused 'Bad typed address refused'
Assert-Equal 'https://usaf.dps.mil' ((Get-FakePnPConnection) -join ' ') 'Bad typed address: no sign-in to a site (only the lookup at the tenant root)'
Assert-Equal $true (Get-ScanTarget $channelLink).IsChannel 'Channel link marked as a channel'
Assert-Equal $groupId (Get-ScanTarget ('https://teams.example.com/l/channel/19%3aabc/My Channel?groupId=' + $groupId)).GroupId `
    'Channel link with a raw space read'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA'
$printed = (Connect-TeamSite (Get-ScanTarget $channelLink) 6>&1 | Where-Object { $_ -is [System.Management.Automation.InformationRecord] }) -join ' '
Assert-Equal $true ($printed -match 'private or shared channel has its own site') 'Channel link prints the private-channel note'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'cancel'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = ($_.Exception.Message -match 'did not finish') }  # expected
Assert-Equal $true ($refused -and (Test-FakePnPCall 'Read-Host') -and -not (Test-FakePnPCall 'Disconnect-PnPOnline')) 'Failed address prompt: refused, no sign-out here (the run signs out at its end)'
$twoLinks = 'https://teams.example.com/l/team/A see also https://teams.example.com/l/team/B?groupId=' + $groupId
$refused = $false
try { Get-ScanTarget $twoLinks | Out-Null } catch { $refused = $true }  # expected: the first link has no groupId
Assert-Equal $true $refused 'A second link pasted after the first is not read'
$refused = $false
try { Get-ScanTarget ('https://teams.example.com/l/team/A see also teams.example.com/l/team/B?groupId=' + $groupId) | Out-Null } catch { $refused = $true }
Assert-Equal $true $refused 'A second link without https:// is not read either'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA'
$null = Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null
Assert-Equal $false (Test-FakePnPCall 'Disconnect-PnPOnline') 'Teams link: still signed in after success'
Reset-FakePnP @{}
Assert-Equal $Sample.Site (Connect-TeamSite (Get-ScanTarget $Sample.Site) 6>$null) 'Site address: used as given'
Assert-Equal $false (Test-FakePnPCall 'Read-Host') 'Site address: no address prompt'
Assert-Equal $false (Test-FakePnPCall 'Disconnect-PnPOnline') 'Site address: still signed in after success'
Assert-Equal $Sample.Site ((Get-FakePnPConnection) -join ' ') 'Site address: one sign-in, to that site'
# A sign-in without a sign-out keeps the last connection when it fails: a web of another site is refused, nothing read.
Reset-FakePnP @{}
Set-FakePnPOption StaleWeb $true
$staleError = ''
try { Connect-TeamSite (Get-ScanTarget $Sample.Site) 6>$null | Out-Null } catch { $staleError = $_.Exception.Message }  # expected: another site's web
Assert-Equal $true ($staleError -match 'did not reach this site') 'Connection reaches another site: refused before any read'
$meter = @{ Total = 40; Done = 10; Read = 8; Skipped = 2; Hits = 3; HitItems = 2; Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(30) } }
$expected = '10 of 40 files (25%) | read 8 | not scanned 2 | possible PII 3 in 2 files | 00:30 elapsed, about 01:30 left'
Assert-Equal $expected (Get-ScanMeterText $meter 'file') 'Meter text (the scan bar middle line): done of total with percent, counts, time elapsed and left'
Assert-Equal 25 (Get-MeterPercent $meter) 'Meter percent'
$meter.Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromMinutes(65) }
Assert-Equal $true ((Get-ScanMeterText $meter 'file') -match '\| 65:00 elapsed, about 195:00 left$') 'Meter time past an hour does not wrap'
$meter.Total = 0
Assert-Equal 0 (Get-MeterPercent $meter) 'Meter percent with no total'
Assert-Equal '10 posts | read 8 | not scanned 2 | possible PII 3 in 2 posts | 65:00 elapsed' (Get-ScanMeterText $meter 'post') 'Meter text: no total, no time left'
$meter.HitItems = 1
Assert-Equal $true ((Get-ScanMeterText $meter 'file') -match 'possible PII 3 in 1 file \|') 'Meter text: "in 1 file", not "in 1 files"'
$oneMeter = @{ Total = 1; Done = 1; Read = 1; Skipped = 0; Hits = 0; HitItems = 0; Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(3) } }
Assert-Equal $true ((Get-ScanMeterText $oneMeter 'file') -match '^1 of 1 file \(100%\)') 'Meter text: "1 of 1 file"'
$nearMeter = @{ Total = 400; Done = 399; Read = 399; Skipped = 0; Hits = 0; HitItems = 0; Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(3) } }
Assert-Equal 99 (Get-MeterPercent $nearMeter) 'Meter percent rounds down: 100% only when every file is done'

# ---- File scan ----
$library = '/sites/TeamA/Shared Documents'
$ssnFileName = "Roster $($Sample.Ssn).txt"
Reset-FakePnP @{
    "$library/General/Roster.docx"  = $Docx
    "$library/General/notes.txt"    = $Txt
    "$library/General/$ssnFileName" = [System.Text.Encoding]::UTF8.GetBytes('no PII inside')
    "$library/Ops/scan #A.pdf"      = $Txt
}
$lib = @(Get-SiteDocumentLibraries)[0]
Assert-Equal 'Documents' (@(Get-SiteDocumentLibraries | ForEach-Object { $_.Title }) -join '|') 'Libraries: only visible document libraries, not Style Library'
Assert-Equal 'General Ops' (Get-ChannelFolderNames $lib) 'Channel folders without Forms'
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>$null) 'Files scanned'
Assert-Equal 4 (Get-FakeCallCount 'Progress done 2') 'Per-file bar closed after each file, also a skipped one'
Write-FileStep (2MB) 'Downloading' 10
Assert-Equal $false (Test-FakePnPCall 'Progress 2 under 1') 'Per-file bar is not a child bar (ISE would indent it)'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'SSN found in docx'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'Phone' -and $_.Item -eq 'notes.txt' }).Count 'Phone found in txt'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'NOT SCANNED' -and $_.Item -eq 'scan #A.pdf' }).Count 'pdf listed'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General/Roster.docx' `
    ($rows | Where-Object { $_.Item -eq 'Roster.docx' } | Select-Object -First 1).Link 'Link built'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/Ops/scan%20%23A.pdf' `
    ($rows | Where-Object { $_.Item -eq 'scan #A.pdf' }).Link 'Link escapes #'
Assert-Equal $true ((Get-FakeProgressText) -match 'File: Roster\.docx') 'Progress bar names the file in work'

# A plain SharePoint site: files at the top of the library are scanned too, also when there is no folder.
Reset-FakePnP @{ "$library/top.txt" = $Txt; "$library/General/notes.txt" = $Txt; "$library/Forms/AllItems.txt" = $Txt }
Set-FakeFolders @('Forms')
Assert-Equal 0 @(Get-ChannelFolderNames $lib).Count 'A site with no folders is not an error'
$rootRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 1 (Invoke-FileScan -Library $lib -Channels @() -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rootRows 6>$null) 'Top-of-library file scanned'
Assert-Equal '(top of the library)' (($rootRows | Select-Object -ExpandProperty Channel -Unique) -join ',') 'Top-of-library files named as such'
Assert-Equal 'top.txt' (($rootRows | Select-Object -ExpandProperty Item -Unique) -join ',') 'Only the top file, not the folder file'
Write-FileStep (6MB) 'Downloading..' (3MB) ([TimeSpan]::FromSeconds(2))
Assert-Equal $true ((Get-FakeProgressText) -match 'Downloading\.\. 3\.0 of 6\.0 MB \(50%\) \| 1\.5 MB/s \| about 00:02 left') `
    'Per-file bar: MB done of the size, percent, speed and time left'
Write-FileStep (6MB) 'Downloading..'
Assert-Equal $true ((Get-FakeProgressText) -match 'Downloading\.\. 6\.0 MB - SharePoint gives no byte count') 'Per-file bar: says when the download gives no byte count'

# ---- The streamed download of a large file: the bar counts the bytes ----
$chunked = New-Object System.IO.MemoryStream(, [byte[]](1..200 | ForEach-Object { [byte]$_ }))
$copied = Copy-StreamWithProgress $chunked 200 'Downloading'
Assert-Equal '200|200' ('{0}|{1}' -f $copied.Length, $copied[199]) 'Stream copy: all bytes, in order'
Assert-Equal $true ($copied -is [byte[]]) 'Stream copy: a byte array, not unrolled'
$largeFile = New-Object byte[] (2MB); [System.Text.Encoding]::UTF8.GetBytes('Call ' + $Sample.Phone).CopyTo($largeFile, 0)
$channelScan = Get-Module ChannelScan
$normalOpener = & $channelScan { $script:OpenDirectStream }
# The fake stream opener: the file bytes, or (Short) fewer bytes than the list says, as a sign-in page would give.
& $channelScan { $script:OpenDirectStream = { param($url) Get-FakeFileStream $url } }
Reset-FakePnP @{ "$library/General/large.txt" = $largeFile }
$streamRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 1 (Invoke-FileScan -Library $lib -Channels @('General') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $streamRows 6>$null) 'Streamed: the large file is read'
Assert-Equal 1 @($streamRows | Where-Object { $_.Detector -eq 'Phone' }).Count 'Streamed: its PII is found'
Assert-Equal $true ((Get-FakeProgressText) -match 'Downloading\.\. 2\.0 of 2\.0 MB \(100%\)') 'Streamed: the bar reaches the full size'
# Now the opener gives two bytes of a web page: not the size the list gives.
& $channelScan { $script:OpenDirectStream = { param($url) Get-FakeFileStream $url -Short } }
Reset-FakePnP @{ "$library/General/large.txt" = $largeFile; "$library/General/large2.txt" = $largeFile }
$shortRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 2 (Invoke-FileScan -Library $lib -Channels @('General') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $shortRows 6>$null) 'Wrong size streamed: the normal download is used'
Assert-Equal 2 @($shortRows | Where-Object { $_.Detector -eq 'Phone' }).Count 'Wrong size streamed: the PII is still found'
Assert-Equal $false (& $channelScan { $script:DirectDownload }) 'Wrong size streamed: no more streamed downloads in this scan'
# The opener fails (no client library, or the sign-in is refused): the normal download, and no more tries.
& $channelScan { $script:OpenDirectStream = { param($url) throw ('FAKE: no stream for ' + $url) } }
Reset-FakePnP @{ "$library/General/large.txt" = $largeFile }
$failRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 1 (Invoke-FileScan -Library $lib -Channels @('General') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $failRows 6>$null) 'Stream fails: the normal download is used'
Assert-Equal $true ((Get-FakeProgressText) -match 'no byte count') 'Stream fails: the bar says it has no byte count'
Assert-Equal 1 @($failRows | Where-Object { $_.Detector -eq 'Phone' }).Count 'Stream fails: the PII is still found'
Assert-Equal $false (& $channelScan { $script:DirectDownload }) 'Stream fails: no more streamed downloads in this scan'
# A '#' in the path: the normal download, the stream stays on for other files.
& $channelScan { $script:OpenDirectStream = { param($url) Get-FakeFileStream $url } }
Reset-FakePnP @{ "$library/General/a #1.txt" = $largeFile }
Assert-Equal 1 (Invoke-FileScan -Library $lib -Channels @('General') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $failRows 6>$null) "A '#' in the path: read"
Assert-Equal $false ((Get-FakeProgressText) -match ' of 2\.0 MB') "A '#' in the path: not streamed"
Assert-Equal $true (& $channelScan { $script:DirectDownload }) "A '#' in the path: streaming stays on"
& $channelScan { param($opener) $script:OpenDirectStream = $opener } $normalOpener
# All workers busy: the scan bar says what it waits for.
Reset-FakePnP @{}
$waitQueue = @{ Activity = 'Folder 1 of 1: General'; Meter = (New-ScanMeter 4) }
& (Get-Module ScanQueue) { param($q) Write-WaitBar $q 3 } $waitQueue
Assert-Equal $true ((Get-FakeProgressText) -match 'Reading 3 files') 'Workers busy: the scan bar says how many files are being read'

# A file named after an SSN: reported as a hit, and neither the item nor the link shows the SSN.
$nameRows = @($rows | Where-Object { $_.Source -eq 'File name' })
Assert-Equal 1 $nameRows.Count 'SSN in a file name reported'
$allText = ($rows | ForEach-Object { $_.Item + ' ' + $_.Link + ' ' + $_.Context }) -join ' '
Assert-Equal $false $allText.Contains($Sample.Ssn) 'No full SSN in any item, link or snippet'
Assert-Equal $false $allText.Contains($Sample.SsnBare) 'No bare SSN in any link'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General' `
    ($nameRows[0].Link) 'Link of a non-Office file named with PII stops at its folder'
$fileId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
# Get-SafeLink is not exported: call it in the module's scope.
$linkOf = { param($ref, $id) & (Get-Module ChannelScan) { param($r, $i) Get-SafeLink 'https://tenant.example.com' $r $i } $ref $id }
Assert-Equal "https://tenant.example.com/sites/TeamA/_layouts/15/Doc.aspx?sourcedoc=%7B$fileId%7D" `
    (& $linkOf "$library/General/Roster $($Sample.Ssn).xlsx" $fileId) 'Office file named with PII: link by file ID'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/FY27/Plan%202.docx' `
    (& $linkOf "$library/FY27/Plan 2.docx" $fileId) 'Digits that are not PII: link to the file itself'

# A site with several document libraries: every library is scanned, each group named "Library / Folder".
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt; '/sites/TeamA/Plans/top.txt' = $Txt }
Set-FakeLibraries @('Shared Documents', 'Plans')
$libs = @(Get-SiteDocumentLibraries)
Assert-Equal 'Documents|Plans' (@($libs | ForEach-Object { $_.Title }) -join '|') 'Libraries: Documents first, then the others'
$multiRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 2 (Invoke-FileScan -Library $libs -Channels @() -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $multiRows 6>$null) 'Libraries: files of both scanned'
Assert-Equal 'Documents / General|Plans / (top of the library)' (@($multiRows | Select-Object -ExpandProperty Channel -Unique | Sort-Object) -join '|') 'Libraries: groups named Library / Folder'
Set-FakeLibraries @('Shared Documents', 'Plans', 'Preservation Hold Library')
$blockedRows = New-Object System.Collections.Generic.List[object]
$readCount = Invoke-FileScan -Library @(Get-SiteDocumentLibraries) -Channels @() -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $blockedRows 6>$null
Assert-Equal 2 $readCount 'Libraries: a library you cannot open does not stop the scan'
Assert-Equal 'Preservation Hold Library' (@($blockedRows | Where-Object { $_.Detector -eq 'BLOCKED' }).Channel -join '|') 'Libraries: it is one BLOCKED row'

# A label in a parent folder or in the channel name, the number in the file name: masked, reported, not in the link.
Reset-FakePnP @{ "$library/General/CAC/$($Sample.DodId).docx" = $Txt; "$library/CAC/$($Sample.DodId).txt" = $Txt }
$rows = New-Object System.Collections.Generic.List[object]
$null = Invoke-FileScan -Library $lib -Channels @('General', 'CAC') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>$null
$splitText = (($rows | ForEach-Object { $_.Item + ' ' + $_.Link }) -join ' ') + ' ' + (Get-FakeProgressText)
Assert-Equal $false $splitText.Contains($Sample.DodId) 'Label in a parent folder or the channel: no full DoD ID in item, link or bar'
Assert-Equal 2 @($rows | Where-Object { $_.Source -eq 'File name' -and $_.Detector -eq 'DoD ID' }).Count 'Label in a parent folder or the channel: reported as a hit'
Assert-Equal $true (@($rows | Where-Object { $_.Link -match 'Doc\.aspx\?sourcedoc=' }).Count -gt 0) 'Label in a parent folder: Office file linked by its ID'

# A channel named with PII: masked in every row.
Reset-FakePnP @{ "$library/SSN $($Sample.Ssn)/a.txt" = $Txt }
$rows = New-Object System.Collections.Generic.List[object]
$shown = (Invoke-FileScan -Library $lib -Channels @("SSN $($Sample.Ssn)") -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $false $shown.Contains($Sample.Ssn) 'Channel lines never show the SSN'
Assert-Equal 'SSN ***-**-6789' (($rows | Select-Object -ExpandProperty Channel -Unique) -join ',') 'Channel name: the SSN masked like a found value'
Assert-Equal $false ((($rows | ForEach-Object { $_.Link }) -join ' ').Contains('6789')) 'Channel link holds no SSN digits'
Assert-Equal $true ((Get-FakeProgressText) -match 'SSN \*\*\*-\*\*-6789') 'Progress bar shows the masked channel name'
Assert-Equal $false (Get-FakeProgressText).Contains($Sample.Ssn) 'Progress bar never shows the SSN'
Reset-FakePnP @{
    "$library/General/Roster.docx"  = $Docx
    "$library/General/notes.txt"    = $Txt
    "$library/General/$ssnFileName" = [System.Text.Encoding]::UTF8.GetBytes('no PII inside')
    "$library/Ops/scan #A.pdf"      = $Txt
}

# ---- Download fallback: a PnP version without -AsMemoryStream ----
Remove-FakeMemoryStreamSwitch
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>$null) 'Fallback download scans'
Assert-Equal $true (Test-FakePnPCall 'Get-PnPContext') 'Fallback used the client context'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'Fallback finds the SSN'
. (Join-Path $PSScriptRoot 'FakePnP.ps1')  # put the full fake Get-PnPFile back for the next tests

# ---- Error wording: a sign-in error a person can act on; the original always follows; other errors as they are ----
Assert-Equal $true ((ConvertTo-FriendlyError 'The remote server returned an error: (403) Forbidden.') -match '^No access: .*\(The remote server') 'A 403 says No access, with the original after it'
$closed = 'The underlying connection was closed: An unexpected error occurred on a send.'
Assert-Equal $closed (ConvertTo-FriendlyError $closed) 'A closed connection is not called a closed sign-in window'
Assert-Equal 'File Not Found.' (ConvertTo-FriendlyError 'File Not Found.') 'A missing file is not called a missing site'
# Sign-in blocked on this computer (the PAM, 2026-11-23: Connect-PnPOnline blocked outside the US): no stop, but the steps
# to scan a synced copy as a folder; "no access" keeps its own words only.
Set-FakeSiteAnswer
Set-FakePnPOption ConnectFails $true
$signIn = ''
try { Connect-TeamSite (Get-ScanTarget 'https://tenant.example.com/teams/TeamA') 6>$null | Out-Null } catch { $signIn = $_.Exception.Message }
Assert-Equal 'True True' ('{0} {1}' -f $signIn.StartsWith('FAKE: sign-in refused'), $signIn.Contains('Click Sync')) `
    ('Sign-in blocked: the steps to scan a synced folder follow the error: ' + $signIn.Split([char]10)[0])
Set-FakePnPOption ConnectFails $false

# ---- Live counts line and the timing ----
Reset-FakePnP @{}
$liveMeter = @{ Total = 4; Done = 2; Read = 1; Skipped = 1; Hits = 2; HitItems = 1; Types = @{ 'SSN' = 1; 'Phone' = 1 }; LastLine = [TimeSpan]::Zero
    Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(20) } }
$live = (Write-MeterLine $liveMeter 'file' -Force 6>&1 | ForEach-Object { [string]$_ }) -join ''
$expected = '  Total | files     2 /     4  50% | read     1 | not scanned     1 | possible PII     2 | SSN 1, Phone 1'
Assert-Equal $expected $live 'Live line: Total first, fixed-width counts, then types worst first; no time (it is on the scan bar)'
$folderMeter = $liveMeter.Clone(); $folderMeter.NameWidth = 12; $liveMeter.NameWidth = 12
$row = (Write-ChannelLine 'General' @{ Done = 0; Read = 0; Skipped = 0; Hits = 0 } $folderMeter 'file' 6>&1 | ForEach-Object { [string]$_ }) -join ''
$totalRow = (Write-MeterLine $liveMeter 'file' -Force 6>&1 | ForEach-Object { [string]$_ }) -join ''
Assert-Equal $true ($row.StartsWith('  General      | files')) 'Folder row: the folder name first, padded to the name column'
foreach ($column in @('| read', '| not scanned', '| possible PII')) {
    Assert-Equal $totalRow.IndexOf($column) $row.IndexOf($column) "Folder row and total row line up at [$column]"
}
$liveMeter.NameWidth = 0
$liveMeter.Done = 1234; $liveMeter.Read = 999; $liveMeter.Skipped = 12; $liveMeter.Hits = 0; $liveMeter.Types = @{}
$wide = (Write-MeterLine $liveMeter 'file' -Force 6>&1 | ForEach-Object { [string]$_ }) -join ''
Assert-Equal $live.IndexOf('| read') $wide.IndexOf('| read') 'Live line: columns stay in place when the counts grow'
Assert-Equal $live.IndexOf('| possible PII') $wide.IndexOf('| possible PII') 'Live line: PII column stays in place'
$liveMeter.Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(25) }
Assert-Equal '' ((Write-MeterLine $liveMeter 'file' 6>&1 | ForEach-Object { [string]$_ }) -join '') 'Live line: not again within 10 seconds'
Assert-Equal 'Time downloading|Time reading files (all readers)' (@((Get-ScanTiming).Keys) -join '|') 'Timing: download and reading times'

# ---- Team or plain site: names the folder column ----
Reset-FakePnP @{}
Assert-Equal $true (Test-TeamSite) 'A team site is a team'
Set-FakeSiteAnswer -TypedSite '' -SiteGroupId ([guid]::Empty)
Assert-Equal $false (Test-TeamSite) 'A plain SharePoint site is not a team'
