<#
DriveScan.Tests.ps1 - tests the drive and network folder scan and the reader workers.

  Purpose:     A pasted folder path is recognised; a folder and its subfolders are scanned with the reader workers;
               files with the same content are read once and still each get their rows; the worker count stays
               inside its limits. Invented files in a new temporary folder only (AGENTS.md "Data boundary").
  Inputs:      $Sample, $Txt, $Docx (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1, after ChannelScan.Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none (a local temporary folder stands in for a network share).
  Writes:      invented files and the report in new temporary folders, deleted at the end.
  Credentials: none.
#>

# A pasted folder is a drive target; a site address and a Teams link are not.
Assert-Equal 'X:\Unit\Shared' (Get-ScanTarget 'X:\Unit\Shared').DrivePath 'Drive: a mapped-drive folder is a drive target'
Assert-Equal '\\server\share\Unit' (Get-ScanTarget '"\\server\share\Unit"').DrivePath 'Drive: a UNC folder in quotes is a drive target'
Assert-Equal '' (Get-ScanTarget $Sample.Site).DrivePath 'Drive: a site address is not a drive target'

# A folder with subfolders, a duplicate file in two places, a file at its top and a fake PDF.
# Letters only in the folder name: a path with a number that looks like PII gets no link (Get-DriveLink).
$letters = -join @(1..12 | ForEach-Object { [char](97 + (Get-Random -Maximum 26)) })
$driveRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-drive-' + $letters)
$driveOut = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-driveout-' + [guid]::NewGuid())
foreach ($folder in @('Ops', 'Ops\old', 'Admin')) { New-Item -ItemType Directory -Path (Join-Path $driveRoot $folder) -Force | Out-Null }
[System.IO.File]::WriteAllBytes((Join-Path $driveRoot 'Ops\notes.txt'), $Txt)
[System.IO.File]::WriteAllBytes((Join-Path $driveRoot 'Ops\old\notes copy.txt'), $Txt)
[System.IO.File]::WriteAllBytes((Join-Path $driveRoot 'Admin\Roster.docx'), $Docx)
[System.IO.File]::WriteAllBytes((Join-Path $driveRoot 'Admin\scan.pdf'), $Txt)
[System.IO.File]::WriteAllBytes((Join-Path $driveRoot 'top.txt'), $Txt)
# A file link to a file outside the folder is never read (the Administrator, 2026-11-21).
$outside = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-outside-' + $letters + '.txt')
[System.IO.File]::WriteAllText($outside, "SSN: $($Sample.Ssn)")
New-Item -ItemType SymbolicLink -Path (Join-Path $driveRoot 'Admin\pointer.txt') -Target $outside | Out-Null
try {
    $found = Get-DriveGroups $driveRoot
    Assert-Equal 'Admin/pointer.txt' (@($found.FileLinks) -join '|').Replace('\', '/') 'Drive: a file link is listed, not read'
    Assert-Equal 'Admin=2|Ops=2|(top of the folder)=1' (@($found.Groups | ForEach-Object { $_.Name + '=' + $_.Files.Count }) -join '|') 'Drive: one group per top-level folder, then the top'
    Set-FakeSiteAnswer -ScanAnswer 'no'
    $declined = Invoke-DriveScan -Path $driveRoot -OutputFolder (Join-Path $driveOut 'declined') -MaxFileMB 25 -Init $null 6>$null
    Assert-Equal $true ($null -eq $declined -and -not (Test-Path -LiteralPath (Join-Path $driveOut 'declined'))) 'Drive: "no" to "Scan now?" saves nothing'
    Set-FakeSiteAnswer
    $book = Invoke-DriveScan -Path $driveRoot -OutputFolder $driveOut -MaxFileMB 25 -Init $null 6>$null
    Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $book) 'Drive: every planted type found'
    $linkRow = @((Get-XlsxSheets $book)['Not scanned'] | Where-Object { $_[3] -match '^A link to another file is not followed' })
    Assert-Equal 1 $linkRow.Count 'Drive: a file link gets a NOT SCANNED row'
    $sheets = Get-XlsxSheets $book
    $phoneRows = @($sheets['Hits'] | Select-Object -Skip 1 | Where-Object { $_[1] -eq 'Phone' })   # ID, Type, Confidence, Folder, File ... Link last
    $phoneItems = @($phoneRows | ForEach-Object { $_[4] } | Sort-Object -Unique)
    Assert-Equal 3 $phoneItems.Count 'Drive: a file with the same content as another still gets its own rows'
    $pdfRow = @($sheets['Not scanned'] | Where-Object { $_[2] -eq 'scan.pdf' })
    Assert-Equal 1 $pdfRow.Count 'Drive: a fake PDF with the bytes of a text file is not taken as its duplicate'
    Assert-Equal $true (@($phoneRows | Where-Object { $_[22] -like 'file:///*' }).Count -ge 1) 'Drive: rows link to the file'
}
finally {
    Remove-Item -LiteralPath $driveRoot -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $outside -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $driveOut -Recurse -Force -ErrorAction SilentlyContinue
}

# The reader pool: the worker count stays between 1 and its maximum; a failed read lowers it.
$pool = New-ReaderPool $null
try {
    Assert-Equal $true ($pool.Target -ge 1 -and $pool.Target -le $pool.Max -and $pool.Max -le 8) 'Pool: workers between 1 and 8'
    Add-PoolJob $pool @{ Id = 1; Name = 'gone.txt'; Bytes = $null; Path = (Join-Path $driveRoot 'gone.txt') }
    $gone = @(Receive-PoolResult $pool -Wait)
    Assert-Equal 'NOT SCANNED' $gone[0].Status 'Pool: a file that cannot be opened is NOT SCANNED, not an error'
}
finally { Close-ReaderPool $pool }

# OneDrive (the Administrator, 2026-11-23): a synced OneDrive item is a reparse point but no link: it is read; a symbolic link or
# junction is not followed; when its kind cannot be read, it stays a link.
$odRoots = @('C:\Users\me\OneDrive - Org')
$inOneDrive = 'C:\Users\me\OneDrive - Org\Ops\a.docx'
Assert-Equal 'Cloud' (Get-ReparseKind ([pscustomobject]@{ FullName = $inOneDrive; LinkType = $null; Target = $null }) $odRoots) 'OneDrive item (in a sync folder, no link type, no target): read'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = 'D:\Share\a.docx'; LinkType = $null; Target = $null }) $odRoots) 'Reparse point outside OneDrive (dedup, Azure File Sync): not read'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = 'C:\Users\me\OneDrive - Orgx\a.docx'; LinkType = $null; Target = $null }) `
    $odRoots) 'A folder that only starts like a sync folder: not read'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = $inOneDrive; LinkType = 'SymbolicLink'; Target = @('x') }) $odRoots) 'Symbolic link in OneDrive: not followed'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = $inOneDrive; LinkType = 'Junction'; Target = $null }) $odRoots) 'Junction in OneDrive: not followed'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = $inOneDrive }) $odRoots) 'Kind not known: not followed'
# An online-only file: 'Cloud' from its attributes; LinkType and Target are not read (on 5.1 that read makes OneDrive
# download the file; device test 2026-11-27). Outside a sync folder the attributes do not make it a cloud item.
$script:linkReads = 0
$onlineEntry = [pscustomobject]@{ FullName = $inOneDrive; Attributes = 0x400400 }
$onlineEntry | Add-Member -MemberType ScriptProperty -Name LinkType -Value { $script:linkReads++; $null }
$onlineEntry | Add-Member -MemberType ScriptProperty -Name Target -Value { $script:linkReads++; $null }
Assert-Equal 'Cloud' (Get-ReparseKind $onlineEntry $odRoots) 'Online-only OneDrive file: read as a cloud item'
Assert-Equal 0 $script:linkReads 'Online-only OneDrive file: LinkType and Target not read (no download by Windows)'
Assert-Equal 'Link' (Get-ReparseKind ([pscustomobject]@{ FullName = 'D:\Share\a.docx'; Attributes = 0x400400 }) $odRoots) 'Online-only attribute outside OneDrive: not read'
$linkEntry = [pscustomobject]@{ FullName = $inOneDrive; Attributes = 0x1410; LinkType = 'Junction'; Target = @('C:\Users') }
Assert-Equal 'Link' (Get-ReparseKind $linkEntry $odRoots) 'Junction with Offline set by hand in OneDrive: not followed'
# Online-only files: asked first (the Administrator, 2026-11-23: "Ask each run"); Enter (no) leaves them NOT SCANNED, y reads them.
function New-OnlineGroups {
    # Two files in one group: one online only, one on the disk.
    $files = New-Object System.Collections.Generic.List[object]
    $files.Add(@{ Path = 'C:\od\Ops\cloud.docx'; Inner = 'cloud.docx'; Name = 'cloud.docx'; Size = 2048L; Modified = '2026-11-23 10:00'; Online = $true })
    $files.Add(@{ Path = 'C:\od\Ops\local.docx'; Inner = 'local.docx'; Name = 'local.docx'; Size = 1024L; Modified = '2026-11-23 10:00'; Online = $false })
    return , @(@{ Name = 'Ops'; Files = $files })
}
$onlineGroups = New-OnlineGroups
$onlineRows = New-Object System.Collections.Generic.List[object]
Set-FakeOnlineAnswer ''
Select-OnlineFiles $onlineGroups $onlineRows 25 6>$null
Assert-Equal '1 local.docx' ('{0} {1}' -f $onlineGroups[0].Files.Count, $onlineGroups[0].Files[0].Name) 'Online only, answer Enter: only the file on the disk is read'
Assert-Equal 1 @($onlineRows | Where-Object { $_.Context -like 'Online only*' }).Count 'Online only, answer Enter: one NOT SCANNED row'
$onlineGroups = New-OnlineGroups
Set-FakeOnlineAnswer 'y'
Select-OnlineFiles $onlineGroups (New-Object System.Collections.Generic.List[object]) 25 6>$null
Assert-Equal 2 $onlineGroups[0].Files.Count 'Online only, answer y: both files are read'
Set-FakeOnlineAnswer ''

# A normal old-version read while a scan's pool is open, then the close: no error, and the next scan starts clean
# (second review of PR 58: the close threw, and every later scan ended INCOMPLETE).
$livePool = New-ReaderPool $null
$closeError = ''
try {
    $versionText = Invoke-LimitedRead ([System.Text.Encoding]::UTF8.GetBytes('old version text')) 'old.txt'
    Assert-Equal 'OK' $versionText.Status 'Old version: read in its own runspace while the scan runs'
}
finally { try { Close-ReaderPool $livePool } catch { $closeError = $_.Exception.Message } }
Assert-Equal '' $closeError 'Pool: closes without error after an old-version read'
$livePool = New-ReaderPool $null
try { Assert-Equal 'OK' (Invoke-LimitedRead ([System.Text.Encoding]::UTF8.GetBytes('next scan')) 'next.txt').Status 'Old version: the next scan reads again' }
finally { Close-ReaderPool $livePool }

# A worker stuck inside a .NET call (it does not stop; here for 15 s, so the test run can end): the file is NOT SCANNED after the time limit, and closing the pool
# (as after Ctrl+C) does not wait for it (the Administrator's gov run, 2026-11-23: a scan hung and Ctrl+C did not stop it).
$readerPool = Get-Module ReaderPool
$savedLimit = & $readerPool { $script:FileSecondsLimit }
& $readerPool { $script:FileSecondsLimit = 2 }
$stuckCode = 'function Invoke-PooledRead { param($Job, $Seen) if ($Job.Id -eq 7) { [System.Threading.Thread]::Sleep(15000) }; @{ Id = $Job.Id; Status = "OK" } }' +
    "`nfunction Get-TextFromBytes { param(`$Bytes, `$FileName) [System.Threading.Thread]::Sleep(15000) }"
$stuckPool = New-ReaderPool @{ Code = $stuckCode }
# One worker only: the next file can run only when the pool gets a worker in place of the stuck one.
[void]$stuckPool.Pool.SetMaxRunspaces(1); $stuckPool.Max = 1; $stuckPool.Target = 1
$clock = [Diagnostics.Stopwatch]::StartNew()
try {
    Add-PoolJob $stuckPool @{ Id = 7; Name = 'stuck.txt'; Bytes = $null; Path = 'stuck.txt' }
    $slow = @(Receive-PoolResult $stuckPool -Wait)
    Assert-Equal '7 NOT SCANNED True' ('{0} {1} {2}' -f $slow[0].Id, $slow[0].Status, $slow[0].Reason.StartsWith('Reading took longer')) 'Pool: a stuck worker gives NOT SCANNED after the limit'
    Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15) 'Pool: the stuck worker does not hold the scan'
    Add-PoolJob $stuckPool @{ Id = 8; Name = 'next.txt'; Bytes = $null; Path = 'next.txt' }
    $next = @(Receive-PoolResult $stuckPool -Wait)
    Assert-Equal '8 OK' ('{0} {1}' -f $next[0].Id, $next[0].Status) 'Pool: the next file runs after a stuck one (one worker)'
    # An old SharePoint version is read off the main thread, with the same limit (review of PR 58).
    $clock.Restart()
    $timedOut = $false
    try { $null = Invoke-LimitedRead ([byte[]](1, 2, 3)) 'old.pdf' } catch [System.TimeoutException] { $timedOut = $true }
    Assert-Equal 'True True' ('{0} {1}' -f $timedOut, ($clock.Elapsed.TotalSeconds -lt 10)) 'Old version: a stuck read stops at the limit'
}
finally {
    $clock.Restart()
    Close-ReaderPool $stuckPool
    Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 5) 'Pool: closing does not wait for stuck workers (Ctrl+C ends)'
    & $readerPool { param($limit) $script:FileSecondsLimit = $limit } $savedLimit
}

# A whole OneDrive folder (the Administrator, 2026-11-23: "Warn, then ask"): its top folders can be shortcuts to other teams'
# libraries ("No reconnaissance"), so Ion asks first; Enter = no. A folder inside it, or any other folder: no question.
$odRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ion-od-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $odRoot 'Team A') -Force | Out-Null
$savedOneDrive = $env:OneDrive; $savedCommercial = $env:OneDriveCommercial; $savedConsumer = $env:OneDriveConsumer
try {
    $env:OneDrive = $odRoot
    # Two variables for one folder (common on a work PC) and one for a folder that is gone: each top folder once, no error.
    $env:OneDriveCommercial = $odRoot
    $env:OneDriveConsumer = Join-Path $odRoot 'gone'
    Set-FakeRootAnswer ''
    $shown = @(Confirm-OneDriveRoot (Split-Path $odRoot -Parent) 6>&1 | ForEach-Object { [string]$_ })
    $topLine = @($shown | Where-Object { $_ -like '*top folders in OneDrive*' })
    Assert-Equal '1 False' ('{0} {1}' -f $topLine.Count, $shown[-1]) 'Same folder twice and a gone folder: asked once, no error'
    Assert-Equal 1 ([regex]::Matches($topLine[0], 'Team A').Count) 'Same folder twice: each top folder shown once'
    Set-FakeRootAnswer ''
    Assert-Equal $false (Confirm-OneDriveRoot $odRoot 6>$null) 'Whole OneDrive folder, answer Enter: not scanned'
    Set-FakeRootAnswer 'y'
    Assert-Equal $true (Confirm-OneDriveRoot ($odRoot + [System.IO.Path]::DirectorySeparatorChar) 6>$null) 'Whole OneDrive folder, answer y: scanned'
    Set-FakeRootAnswer ''
    Assert-Equal $true (Confirm-OneDriveRoot (Join-Path $odRoot 'Team A') 6>$null) 'A folder inside OneDrive: no question'
    Assert-Equal $false (Confirm-OneDriveRoot (Join-Path (Join-Path $odRoot 'Team A') '..') 6>$null) 'The OneDrive folder written with ..: asked (Enter = no)'
    Assert-Equal $false (Confirm-OneDriveRoot (Split-Path $odRoot -Parent) 6>$null) 'A folder that holds the OneDrive folder: asked (Enter = no)'
    $sibling = $odRoot + '-library'   # a synced team library next to OneDrive (registry list, not the variable)
    New-Item -ItemType Directory -Path $sibling -Force | Out-Null
    Assert-Equal $true (Confirm-OneDriveRoot $sibling 6>$null) 'A synced team library folder: no question'
    $env:OneDrive = 'bad' + [char]0 + 'path'   # a path .NET cannot read: skipped, the scan is not stopped
    Assert-Equal $true (Confirm-OneDriveRoot $sibling 6>$null) 'A bad OneDrive variable: skipped, no error'
    $env:OneDrive = $odRoot
    Remove-Item -LiteralPath $sibling -Recurse -Force -ErrorAction SilentlyContinue
    # A cancel at the question saves nothing (as at "Scan now?"): Invoke-DriveScan returns no workbook.
    Set-FakeRootAnswer 'cancel'
    $odOut = Join-Path $odRoot 'out-not-made'
    Assert-Equal $null (Invoke-DriveScan -Path $odRoot -OutputFolder $odOut -MaxFileMB 25 -Init $null 6>$null) 'Whole OneDrive folder, a cancel at the question: no workbook'
    Assert-Equal $false (Test-Path -LiteralPath $odOut) 'Whole OneDrive folder, a cancel at the question: nothing saved'
}
finally {
    $env:OneDrive = $savedOneDrive; $env:OneDriveCommercial = $savedCommercial; $env:OneDriveConsumer = $savedConsumer
    Remove-Item -LiteralPath $odRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---- Accountability (the Administrator, 2026-11-21): who changed the file, who added the PII and when ----
Assert-Equal $Sample.Email (Get-PersonText ([pscustomobject]@{ Email = $Sample.Email; LookupValue = 'Tom Mims' })) 'Who: a SharePoint person is shown by e-mail address'
# Version history: the first version that holds the SSN dates it, and names who saved that version.
$plain = [System.Text.Encoding]::UTF8.GetBytes('Roster draft, no numbers yet')
$withSsn = [System.Text.Encoding]::UTF8.GetBytes("Roster SSN: $($Sample.Ssn)")
function global:New-FakeVersion {
    # One old version: its time, who saved it, and its bytes.
    param([datetime]$Created, [string]$By, [byte[]]$Bytes)
    $version = [pscustomobject]@{ Created = $Created; CreatedBy = [pscustomobject]@{ Email = $By } }
    $version | Add-Member -NotePropertyName Bytes -NotePropertyValue $Bytes
    $version | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value { [pscustomobject]@{ Value = (New-Object System.IO.MemoryStream(, $this.Bytes)) } }
    return $version
}
$versionList = @((New-FakeVersion ([datetime]'2026-01-02 03:04') $Sample.Email $withSsn), (New-FakeVersion ([datetime]'2026-01-01 08:00') 'first.saver' $plain))
function global:Get-PnPFileVersion {
    # Fake PnP: the file's old versions (newest first, as SharePoint lists them).
    param([string]$Url)
    $versionList
}
$realContext = ${function:global:Get-PnPContext}
function global:Get-PnPContext {
    # Fake PnP context: Load and ExecuteQuery do nothing.
    $context = [pscustomobject]@{}
    $context | Add-Member -MemberType ScriptMethod -Name Load -Value { param($item) }
    $context | Add-Member -MemberType ScriptMethod -Name ExecuteQuery -Value { }
    return $context
}
try {
    $hits = @(Find-PiiInText "Roster SSN: $($Sample.Ssn)")
    $who = @{ Modified = '2026-02-01 10:00'; ModifiedBy = 'last.editor'; Created = '2026-01-01 08:00'; CreatedBy = 'first.saver' }
    $byHit = Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who
    $added = $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal ('2026-01-02 03:04|' + $Sample.Email) ($added.AddedOn + '|' + $added.AddedBy) 'Who: the first version with the PII names who added it and when'
    # More old versions than are read, the SSN first in a later one: who added it is not recorded, never the last editor.
    $versionList = @(1..30 | ForEach-Object { New-FakeVersion ([datetime]'2026-01-01').AddDays($_) 'early.saver' $plain }) +
        @(New-FakeVersion ([datetime]'2026-03-01') 'real.adder' $withSsn)
    $byHit = Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who
    $pastEntry = $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    $pastText = $pastEntry.AddedOn + '|' + $pastEntry.AddedBy + '|' + $pastEntry.Note
    Assert-Equal ('Not recorded|Not recorded|' + "Not in the old versions read: check the file's version history by hand") $pastText `
        'Who: past the versions read, who added it is not recorded, and the note says why'
    # An older version that cannot be read: the later version with the SSN is named, marked "may have been added earlier".
    $unreadable = New-FakeVersion ([datetime]'2026-01-01') 'unread.saver' $withSsn
    $unreadable | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value { throw 'version not available' } -Force
    $versionList = @($unreadable, (New-FakeVersion ([datetime]'2026-01-05') 'later.saver' $withSsn))
    $byHit = Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who
    $gapAdded = $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'later.saver|An older version could not be read: it may have been added earlier' ($gapAdded.AddedBy + '|' + $gapAdded.Note) `
        'Who: after an unread version, the person alone in "PII first saved by", and the note says it may be earlier'
    # An old version over the size cap (30 MB): not downloaded, so it counts as not read (plan step 5, 2026-11-26).
    $bigVersion = New-FakeVersion ([datetime]'2026-01-01') 'big.saver' $withSsn
    $bigVersion | Add-Member -NotePropertyName Size -NotePropertyValue (30MB)
    $bigVersion | Add-Member -NotePropertyName Opened -NotePropertyValue $false
    $bigVersion | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value { $this.Opened = $true; throw 'should not be downloaded' } -Force
    $versionList = @($bigVersion, (New-FakeVersion ([datetime]'2026-01-05') 'later.saver' $withSsn))
    $bigEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'False|later.saver|An older version could not be read: it may have been added earlier' `
        ('{0}|{1}|{2}' -f $bigVersion.Opened, $bigEntry.AddedBy, $bigEntry.Note) 'Who: an old version over 25 MB is not downloaded and counts as not read'
    # More old versions than are read, all readable, the SSN in the oldest: a plain credit, no "could not be read".
    $versionList = @(New-FakeVersion ([datetime]'2026-01-01') 'true.adder' $withSsn) + @(2..31 | ForEach-Object { New-FakeVersion ([datetime]'2026-01-01').AddDays($_) 'other' $withSsn })
    $plainEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'true.adder|Versions after the 30 oldest were not checked' `
        ($plainEntry.AddedBy + '|' + $plainEntry.Note) 'Who: past the version cap, the oldest version with the hit is credited, with a note'
    # Removed, then put back: the first adder is named, and who put it back (the Administrator, 2026-11-21).
    $versionList = @((New-FakeVersion ([datetime]'2026-01-01') 'orig.adder' $withSsn), (New-FakeVersion ([datetime]'2026-01-02') 'remover' $plain),
        (New-FakeVersion ([datetime]'2026-01-03') 'readder' $withSsn))
    $backEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'orig.adder|Put back by readder on 2026-01-03 00:00' ($backEntry.AddedBy + '|' + $backEntry.Note) 'Who: a hit put back after a removal names who put it back'
    # Removed, then a version not read, then back: the put-back cannot be named, and the report says so.
    $gapVersion = New-FakeVersion ([datetime]'2026-01-03') 'unread' $withSsn
    $gapVersion | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value { throw 'version not available' } -Force
    $versionList = @((New-FakeVersion ([datetime]'2026-01-01') 'orig.adder' $withSsn), (New-FakeVersion ([datetime]'2026-01-02') 'remover' $plain),
        $gapVersion, (New-FakeVersion ([datetime]'2026-01-04') 'back' $withSsn))
    $unsureEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'orig.adder|It may have been removed and put back (a version could not be read)' ($unsureEntry.AddedBy + '|' + $unsureEntry.Note) 'Who: a put-back next to an unread version is noted'
    # A version not read right after the first one with the SSN (a removal and put-back may hide in it): noted.
    $hidden = New-FakeVersion ([datetime]'2026-01-02') 'unread' $plain
    $hidden | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value { throw 'version not available' } -Force
    $versionList = @((New-FakeVersion ([datetime]'2026-01-01') 'first.adder' $withSsn), $hidden, (New-FakeVersion ([datetime]'2026-01-03') 'back' $withSsn))
    $hiddenEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'first.adder|It may have been removed and put back (a version could not be read)' `
        ($hiddenEntry.AddedBy + '|' + $hiddenEntry.Note) 'Who: an unread version after the first sighting is noted'
    # Not seen before the unread version: no removal is known, so no put-back note.
    $versionList = @((New-FakeVersion ([datetime]'2026-01-01') 'none' $plain), $hidden, (New-FakeVersion ([datetime]'2026-01-03') 'a3' $withSsn))
    $laterNote = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue].Note
    Assert-Equal $false ($laterNote -match 'put back') 'Who: no put-back note for a hit first seen after an unread version'
    # Every old version read, none holds the SSN: the current version is named, and the note says so.
    $versionList = @((New-FakeVersion ([datetime]'2026-01-01') 'early.saver' $plain), (New-FakeVersion ([datetime]'2026-01-02') 'early.saver' $plain))
    $currentEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'last.editor|In no older version: the current version is named' ($currentEntry.AddedBy + '|' + $currentEntry.Note) 'Who: the current-version fallback is noted'
    $versionList = @()
    $byHit = Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who
    Assert-Equal 'first.saver' $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue].AddedBy 'Who: with no older version, the maker added it'
    Assert-Equal "No older version: the file's creator is named" $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue].Note 'Who: the no-older-version fallback is noted'
}
finally {
    Remove-Item function:global:Get-PnPFileVersion, function:global:New-FakeVersion
    Set-Item function:global:Get-PnPContext $realContext
}
# Rows get the columns; a drive file says what Windows does not record.
$row = New-ReportRow 'File' 'Ops' 'a.txt' '' '' 'SSN' 'High' '***-**-0000' 'x'
Set-RowAccountability @($row) (Get-DriveWho (Join-Path ([System.IO.Path]::GetTempPath()) 'no-such-file.txt') $null)
Assert-Equal $true ($row.AddedBy -eq 'Not recorded' -and $row.ModifiedBy -match '^Not recorded') 'Who: a drive file says what is not recorded'
Assert-Equal 'A drive keeps no file history' $row.HistoryNote 'Who: a drive file says why in the history note'
# A file link whose name holds PII: its name gets a hit row (masked), and it is still not read.
$nameRows = New-Object System.Collections.Generic.List[object]
& (Get-Module DriveScan) { param($r, $n) Add-DriveListRows @{ Blocked = @(); Links = @(); FileLinks = @("Ops/memo $n.txt") } $r } $nameRows $Sample.Ssn
Assert-Equal 'SSN|NOT SCANNED' ((@($nameRows | ForEach-Object { $_.Detector }) | Sort-Object -Descending) -join '|') 'Drive: a file link named with PII gets a name hit and a NOT SCANNED row'
