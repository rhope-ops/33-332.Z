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
# A file link to a file outside the folder is never read (the PAM, 2026-11-21).
$outside = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-outside-' + $letters + '.txt')
[System.IO.File]::WriteAllText($outside, "SSN: $($Sample.Ssn)")
New-Item -ItemType SymbolicLink -Path (Join-Path $driveRoot 'Admin\pointer.txt') -Target $outside | Out-Null
try {
    $found = Get-DriveGroups $driveRoot
    Assert-Equal 'Admin/pointer.txt' (@($found.FileLinks) -join '|').Replace('\', '/') 'Drive: a file link is listed, not read'
    Assert-Equal 'Admin=2|Ops=2|(top of the folder)=1' (@($found.Groups | ForEach-Object { $_.Name + '=' + $_.Files.Count }) -join '|') 'Drive: one group per top-level folder, then the top'
    $book = Invoke-DriveScan -Path $driveRoot -OutputFolder $driveOut -MaxFileMB 25 -Init $null 6>$null
    Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $book) 'Drive: every planted type found'
    $linkRow = @((Get-XlsxSheets $book)['Not scanned'] | Where-Object { $_[3] -match '^A link to another file is not followed' })
    Assert-Equal 1 $linkRow.Count 'Drive: a file link gets a NOT SCANNED row'
    $sheets = Get-XlsxSheets $book
    $phoneItems = @($sheets['Phone'] | Select-Object -Skip 1 | ForEach-Object { $_[2] } | Sort-Object -Unique)   # Confidence, Folder, File ...
    Assert-Equal 3 $phoneItems.Count 'Drive: a file with the same content as another still gets its own rows'
    $pdfRow = @($sheets['Not scanned'] | Where-Object { $_[2] -eq 'scan.pdf' })
    Assert-Equal 1 $pdfRow.Count 'Drive: a fake PDF with the bytes of a text file is not taken as its duplicate'
    Assert-Equal $true (@($sheets['Phone'] | Select-Object -Skip 1 | Where-Object { $_[6] -like 'file:///*' }).Count -ge 1) 'Drive: rows link to the file'
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

# ---- Accountability (the PAM, 2026-11-21): who changed the file, who added the PII and when ----
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
        'Who: after an unread version, the person alone in "PII added by", and the note says it may be earlier'
    # More old versions than are read, all readable, the SSN in the oldest: a plain credit, no "could not be read".
    $versionList = @(New-FakeVersion ([datetime]'2026-01-01') 'true.adder' $withSsn) + @(2..31 | ForEach-Object { New-FakeVersion ([datetime]'2026-01-01').AddDays($_) 'other' $withSsn })
    $plainEntry = (Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who)[$hits[0].Detector + '|' + $hits[0].MaskedValue]
    Assert-Equal 'true.adder|Versions after the 30 oldest were not checked' `
        ($plainEntry.AddedBy + '|' + $plainEntry.Note) 'Who: past the version cap, the oldest version with the hit is credited, with a note'
    # Removed, then put back: the first adder is named, and who put it back (the PAM, 2026-11-21).
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
    $versionList = @()
    $byHit = Find-PiiAddedVersion '/sites/x/Shared Documents/a.txt' 'a.txt' $hits $who
    Assert-Equal 'first.saver' $byHit[$hits[0].Detector + '|' + $hits[0].MaskedValue].AddedBy 'Who: with no older version, the maker added it'
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
