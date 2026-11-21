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
try {
    $found = Get-DriveGroups $driveRoot
    Assert-Equal 'Admin=2|Ops=2|(top of the folder)=1' (@($found.Groups | ForEach-Object { $_.Name + '=' + $_.Files.Count }) -join '|') 'Drive: one group per top-level folder, then the top'
    $book = Invoke-DriveScan -Path $driveRoot -OutputFolder $driveOut -MaxFileMB 25 -Init $null 6>$null
    Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $book) 'Drive: every planted type found'
    $sheets = Get-XlsxSheets $book
    $phoneItems = @($sheets['Phone'] | Select-Object -Skip 1 | ForEach-Object { $_[2] } | Sort-Object -Unique)   # Confidence, Folder, File ...
    Assert-Equal 3 $phoneItems.Count 'Drive: a file with the same content as another still gets its own rows'
    $pdfRow = @($sheets['Not scanned'] | Where-Object { $_[2] -eq 'scan.pdf' })
    Assert-Equal 1 $pdfRow.Count 'Drive: a fake PDF with the bytes of a text file is not taken as its duplicate'
    Assert-Equal $true (@($sheets['Phone'] | Select-Object -Skip 1 | Where-Object { $_[6] -like 'file:///*' }).Count -ge 1) 'Drive: rows link to the file'
}
finally {
    Remove-Item -LiteralPath $driveRoot -Recurse -Force -ErrorAction SilentlyContinue
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
