<#
OneDriveWeb.Tests.ps1 - tests the OneDrive hybrid (plan step 4): the local-to-web map, the kinds of OneDrive folder,
the fall-back to a local scan, and a scan that reads the history and a web-only file from the (fake) web.

  Purpose:     The map follows probe 4 (web path = UrlNamespace path + the local path after MountPoint; the longest
               MountPoint wins). Only the own OneDrive folder uses the web until probe 4 runs A and B pass. A failed
               sign-in scans local only and says so. The web side reads only files under the pasted folder.
  Inputs:      $Txt, $Sample (TestKit.ps1); the fake PnP (FakePnP.ps1). Invented paths and files only.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1, after DriveScan.Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none (fake PnP; a local temporary folder stands in for the OneDrive folder).
  Writes:      invented files and the report in new temporary folders, deleted at the end.
  Credentials: none.
#>

# ---- The map (probe 4) ----
$odHome = 'C:\Users\me\OneDrive - Org'
$odEntries = @(
    @{ MountPoint = $odHome; UrlNamespace = 'https://org-my.example.com/personal/me_org/Documents/' },
    @{ MountPoint = 'C:\Users\me\OneDrive - Org\'; UrlNamespace = 'https://org-my.example.com/personal/me_org/Documents' },
    @{ MountPoint = 'C:\Users\me\OneDrive - Org\Team X - Files'; UrlNamespace = 'https://org.example.com/sites/TeamX/Shared Documents/Plans' },
    @{ MountPoint = 'C:\Users\me\Org\Team Y - Docs'; UrlNamespace = 'https://org.example.com/teams/TeamY/Shared%20Documents' })
$m = Get-OneDriveWebMap ($odHome + '\Ops\Q1') $odEntries @($odHome)
Assert-Equal 'Home|https://org-my.example.com/personal/me_org|/personal/me_org/Documents|/personal/me_org/Documents/Ops/Q1' `
    ('{0}|{1}|{2}|{3}' -f $m.Kind, $m.Site, $m.LibraryPath, $m.WebFolder) 'OneDrive map: own OneDrive, site, library and web folder'
$m = Get-OneDriveWebMap ($odHome + '\Team X - Files\Week 2') $odEntries @($odHome)
Assert-Equal 'Shortcut|https://org.example.com/sites/TeamX|/sites/TeamX/Shared Documents/Plans/Week 2' ('{0}|{1}|{2}' -f $m.Kind, $m.Site, $m.WebFolder) `
    'OneDrive map: a shortcut in the own OneDrive (the longest mount point wins)'
$m = Get-OneDriveWebMap 'C:\Users\me\Org\Team Y - Docs' $odEntries @($odHome)
Assert-Equal 'Library|/teams/TeamY/Shared Documents|/teams/TeamY/Shared Documents' ('{0}|{1}|{2}' -f $m.Kind, $m.LibraryPath, $m.WebFolder) `
    'OneDrive map: a synced library, its own top folder (escaped address read)'
Assert-Equal $true ($null -eq (Get-OneDriveWebMap 'C:\Users\me\OneDrive - Orgx\a' $odEntries @($odHome))) 'OneDrive map: a folder with a longer name is not inside'
Assert-Equal $true ($null -eq (Get-OneDriveWebMap 'D:\Share\Ops' $odEntries @($odHome))) 'OneDrive map: a folder OneDrive does not sync has no map'
Assert-Equal $true ($null -eq (Get-OneDriveWebMap ($odHome + '\a') @(@{ MountPoint = $odHome; UrlNamespace = 'not an address' }) @($odHome))) `
    'OneDrive map: an entry with no web address has no map'

# ---- Kinds and fall-back (no scan) ----
Reset-FakePnP @{}
$web = Connect-OneDriveWeb ($odHome + '\Team X - Files') $odEntries @($odHome) 6>$null
Assert-Equal 'False|0' ('{0}|{1}' -f $web.Used, @(Get-FakePnPConnection).Count) 'OneDrive web: a shortcut is not used yet, no sign-in'
Assert-Equal $true ($web.Note -like 'Not used: this kind of OneDrive folder*history is not available*') 'OneDrive web: the note says why'
$web = Connect-OneDriveWeb 'D:\Share\Ops' $odEntries @($odHome) 6>$null
Assert-Equal 'False|' ('{0}|{1}' -f $web.Used, $web.Note) 'OneDrive web: a plain drive folder gets no note'
Set-FakePnPOption 'ConnectFails' $true
$web = Connect-OneDriveWeb ($odHome + '\Ops') $odEntries @($odHome) 6>$null
Assert-Equal 'False|True' ('{0}|{1}' -f $web.Used, (Test-OneDriveSignIn)) 'OneDrive web: a failed sign-in is not used, and the run still signs out'
Assert-Equal $true ($web.Note -like 'Not used (*). The who/when history is not available*') 'OneDrive web: a failed sign-in is named in the note'

# ---- A scan: the history from the web, a web-only file read into memory, nothing outside the folder ----
$letters = -join @(1..12 | ForEach-Object { [char](97 + (Get-Random -Maximum 26)) })
$odRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-onedrive-' + $letters)
$odOut = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-onedriveout-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path (Join-Path $odRoot 'Ops') -Force | Out-Null
[System.IO.File]::WriteAllBytes((Join-Path (Join-Path $odRoot 'Ops') 'notes.txt'), $Txt)
$lib = '/sites/TeamA/Shared Documents'
Reset-FakePnP @{
    ($lib + '/Ops/notes.txt') = $Txt
    ($lib + '/Ops/Sub/webonly.txt') = [System.Text.Encoding]::UTF8.GetBytes("Card on file. SSN: $($Sample.Ssn)")
    ($lib + '/Other/outside.txt') = [System.Text.Encoding]::UTF8.GetBytes("SSN: $($Sample.Ssn)")
}
Set-FakeSiteAnswer
$odVersionUrls = New-Object System.Collections.Generic.List[string]
Set-Item function:global:Get-PnPFileVersion -Value ({ param([string]$Url) $odVersionUrls.Add($Url) }.GetNewClosure())   # no old versions; records the files
try {
    $odSync = @(@{ MountPoint = $odRoot; UrlNamespace = 'https://tenant.example.com' + $lib })
    $book = Invoke-DriveScan -Path (Join-Path $odRoot 'Ops') -OutputFolder $odOut -MaxFileMB 25 -Init $null -SyncEntries $odSync -OneDriveHomes @($odRoot) 6>$null
    $sheets = Get-XlsxSheets $book
    $hits = @($sheets['Hits'] | Select-Object -Skip 1)
    $files = @($hits | ForEach-Object { $_[4] } | Sort-Object -Unique)
    Assert-Equal 'notes.txt|webonly.txt' ($files -join '|') 'OneDrive scan: the disk file and the web-only file are read; the file outside the folder is not'
    Assert-Equal 'https://tenant.example.com/sites/TeamA' (@(Get-FakePnPConnection) -join ' ') 'OneDrive scan: one sign-in, to the site of the folder'
    Assert-Equal '0|0' ('{0}|{1}' -f @($odVersionUrls | Where-Object { $_ -like '*outside.txt' }).Count, (Get-FakeCallCount 'Online question')) 'OneDrive scan: nothing outside the folder is read, no download question'
    Assert-Equal 2 @($odVersionUrls | Sort-Object -Unique).Count 'OneDrive scan: the history of each file with a hit is read from the web'
    $header = @($sheets['Hits'])[0]; $noteColumn = [array]::IndexOf(@($header), 'History note')
    Assert-Equal 0 @($hits | Where-Object { $_[$noteColumn] -eq 'A drive keeps no file history' }).Count 'OneDrive scan: no hit says "a drive keeps no history"'
    $summary = (@($sheets['Summary']) | ForEach-Object { @($_) -join ' ' }) -join "`n"
    Assert-Equal $true ($summary -match 'OneDrive on the web Used: 2 files in this folder on the web') 'OneDrive scan: the Summary says the web was used'
    $linkColumn = [array]::IndexOf(@($header), 'Link')
    $webLinks = @($hits | Where-Object { $_[3] -eq 'Sub' -and $_[4] -eq 'webonly.txt' -and $_[$linkColumn] -like 'https://tenant.example.com/*' })
    Assert-Equal $true ($webLinks.Count -ge 1) 'OneDrive scan: a web-only file links to its web copy'
}
finally {
    Remove-Item function:global:Get-PnPFileVersion -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $odRoot -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $odOut -Recurse -Force -ErrorAction SilentlyContinue
}

# An online-only file with a web copy is not asked about (it is read from the web); one with no web copy still is.
Reset-FakePnP @{}
$groups = New-Object System.Collections.Generic.List[object]
$groups.Add(@{ Name = 'Ops'; Files = (New-Object System.Collections.Generic.List[object]) })
$groups[0].Files.Add(@{ Path = 'x'; Inner = 'a.txt'; Name = 'a.txt'; Size = 10; Modified = ''; Online = $true; Web = @{ FileRef = '/x/Ops/a.txt' } })
$rowsOnline = New-Object System.Collections.Generic.List[object]
Select-OnlineFiles $groups.ToArray() $rowsOnline 25 6>$null
Assert-Equal '0|1' ('{0}|{1}' -f (Get-FakeCallCount 'Online question'), $groups[0].Files.Count) 'OneDrive web: an online-only file with a web copy is not asked about'
$groups[0].Files.Add(@{ Path = 'y'; Inner = 'b.txt'; Name = 'b.txt'; Size = 10; Modified = ''; Online = $true })
Select-OnlineFiles $groups.ToArray() $rowsOnline 25 6>$null
Assert-Equal '1|1' ('{0}|{1}' -f (Get-FakeCallCount 'Online question'), $groups[0].Files.Count) 'OneDrive web: one with no web copy is still asked about (Enter = not read)'
