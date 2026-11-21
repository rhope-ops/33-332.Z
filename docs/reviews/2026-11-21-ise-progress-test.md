# ISE progress-bar layout test (2026-11-21)

Why: on the gov run of 2026-11-21 the PAM saw the scan bar as 2 lines in Windows PowerShell ISE: line 1 the folder with a "." after it, line 2 indented with the counts and the file at its end. ISE draws the bar itself, so the "." and the indent may come from ISE, not from Ion. This test shows 5 layouts, 6 seconds each, so the PAM can say which one ISE draws as wanted: no "." after the folder, no indent, the file on its own line, and "Downloading...".

## What it does

- Nothing but progress bars with invented text (a folder "General", a file "Roster.docx"). No sign-in, no network, no file read or written.
- Before each layout it prints its letter on screen ("Layout A now"). It ends by itself after about 30 seconds.

## How to run it

Test 1 is done (result below). These were its steps:

1. On the gov device, open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, paste the test 1 block into the script pane. You see the script. Go on.
3. Press **F5**. You see "Layout A now" in the console pane. Wait until "Layout test done." shows.
4. In the chat, type the letter of the layout that looked right. You see your message in the chat. The task is complete.

```powershell
# Ion ISE progress-bar layout test: invented text only, no sign-in, no network, no files. About 30 seconds.
$stats = '12 of 40 files (30%) | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 elapsed, about 01:48 left'
$file = 'File: General/Roster.docx'
$layouts = @(
    @{ Name = 'A (the new Ion build: the file after a line break in the status)'; Activity = 'Folder 1 of 2: General'; Status = ($stats + "`n" + $file) },
    @{ Name = 'B (all 3 lines in the activity, status one space)'; Activity = ('Folder 1 of 2: General' + "`n" + $stats + "`n" + $file); Status = ' ' },
    @{ Name = 'C (folder line ends in a space; file as CurrentOperation, the old build)'; Activity = 'Folder 1 of 2: General '; Status = $stats; Operation = $file },
    @{ Name = 'D (download bar, the new build)'; Activity = 'Downloading...'; Status = '6.0 of 12.0 MB (50%) | 1.5 MB/s | about 00:04 left' },
    @{ Name = 'E (download bar, two dots)'; Activity = 'Downloading..'; Status = '6.0 of 12.0 MB (50%) | 1.5 MB/s | about 00:04 left' }
)
foreach ($layout in $layouts) {
    Write-Host ('Layout ' + $layout.Name + ' now')
    for ($step = 0; $step -lt 6; $step++) {
        $percent = 20 + 10 * $step
        if ($layout.ContainsKey('Operation')) {
            Write-Progress -Id 1 -Activity $layout.Activity -Status $layout.Status -PercentComplete $percent -CurrentOperation $layout.Operation
        }
        else { Write-Progress -Id 1 -Activity $layout.Activity -Status $layout.Status -PercentComplete $percent }
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $layout.Activity -Completed
}
Write-Host 'Layout test done.'
```

## Result of test 1 (the PAM, 2026-11-21)

- "e looks right, the others don't put 'file' on its own 3rd line. the break doesn't work."
- So ISE adds a "." after a bar's first line: the download bar now says "Downloading.." and shows "Downloading..." (src/ChannelScan.psm1 $DownloadStep).
- A line break in the text does not show in ISE. The scan bar keeps the file as CurrentOperation until test 2 decides.

## Test 2: the file on a line of its own

Each layout uses a second bar for the file. The test takes about 24 seconds.

1. On the gov device, open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, paste the test 2 block below into the script pane. You see the script. Go on.
3. Press **F5**. You see "Layout F now" in the console pane and a bar at the top. Go on.
4. Look at the bar at the top while each letter shows. You see the next letter every 6 seconds. Wait until "Layout test 2 done." shows.
5. In the chat, type the letter (F to I) that put the file on its own line. You see your message in the chat. Go on.
6. If that layout's folder line ended in a ".", type `dot` in the chat. You see your message in the chat. The task is complete.

```powershell
# Ion ISE progress-bar layout test 2: invented text only, no sign-in, no network, no files. About 24 seconds.
$stats = '12 of 40 files (30%) | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 elapsed, about 01:48 left'
$path = 'General/Roster.docx'
$layouts = @(
    @{ Name = 'F (second bar: the file as its heading, status one space)'; FileActivity = ('File: ' + $path); FileStatus = ' ' },
    @{ Name = 'G (second bar: heading "File", the path below it)'; FileActivity = 'File'; FileStatus = $path },
    @{ Name = 'H (second bar: the file as its heading, no status given)'; FileActivity = ('File: ' + $path); FileStatus = $null },
    @{ Name = 'I (one bar: folder and counts on line 1, the file on line 2)'; FileActivity = $null; FileStatus = $null }
)
foreach ($layout in $layouts) {
    Write-Host ('Layout ' + $layout.Name + ' now')
    for ($step = 0; $step -lt 6; $step++) {
        $percent = 20 + 10 * $step
        if ($layout.Name.StartsWith('I')) {
            Write-Progress -Id 1 -Activity ('Folder 1 of 2: General | ' + $stats) -Status ('File: ' + $path) -PercentComplete $percent
        }
        else {
            Write-Progress -Id 1 -Activity 'Folder 1 of 2: General' -Status $stats -PercentComplete $percent
            if ($null -ne $layout.FileStatus) { Write-Progress -Id 3 -Activity $layout.FileActivity -Status $layout.FileStatus }
            else { Write-Progress -Id 3 -Activity $layout.FileActivity }
        }
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 3 -Activity 'File' -Completed
    Write-Progress -Id 1 -Activity 'Folder' -Completed
}
Write-Host 'Layout test 2 done.'
```

## Result of test 2

- Waiting on the PAM.
