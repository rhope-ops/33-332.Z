# ISE progress-bar layout test (2026-11-21)

Why: on the gov run of 2026-11-21 the PAM saw the scan bar as 2 lines in Windows PowerShell ISE: line 1 the folder with a "." after it, line 2 indented with the counts and the file at its end. ISE draws the bar itself, so the "." and the indent may come from ISE, not from Ion. This test shows 5 layouts, 6 seconds each, so the PAM can say which one ISE draws as wanted: no "." after the folder, no indent, the file on its own line, and "Downloading...".

## What it does

- Nothing but progress bars with invented text (a folder "General", a file "Roster.docx"). No sign-in, no network, no file read or written.
- Before each layout it prints its letter on screen ("Layout A now"). It ends by itself after about 30 seconds.

## How to run it

1. Open Windows PowerShell ISE on the gov device.
2. Copy the whole block below, paste it into the script pane, and press F5.
3. Watch the bar at the top of the console pane while each letter is shown.
4. Tell the agent, for each letter A to E: how many lines, any "." added, any indent, and where the file line is. Or just the letter that looks right.

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

## Result

- Waiting on the PAM. When the answer is in, Ion's scan bar (src/ScanQueue.psm1 Write-ScanBar) and download bar (src/ChannelScan.psm1 $DownloadStep) take the layout that matches.
