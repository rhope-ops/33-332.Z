# E-signature shape test (2026-11-23)

The PAM approved this test with the plan of 2026-11-23 (docs/reviews/2026-11-23-ion-esignature-onedrive-plan.md, step 4 of
the request). Goal: check the e-signature fix against the text that Ion's own PDF reader gets from 2 or 3 real e-signed
PDFs, without a name or a number leaving the gov device.

## What it does

- Read-only. It reads at most 3 PDFs in the one folder that you give (no subfolders). It writes nothing.
- It uses the Ion text that you copy (the fixed version on the branch of PR 57), so it uses Ion's own PDF reader and
  Ion's own detectors. It loads them in memory only.
- For each run of 9 or 10 digits (at most 10 for each file), it prints the shape of the text around it: A = capital
  letter, a = small letter, 9 = digit, _ = space, | = line break, ? = any other character. Signature words (Digitally,
  signed, by, Date, Reason, CN, OU and others) stay as words. Then it prints what Ion finds there (for example
  "DoD ID/Low"). It never prints a name, a number or a file name (Data boundary rule).

## How to run it

1. On the gov device, open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
3. On this page, do these actions:
   1. Select all the text in the box below.
   2. Press **Ctrl+C**.
   You see no change: the test text is on the clipboard. Go on.
4. Do not press **F5** until the Ion text is on the clipboard. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the test text in the pane. Go on.
5. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/ccr-9805d725-onoreq/paste/Ion-script.txt`. You see the Ion script as text. Go on.
6. Do not copy other text until ISE shows the question **Folder with 2 or 3 e-signed PDFs**. In the browser page with the Ion script, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the Ion text selected. Go on.
7. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Folder with 2 or 3 e-signed PDFs (only this folder is read)**. Go on.
8. On the question **Folder with 2 or 3 e-signed PDFs**, do these actions:
   1. Type the path of a folder with 2 or 3 e-signed PDFs, for example `C:\Users\<you>\Documents\EsignTest`.
   2. Press **Enter**.
   You see the line **File 1: ...** and shape lines. Wait until the line **Done.** shows.
9. In the ISE output pane, do these actions:
   1. Select the lines from **File 1** to **Done.**.
   2. Press **Ctrl+C**.
   You see the lines selected. Go on.
10. In the chat with the agent, do these actions:
    1. Press **Ctrl+V**.
    2. Send the message.
    You see the lines in the chat. The task is complete.

If you see **Result: the Ion script text is not on the clipboard**, do the actions of the browser page and of **F5** again: they put the Ion text on the clipboard and start the test.

```powershell
# Ion e-signature shape test: read-only, one folder, at most 3 PDFs. Prints shapes only: A = capital, a = small letter,
# 9 = digit, _ = space, | = line break. No name, no number and no file name is shown. Nothing is saved.
$ionText = Get-Clipboard -Raw
$block = [regex]::Match([string]$ionText, '(?s)\$IonReaderBlock = \{\r?\n(.*)\r?\n\}\r?\n\$null = New-Module -Name IonReaders')
if (-not $block.Success) { Write-Host 'Result: the Ion script text is not on the clipboard. Copy all of it again, then press F5.'; return }
$null = New-Module -Name IonShapeTest -ScriptBlock ([scriptblock]::Create($block.Groups[1].Value)) | Import-Module -Force
$folder = (Read-Host 'Folder with 2 or 3 e-signed PDFs (only this folder is read)').Trim().Trim('"')
$pdfs = @(Get-ChildItem -LiteralPath $folder -Filter '*.pdf' -File -ErrorAction Stop | Select-Object -First 3)
if ($pdfs.Count -eq 0) { Write-Host 'Result: no PDF in that folder.'; return }
$keep = '(?i)^(?:digitally|signed|by|date|reason|signature|location|cn|ou|dn|o|c|dod|pki|usaf|usa|jr|sr|ii|iii)$'
$fileNumber = 0
foreach ($pdf in $pdfs) {
    $fileNumber++
    $read = Get-TextFromBytes ([System.IO.File]::ReadAllBytes($pdf.FullName)) $pdf.Name
    if ($read.Status -ne 'OK') { Write-Host ("File {0}: not read ({1})" -f $fileNumber, $read.Status); continue }
    $text = [string]$read.Text
    $hits = @(Get-PiiSpan $text)
    $runs = @([regex]::Matches($text, '\d(?:(?:[ \t-]{1,2}|[ \t]*\r?\n[ \t]*)?\d){8,9}') | Select-Object -First 10)
    Write-Host ("File {0}: {1} characters, {2} runs of 9 or 10 digits" -f $fileNumber, $text.Length, $runs.Count)
    foreach ($run in $runs) {
        $start = [Math]::Max(0, $run.Index - 60); $end = [Math]::Min($text.Length, $run.Index + $run.Length + 30)
        $window = $text.Substring($start, $end - $start)
        $shape = [regex]::Replace($window, '[A-Za-z]+', { param($w) if ($w.Value -match $keep) { $w.Value } else { ($w.Value -creplace '[A-Z]', 'A') -creplace '[a-z]', 'a' } })
        $shape = ((($shape -replace '\d', '9') -replace '\r?\n', '|') -replace ' ', '_') -replace '[^\x21-\x7E]', '?'
        $found = @($hits | Where-Object { $_.Index -lt $run.Index + $run.Length -and $run.Index -lt $_.Index + $_.Length } | ForEach-Object { $_.Detector + '/' + $_.Confidence })
        if ($found.Count -eq 0) { $found = @('no hit') }
        Write-Host ("  {0}  ->  Ion: {1}" -f $shape, ($found -join ', '))
    }
}
Write-Host 'Done. The lines above show no name and no number. Copy them and paste them in the chat.'
```
