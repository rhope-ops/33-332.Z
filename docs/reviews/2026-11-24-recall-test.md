# Recall test (2026-11-24)

The PAM approved this test on 2026-11-24. Goal: the first measured catch rate of Ion. The test plants known invented
values in 12 files. Ion scans the files. Then the test compares Ion's workbook with the key of planted values.

## What it does

- The test is one .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-24-recall-test.txt`. It has two modes.
- Make mode writes 12 invented files into one empty folder: .txt, .csv, .docx, .xlsx and .pdf. They hold 53 planted values of 8 types, and decoys.
- The values come from a fixed seed. The files and the key are the same on each run. The text of the test holds no whole value.
- Score mode reads the Hits and Not scanned tabs of Ion's workbook. It does not use Excel, and it does not change the workbook.
- Score matches each hit to the key by file and type. It uses the last 4 digits, the last 3 passport characters, or the e-mail domain.
- Ion masks a date of birth and a home address completely. For these 2 types, Score counts the hits in each file.
- A hit that does not match the key is a false hit. A decoy that Ion reports is a false hit.
- Score prints counts and type names only. It prints no value, no path and no name.

| Type | Planted | Forms |
|---|---|---|
| SSN | 9 | labelled; dashed with no label; 9 digits with no label; dashed under a header row; a number cell with a label |
| DoD ID | 10 | labelled ("DoD ID", "EDIPI"); e-signature (LAST.FIRST.M. then the ID); after "Digitally signed by"; a number cell; under a header row |
| Date of birth | 7 | after "DOB", "Date of birth" or "Born"; under a header row |
| Payment card | 4 | Visa, Mastercard and American Express numbers that pass the Luhn check, with and without the label "Card" |
| Passport | 3 | a letter and 8 digits (2), 9 digits (1), each after "Passport" |
| Phone | 8 | fictional 555-01xx numbers, with "Cell", "Mobile" or "Home", and in a CSV column |
| Personal e-mail | 7 | addresses at public mail providers |
| Home address | 5 | number, street, city, state and ZIP code, with and without a label |

Decoys (they must not be hits): order numbers, dates with no label, 10-digit tracking numbers in text, an SSN shape with area 9xx, and a 16-digit number that fails the Luhn check.

## How to run it

Part 1: make the test folder.

1. On the gov device, open **File Explorer**. You see the window **File Explorer**. Go on.
2. In the left pane of **File Explorer**, click your OneDrive folder. You see the files in your OneDrive. Go on.
3. In the window **File Explorer**, do these actions:
   1. Click **New**.
   2. Click **Folder**.
   3. Type `RecallTest`.
   4. Press **Enter**.
   You see the new folder **RecallTest**. Go on.

Part 2: make the test files. If you do not see **Copy as path** in a menu, close the menu. Then hold **Shift** and right-click again.

4. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
5. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
6. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-24-recall-test.txt`. You see the test as text. Go on.
7. In the browser page with the test, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the test text selected. Go on.
8. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the test text in the pane. Go on.
9. In the window **File Explorer**, do these actions:
   1. Right-click the folder **RecallTest**.
   2. Click **Copy as path**.
   The menu closes. Go on.
10. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Type M to make the test files, or S to score the Ion workbook**. Go on.
11. On the question **Type M to make the test files**, do these actions:
    1. Type `M`.
    2. Press **Enter**.
    You see the question **Path of the empty test folder**. Go on.
12. On the question **Path of the empty test folder**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see the line **Done: 12 test files with 53 planted values are in the folder**. Go on.

If you see **Result: that folder is not empty**, delete all files in the folder **RecallTest**. Then press **F5** in the window **Windows PowerShell ISE**, type `M`, press **Enter**, press **Ctrl+V** and press **Enter**.

Part 3: scan the test folder with Ion. If you see the question **Download and read them? [y/N]**, type `y` and press **Enter**. Ion asks this when OneDrive keeps a test file online only.

13. In the window **Windows PowerShell ISE**, do these actions:
    1. Click **File**.
    2. Click **New**.
    You see an empty script pane **Untitled2.ps1**. Go on.
14. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/paste/Ion-script.txt`. You see the Ion script as text. Go on.
15. In the browser page with the Ion script, do these actions:
    1. Press **Ctrl+A**.
    2. Press **Ctrl+C**.
    You see all the Ion text selected. Go on.
16. In the script pane **Untitled2.ps1**, press **Ctrl+V**. You see the Ion text in the pane. Go on.
17. In the window **File Explorer**, do these actions:
    1. Right-click the folder **RecallTest**.
    2. Click **Copy as path**.
    The menu closes. Go on.
18. In the window **Windows PowerShell ISE**, press **F5**. You see the question **What to scan**. Go on.
19. On the question **What to scan**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see the question **Add another (1 so far), or press Enter to start**. Go on.
20. On the question **Add another**, press **Enter**. You see the question **Scan now? [Y/n]**. Go on.
21. On the question **Scan now?**, press **Enter**. You see the progress bars. Wait until the question **Open the workbook now? [Y/n]** shows.
22. On the question **Open the workbook now?**, do these actions:
    1. Type `n`.
    2. Press **Enter**.
    You see the question **Scan more: paste it, or press Enter to exit**. Go on.
23. On the question **Scan more**, press **Enter**. You see the line **Ion is done.** Go on.

Part 4: score the Ion workbook. If you do not see **Copy as path** in a menu, close the menu. Then hold **Shift** and right-click again.

24. In the window **File Explorer**, do these actions:
    1. In the left pane, click **Documents**.
    2. Double-click the folder **Ion**.
    3. Double-click the folder **RecallTest**.
    You see the file **Ion RecallTest** with the date and time. Go on.
25. In the window **File Explorer**, do these actions:
    1. Right-click the file **Ion RecallTest**.
    2. Click **Copy as path**.
    The menu closes. Go on.
26. In the window **Windows PowerShell ISE**, click the tab **Untitled1.ps1**. You see the test text. Go on.
27. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Type M to make the test files, or S to score the Ion workbook**. Go on.
28. On the question **Type M to make the test files**, do these actions:
    1. Type `S`.
    2. Press **Enter**.
    You see the question **Path of the Ion workbook (.xlsx)**. Go on.
29. On the question **Path of the Ion workbook**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see the line **Ion recall test** and a table. Wait until the line **Done.** shows.
30. In the ISE output pane, do these actions:
    1. Select the lines from **Ion recall test** to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
31. In an e-mail to yourself, do these actions:
    1. Press **Ctrl+V**.
    2. Send the message.
    You see the message in your sent mail. The task is complete.

## What the lines mean

- Planted: the values of this type in the key. Found: the hits that match a planted value. Missed: Planted minus Found.
- False: the hits that match no planted value (a decoy, a wrong type, or a file that is not a test file).
- Recall: Found divided by Planted. Precision: Found divided by all hits.
- "By form" shows each type again, split by how the value is written (for example "labelled" or "header row only").

## Baseline (pwsh 7, Linux, invented files)

Run on 2026-11-24 by `tests/Recall.Tests.ps1` (part of `tests/Run-Tests.ps1`): Make into a temporary folder, then
`Invoke-DriveScan` with the reader block of `paste/Ion-script.txt` in the reader workers (as on the gov device), then
Score. Ion: `src/` of origin/main 9ded821, not changed. PowerShell 7.4.6 on Linux. The test prints the line
`RECALL found=46/53 false=1`.

| Type | Planted | Found | Missed | False | Recall |
|---|---|---|---|---|---|
| SSN | 9 | 9 | 0 | 1 | 100% |
| DoD ID | 10 | 7 | 3 | 0 | 70% |
| Date of birth | 7 | 4 | 3 | 0 | 57% |
| Payment card | 4 | 4 | 0 | 0 | 100% |
| Passport | 3 | 2 | 1 | 0 | 67% |
| Phone | 8 | 8 | 0 | 0 | 100% |
| Personal e-mail | 7 | 7 | 0 | 0 | 100% |
| Home address | 5 | 5 | 0 | 0 | 100% |
| All types | 53 | 46 | 7 | 1 | 87% |

Recall 87% (46 of 53). Precision 98% (46 of 47 hits). All 12 files were read (0 in the Not scanned tab). No decoy was a hit.

After the passport rule (PR 97, 2026-11-24: 9 digits after a Passport label are a Passport, approved by the PAM): the test prints `RECALL found=47/53 false=0`. The baseline table and the list below describe the run before this change.

Without the 9 values under a header row only: recall 98% (43 of 44), precision 98% (43 of 44 hits). Without the 9-digit passport too: 43 of 43, no false hit.

The 7 misses and the 1 false hit:

- DoD ID, header row only: 0 of 3. The workbook has the label "DoD ID" only in the header row of its column. Ion looks for the label at most 40 characters before the number.
- Date of birth, header row only: 0 of 3. The same layout; Ion looks for a date-of-birth label at most 30 characters before the date.
- Passport, 9 digits: 0 of 1. Ion reports the 9 digits as an SSN (Low), because SSN wins over Passport when two hits overlap. This is also the 1 false hit. Fixed by PR 97.
- Every other form: all found. This includes both e-signature forms, numbers in number cells, the CSV columns and the PDF text.

The baseline is a measurement only: the detectors were not changed for it. PR 97 changed the passport rule after it.

After the roster fix (2026-11-24, branch `claude/roster-labels`; the detectors are not changed): `RECALL found=52/53 false=1`.
SSN 9 of 9, DoD ID 10 of 10, Date of birth 7 of 7, Passport 2 of 3; the other types as above. Recall 98% (52 of 53),
precision 98% (52 of 53 hits). A PII label only in a column header (worksheet, .csv, .tsv) is now written before each value
below it (`src/ContainerText.psm1`). The 1 miss and the 1 false hit are the 9-digit passport (see above).

Method note: with `-Init $null` (the reader modules loaded one by one in the workers), the PDFs were NOT SCANNED,
because `$script:PdfTokenPattern` was not set in the worker. The bundle's reader block reads them. The baseline uses
the bundle's reader block, the same code as on the gov device.

## Result (gov device)

Not run yet.
