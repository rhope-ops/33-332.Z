# E-signature shape test (2026-11-23)

The PAM approved this test with the plan of 2026-11-23 (docs/reviews/2026-11-23-ion-esignature-onedrive-plan.md, step 4 of
the request). Goal: check the e-signature fix against the text that Ion's own PDF reader gets from 2 or 3 real e-signed
PDFs, without a name or a number leaving the gov device.

## What it does

- Read-only. It reads at most 3 PDFs in the one folder that you give (no subfolders). It writes nothing.
- It uses the Ion text that you copy (the fixed version on the branch of this follow-up pull request), so it uses Ion's own PDF reader and
  Ion's own detectors. It loads them in memory only.
- For each run of 9 or 10 digits (at most 10 for each file), it prints the shape of the text around it: A = capital
  letter, a = small letter, 9 = digit, _ = space, | = line break, ? = any other character. Signature words (Digitally,
  signed, by, Date, Reason, CN, OU and others) stay as words. Then it prints what Ion finds there (for example
  "DoD ID/Low"). It never prints a name, a number or a file name (Data boundary rule).

## How to run it

The test is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-23-esignature-shape-test.txt`.

1. On the gov device, open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
3. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/claude/ion-esig-gaps/docs/reviews/2026-11-23-esignature-shape-test.txt`. You see the test as text. Go on.
4. In the browser page with the test, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the test text selected. Go on.
5. Do not press **F5** until the Ion text is on the clipboard. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the test text in the pane. Go on.
6. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/claude/ion-esig-gaps/paste/Ion-script.txt`. You see the Ion script as text. Go on.
7. Do not copy other text until ISE shows the question **Folder with 2 or 3 e-signed PDFs**. In the browser page with the Ion script, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the Ion text selected. Go on.
8. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Folder with 2 or 3 e-signed PDFs (only this folder is read)**. Go on.
9. On the question **Folder with 2 or 3 e-signed PDFs**, do these actions:
   1. Type the path of a folder with 2 or 3 e-signed PDFs, for example `C:\Users\<you>\Documents\EsignTest`.
   2. Press **Enter**.
   You see the line **File 1: ...** and shape lines. Wait until the line **Done.** shows.
10. In the ISE output pane, do these actions:
    1. Select the lines from **File 1** to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
11. In the chat with the agent, do these actions:
    1. Press **Ctrl+V**.
    2. Send the message.
    You see the lines in the chat. The task is complete.

If you see **Result: the Ion script text is not on the clipboard**, do the actions of the browser page with the Ion script and of **F5** again: they put the Ion text on the clipboard and start the test.
