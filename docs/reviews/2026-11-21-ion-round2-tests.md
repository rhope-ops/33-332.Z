# Ion round 2 tests: parallel downloads and view counts (2026-11-21)

The PAM approved this plan on 2026-11-21 (docs/DECISIONS.md, newest row). Step 1 asks two things of DoD365 before Ion is changed:

- **Test 3, parallel downloads:** can a copy of your sign-in download a file in a worker (a PowerShell runspace), and do 3 workers at once go faster than one after another? If yes, Ion downloads 2 to 4 files at once (step 3 of the plan).
- **Test 4, view counts:** does SharePoint give the view count of one file on usaf.dps.mil? If yes, Ion can count the views after the "PII first saved" date (SharePoint keeps about the last 90 days).

## What they do

- Read-only. Each test signs in to your own team's site with your CAC, as Ion does, reads one file you name, and signs out.
- Each test touches that one site and that one file: nothing is listed ("No reconnaissance"). Test 4 makes no Microsoft Graph call (commercial Graph is outside the data boundary): its addresses are on usaf.dps.mil only.
- Each prints only result lines: "works" with sizes, times or a count, "no answer", or "failed" with an error type name. It never prints an address, a file name or an error message, so the lines are safe to send in an e-mail.
- If a test hangs for more than a minute, press the red Stop button, close PowerShell ISE (this ends the sign-in), and report "hung".

## How to run test 4

Test 3 does not run (Results, below); its text is in the git history of this file (before 2026-11-25). Test 4 is a
.txt file (Project rules, "Paste format", 2026-11-23): `docs/reviews/2026-11-25-view-count-test-4.txt`. Before you
start, choose 1 file on your team site that people opened in the last 90 days.

1. In your web browser, open the folder of the file on your team site. You see the files of the folder. Write down the name of the file as the column **Name** shows it, for example `Report.pdf`. Go on.
2. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
3. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
4. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
5. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-view-count-test-4.txt`. You see the test as text. The first line starts with `# Ion test 4`. Go on.
6. In the browser page with the test, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the test text selected. Go on.
7. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the test text in the pane. Go on.
8. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Question 1: the address of the folder page in your web browser**. Go on.
9. In your web browser, on the page of the folder, do these actions:
   1. Click the address bar.
   2. Press **Ctrl+C**.
   You see the address selected. Go on.
10. On the question **Question 1** in **Windows PowerShell ISE**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see **Question 2**. Go on.
11. On the question **Question 2**, do these actions:
    1. Type the file name that you wrote down, for example `Report.pdf`.
    2. Press **Enter**.
    You see that the screen clears. Then you see the line `T4 test version=2026-11-25d`. Go on.
12. If a window asks if you want to continue running scripts, click **Yes**. The window closes. Go on.
13. If a sign-in window opens, sign in with your CAC. The window closes. The sign-in can take 2 minutes or more. Go on.
14. In the ISE output pane, wait until the line **Done.** shows. Do not copy before then.
15. In the ISE output pane, do these actions:
    1. Select the lines from the first line to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
16. In an e-mail to yourself, do these actions:
    1. Type `test 4`.
    2. Press **Enter**.
    3. Press **Ctrl+V**.
    4. Click **Send**.
    You see the message in your sent mail. The task is complete.

## What the results mean

| Result | Meaning | Next step |
|---|---|---|
| 3a and 3b work | A copy of the sign-in downloads in workers. The two times show whether 3 at once beat one by one. | Step 3: 2 to 4 parallel downloads, largest first. |
| 3a works, 3b failed | The copy of the sign-in does not work in a worker. | Step 3 needs another way (the agent proposes one); downloads stay one by one. |
| T4a or T4b works | SharePoint on DoD365 gives view counts. | Ion adds a "Views since PII first saved" column (last 90 days at most). |
| T4a and T4b failed or no answer | No view counts on DoD365 through usaf.dps.mil. | No views column. |

## Results

- Test 3: not run. The PAM, 2026-11-24: build with fallback; the header line "Parallel downloads" of a normal Ion run
  is the evidence (docs/reviews/2026-11-24-parallel-downloads-design.md). Test 4: as a .txt file since 2026-11-25; version 2026-11-25d takes the folder from the address of the folder page and asks for the file name only (the web page shows the library as "Documents", the address has "Shared Documents"; review of PR 125). Stub run on PowerShell 7.4 with invented data: 4 cases end with "Done.". Waiting on the PAM's run.
