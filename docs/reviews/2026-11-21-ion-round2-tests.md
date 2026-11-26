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
.txt file (Project rules, "Paste format", 2026-11-23): `docs/reviews/2026-11-25-view-count-test-4.txt`.

1. In your web browser, on your team site, find a file that people opened in the last 90 days.
2. In your web browser, open the folder of that file. Write down the file name as the column **Name** shows it, for example `Report.pdf`.
3. Close every **Windows PowerShell ISE** window.
4. Open **Windows PowerShell ISE**, then click **File** > **New**. The empty script pane **Untitled1.ps1** opens.
5. In your web browser, open a new tab, type `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-view-count-test-4.txt`, then press **Enter**. The first line starts with `# Ion test 4`.
6. In the tab with the test, press `Ctrl+A`, then press `Ctrl+C`.
7. In the script pane **Untitled1.ps1**, press `Ctrl+V`, then press `F5`. The **Console pane** shows **Question 1**.
8. In your web browser, click the tab of the folder page, click the address bar, then press `Ctrl+C`.
9. In the **Console pane** of ISE, below **Question 1**, press `Ctrl+V`, then press **Enter**. **Question 2** shows.
10. In the **Console pane**, below **Question 2**, type the file name that you wrote down, then press **Enter**.
11. Make sure that the **Console pane** clears and then shows `T4 test version=2026-11-25e`. If it shows another version, close ISE and start this list again.
12. While it runs, if a window asks if you want to continue running scripts, click **Yes**.
13. While it runs, if a sign-in window opens, sign in with your CAC. This can take 2 minutes or more.
14. Wait until the line **Done.** shows in the **Console pane** of ISE. Do not copy before then.
15. In the **Console pane**, select the lines from `T4 test version=2026-11-25e` to **Done.**, then press `Ctrl+C`.
16. In Outlook, click **New Email**. In **To**, type your own e-mail address. In **Subject**, type `test 4`.
17. Click in the body of the e-mail, press `Ctrl+V`, then click **Send**.

## What the results mean

| Result | Meaning | Next step |
|---|---|---|
| 3a and 3b work | A copy of the sign-in downloads in workers. The two times show whether 3 at once beat one by one. | Step 3: 2 to 4 parallel downloads, largest first. |
| 3a works, 3b failed | The copy of the sign-in does not work in a worker. | Step 3 needs another way (the agent proposes one); downloads stay one by one. |
| T4a or T4b works | SharePoint on DoD365 gives view counts. | Ion adds a "Views since PII first saved" column (last 90 days at most). |
| T4a and T4b failed or no answer | No view counts on DoD365 through usaf.dps.mil. | No views column. |

## Results

- Test 3: not run. The PAM, 2026-11-24: build with fallback; the header line "Parallel downloads" of a normal Ion run
  is the evidence (docs/reviews/2026-11-24-parallel-downloads-design.md). Test 4: as a .txt file since 2026-11-25; version 2026-11-25e takes the folder from the address of the folder page and asks for the file name only (the web page shows the library as "Documents", the address has "Shared Documents"; review of PR 125). Stub run on PowerShell 7.4 with invented data: 4 cases end with "Done.". Waiting on the PAM's run.
- Test 4 result (version 2026-11-25e, the PAM's e-mail, received 2026-11-25 19:51): folder from id, sign-in ok,
  file found by name; `T4a failed: PSInvalidOperationException`, `T4b failed: PSInvalidOperationException`. By the
  table above: no view counts on DoD365 through usaf.dps.mil, so no views column. Confidence: medium (1 run, 1 file;
  the test prints the error type only, so a 403 and a 404 look the same).
