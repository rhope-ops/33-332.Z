# Device probe 3 (2026-11-24)

Probe 2 showed: the OneDrive folder maps to its web address, but the web path did not find the file; and a second
sign-in after a sign-out failed (PSInvalidOperationException). The PAM also saw a 404 with a Teams link and a PIN
prompt. Probe 3 tests how Ion can sign in once for several sites on the same host, which site each read really
reaches, and which web path finds a OneDrive file. Ion's next sign-in change uses only the method that passes here.

## What it does

- Read-only. It writes nothing and changes no setting.
- It asks 4 questions first, then clears the screen. The lines you copy hold no answer.
- It prints only True or False, counts, error kinds (for example `PSInvalidOperationException/NoConnection`) and times.
  It prints no name, no address, no path and no value from a file.
- It touches only the folder and the sites you give.

## Before you run it: get 2 site addresses and 1 Teams link

Use 2 teams that you scan. Keep the results in Notepad.

1. On the gov device, open Notepad. You see an empty Notepad window. Go on.
2. In Teams, in the channel **General** of the first team, click **Files**. You see the files of the channel. Go on.
3. Click **Open in SharePoint**. You see the files in your web browser. Go on.
4. In the web browser, do these actions:
   1. Click the address bar.
   2. Press **Ctrl+C**.
   You see the address selected. Go on.
5. In Notepad, do these actions:
   1. Press **Ctrl+V**.
   2. Press **Enter**.
   You see address 1. Go on.
6. In Teams, in the channel **General** of the second team, click **Files**. You see the files of the channel. Go on.
7. Click **Open in SharePoint**. You see the files in your web browser. Go on.
8. In the web browser, do these actions:
   1. Click the address bar.
   2. Press **Ctrl+C**.
   You see the address selected. Go on.
9. In Notepad, do these actions:
   1. Press **Ctrl+V**.
   2. Press **Enter**.
   You see address 2. Go on.
10. In Teams, in the list of teams, do these actions:
    1. Click **...** next to the name of the first team.
    2. Click **Get link to team** (new Teams: **Copy link**).
    3. Click **Copy**.
    You see that the link is copied. Go on.
11. In Notepad, press **Ctrl+V**. You see the Teams link. The task is complete.

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-24-device-probe-3.txt`.

1. On the gov device, close every Windows PowerShell ISE window. You see no ISE window. Go on.
2. Open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
3. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
4. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-24-device-probe-3.txt`. You see the probe as text. Go on.
5. In the browser page with the probe, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the probe text selected. Go on.
6. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
7. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Question 1: your OneDrive folder**. Go on.
8. On the question **Question 1**, do these actions:
    1. Type the path of your OneDrive folder.
    2. Press **Enter**.
    You see **Question 2**. Go on.
9. On the question **Question 2**, do these actions:
    1. Paste address 1 from Notepad.
    2. Press **Enter**.
    You see **Question 3**. Go on.
10. On the question **Question 3**, do these actions:
    1. Paste address 2 from Notepad.
    2. Press **Enter**.
    You see **Question 4**. Go on.
11. On the question **Question 4**, do these actions:
    1. Paste the Teams link from Notepad.
    2. Press **Enter**.
    You see that the screen clears. Go on.
12. In each CAC sign-in window, sign in with your CAC. Write down the number of sign-in windows and PIN prompts. Wait
    until the line **Done.** shows. Do not copy before then.
13. In the ISE output pane, do these actions:
    1. Select the lines from the first line to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
14. In an e-mail to yourself, do these actions:
    1. Press **Ctrl+V**.
    2. Type the number of sign-in windows and the number of PIN prompts.
    3. Send the message.
    You see the message in your sent mail. The task is complete.

To skip a question, press **Enter** without an answer.

## Result

- First run, 2026-11-24 (the PAM, by e-mail; kinds and counts only): Questions 2-4 were not used (the probe accepted
  only a bare site address, and the Teams link had no groupId), so T1-T3 did not run. T4: 2 matching entries; sign-in
  ok; paths A, B and C found=False; web read failed (PSInvalidOperationException); disk read "0 KB" (rounded: under
  512 bytes). The cause of found=False is not known yet (assumption, not proven: the file is not on the web, or the
  path is built wrong). Fixed: any address inside a site works; empty, hidden and system files are skipped; a new line
  "T4 name in library" shows if the file is on the web, and if a match is at the path built from the address; the size is in bytes.
- Second run, 2026-11-24 (the PAM, by e-mail; kinds, counts and times only). T2 did not run, because T1 ran.

  | Line | Result |
  |---|---|
  | T4 matching entries / file on disk / library part in the address | 2 / True / True |
  | T4 sign-in | ok, 7567 ms |
  | T4 paths A, B, C | all ok=True found=False (1056, 399, 1190 ms) |
  | T4 name in library | listed=4259, stopped at 5000=False, matches=0, match equals path A=False, library root equals address=True (8568 ms) |
  | T4 web read / disk read | ok=False PSInvalidOperationException/Other / ok=True bytes=215 |
  | T1 root sign-in / site lookup (GetSiteStatus) | ok, 2513 ms / ok, site found=True (1419 ms) |
  | T1 copy to site A: web / rest / list | is target=True / is target=True / ok=False PSInvalidOperationException/Other |
  | T3 copy to site B: web / rest / list | is target=True / is target=True / ok=False PSInvalidOperationException/Other |
  | T3 sign-in B, no sign-out | ok, 600 ms |
  | T3 new B: web / rest / list | is target=True / is target=True / ok=True "no file in the first library" (1213 ms) |
  | T3 sign-out / sign-in B after sign-out | ok, 22 ms / ok, 617 ms |

  The PAM's notes: 2 sign-in windows for 4 sign-ins. Window 1 asked "continue running scripts" (the IE engine of
  `-UseWebLogin`; answer Yes); window 2 did not. No PIN prompt.

  What it proves:
  1. The session copy (`ClientContextExtensions.Clone` + `Set-PnPContext`) reaches the right site for the web and REST
     reads, but the list read fails at both sites. Ion cannot use it.
  2. A plain `Connect-PnPOnline -UseWebLogin` to the next site, without a sign-out, works in 0.6 s, and its web, REST
     and list reads reach that site. This is the method for "one sign-in per host": connect to each site with no
     sign-out between sites, and sign out one time at the end of the run.
  3. The Teams-link path works today: root sign-in, then GetSiteStatus finds the site (no 404 in this run).
  4. The OneDrive address maps to the library: its root equals the address. But the 215-byte file is not in that
     library: 4259 items listed (all of them, under the 5000 cap), 0 matches by folder and name. Paths A, B and C
     found nothing, and the web read fails for the same reason. The probe's path parts are correct (the folder has no
     trailing `\`, so the web path is `/<folder part>/<name>`).
  5. The sign-in after the sign-out worked (617 ms, no window). Probe 2's failure of a second sign-in after a sign-out
     did not occur again.

  What stays open:
  - Which sign-in opened which window is an inference from the times (7.5 s and 2.5 s against 0.6 s), not a fact:
    assumption, medium confidence: one window per host (the OneDrive host for T4, the tenant root host for T1; T3 on
    the same host as T1 needed none).
  - The no-sign-out list read found no file, so file versions, the file location check and a download did not run on
    that connection. Ion's first real run with the change tests them.
  - The order "site A, then site B" with two plain sign-ins did not run (T2 skipped). The tested orders were OneDrive
    host, then root, then site B; Ion's Teams-link path already signs in at the root and then the site.
  - The sign-in lasts after `Disconnect-PnPOnline` (617 ms, no window). So "sign out" ends the PnP connection, but
    the sign-in cookie of the window stays until ISE closes (assumption, medium confidence).
  - Why the OneDrive file is not in the library is not known (assumption, not proven: not uploaded yet, or the folder
    is a shortcut to another library). The web path forms (A, B, C) are still untested. The next probe takes a file
    from the web listing and checks that it exists at the mapped local path.
