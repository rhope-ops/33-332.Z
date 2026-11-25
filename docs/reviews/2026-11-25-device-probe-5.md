# Device probe 5 (2026-11-25)

Probe 4, runs A and B (2026-11-25), showed two things. A shortcut in OneDrive maps to its library, but the library has
more than 5000 items, and probe 4 lists only the first 5000. A synced team library has no entry in the OneDrive list
that Ion reads. Ion's OneDrive step (PR 109) asks SharePoint for the items of the one mapped folder only. Probe 5 tests
that call on the gov device, and it checks the other OneDrive list for a synced team library.

## What it does

- Read-only. It writes nothing and changes no setting.
- It prints only True or False, counts, error kinds and times. It prints no name, no address, no path and no value.
- It asks 1 question first, then clears the screen.
- It touches only the folder you give and the one library of that folder. It reads the list details of at most 5000 items in that folder, checks at most 20 files, and reads 1 file of at most 5 MB into memory.
- It reads the OneDrive app's 2 lists in the registry (read-only).

| Line | It shows |
|---|---|
| P5 matching entries | how many OneDrive sync entries hold the folder |
| P5 no OneDrive entry holds this folder | the folder has no sync entry; the next line follows |
| P5 the OneDrive account list holds this folder | True: the other OneDrive list has the folder |
| P5 entry is the OneDrive home folder | True: your own OneDrive; False: a shortcut or a synced library |
| P5 local files | the files on the disk, and how many are online only |
| P5 sign-in | the sign-in to the site of the folder |
| P5 library by its address | the library is found by its address, and its root is the mapped address (another root: no listing) |
| P5 folder listing | the items that SharePoint gives for the one folder: all, under the folder, files at any depth, files in the folder itself |
| P5 names | the file names on both sides, on the web only, on the disk only |
| P5 web files checked | of at most 20 web files: at the mapped local path, same size, found by web path, errors |
| P5 folder found=False, P5 no web listing | the probe stops at this line (fixed text) |
| P5 stopped | an error stopped the probe: its error kind |
| P5 disk read skipped | the small file is not at the mapped local path, or it is online only |
| P5 entries with the chosen folder, P5 web read, P5 disk read | as in probe 4: entries for the same folder; 1 file of at most 5 MB read into memory |

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-25-device-probe-5.txt`. Do the
steps 2 times: run A with a shortcut folder in your OneDrive, run B with a folder of a synced team library.

1. Open **File Explorer**. You see the window **File Explorer**. Go on.
2. In the left pane of **File Explorer**, click the folder for this run. For run A, click a shortcut folder in your OneDrive. For run B, click a team library under the name of your organization. You see the files of that folder. Go on.
3. In the window **File Explorer**, double-click a folder that has 5 or more files. You see the files of that folder. Go on.
4. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
5. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
6. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
7. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-5.txt`. You see the probe as text. Go on.
8. In the browser page with the probe, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the probe text selected. Go on.
9. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
10. In the window **File Explorer**, do these actions:
    1. Click the address bar.
    2. Press **Ctrl+C**.
    You see the path of the folder selected. Go on.
11. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Question 1: a OneDrive folder that has files**. Go on.
12. On the question **Question 1**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see that the screen clears. Go on.
13. If a window asks if you want to continue running scripts, click **Yes**. You see the sign-in window. Go on.
14. If a sign-in window opens, sign in with your CAC. The window closes. Go on.
15. In the ISE output pane, wait until the line **Done.** shows. Do not copy before then.
16. In the ISE output pane, do these actions:
    1. Select the lines from the first line to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
17. In an e-mail to yourself, do these actions:
    1. Type `run A` or `run B` (for run B: under the lines of run A).
    2. Press **Enter**.
    3. Press **Ctrl+V**.
    You see the lines in the e-mail. Go on.
18. If you did run A only, start this list again from the top for run B. You see this list from the top. Go on.
19. In the e-mail with run A and run B, click **Send**. You see the message in your sent mail. The task is complete.

## Result

- Run A and run B, 2026-11-25 (the PAM, by e-mail; counts only):
  1. Run A, a shortcut in the own OneDrive (no scripts question, no PIN): entries 3, 1 with the chosen folder (1 web
     address); not the home folder; local files 22, online only 17; sign-in ok (6688 ms); library by its address found,
     root matches; folder listing: listed 5000, stopped at 5000, under the folder 5000, files at any depth 3898, files in
     the folder itself 0; names on both 0, local only 22.
  2. Run B, a synced Teams library (the PAM: "a sync'd teams"): entries 1; not the home folder; local files 4, online
     only 4; sign-in ok (623 ms); library found, root matches; folder listing: listed 0; names on both 0, local only 4.

  What it proves:
  1. `Get-PnPListItem -FolderServerRelativeUrl` gives the items at any depth under the folder (run A: 3898 files
     below the folder, 0 in it). Assumption A7 of the design holds (high confidence).
  2. `Get-PnPList -Identity <site-relative address>` finds the library, with the right root, for both kinds.
  3. The folder map (library address + the local path after MountPoint) is wrong for a shortcut and for a synced Teams
     library: run A maps to a folder with over 5000 items and none of the 22 local files; run B maps to an empty place.
     Assumption (medium confidence): for these kinds, UrlNamespace names the library, and MountPoint is the synced
     folder inside it, so the folder's own path in the library is missing from the map.
  4. Ion stays as built: only the own OneDrive uses the web; a shortcut and a synced library are scanned on the
     computer only, with their history "A drive keeps no file history".
