# Device probe 7 (2026-11-25)

Probe 6 (docs/reviews/2026-11-25-device-probe-6.md) stopped with `ArgumentException/Other` in both runs, before its
first candidate. Cause, reproduced on PowerShell 7.4 with invented data (not yet on 5.1; the device runs showed the
same error type): `@($candidates)` on a `List[object]` made by `New-Object` throws "Argument types do not match". A
list made by `::new()`, a `List[string]`, and the pipeline form `@($list | ...)` do not throw. Ion's `src/` uses its
`New-Object` lists only in the pipeline form, so Ion has no fault of this kind (search of 2026-11-25).

Probe 7 keeps plain arrays and drops the settings-file scan (0 lines named the folder in both runs of probe 6). It
reads the registry values `FullRemotePath` and `WebUrl` of the OneDrive app's entry for the folder, prints their
shape only, makes the web folder from each one, and counts the local file names that are in that one web folder. Ion's
folder map fix (plan step 4: shortcuts and synced libraries) uses the value whose candidate has "names on both" equal
to "local files".

## What it does

- Read-only. It writes nothing and changes no setting.
- It prints only True or False, counts, kinds, error kinds and times. It prints no name, no address, no path and no value.
- It asks 1 question first, then clears the screen.
- It touches only the folder you give and the one library of that folder ("No reconnaissance"). It lists at most 3
  web folders (the probe 5 map and 1 for each value), and at most 500 items of each.
- A stop prints the step where it stopped (`P7 stopped at step N (name)`).

| Line | It shows |
|---|---|
| P7 probe version | 2026-11-25a: the text you pasted is this version |
| P7 entry is the OneDrive home folder | True: your own OneDrive; False: a shortcut or a synced library |
| P7 entry is the OneDrive home folder (second part) | folder is the synced folder itself: True when you gave the top folder of the shortcut or library |
| P7 local files | the files directly in the folder you gave, and its depth below the synced folder; with 0 files the probe stops before the sign-in |
| P7 value FullRemotePath, P7 value WebUrl | present; kind (https address, http address, rooted path, text, text with a colon, not text); same host as the library; path parts; where it is against the library root; depth below the root; if its last part is the synced folder's name, or the end of that name; if it is the library address |
| P7 folder found=False | the folder you gave is not on this computer: the probe stops (fixed text) |
| P7 no OneDrive entry holds this folder | the folder is not in OneDrive: the probe stops (fixed text) |
| P7 the folder has no files | the folder has 0 files: the probe stops before the sign-in (fixed text) |
| P7 FullRemotePath same as WebUrl | True when the 2 values are the same |
| P7 candidates | the web folders to try: the probe 5 map, and 1 for each value that is below the library root |
| P7 sign-in | the sign-in to the site of the folder |
| P7 library found | the library by its address, and if its root is the mapped address |
| P7 candidate | for each web folder: its items, if the list stopped at 500 items (then the count can be low), the files in it, and the local file names that are on both sides |
| P7 stopped at step | an error stopped the probe: the step and the error kind |

## Test before the device run

Stub run on PowerShell 7.4 (2026-11-25, invented data, fake registry entry and fake PnP commands): 6 cases, each ends
with "Done." and prints no value: a full address below the library root (candidate 2: names on both 2 of 3), a text
path, a path that starts with the library part, a missing value, a value on another site with a list error
(`RuntimeException/NotFound` on the candidate line), and an error in the library step (`P7 stopped at step 6 (library)`).

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-25-device-probe-7.txt`. Do the
steps 2 times: run A with a folder of a synced Teams library, run B with a shortcut folder in your OneDrive.

1. Open **File Explorer**.
2. In the left pane of **File Explorer**, click the folder for this run. For run A, click a synced Teams library under the name of your organization.
3. For run B, click a shortcut folder in your OneDrive.
4. In **File Explorer**, double-click a folder that has 5 or more files.
5. Close every **Windows PowerShell ISE** window.
6. Open **Windows PowerShell ISE**, then click **File** > **New**. The empty script pane **Untitled1.ps1** opens.
7. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-7.txt`. The first line of the page starts with `# Ion device probe 7`.
8. In the browser page, press `Ctrl+A`, then press `Ctrl+C`.
9. In the script pane **Untitled1.ps1**, press `Ctrl+V`.
10. In **File Explorer**, click the address bar, then press `Ctrl+C`.
11. In **Windows PowerShell ISE**, press `F5`. The question **Question 1: a OneDrive folder that has files** shows.
12. On **Question 1**, press `Ctrl+V`, then press **Enter**. The screen clears and shows `P7 probe version=2026-11-25a`.
13. If a window asks if you want to continue running scripts, click **Yes**.
14. If a sign-in window opens, sign in with your CAC. The sign-in can take 2 minutes or more.
15. Wait until the line **Done.** shows in the ISE output pane. Do not copy before then.
16. If the line **P7 the folder has no files** shows, start this list again with a folder of 5 or more files.
17. In the ISE output pane, select the lines from the first line to **Done.**, then press `Ctrl+C`.
18. In an e-mail to yourself, type `run A` or `run B` , press **Enter**, then press `Ctrl+V` (run B goes under run A).
19. If you did run A only, start this list again from the top for run B.
20. In the e-mail with run A and run B, click **Send**.

## Open (review of PR 122, MINOR, not verified)

- `Get-PnPListItem -FolderServerRelativeUrl` can return items of subfolders too (assumption, low confidence: not checked
  in the PnP.PowerShell 1.12.0 source). Then a 500-item cap can drop files of the folder itself. Probe 7 prints
  "stopped at 500" for this. `src/DriveScan.psm1` uses the same call: check it in the folder map fix.

## Result

Waits on the PAM's runs.
