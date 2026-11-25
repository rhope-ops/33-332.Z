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
| P7 local files | the files directly in the folder you gave, and its depth below the synced folder |
| P7 value FullRemotePath, P7 value WebUrl | present; kind (https address, rooted path, text); same host as the library; path parts; where it is against the library root; depth below the root; if its last part is the synced folder's name, or the end of that name; if it is the library address |
| P7 candidates | the web folders to try: the probe 5 map, and 1 for each value that is below the library root |
| P7 sign-in | the sign-in to the site of the folder |
| P7 library found | the library by its address, and if its root is the mapped address |
| P7 candidate | for each web folder: its items, the files in it, and the local file names that are on both sides |
| P7 stopped at step | an error stopped the probe: the step and the error kind |

## Test before the device run

Stub run on PowerShell 7.4 (2026-11-25, invented data, fake registry entry and fake PnP commands): 6 cases, each ends
with "Done." and prints no value: a full address below the library root (candidate 2: names on both 2 of 3), a text
path, a path that starts with the library part, a missing value, a value on another site with a list error
(`RuntimeException/NotFound` on the candidate line), and an error in the library step (`P7 stopped at step 6 (library)`).

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-25-device-probe-7.txt`. Do the
steps 2 times: run A with a folder of a synced Teams library, run B with a shortcut folder in your OneDrive. Use the
same 2 folders as in probe 6.

1. Open **File Explorer**. You see the window **File Explorer**. Go on.
2. In the left pane of **File Explorer**, click the folder for this run. For run A, click a synced Teams library under the name of your organization. For run B, click a shortcut folder in your OneDrive. You see the files of that folder. Go on.
3. In the window **File Explorer**, double-click a folder that has 5 or more files. You see the files of that folder. Go on.
4. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
5. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
6. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
7. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-7.txt`. You see the probe as text. The first line starts with `# Ion device probe 7`. Go on.
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
    You see that the screen clears. Then you see the line `P7 probe version=2026-11-25a`. Go on.
13. If a window asks if you want to continue running scripts, click **Yes**. The window closes. Go on.
14. If a sign-in window opens, sign in with your CAC. The window closes. The sign-in can take 2 minutes or more. Go on.
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

Waits on the PAM's runs.
