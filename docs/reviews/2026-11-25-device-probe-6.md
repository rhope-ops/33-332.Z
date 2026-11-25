# Device probe 6 (2026-11-25)

Probe 5 (2026-11-25) showed that the OneDrive address of a shortcut and of a synced Teams library names the library
only, so Ion's folder map misses the synced folder's own path in the library. Probe 6 looks for that path in 2 places
on the gov device: the OneDrive app's registry entry, and its settings files (.ini). It tries each path that it finds
against the web, and prints which source fits. Ion's map fix (plan step 4) uses only the source that fits here.

## What it does

- Read-only. It writes nothing and changes no setting.
- It prints only True or False, counts, the value names of the OneDrive app's registry entry, error kinds and times. It prints no name, no address, no path and no value from a file.
- It asks 1 question first, then clears the screen.
- It reads the OneDrive app's registry entry and its settings files (.ini) on this computer.
- It touches only the folder you give and the one library of that folder. For each candidate path, it reads the list details of at most 500 items of that one web folder.

| Line | It shows |
|---|---|
| P6 folder found=False, P6 no OneDrive entry holds this folder | the probe stops at this line (fixed text) |
| P6 entry value names | the names of the values in the OneDrive app's entry for this folder |
| P6 entry is the OneDrive home folder | True: your own OneDrive; False: a shortcut or a synced library |
| P6 local files | the files directly in the folder you gave |
| P6 settings files | the .ini files of the OneDrive app, the lines that name this folder, and the candidate paths found |
| P6 sign-in | the sign-in to the site of the folder |
| P6 library found | the library by its address, and if its root is the mapped address |
| P6 candidate | for each candidate path: where it came from, the items of that web folder, the files in it, the local file names that are on both sides, and True when the path is the name of the synced folder |
| P6 stopped | an error stopped the probe: its error kind |

The candidate with "names on both" equal to "local files" is the right source.

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-25-device-probe-6.txt`. Do the
steps 2 times: run A with a shortcut folder in your OneDrive, run B with a folder of a synced Teams library.

1. Open **File Explorer**. You see the window **File Explorer**. Go on.
2. In the left pane of **File Explorer**, click the folder for this run. For run A, click a shortcut folder in your OneDrive. For run B, click a synced Teams library under the name of your organization. You see the files of that folder. Go on.
3. In the window **File Explorer**, double-click a folder that has 5 or more files. You see the files of that folder. Go on.
4. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
5. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
6. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
7. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-6.txt`. You see the probe as text. Go on.
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

- Not run yet.
