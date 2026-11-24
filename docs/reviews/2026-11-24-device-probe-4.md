# Device probe 4 (2026-11-24)

Probe 3 (second run) showed: the OneDrive address maps to the root of its web library, but the file that the probe
took from the disk was not in that library (4259 items listed, 0 matches). So the web path forms are still untested.
Probe 4 starts from the other side: it takes the files of one folder from the web listing and checks that each one is
at the local path that the OneDrive address maps to. Ion's OneDrive step (plan step 4) uses only the mapping that
passes here.

## What it does

- Read-only. It writes nothing and changes no setting.
- It asks 1 question first, then clears the screen. The lines you copy hold no answer.
- It prints only True or False, counts, error kinds and times. It prints no name, no address, no path and no value
  from a file.
- It touches only the folder you give and the site of its OneDrive address: it reads the list of that site's
  libraries, at most 5000 items of the one library whose root is the address, and looks up at most 20 of its files
  by their web address. It reads at most 1 file of at
  most 5 MB from the web, into memory only. It reads that file from the disk only when it is on the disk (an
  online-only file is not read, so OneDrive downloads nothing).

| Line | It shows |
|---|---|
| P4 folder found=False | the path in Question 1 is not a folder: the probe stops |
| P4 no OneDrive entry holds this folder | the folder is not synced by OneDrive: the probe stops |
| P4 matching entries | how many OneDrive sync entries hold the folder |
| P4 entries with the chosen folder | of these, how many have the chosen folder, and how many different web addresses they have |
| P4 entry is the OneDrive home folder | True: your own OneDrive; False: a synced library or a shortcut |
| P4 local files | the files on the disk (online-only files too) |
| P4 sign-in | the sign-in to the site of the OneDrive address |
| P4 library with this root | a web library whose root is the mapped address |
| P4 web listing | the items listed, the item kinds (FSObjType 0 = file, 1 = folder), and the files of the same folder on the web |
| P4 no web listing | the listing failed: the probe prints no counts after it |
| P4 names | the file names on both sides, on the web only, on the disk only |
| P4 web files checked | of at most 20 web files: how many are at the mapped local path, with the same size, found by their web path, and errors |
| P4 web read, P4 disk read | the smallest web file, read from the web and from the disk: the same size as on the web (a "P4 disk read skipped" line shows when the file is not at the mapped local path, or is online only) |
| P4 stopped | an error stopped the probe: its kind |

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-24-device-probe-4.txt`.

1. On the gov device, open **File Explorer**. You see the window **File Explorer**. Go on.
2. In the left pane of **File Explorer**, click your OneDrive folder. You see the files in your OneDrive. Go on.
3. In the window **File Explorer**, double-click a folder that has 5 or more files. You see the files of that folder. Go on.
4. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
5. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
6. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
7. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-24-device-probe-4.txt`. You see the probe as text. Go on.
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
16. On paper, write down the number of sign-in windows and the number of PIN prompts. If no sign-in window opened, write down `0`. You see the 2 numbers on the paper. Go on.
17. In the ISE output pane, do these actions:
    1. Select the lines from the first line to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
18. In an e-mail to yourself, do these actions:
    1. Press **Ctrl+V**.
    2. Type the number of sign-in windows and the number of PIN prompts that you wrote down.
    3. Send the message.
    You see the message in your sent mail. The task is complete.

## Result

Not run yet.
