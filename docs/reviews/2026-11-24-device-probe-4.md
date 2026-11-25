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

- Run 1, 2026-11-24 (the PAM, by e-mail; counts only; 1 sign-in window with the scripts prompt, no PIN):

  | Line | Result |
  |---|---|
  | P4 matching entries / entries with the chosen folder / different web addresses | 2 / 2 / 1 |
  | P4 entry is the OneDrive home folder | True |
  | P4 local files / online only | 9 / 0 |
  | P4 sign-in | ok, 10206 ms |
  | P4 library with this root | libraries=4, found=True (1963 ms) |
  | P4 web listing | listed=4259, stopped at 5000=False, FSObjType 0=3608 1=651 other=0, web folder found=True, web files in it=9 (8134 ms) |
  | P4 names | on both=9, web only=0, local only=0 |
  | P4 web files checked | 9: at the mapped local path=9, same size=8, found by their web path=9, errors=0 |
  | P4 web read / disk read | bytes equal web size=True / True (1816 / 103 ms) |

  What it proves:
  1. The mapping is right: local path = sync folder (MountPoint) + (web path minus UrlNamespace path). 9 of 9 web
     files are at the mapped local path, and path A (address + folder + name) finds 9 of 9.
  2. The 2 sync entries hold the same folder and the same web address: the longest-match choice is safe.
  3. The address is the root of one of the site's 4 libraries; the FSObjType filter works (3608 files, 651 folders).
  4. A file read from the web and from the disk has the web size.
  5. Plan step 4 (OneDrive hybrid) can use this mapping.

  What stays open:
  - 1 of 9 files has a different size on the web and on the disk. Assumption, medium confidence: SharePoint writes its
    properties into the web copy of an Office file. Ion's hybrid step reads the disk copy for content and uses the
    web only for the who/when history.
  - Probe 3's 215-byte file was not in the library. With probe 4's result the cause is that file (not synced, or in
    another folder), not the path rule (assumption, high confidence).

- Run 2, 2026-11-24 (the PAM, by e-mail; counts only): the same result as run 1. Only these lines differ: sign-in ok,
  11043 ms; web listing listed=4278, FSObjType 0=3624 1=654 other=0 (files were added to the library between the runs).
  Entries 2 (1 web address), home folder True, local 9 / online only 0, libraries=4 (found), web folder found with 9
  web files, names on both 9, checked 9 (at the mapped path 9, same size 8, found by web path 9, errors 0), web and disk
  read sizes equal. What it proves: run 1 repeats; the mapping and path A are stable.

- Runs A and B, 2026-11-25 (the PAM, by e-mail; counts only). The PAM wrote: "The 2nd one was the sync under my
  organization (sharepoint) and the first was a shortcut under my onedrive".
  1. A shortcut in the own OneDrive: matching entries 3, 1 with the chosen folder (1 web address); not the home
     folder; local files 12, online only 12; sign-in ok (6714 ms); library found (libraries=12). Web listing: 5000
     items, stopped at 5000 = True (FSObjType 0=3947 1=1053), web folder found = False, names on both 0. The map and
     the sign-in work. The library has more than 5000 items; assumption (medium confidence): the first 5000 did
     not reach the folder (a wrong subfolder map is not ruled out; probe 5 shows which).
  2. A folder of a synced team library: matching entries 0 ("no OneDrive entry holds this folder").
  What it proves: a shortcut needs a folder-scoped listing (PR 109 asks for the folder only); a synced team library
  is not in this OneDrive list. Next: probe 5 (docs/reviews/2026-11-25-device-probe-5.md).
