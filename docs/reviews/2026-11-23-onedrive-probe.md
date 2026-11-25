# OneDrive probe (2026-11-23)

The PAM approved PR 3 of the plan of 2026-11-23 (docs/reviews/2026-11-23-ion-esignature-onedrive-plan.md on the branch
of PR 57): Ion reads a synced OneDrive or SharePoint folder. Ion tells a OneDrive item from a real link with
PowerShell's own `LinkType` and `Target` (a real link has one; a OneDrive item has none), and reads such an item only
inside a folder that the OneDrive app syncs. This probe checks both on the gov device, on Windows PowerShell 5.1,
before PR 3 merges.

## What it does

- Read-only. It lists the one folder that you give and its direct subfolders (no deeper), at most 2,000 items. It opens
  no file and writes nothing.
- It prints counts only, one line for each kind: file or folder, reparse point or not, online only or not, and the
  `LinkType` and `Target` that PowerShell gives, and whether the folder is one that OneDrive syncs. It prints no name,
  no path and no number from a file.
- Expected result: `in OneDrive=True`, and the OneDrive items show `reparse=True LinkType=(none) Target=(none)`. If
  not, Ion skips them as links, and PR 3 needs another test.

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-23-onedrive-probe.txt`.

1. On the gov device, open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
3. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-23-onedrive-probe.txt`. You see the probe as text. Go on.
4. In the browser page with the probe, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the probe text selected. Go on.
5. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
6. In the window **Windows PowerShell ISE**, press **F5**. You see the question **OneDrive or synced folder**. Go on.
7. On the question **OneDrive or synced folder**, do these actions:
   1. Type the path of your OneDrive folder or of a synced library, for example `C:\Users\<you>\OneDrive - <organization>`.
   2. Press **Enter**.
   You see the line **in OneDrive** and lines that start with a number. Wait until the line **Done.** shows.
8. In the ISE output pane, do these actions:
   1. Select the lines from **in OneDrive** to **Done.**.
   2. Press **Ctrl+C**.
   You see the lines selected. Go on.
9. In the chat with the agent, do these actions:
   1. Press **Ctrl+V**.
   2. Send the message.
   You see the lines in the chat. The task is complete.

## Result

- Date: 2026-11-24. The PAM ran the probe on the gov device, in Windows PowerShell 5.1 ISE.
- Ion version: main at commit `4f23703` (PR 71). The OneDrive code is PR 59 (`src/DriveScan.psm1`).
- Output: 4 count lines (below). The first line `in OneDrive=<True|False> (sync folders found: N)` was not in the
  PAM's message.

| Count | Item | reparse | online | LinkType | Target |
|---|---|---|---|---|---|
| 61 | folder | True | False | (none) | (none) |
| 33 | file | True | False | (none) | (none) |
| 2 | file | False | False | (none) | (none) |
| 1 | file | True | True | (none) | (none) |

What it proves:

- All 94 reparse items (61 folders, 33 files) have no `LinkType` and no `Target` on Windows PowerShell 5.1. So
  `Get-ReparseKind` reads them as OneDrive items, not as links, as designed.
- 1 file is online only. Ion asks before it downloads it.
- 2 files are normal files (not reparse points). Ion reads them as usual.

Follow-up answers (the PAM, question dialog, 2026-11-24):

- The probe's first line showed 3 sync folders found. So the logic of `Get-OneDriveRoots` (the OneDrive environment
  variables and the value names under `HKCU:\Software\Microsoft\OneDrive\Accounts\*\Tenants`) finds the sync folders.
  The exact `in OneDrive=True` text was not quoted.
- A OneDrive link pasted into Ion started file downloads. That is the web (SharePoint) path of Ion.

Still not verified: an Ion scan of a local synced folder path (for example `C:\Users\<you>\OneDrive - <organization>`)
reads the files (files read, not "A link to another file is not followed").
