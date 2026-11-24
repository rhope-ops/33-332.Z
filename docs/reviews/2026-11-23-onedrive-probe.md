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
3. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/claude/ion-onedrive/docs/reviews/2026-11-23-onedrive-probe.txt`. You see the probe as text. Go on.
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
