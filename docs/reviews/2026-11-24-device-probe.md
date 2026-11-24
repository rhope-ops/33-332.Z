# Device probe 2 (2026-11-24)

Step 0 of the plan of 2026-11-24 (docs/reviews/2026-11-24-no-persistence-plan.md). Steps 4 and 5 wait on its result.

## What it does

- Read-only. It writes nothing and changes no setting.
- Part A (no sign-in): it reads the OneDrive app's own list of synced folders in the registry and checks if your OneDrive
  folder has a web address there. Ion needs that address to read online-only files and the file history through PnP.
- Part B (one CAC sign-in per site): on one file of your OneDrive and, if you give one, one file of a team site, it
  checks which permission reads work for you without admin, and it times one read from the web and one from the disk.
- It prints only True or False, counts, error type names and times. It prints no name, no address, no path and no value
  from a file.

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-24-device-probe.txt`.

1. On the gov device, open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
3. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-24-device-probe.txt`. You see the probe as text. Go on.
4. In the browser page with the probe, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the probe text selected. Go on.
5. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
6. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Part A: your OneDrive folder**. Go on.
7. On the question **Part A: your OneDrive folder**, do these actions:
   1. Type the path of your OneDrive folder, for example `C:\Users\<you>\OneDrive - <organization>`.
   2. Press **Enter**.
   You see lines that start with **A1** to **A4** and the question **Part B: a team site address**. Go on.
8. On the question **Part B: a team site address**, do these actions:
   1. Type the address of a team site that you scan, for example `https://usaf.dps.mil/teams/<team-site>`.
   2. Press **Enter**.
   You see a CAC sign-in window. Go on.
9. In each CAC sign-in window, sign in with your CAC. You see lines that start with **B-onedrive** and **B-team**. Wait
   until the line **Done.** shows.
10. In the ISE output pane, do these actions:
    1. Select the lines from **A1** to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
11. In the chat with the agent, do these actions:
    1. Press **Ctrl+V**.
    2. Send the message.
    You see the lines in the chat. The task is complete.

## Result

Not run yet.
