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
    1. Type the address of a team site that you scan.
    2. Press **Enter**.
    You see **Question 3**. Go on.
10. On the question **Question 3**, do these actions:
    1. Type the address of a second team site on the same host.
    2. Press **Enter**.
    You see **Question 4**. Go on.
11. On the question **Question 4**, do these actions:
    1. Type a Teams link of a team that you scan.
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

To skip a question, press **Enter** without an answer. To get a Teams link: in Teams, click **...** next to the team name, then click **Copy link**.

## Result

Not run yet.
