# Device probe 8 (2026-11-25)

Plan step 5 (permission scope column: Anyone link, Whole organization, Site members, Specific people; counts only, no
names) needs to know which permission reads work for a member without admin. No probe read a permission yet: probe 2
did not find its file, and probe 3 tested sign-in and paths only. The PAM chose a separate probe 8 on 2026-11-25
(docs/DECISIONS.md). Ion's step 5 uses only the reads that work here.

## What it does

- Read-only. It writes nothing and changes no setting. The sharing read is a POST call that only reads.
- It asks 2 questions first, then clears the screen.
- It touches only the one site and the one file you give ("No reconnaissance"). It lists no other file, site or user.
- It prints only True or False, counts, kinds, error kinds and times. It prints no name, no address, no path, no group
  name and no value. It reads the sign-in name and the type of each permission entry of the file, in memory only, to
  sort it into a kind, and prints only the count of each kind (the PAM's decision, 2026-11-25).

| Line | It shows |
|---|---|
| P8 probe version | 2026-11-25e: the text you pasted is this version |
| P8 answer 1 is a DoD365 site address=False | the address is not on usaf.dps.mil: the probe stops |
| P8 folder from | where the probe found the folder in the address: id (a folder page), library page, path |
| P8 folder in the address found=False | the address does not name a folder of this site: the probe stops |
| P8 sign-in | the sign-in to the site |
| P8 file found | the file in that folder, found by its name, or (only for a name with no type) by its name without its type: then the probe reads the file names of that one folder; False: the probe stops |
| P8 answer with .. or \ found=True | an answer has `..`, `.` as a folder, or `\` in it: the probe stops |
| P8 files with this name in the folder, P8 folder read | the count of files with that name when it is not exactly 1, or why the folder could not be read |
| P8 library read=False | the library of the file could not be read: the sharing links are not read |
| P8 unique permissions | True: the file has its own permissions; False: it has the permissions of its folder or library |
| P8 site groups | the site's own groups (owners, members, visitors) that the probe can read |
| P8 permission entries | the permission entries of the file (people and groups with access) |
| P8 entry kinds | the count of each kind: person, site group, team members or team owners (Microsoft 365 group), SharePoint group, sharing link group, security group, mail list, whole organization; "(limited access only)" when the entry gives no access to this file |
| P8 sharing links | the sharing links of the file, counted by kind (SharePoint's SharingLinkKind): direct, organization view or edit, anyone view or edit, specific people (flexible), and others |
| P8 Get-PnPFileSharingLink exists | if PnP.PowerShell 1.12.0 has this command |
| P8 stopped at step | an error stopped the probe: the step and the error kind |

A read that fails shows `read=False` and the error kind, for example `NoAccess`: then Ion cannot use that read.

## Test before the device run

Stub run on PowerShell 7.4 (2026-11-25, invented data, fake PnP commands that give each collection as one object and nothing
for a missing file, as PnP.PowerShell 1.12.0 does): 11 cases: the 6 below, a missing file, a missing file named
without its type, an answer with `..`, a site name `..`, and an answer with `\`. The 6 cases, each ends with "Done." and prints no value: a folder page address (8 entries in 7
kinds, 3 links in 3 kinds), a name without its type, a library page address, each read refused (`NoAccess`), an
address that is not on usaf.dps.mil, and a folder of another site (the probe stops).

## How to run it

The probe is a .txt file (Project rules, "Paste format"): `docs/reviews/2026-11-25-device-probe-8.txt`. Do the steps
2 times: run A with a file that you shared with a link, run B with a file that you did not share. Both files are on
your team site.

1. In your web browser, open the folder of the file for this run on your team site. You see the files of the folder. Write down the name of the file as the column **Name** shows it, for example `Report.pdf`. Go on.
2. Close every **Windows PowerShell ISE** window. You see no ISE window. Go on.
3. Open **Windows PowerShell ISE**. You see the window **Windows PowerShell ISE**. Go on.
4. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
5. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-8.txt`. You see the probe as text. The first line starts with `# Ion device probe 8`. Go on.
6. In the browser page with the probe, do these actions:
   1. Press **Ctrl+A**.
   2. Press **Ctrl+C**.
   You see all the probe text selected. Go on.
7. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
8. In the window **Windows PowerShell ISE**, press **F5**. You see the question **Question 1: the address of the folder page in your web browser**. Go on.
9. In your web browser, on the page of the folder, do these actions:
   1. Click the address bar.
   2. Press **Ctrl+C**.
   You see the address selected. Go on.
10. On the question **Question 1** in **Windows PowerShell ISE**, do these actions:
    1. Press **Ctrl+V**.
    2. Press **Enter**.
    You see **Question 2**. Go on.
11. On the question **Question 2**, do these actions:
    1. Type the file name that you wrote down, for example `Report.pdf`.
    2. Press **Enter**.
    You see that the screen clears. Then you see the line `P8 probe version=2026-11-25e`. Go on.
12. If a window asks if you want to continue running scripts, click **Yes**. The window closes. Go on.
13. If a sign-in window opens, sign in with your CAC. The window closes. The sign-in can take 2 minutes or more. Go on.
14. In the ISE output pane, wait until the line **Done.** shows. Do not copy before then.
15. In the ISE output pane, do these actions:
    1. Select the lines from the first line to **Done.**.
    2. Press **Ctrl+C**.
    You see the lines selected. Go on.
16. In an e-mail to yourself, do these actions:
    1. Type `probe 8 run A` or `probe 8 run B` (for run B: under the lines of run A).
    2. Press **Enter**.
    3. Press **Ctrl+V**.
    You see the lines in the e-mail. Go on.
17. If you did run A only, start this list again from the top for run B. You see this list from the top. Go on.
18. In the e-mail with run A and run B, click **Send**. You see the message in your sent mail. The task is complete.

## Result

Waits on the PAM's runs.
