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
| P8 probe version | 2026-11-25f: the text you pasted is this version |
| P8 answer 1 is a DoD365 site address=False | the address is not on usaf.dps.mil: the probe stops |
| P8 folder from | where the probe found the folder in the address: id (a folder page), library page, path |
| P8 folder in the address found=False | the address does not name a folder of this site: the probe stops |
| P8 sign-in | the sign-in to the site |
| P8 file found | the file in that folder, found by its name, or (only for a name with no type) by its name without its type: then the probe reads the file names of that one folder; False: the probe stops |
| P8 answer with .. or \ found=True | an answer has a part of only dots (`..`), `\`, or a code like `%2e` left after one decode: the probe stops |
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

1. In your web browser, open the folder of the file for this run on your team site. Write down the file name as the column **Name** shows it, for example `Report.pdf`.
2. Close every **Windows PowerShell ISE** window.
3. Open **Windows PowerShell ISE**, then click **File** > **New**. The empty script pane **Untitled1.ps1** opens.
4. In your web browser, open `https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/docs/reviews/2026-11-25-device-probe-8.txt`. The first line of the page starts with `# Ion device probe 8`.
5. In the browser page, press `Ctrl+A`, then press `Ctrl+C`.
6. In the script pane **Untitled1.ps1**, press `Ctrl+V`, then press `F5`. The question **Question 1: the address of the folder page in your web browser** shows.
7. In your web browser, on the page of the folder, click the address bar, then press `Ctrl+C`.
8. In **Windows PowerShell ISE**, on **Question 1**, press `Ctrl+V`, then press **Enter**. **Question 2** shows.
9. On **Question 2**, type the file name that you wrote down, then press **Enter**. The screen clears and shows `P8 probe version=2026-11-25f`.
10. If a window asks if you want to continue running scripts, click **Yes**.
11. If a sign-in window opens, sign in with your CAC. The sign-in can take 2 minutes or more.
12. Wait until the line **Done.** shows in the ISE output pane. Do not copy before then.
13. In the ISE output pane, select the lines from the first line to **Done.**, then press `Ctrl+C`.
14. In an e-mail to yourself, type `probe 8 run A` or `probe 8 run B` , press **Enter**, then press `Ctrl+V` (run B goes under run A).
15. If you did run A only, start this list again from the top for run B.
16. In the e-mail with run A and run B, click **Send**.

## Known limits (fourth review of PR 125, MINOR, kept)

- A file name with `%` and 2 hex digits (for example `80%AB`) is refused: choose another file.
- A subsite of the same team site can be reached through the address or answer 2: it is inside the one team.
- The probe counts the other entry types too (mail list, security group, SharePoint group): the PAM's words of
  2026-11-25 cover "certain permissions or user groups" (docs/DECISIONS.md).

## Result

Waits on the PAM's runs.
