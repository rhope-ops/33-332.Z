# Design: OneDrive hybrid (plan step 4, 2026-11-24)

Status: approved design, 2026-11-24 (the PAM, question dialog; docs/DECISIONS.md). Build waits on recall run 3 (the PAM, 2026-11-24: "Design now"; build after
run 3 shows why 3 Office files in a OneDrive folder were not scanned). Plan row 4:
docs/reviews/2026-11-24-no-persistence-plan.md.

## Goal

When the pasted folder is in a folder that OneDrive syncs, Ion reads:

- each file that is on the disk, from the disk (as now);
- each online-only file, and each file in the mapped web folder that is not on the disk, through PnP, into memory
  (nothing goes to the disk, so no download question);
- the who/when history ("PII first saved", "PII first saved by") of each file with a hit through PnP, as for a team site.

If the sign-in or the mapping fails, Ion scans as now (local only, with the download question) and the Summary says so.

## Evidence

- Probe 4, runs 1 and 2 (docs/reviews/2026-11-24-device-probe-4.md): web path = UrlNamespace path + (local path minus
  MountPoint); 9 of 9 web files at the mapped local path; path A (address + folder + name) finds 9 of 9; 2 sync entries
  with the same address (longest MountPoint match is safe); 1 of 9 Office files has another size on the web.
- Ion already reads a SharePoint file into memory (`Get-PnPFile -AsMemoryStream`, src/ChannelScan.psm1) and its version
  history (src/Accountability.psm1 Find-PiiAddedVersion), with one sign-in per host (src/SiteTarget.psm1).

## Flow (src/DriveScan.psm1 Invoke-DriveScan)

1. Find the sync entry: the entries in `HKCU:\Software\SyncEngines\Providers\OneDrive` (read-only), the longest MountPoint
   that holds the pasted folder. None: local only, as now.
2. Map: site = the site part of UrlNamespace; web folder = UrlNamespace path + the rest of the pasted path ('\' to '/').
3. Sign in once to that site (Connect-TeamSite, Assert-ConnectedSite). It counts as the run's sign-in to that host.
4. Find the one library whose root is the UrlNamespace path (as probe 4). List only the items under the mapped web
   folder (folder-scoped query, not the whole library). "No reconnaissance": no other library, site or user is read.
5. Match the local list and the web list by relative path:
   - on the disk: read from the disk (content); the web item gives the history;
   - online only and on the web: read through PnP into memory;
   - on the disk only (not synced yet): read from the disk; history "Not recorded (not synced yet)";
   - on the web only (not synced down): read through PnP into memory (the PAM, 2026-11-24, Q4). It is inside the
     pasted folder's own mapped web folder, so the scan does not widen.
6. History for each file with a hit: Find-PiiAddedVersion with the mapped server-relative path. When the disk copy is
   newer than the web version, a History note says "The disk copy is newer than the last synced version".
7. Sign-out at the end of the run, as now. Masked output, read-only and the frozen detectors stay as they are.

## Assumptions

| # | Assumption | Source | Confidence |
|---|---|---|---|
| A1 | The mapping and path A work for the own OneDrive folder | probe 4, 2 runs | high |
| A2 | They work the same for a synced team library and for a shortcut ("Add shortcut to My files") | not tested | low |
| A3 | The disk copy has the same content as the web copy; the size difference is SharePoint properties | probe 4 (1 of 9) | medium |
| A4 | The own CAC sign-in can read the version history of files in the own OneDrive | Accountability works on team sites | medium |
| A5 | Recall run 2 (3 Office files not scanned in a OneDrive folder) does not change this design | not known until run 3 | low |

## Decisions (the PAM, question dialog, 2026-11-24)

1. Online-only files: read through PnP into memory, with no download question (REC).
2. Scope: the own OneDrive, a synced team library and a shortcut; a probe line first for the 2 untested kinds (REC).
3. A failed sign-in or mapping: scan local only, and a Summary line says the who/when history is not available (REC).
4. Files on the web only, in the mapped folder: read through PnP (not the REC, which was a Not scanned row).

## Tests (tests/FakePnP.ps1, tests/DriveScan.Tests.ps1)

Fake sync entries and a fake library: mapping with 2 entries (longest wins); each of the 4 match cases; history called
with the mapped path; sign-in failure falls back to local only; nothing outside the mapped folder is listed; masked
names in every new row.

## Build order (after recall run 3)

1. Probe line for a synced team library and a shortcut (Decision 2), as a .txt test snippet; the PAM runs it.
2. Build with tests, fresh review, lane B. Open follow-ups of PR 105 (recall Score, MINOR): an exact-output test for a
   decomposed name; a soft hyphen (\p{Cf}) can split a word.
