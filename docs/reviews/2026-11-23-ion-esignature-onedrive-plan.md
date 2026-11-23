# Ion plan: e-signature DoD IDs, the 50-minute hang, OneDrive scans (2026-11-23)

Status: APPROVED 2026-11-23 by the PAM (question dialog: all 3 PRs; Q1 name or signature word; Q2 dotted any case, no dots CAPS;
Q3 online-only files: ask each run; Q4 SharePoint scan, Ctrl+C did not stop it). The PRs go one after another on the one allowed branch. Base: main 6cbc517. Branch: ccr-9805d725-onoreq.
Kept: docs/DECISIONS.md rows of 2026-11-21 (e-signature DoD ID = DoD ID / Low, Low tab, no version history).
Notation: `#` = one digit (tests build invented values at run time); `\n` = line break.

## Progress

- PR 57 (PR 1): built, fresh-reviewed; waits on the PAM's shape lines (docs/reviews/2026-11-23-esignature-shape-test.md) before merge.
- PR 58 (PR 2, hang): on branch claude/ion-hang-limits (the PAM, 2026-11-23: a new branch per PR). PR 3 (OneDrive and the
  Connect-PnPOnline fallback): on branch claude/ion-onedrive.
- PR 57 known limits after 3 fresh reviews (accepted by the author, for the PAM to see): a lower-case 3-part dotted
  word with a one-letter last part ("www.site.a." then 10 digits) reads as a name; a row number on the line before a
  dotted name, or a capitals line such as "GRAND TOTAL" then gapped 9+1 digits on the same line, reads as an
  e-signature (DoD ID Low, not SSN); a run of 40,000+ capitals joined by hyphens with no space is slow (5 s at 78k).
  Fourth review: a card split so that the 10 digits start inside it ("Signed by ###### ##########", a 6-10 split) is a
  DoD ID, not a card; a DoD ID followed on the same line by 3+ digits that pass the card check reads as a card.

## Three pull requests, disjoint files

| PR | Scope | Files | Waits on |
|---|---|---|---|
| 1 | E-signature DoD IDs | src/PiiPatterns.psm1, tests/Detectors.Tests.ps1, paste/Ion-script.txt | the PAM's answers to Q1, Q2 |
| 2 | Hang: no file or download can stop the scan | src/ReaderPool.psm1, src/ScanQueue.psm1, src/ChannelScan.psm1, tests | the PAM's hang details (Q4) |
| 3 | OneDrive / synced folders (OCONUS route) | src/DriveScan.psm1, tests, README.md, docs/DECISIONS.md | probe result, Q3 |

PR 1 does not wait on PR 2 or 3.

## PR 1: e-signature DoD IDs

Cause (src/PiiPatterns.psm1:162-184, :43-46): a 9-digit run is part of a DoD ID only when (a) a dotted capital CAC name
touches it with no gap, or (b) a signature word comes before "### ## #### #". PDF extraction drops dots, adds spaces and
line breaks. The CAC name pattern has the same limits, so the DoD ID is also lost.

Design: one shared recogniser finds each e-signature span (name shape or signature word + 10 digits with gaps).
Find-DodIdHits reports each span as DoD ID / Low; Find-SsnHits skips any 9-digit match inside a span. So each SSN that
is suppressed becomes a DoD ID hit (no PII lost).

Name shape before the digits (gap allowed: spaces, tabs, one line break, `\r\n` too):
- Dotted: `LAST.FIRST[.M][.]`, any letter case, apostrophe and hyphen in a part (O'BRIEN, SMITH-JONES), JR/SR/II/III.
- Not dotted: `LAST FIRST [M][.]` or `LAST, FIRST [M][.]` in CAPITALS only (proposal, Q2).
- Or a signature word (signed, CN=, DN:, Signature:) within 40 characters before.
- The digits: 10, with single spaces, dashes or one line break between digit groups; or the 10 digits on the line before
  the name.
- The existing guards stay: NO./REF./ACCT./... are no name part; the last name has 2 letters or more.

Expected results (each one becomes a test):

| # | Layout | Today | Expected |
|---|---|---|---|
| 1 | `DOE JOHN A ### ## #### #` | SSN/Medium | DoD ID/Low |
| 2 | `DOE.JOHN.A. ######### #` | SSN/Low | DoD ID/Low |
| 3 | `Signature: DOE, JOHN A. ######### #` | SSN/Low | DoD ID/Low |
| 4 | `DOE.JOHN.A. ##########` (1 and 2 spaces) | none | DoD ID/Low |
| 5 | `DOE.JOHN.A.\n##########` (`\n` and `\r\n`) | none | DoD ID/Low |
| 6 | `DOE.JOHN.A.#####\n#####` | none | DoD ID/Low |
| 7 | `DOE.JOHN.A.#########\n#` | not tested | DoD ID/Low |
| 8 | `##########\nDOE.JOHN.A.` | none | DoD ID/Low |
| 9 | `doe.john.a.##########` | none | DoD ID/Low |
| 10 | `O'BRIEN.JOHN.A.##########` | not tested | DoD ID/Low |
| 11 | `DOE.JOHN.A.###-##-####-#` | none | DoD ID/Low |
| 12 | `Digitally signed by DOE.JOHN.A ##########` | none | DoD ID/Low |
| 13 | `Digitally signed by ######### # Date` | none (test :32) | DoD ID/Low (test :32 changes: PII no longer lost) |
| 14 | the 9 layouts that pass today (JR, hyphen, 2-part first name, gaps, date after, CN=/OU=, Reason:) | DoD ID/Low | unchanged |

Controls (must not change):

| # | Layout | Expected |
|---|---|---|
| C1 | `SSN ###-##-####, 6 copies` | SSN/High |
| C2 | `######### 3 pages`, `assigned ######### 1 slot`, `Unsigned ### ## #### 2`, `### ## #### 2 copies` (PR 49) | SSN |
| C3 | `###-##-####-#` alone (no name, no signature word) | no hit, as today (correction 2026-11-23: not SSN/Medium) |
| C4 | `tracking ########## shipped`; `Acct.No.`, `REF.NO.`, `backup.config.`, `e.g.` + 10 digits | no DoD ID |
| C5 | `assigned to bob a ######### #` (small letters, no dots) | SSN/Low (Q2) |
| C7 | `THE UNIT A ##########` (capitals, no dots, plain digits; build finding: 2,861 false DoD IDs in 2 MB of prose) | no DoD ID |
| C6 | Name hit next to an e-signature stays Low (test :43) | unchanged |

Also checked: the masked value and Protect-PiiText for gapped, dashed and line-split digits (no part of the digits
shown); Find-PiiInText (all detectors, overlaps); detector time on a 5 MB text before and after (no regex blow-up);
src/PiiPatterns.psm1 under the 450-line cap (394 now); the 506-text benchmark (rhope-ops/archimedes_lever, read-only)
before and after: precision must not drop.

## Proof from real data (step 4 of the request)

APPLIED: Data boundary (Project rules, 2026-11-20, cannot be changed) bars real names and IDs in chat, not only digits.
So the snippet prints the shape only: capital letter = A, small letter = a, digit = 9, space and line break as visible
marks, and only a 60-character window around each 9-10 digit run. It reads the PDF with Ion's own reader (the pasted
bundle), shows nothing else and writes nothing. Written with the instructions skill, saved as
docs/reviews/2026-11-23-esignature-shape-test.md. The PAM runs it on 2-3 e-signed PDFs and pastes the shape lines; each
shape becomes a test before PR 1 merges.

## PR 2: the hang (cause not known yet)

What the code shows: no time limit per file in a worker (only PDFs have one), no time limit on a SharePoint download,
and Wait-ScanQueue waits with no end. Plan: (1) a failing test with a job that never ends; (2) a per-file limit (proposal:
120 s): the worker is stopped, the file gets a Not scanned row "took longer than 120 s: check by hand", and a new worker
starts; (3) a download time limit; (4) the scan bar names the file that waits longest. Risk: stopping a runspace can block
inside a .NET call; the main thread then drops that runspace and goes on. Built after the PAM's details (Q4).

## PR 3: OneDrive and synced SharePoint folders (OCONUS)

Cause: src/DriveScan.psm1:53 and :58 skip every reparse point. OneDrive "Files On-Demand" makes synced folders and files
reparse points, so Ion lists them as links and reads nothing. No rule change is needed: "No reconnaissance" stays; a
synced file is the file itself, not a link. The DECISIONS row of 2026-11-21 (any reparse point = file link) changes:
symbolic links and junctions stay skipped; OneDrive cloud files are read.
Not verified: how PS 5.1 tells a cloud file from a link without Add-Type (candidates: LinkType/Target empty; attributes
Offline, RecallOnOpen, RecallOnDataAccess). Step 1 = a read-only probe that prints counts per attribute set only.
Local scans have no version history: "PII first saved by" = file owner/created data only (README row).

## GAPS

- No regex time limit in the detectors (cost: small; benefit: a bad text cannot hang a worker). Risk while PR 2 is open.
- No real-data layout corpus (cost: the snippet run; benefit: tests match real PDFs, not guesses).
- No per-file time limit / watchdog (PR 2). Risk: a 50-minute scan can be lost.
- Benchmark not in CI (cost: a script; benefit: precision/recall measured on each detector change). Not in this scope.

## Assumptions

- The PAM's layouts are invented and match the real ones in shape (source: the request; confidence medium; the snippet checks it).
- The hang was a drive or SharePoint scan in ISE on 5.1 (source: none; confidence low; Q4).
- "Cmdlet blocked" means Connect-PnPOnline only, not the language mode (confidence low; the drive scan would also fail).
