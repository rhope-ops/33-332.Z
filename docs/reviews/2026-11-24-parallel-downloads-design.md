# Design: parallel downloads in a planned order (step 6, 2026-11-24)

Plan: docs/reviews/2026-11-24-no-persistence-plan.md, step 6. Approval: docs/DECISIONS.md, 2026-11-21 row "Ion round 2
after the gov run", item (3): "2-4 parallel downloads, ordered from the full file list" (the PAM: "orient files for
download to appear at the beginning"). It replaces the 2026-11-20 row's "No parallel downloads: each would need its own
CAC sign-in", if test 3 (below) or the first device run shows that a copy of the sign-in works in a worker.

Status: design written; open questions go to the PAM in one dialog. Nothing is built yet.

## Today

The SharePoint and Teams scan downloads one file at a time on the main thread (src/ChannelScan.psm1 Get-FileBytes,
Invoke-OneFileScan). The readers (src/ReaderPool.psm1, 1 to 8 runspaces) read while the main thread downloads the next
file. On a slow tenant, the main thread waits on the network and the readers wait for it.

## Design

1. **Copies of the sign-in.** After the one CAC sign-in, the main thread makes one copy of the PnP client context per
   download worker (PnP.Framework `ClientContextExtensions.Clone`, same site address; test 3 of
   docs/reviews/2026-11-21-ion-round2-tests.md does the same). No new sign-in, no stored credential. A worker calls no
   PnP cmdlet: it uses only its own copy (a client context is not safe on two threads).
2. **Download workers.** A runspace pool of N workers (N = the PAM's answer, 2 to 4) in ChannelScan.psm1 (no new
   module: scripts/build_paste.py and Find-ChannelPII.ps1 list the modules). Each worker gets its copy, the file's
   server-relative path and the list size, reads the file with `OpenBinaryStream` into memory (never to disk), and puts
   its bytes done into a shared table. The main thread draws the bar from that table (a worker's Write-Progress does
   not reach ISE).
3. **Network limit.** Windows PowerShell 5.1 (.NET Framework) allows 2 connections per host by default. Ion sets the
   connection limit of the one site host to N + 1 (in memory; it ends when the PowerShell window closes). No other host.
4. **Planned order.** Before the first download, Ion plans the order from the full file list (the PAM's answer to
   question 2). A pure function makes the plan, so a test checks the order without timing.
5. **Check every download.** A download whose byte count is not the list size (for example a sign-in page) is a
   failure, as the streamed download is today. Size 0 in the list: the count is not checked.
6. **Fallback.** A failed worker download is tried one more time on the main thread (Get-FileBytes, as today). If that
   works, the copy of the sign-in is at fault: parallel downloads stop for the rest of the scan. If it also fails, the
   file is at fault: its NOT SCANNED row, as today, and parallel downloads go on.
7. **Time limit and cancel.** A download over 180 s (the reader limit) is stopped without waiting and tried on the main
   thread. The download pool closes in the scan's `finally`, before the one sign-out at the end of the run. Ctrl+C still
   stops the scan.
8. **Memory.** At most N downloads in flight plus 1 finished download that waits for a reader. A new download starts
   only after a finished one goes to the readers. Peak with the 25 MB limit and N = 3: 25 MB x (3 + 1 + 8 readers)
   = 300 MB, against 225 MB today.
9. **Screen and report.** One download bar for all files in flight (preview in question 4). "Time downloading" becomes
   the wall time with at least one download in flight. A new header line "Parallel downloads": "3", or "off (<error
   type name>)", with no message text (a message can hold a path). This is the device evidence.
10. **Kept rules.** No reconnaissance: the workers get only the paths of the listed files of the one site. Masked
    output: no path or name leaves the main thread unmasked (the worker gives back bytes and an error type name).
    Read-only. One sign-out at the end of the run. Detectors frozen: no change in src/Pii*.psm1 or the readers.
11. **Out of scope.** OneDrive web files (src/DriveScan.psm1) keep the one-by-one Get-FileBytes, whose signature does
    not change (another session works in that file). Version-history reads stay on the main thread.

## Assumptions

| Assumption | Source | Confidence |
|---|---|---|
| A copy of the sign-in downloads in a runspace on DoD365 | PnP.Framework Clone copies the sign-in handler; test 3 not run on the device (its Results: "Waiting on the PAM") | low: question 1 |
| 3 downloads at once are faster than one by one on the gov network | test 3 not run; the 2026-11-21 slow run (28:49) | medium |
| The connection limit of 2 per host applies in ISE | .NET Framework default for desktop programs | high |
| Nothing in the tests or the workbook needs today's row order inside a folder | grep of tests/ for row indexes: 2 hits, both one-file cases | medium |
| Fake contexts in tests must be script text and plain data in the worker | a ScriptMethod object is tied to the main runspace | high |

## Tests (tests/FakePnP.ps1, tests/ChannelScan.Tests.ps1)

- The order planner on a fixed list (sizes, types, folders).
- A fake worker download (script text, no ScriptMethod) gives the bytes; the rows are the same as the one-by-one scan.
- A worker gives the wrong size: the main thread tries again; parallel stops; header "off (...)".
- A file that fails on both: one NOT SCANNED row; parallel stays on.
- A worker over the time limit: stopped, the scan goes on.
- The pool closes after a cancel; no path or name in the header line.
