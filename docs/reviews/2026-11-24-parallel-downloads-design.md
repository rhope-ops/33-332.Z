# Design: parallel downloads in a planned order (step 6, 2026-11-24)

Plan: docs/reviews/2026-11-24-no-persistence-plan.md, step 6. Approval: docs/DECISIONS.md, 2026-11-21 row "Ion round 2
after the gov run", item (3): "2-4 parallel downloads, ordered from the full file list" (the PAM: "orient files for
download to appear at the beginning"). It replaces the 2026-11-20 row's "No parallel downloads: each would need its own
CAC sign-in", if test 3 (below) or the first device run shows that a copy of the sign-in works in a worker.

Status: approved 2026-11-24 (docs/DECISIONS.md: 3 at once, per folder largest first, build with fallback, one bar).
Built and merged in PR 115 (2026-11-24). Waits on the device: the header line "Parallel downloads" of the next Ion run.

## Today

The SharePoint and Teams scan downloads one file at a time on the main thread (src/ChannelScan.psm1 Get-FileBytes,
Invoke-OneFileScan). The readers (src/ReaderPool.psm1, 1 to 8 runspaces) read while the main thread downloads the next
file. On a slow tenant, the main thread waits on the network and the readers wait for it.

## Design

1. **Copies of the sign-in.** After the one CAC sign-in, the main thread makes one copy of the PnP client context per
   download worker (PnP.Framework `ClientContextExtensions.Clone`, same site address; test 3 of
   docs/reviews/2026-11-21-ion-round2-tests.md does the same). No new sign-in, no stored credential. A worker calls no
   PnP cmdlet: it uses only its own copy (a client context is not safe on two threads).
2. **Download workers.** A runspace pool of N = 3 workers (the PAM, 2026-11-24) in ChannelScan.psm1 (no new
   module: scripts/build_paste.py and Find-ChannelPII.ps1 list the modules). Each worker gets its copy, the file's
   server-relative path and the list size, reads the file with `OpenBinaryStream` into memory (never to disk), and puts
   its bytes done into a shared table. The main thread draws the bar from that table (a worker's Write-Progress does
   not reach ISE).
3. **Network limit.** Windows PowerShell 5.1 (.NET Framework) allows 2 connections per host by default. Ion sets the
   connection limit of the one site host to N + 1 (in memory; it ends when the PowerShell window closes). No other host.
4. **Planned order.** Before the first download, Ion plans the order from the full file list (the PAM, 2026-11-24: per folder, the
   files to download first, largest first, then the files not downloaded). A pure function makes the plan, so a test checks the order without timing.
5. **Check every download.** A download whose byte count is not the list size (for example a sign-in page) is a
   failure, as the streamed download is today. Size 0 in the list: the count is not checked.
6. **Fallback.** A failed worker download is tried one more time on the main thread (Get-FileBytes, as today). If that
   works, the copy of the sign-in is at fault: parallel downloads stop for the rest of the scan. If it also fails, the
   file is at fault: its NOT SCANNED row, as today, and parallel downloads go on. Two more cases keep this rule: the main
   thread also gets a size other than the list size (the list size is wrong: parallel stays on); no worker is left
   (each new copy failed after a time limit): "off (NoWorker)".
7. **Time limit and cancel.** A download over 180 s (the reader limit) is stopped without waiting and tried on the main
   thread. The download pool closes in the scan's `finally`, before the one sign-out at the end of the run. Ctrl+C still
   stops the scan.
8. **Memory.** At most N downloads in flight plus 1 finished download that waits for a reader. A new download starts
   only after a finished one goes to the readers. Peak with the 25 MB limit and N = 3: 25 MB x (3 + 1 + 8 readers)
   = 300 MB, against 225 MB today.
9. **Screen and report.** One download bar for all files in flight (approved preview). "Time downloading" is the wall
   time with a worker download in flight, plus each main-thread download (the download clock pauses for it). A new header line "Parallel downloads": "3", or "off (<error
   type name>)", with no message text (a message can hold a path). This is the device evidence.
10. **Kept rules.** No reconnaissance: the workers get only the paths of the listed files of the one site. Masked
    output: no path or name leaves the main thread unmasked (the worker gives back bytes and an error type name).
    Read-only. One sign-out at the end of the run. Detectors frozen: no change in src/Pii*.psm1 or the readers.
11. **Out of scope.** OneDrive web files (src/DriveScan.psm1) keep the one-by-one Get-FileBytes, whose signature does
    not change (another session works in that file). Version-history reads stay on the main thread.

## Assumptions

| Assumption | Source | Confidence |
|---|---|---|
| A copy of the sign-in downloads in a runspace on DoD365 | PnP.Framework Clone copies the sign-in handler; test 3 not run on the device (its Results: "Waiting on the PAM") | low: the PAM chose to build with fallback; the next device run answers it |
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
- Built as well: two folders (the bar shows in folder 2; review of PR 115, MAJOR, fixed).
- Not built: a cancel test.

## Open after PR 115 (fresh review of the fix, MINOR; fixed 2026-11-25 on branch claude/speed-minors-2026-11-25)

- "Time downloading" is a little too low for a '#' or '%' file while worker downloads run: the download clock stays
  stopped while that file goes to the readers (src/ChannelScan.psm1 Invoke-ParallelGroup); its code comment says
  "burst clock", it is the Busy clock.
- No tests for: "off (NoWorker)", the list-size fault, the per-job byte keys, the timing.
- "NoWorker" is not an error type name (like "WrongSize" and "NoResult" before it).
- The synopsis of Complete-Download does not name the list-size exception.

## Device result (the PAM, 2026-11-25, same site as the day before; counts only)

| | 2026-11-24 (one at a time) | 2026-11-25 (build 5f4437a0, 3 at once) |
|---|---|---|
| Time used | 68:45 | 39:36 (42% less) |
| Time downloading | 45:44 | 38:49 |
| Time reading files (all readers) | not sent | 64:54 |
| Parallel downloads | none | 3 |

The copy of the sign-in works in a worker on DoD365 (the low assumption above is now proven). Downloads still run in
38:49 of 39:36: the network is still the slowest part. The PAM: "much more demanding" (the fan or heat), so speed
modes follow (docs/DECISIONS.md, 2026-11-25): Normal (Enter) = 2 downloads and half the processors for readers; Fast =
3 downloads and one reader per processor; both at below-normal priority.

Speed modes merged 2026-11-25 (PR 119). Open after its fix review (MINOR and NIT, for a later lane B change):
- tests/EndToEnd.Tests.ps1 "Enter gives Normal again in the next run" cannot fail: each run re-imports the modules,
  which sets Normal.
- The Fast pool test proves nothing on a 1-processor runner.
- README Speed line: Normal is also capped at 8 readers; "processors read files" and "reader" name one thing twice.
- src/Find-ChannelPII.ps1 comment "Fast = all": Fast is one reader per processor (at most 8) and 3 downloads.

Fixed 2026-11-25 (branch claude/speed-minors-2026-11-25), all items of both lists: the Busy clock pauses for the
download of a '#' or '%' file only (Invoke-OneFileScan -Pool); "off (<reason>)" names Ion's own reasons WrongSize,
NoResult and NoWorker (New-DownloadPool); the Complete-Download synopsis names the list-size fault; new tests: NoWorker,
the list-size fault, per-job byte keys, the timing (wall time, 3 at once), the '#' file clock; the Enter test sets Fast
right before the answer; the reader pool tests set 4 processors; README Speed line and the speed comment.

Normal mode on the device (the PAM, 2026-11-25): the fan is quiet with mild spikes (Fast: much more intense). No
times: the PAM stopped the run, because the Microsoft sign-in was very slow that day (3-5 min to the window). Screen
note above the speed question, from the PAM's words ("one uses less power and memory and the other is resource
intensive"): "Normal: uses less power and memory." and "Fast: resource intensive."

