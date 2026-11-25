# Plan: no-persistence improvements (approved 2026-11-24)

Approved by the PAM in the question dialog, 2026-11-24: "Approve all". Ion saves nothing between runs except the
workbooks (the PAM: "the script shouldn't save anything between sessions yet"; caching and rescans belong to a later app).

| Step | PR | Who / model | Time | Gate |
|---|---|---|---|---|
| 0 | Device probe (.txt): OneDrive registry shape (masked: no names or addresses printed) and which PnP.PowerShell 1.12.0 permission reads work without admin | orchestrator | 30 min + the PAM's run | the PAM runs it |
| 1 | Month-name dates count as dates before an e-signature ID (PR 76 review MINOR) | worker, Opus | 30 min | lane B |
| 2 | Several sites per run: one target per line, an empty line ends the list; one CAC sign-in and one workbook per site; a final line per site | orchestrator | 1.5 h | lane B |
| 3 | QOL: Review date must be a date; the CAC name masked in Context; a 4-step "Start here" box on Summary | orchestrator | 1 h | lane B |
| 4 | OneDrive hybrid: files on disk read locally; online-only files and the who/when history through PnP (one sign-in, nothing on disk); local only if PnP fails | orchestrator | 2.5 h | probe result, lane B |
| 5 | Permission scope column (Hits, Files): Anyone link, Whole organization, Site members, Specific people (count only); no names | worker, Opus | 2 h | probe result, lane B |
| 6 | Parallel downloads in a planned order (approved 2026-11-21, step 3) | worker, Opus | 2.5 h | lane B |

About 10 agent-hours; at most 2 agents at a time; one fresh reviewer (Opus) per PR. Steps 1-3 first; 4 and 5 after
the probe; 6 last.

Decisions of the same dialog (docs/DECISIONS.md, 2026-11-24): OneDrive hybrid; several sites, one workbook per site;
permissions as scope only ("No reconnaissance": no user names listed); month-name dates are dates.

## Status

- Step 0: probe 2 run; probe 3 run twice (second run recorded, PR 89): connect without sign-out works; the OneDrive
  file was not in its library. Probe 4 (PR 91): OneDrive web files at the mapped local path; the PAM runs it.
- Step 1: merged (PRs 82, 85). Detectors frozen after PR 85 (the PAM, 2026-11-24), except the roster header fix.
- Step 2: merged (PRs 80, 83). Step 3: merged (PR 81).
- One sign-in per host: merged (PRs 90, 92): no sign-out between sites, one at the end of the run; full-address site check.
  First device result (the PAM, 2026-11-24): 2 team sites in one run, 1 sign-in window, no PIN, no script error
  (counts not copied).
- Step 4 (OneDrive hybrid): unblocked. Probe 4 runs 1 and 2 (2026-11-24): the local-to-web mapping and path A work (9 of 9).
  Design approved 2026-11-24 (docs/reviews/2026-11-24-onedrive-hybrid-design.md); built 2026-11-25 for the own OneDrive (PR 109, src/DriveScan.psm1); probe 5 (2026-11-25): the folder
  listing works at any depth, but the map is wrong for shortcuts and synced libraries: they stay local only.
- Step 5 (permission scope): waits on probe 3 (no permission read ran yet). Step 6 (parallel downloads): open.
- Added 2026-11-24: seeded recall test (PR 87; local baseline 46/53 found, 87%); roster header fix merged
  (PR 96) and passport rule merged (PR 97): local recall 53/53, false 0. Build ID on the start screen (PR 95).
  Roster follow-ups (review of PR 96, MINOR): a csv/tsv row that starts with an empty field loses its columns; the
  1 MB cap counts characters and has no note in the report (old versions over the cap can make roster hits look new
  in the version check); README and the Method tab do not explain header labels; roster row has no cell location.
  Recall on the device: run 2 35/53 (3 files got a Purview encryption label after run 1: not scanned, correctly
  reported); run 3 with new files 53/53, false 0 (docs/reviews/2026-11-24-recall-test.md, Result).
