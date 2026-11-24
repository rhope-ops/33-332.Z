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

- Step 0: probe 2 run (OneDrive maps to its address; second sign-in failed); probe 3 merged (PRs 84, 86), the PAM runs it.
- Step 1: merged (PRs 82, 85). Detectors frozen after PR 85 (the PAM, 2026-11-24), except the roster header fix.
- Step 2: merged (PRs 80, 83). Step 3: merged (PR 81).
- Step 4 (OneDrive hybrid) and one sign-in per host (the PAM, 2026-11-24): wait on probe 3's second run.
- Step 5 (permission scope): waits on probe 3 (no permission read ran yet). Step 6 (parallel downloads): open.
- Added 2026-11-24: seeded recall test (PR 87; local baseline 46/53 found, 87%); roster header fix in work
  (claude/roster-headers); the PAM to run the recall test on the device after probe 3.
