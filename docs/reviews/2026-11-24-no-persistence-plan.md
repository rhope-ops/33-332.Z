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

- Step 0: open. Step 1: open. Step 2: open. Step 3: built (claude/workbook-qol), in review. Step 4: open. Step 5: open. Step 6: open.
