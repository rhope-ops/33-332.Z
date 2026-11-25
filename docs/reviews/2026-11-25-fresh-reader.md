# Fresh-reader test (2026-11-25)

Files: the step lists in docs/reviews/2026-11-21-ion-round2-tests.md ("How to run test 4"),
docs/reviews/2026-11-25-device-probe-7.md and docs/reviews/2026-11-25-device-probe-8.md ("How to run it").
Reader: the `reviewer` subagent (strong tier), a new agent each round, given only the step lists.

| Round | Marked steps | Result |
|---|---|---|
| 1 | "ISE output pane" (ISE calls it the Console pane); "On Question N" names no pane; "the first line" not named; the raw link opened in the folder tab closes that page; run A and run B written as 2 steps in order; the e-mail steps (new e-mail, body, keep the draft for run B) | All fixed: Console pane, the version line as the first line, a new tab for the raw link, one step per run, separate e-mail steps |
| 2 | No action when the version line is wrong; the "continue running scripts" window can open during the wait; run B finds the tabs of run A open; test 4 step 1 asked for a fact the page does not show; steps 1-2 of probe 8 named no screen | All fixed: "If it shows another version, close ISE and start this list again"; "While it runs, if ..."; close the tabs of this list after run A; screen named in each first step |

Not tested by a reader: if the "continue running scripts" window opens at F5 or at the sign-in (the device run
shows it; the step is a condition, so both work).
