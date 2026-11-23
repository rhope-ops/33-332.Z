# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

0. Fleet v2 3B: PRs 51, 53, 54 merged; kit 11.0.1 sync (branch claude/fleet-v2-kit-11) waits on the PAM's .check row;
   PR 55 waits on that sync. Evidence: docs/reviews/2026-11-23-fleet-v2-stage3b.md.
1. Done 2026-11-20: kit sync, PII tool, bundle, site, Ion (PRs #2-#13; git log and the PR bodies hold the detail
   and the MINOR follow-ups). GitHub Pages on; denylist secret added.
9. Done 2026-11-21: PRs #14 (PDF reader), #19 (round 3 readers), #20 (network drives, workers, accountability).
10. The PAM to test on the gov device: ISE paste of paste/Ion-script.txt; a drive folder scan (runspaces in ISE, FIPS
   SHA-256); Windows OCR on an image; SharePoint version history ("PII first saved by"); colours; Excel opening.
11. Speed items 3 and 4 not built (measured 2026-11-21): a one-lookup token test gave no gain (11.2 s either way on
   long.pdf: PowerShell's cost per statement, not the comparisons); the xref table would save only the 0.11 s header
   scan of a 2.1 s object read (old object versions must still be read). Workers gave 1.5x on 4 cores (GC bound?).
12. Tracked minors (PR bodies of #14, #19, #20): short spaced line under a heading of another size; OCR time outside
   the PDF limit; a PDF with text and scans not OCR'd; meter misses
   file-name hits; no time limit per worker; old versions not size-capped; masked dates/addresses compare equal.
   A second address pasted after a Teams link without a groupId can supply that groupId (the site check holds).
   Teams-link lookup and OCR wait on the PAM's test snippet results.
13. README rework 2026-11-21: on main at the PAM's request (no wait for network-drives).
    Steps and open items: docs/reviews/2026-11-21-readme-rework.md.
14. Done 2026-11-21: PRs 30-39 (bars, symbol-font hits, glyph names, workbook columns, screen polish, "Scan now?").
    Gov results: bar works; "Wingdings" PDF 1 hit left (a wide PDF: open); Ctrl+C saves; Teams lookup test 1 works.
    Next: Teams test 2 (docs/reviews/2026-11-21-teams-lookup-test.md). Roadmap and open items: docs/DECISIONS.md
    2026-11-21 rows; C2, C4-C6 of docs/reviews/2026-11-21-ion-polish-audit.md; the output folder layout.
15. 2026-11-21 PR 46: gamma on the start line only; file on its own scan-bar line; "Downloading..."; "done / total";
    not-scanned groups split (label, password, PDF time limit, PDF parts). Waits on the PAM: ISE layout test
    (docs/reviews/2026-11-21-ise-progress-test.md); the reason text of the "damaged" PDFs on the next run; decisions
    on item permissions/views and parallel SharePoint downloads (open questions of that chat).
16. Plan approved 2026-11-21 (docs/DECISIONS.md, newest row): step 1 = Teams link alone (built) + 2 test snippets
    (parallel download, view counts: docs/reviews/2026-11-21-ion-round2-tests.md); step 2 = permission columns;
    step 3 = parallel downloads in a planned order, and version-history reads off the main thread (slow run 28:49).
17. 2026-11-23 plan approved (docs/reviews/2026-11-23-ion-esignature-onedrive-plan.md; one branch, PRs in order):
    PR 57 e-signature DoD IDs built and reviewed; waits on the PAM's shape lines (docs/reviews/2026-11-23-esignature-shape-test.md)
    before merge. Next: PR 2 hang (SharePoint download and per-file time limits), PR 3 OneDrive (probe first; online-only: ask).
