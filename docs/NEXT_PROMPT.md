# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Done 2026-11-20: kit sync, PII tool, bundle, site, Ion (PRs #2-#13; git log and the PR bodies hold the detail
   and the MINOR follow-ups). GitHub Pages on; denylist secret added.
9. Done 2026-11-21: PR #14 (PDF reader, fast; 093b2af), PR #19 (round 3 readers: OCR, old Office, .msg, RTF,
   OpenDocument, Visio, zip, e-mail; 7f12d69), PR #20 (network drives, 1-8 reader workers, one read per identical
   file, accountability columns with put-backs; bdb7dd3). Fresh reviews until no MAJOR; each MAJOR fix has a test.
10. The PAM to test on the gov device: ISE paste of paste/Ion-script.txt; a drive folder scan (runspaces in ISE, FIPS
   SHA-256); Windows OCR on an image; SharePoint version history ("PII added by"); colours; Excel opening.
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
14. 2026-11-21 (this session): PRs 30, 31, 34, 36, 37 merged: moving download bar (streamed, MB/s), wait bar, Wingdings and
    picture-font hits gone, PDF glyph names (/Differences), Low-hit noise guard, workbook "Last saved"/"History note",
    screen polish B1-B6, cleanup A1-A5, "Scan now? [Y/n]" (C1). Audit list: docs/reviews/2026-11-21-ion-polish-audit.md.
15. The PAM to test on the gov device: the streamed download bar (falls back if the CAC cookie is not sent); a re-run on
    the "Wingdings" PDF; Ctrl+C during a scan (the last lines); the window-width cut in ISE; the Teams-link test
    (docs/reviews/2026-11-21-teams-lookup-test.md, results A and B).
16. Open: C2, C4-C6 of the audit list (each needs a plan the PAM approves). Review MINORs carried: WinAnsi C1 quotes and
    packed guillemets in the noise guard (PR 31 review).
