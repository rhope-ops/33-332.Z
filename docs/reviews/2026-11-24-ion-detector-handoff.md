# Ion detector fixes: handoff (2026-11-24)

Paused at the PAM's request. Nothing merged or released. Approvals by the PAM in chat, 2026-11-24 (the plan below):
merge PR 60 and PR 61 after a fresh review with no MAJOR; e-signature DoD IDs Low; an SSN swaps to a DoD ID only after a
signature word or when the digits touch the CAC name (S23, S24, E stay SSN); passport over SSN; split the 2 name defects
into their own PR; speed budget: +15% on prose and form rows, 2x on signature-heavy text, extra time accepted where new
IDs are found; workbook changes, all 4 groups (below). Record these in docs/DECISIONS.md (PR 61 already has its row).

## Open, in order

1. PR 60 (claude/ion-esig-dodid, head 013ca82): fresh review of the last commit (the sixth review's fixes: word cut, prose
   speed, dates, binary-search add-pass). Design: main's detectors unchanged; src/PiiDetectors.psm1 Find-SignedIdHits
   lists candidates, Add-SignedIdHits adds them after the overlap removal only where no kept hit is. Evidence:
   docs/reviews/2026-11-23-ion-esignature-layouts.md on that branch. Then add the DECISIONS rows and merge.
2. PR 61 (claude/ion-five-defects, head fd64a8e): fix the seventh review's MAJOR: the passport skip in Find-SsnHits
   must apply only to ASCII digits (`$m.Value -match '^[0-9 ]+$'`), with a test (full-width digits after "Passport:");
   optional MINOR: e-mail grow crosses Unicode spaces not in $MailNameEnd. Then a fresh review, merge main into the
   branch, rebuild paste/Ion-script.txt, merge.
3. The PAM tests the merged paste/Ion-script.txt on the gov device before users paste it again.
4. Workbook PR (preview first, lane C): mask the CAC name in Context; Ion version and Finished time on the Summary;
   CUI header and footer on every tab; review columns next to the hit; a Files tab; "Reviewed x of y" counts; the file
   name as the link; a confidence fill; aligned Summary counts and wrapped text; numeric High/Medium/Low cells;
   snippets cut at a word; 4 "Start here" steps (STE skill); landscape, fit-to-width print; a date check on Review
   date; a sample workbook as the preview and test base; layout tests.
5. Name PR: defects "name runs over a line break / takes the next label" and "middle initial without a dot". Build it
   like PR 60: main's name match first, then only narrow trims or adds, with a no-loss test. Earlier attempts and their
   review findings: PR 61 history up to 7595566.

## Method that worked

- Compare every change against main's modules (`git show origin/main:src/...` into a scratch folder) on generated
  texts; a hit main finds must keep its type, confidence and span. Give each reviewer its own scratch folder.
- Time as the fastest of 10 Get-PiiSpan calls per process, 3 processes, main and branch in turn (single runs vary 25%).
- `scripts/check.sh` before each push; the file (450) and function (60) line caps and the content check (no literal
  e-mail, phone or SSN shapes in tests: build them from parts) failed most often.
