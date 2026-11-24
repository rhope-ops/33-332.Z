# Ion: the 5 detector defects of the R4e benchmark (2026-11-23)

Source: archimedes_lever docs/reviews/2026-11-23-tr-r4e-ion-benchmark.md ("Ion: misses and over-flags", items 1, 3b, 4, 8).
Tests: tests/Detectors.Tests.ps1, section "The 5 Ion defects of the R4e benchmark". All data invented.

| Defect | Before (main b278ab0) | After | Code |
|---|---|---|---|
| Street type in capitals ("42 ELM ST") | no hit | Home address | src/PiiPatterns.psm1 Find-HomeAddressHits: street type as written or in capitals (not in small letters: "4 Games st") |
| 9-digit passport after "passport" | SSN Low | Passport Medium | Find-SsnHits: no SSN for 9 digits in a row with "passport" in the 30 characters before and no SSN label |
| "_" in an e-mail name ("first_last@...") | hit starts after the "_" | the whole address | src/PiiDetectors.psm1 Expand-MailOverUnderscore: an e-mail hit right after a "_" takes the e-mail name before it, but no character of another hit |
| Name runs over a line break, takes the next label | "Pat Lee\nName", second name lost | "Pat Lee", "Sam Roe" | Find-NameHits: a later part starts on a later line only when that line does not start another field (a word, up to 3 small words, then ":" or "="); a label with ":", "=" or a dash and no letter after it stops a name |
| Middle initial without a dot ("JANE M", "Jane M Doe") | initial cut; "Member: Jane M Doe" no hit | initial kept, name found | Find-NameHits: a part can be one capital letter |

## Benchmark corpus (506 texts, 528 labelled spans, seed 20261123; run outside the repo)

| | Before | After |
|---|---|---|
| Labelled spans found (any confidence) | 438 | 448 |
| Spans lost, or found at a lower confidence | - | 0 |
| False positives (all Low) | 29 (SSN 12, Phone 5, e-mail 12) | 26 (SSN 9, Phone 5, e-mail 12) |
| Found spans with the exact start and end | 420 of 438 | 448 of 448 |
| High or Medium hits on the 126 decoys | 0 | 0 |

## Fresh review (2026-11-23) and name matrix

The first review found 4 recall losses against main (no-break spaces, a wrap after the second word, label words used as
names, an e-mail glued by "_" to another hit) and street-type false positives. All are fixed and tested. A name matrix
(3,240 cases: 6 labels, 12 name forms, 6 separators, 9 following texts) then gives: whole name found on main 2,460, on
this branch 3,036; 0 cases lost.

## Second fresh review (2026-11-24)

3 recall losses against main, all fixed and tested ("Second review of PR 61" in tests/Detectors.Tests.ps1): reading the
e-mail with its underscores took a glued name or address ("Name: Jane Doe_jdoe@..."), so the e-mail hit now grows to the
left only over characters no other hit holds; a wrapped part on a line with a label ("Doe SSN:") was cut, so a line is
another field only when a word, small words and ":" start it; label words in hyphenated names ("Poc-Smith") stopped a name.
Also as on main again: "\r" alone and blank lines between parts, and a government address before "_report.pdf".
Wider name matrix (6,240 cases: 16 name forms, 8 separators, 13 following texts): whole name on main 5,100, now 5,778; 0 lost.
Benchmark corpus: unchanged from the table above.

## Decisions to confirm

- Passport over SSN changes which type wins for a bare 9-digit number after "passport" (the overlap order in
  PiiDetectors.psm1 $DetectorOrder is unchanged). This is the agent's recommendation, not a PAM decision yet.
- A one-word name on a line of its own, then a label on the next line, is still not found (a name needs 2 parts).
