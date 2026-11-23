# Ion: the 5 detector defects of the R4e benchmark (2026-11-23)

Source: archimedes_lever docs/reviews/2026-11-23-tr-r4e-ion-benchmark.md ("Ion: misses and over-flags", items 1, 3b, 4, 8).
Tests: tests/Detectors.Tests.ps1, section "The 5 Ion defects of the R4e benchmark". All data invented.

| Defect | Before (main b278ab0) | After | Code |
|---|---|---|---|
| Street type in capitals ("42 ELM ST") | no hit | Home address | src/PiiPatterns.psm1 Find-HomeAddressHits: street type in any case |
| 9-digit passport after "passport" | SSN Low | Passport Medium | Find-SsnHits: no SSN for 9 digits in a row with "passport" in the 30 characters before and no SSN label |
| "_" in an e-mail name ("first_last@...") | hit starts after the "_" | the whole address | src/PiiDetectors.psm1: the e-mail detector reads the text with its underscores |
| Name runs over a line break, takes the next label | "Pat Lee\nName", second name lost | "Pat Lee", "Sam Roe" | Find-NameHits: only the second part can start on the next line; labels stop a name |
| Middle initial without a dot ("JANE M", "Jane M Doe") | initial cut; "Member: Jane M Doe" no hit | initial kept, name found | Find-NameHits: a part can be one capital letter |

## Benchmark corpus (506 texts, 528 labelled spans, seed 20261123; run outside the repo)

| | Before | After |
|---|---|---|
| Labelled spans found (any confidence) | 438 | 448 |
| Spans lost, or found at a lower confidence | - | 0 |
| False positives (all Low) | 29 (SSN 12, Phone 5, e-mail 12) | 26 (SSN 9, Phone 5, e-mail 12) |
| Found spans with the exact start and end | 420 of 438 | 448 of 448 |
| High or Medium hits on the 126 decoys | 0 | 0 |

## Decisions to confirm

- Passport over SSN changes which type wins for a bare 9-digit number after "passport" (the overlap order in
  PiiDetectors.psm1 $DetectorOrder is unchanged). This is the agent's recommendation, not a PAM decision yet.
- A one-word name on a line of its own, then a label on the next line, is still not found (a name needs 2 parts).
