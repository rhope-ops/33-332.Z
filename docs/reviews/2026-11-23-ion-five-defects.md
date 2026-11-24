# Ion: 3 detector defects of the R4e benchmark (2026-11-23, split 2026-11-24)

Source: archimedes_lever docs/reviews/2026-11-23-tr-r4e-ion-benchmark.md ("Ion: misses and over-flags", items 3b, 4, 8).
Tests: tests/Detectors.Tests.ps1, section "3 Ion defects of the R4e benchmark". All data invented.

The benchmark named 5 defects. The 2 name defects (a name runs over a line break or takes the next label; a middle initial
without a dot is cut) caused new losses against main in 4 review rounds. The PAM split them off (2026-11-24): they go to
their own pull request, built so that main's name match is kept and only narrow trims or adds are made.

| Defect | Before (main b278ab0) | After | Code |
|---|---|---|---|
| Street type in capitals ("42 ELM ST") | no hit | Home address | src/PiiPatterns.psm1 Find-HomeAddressHits: main's pattern first; then street types in capitals, only where main found nothing; such a hit is Low, or Medium with ", state ZIP", never High (not in small letters: "4 Games st") |
| 9-digit passport after "passport" | SSN Low | Passport Medium | Find-SsnHits: no SSN for 9 digits in a row with "passport" in the 30 characters before and no SSN label |
| "_" in an e-mail name ("first_last@...") | hit starts after the "_" | the whole address | src/PiiDetectors.psm1 Expand-MailOverUnderscore: after the unreadable-text filter, an e-mail hit right after a "_" grows left over the e-mail name, at most to the end of the hit before it, and not into another address ("@") |

## Checks (2026-11-24)

| | main | This branch |
|---|---|---|
| Benchmark corpus: labelled spans found (506 texts, 528 spans, seed 20261123; run outside the repo) | 438 | 448 |
| Spans lost, or found at a lower confidence | - | 0 |
| False positives (all Low) | 29 | 26 |
| High or Medium hits on the 126 decoys | 0 | 0 |
| Name matrix (6,240 cases): names lost or gained | - | 0 and 0 (names unchanged) |

Speed (budget: +15%, the PAM, 2026-11-24), main / branch: 1.2 MB prose 441 / 455 ms (1.03x) and 5,000 form rows
12.0 / 11.8 s (0.98x), second call in its own process, median of 5; 2,000 e-mails with "_" 203 / 200 ms (0.99x) and with
50-character names 193 / 206 ms (1.07x), fastest of 10 calls in each of 3 processes (single runs vary by 25% on this host).

## Review rounds

- First to fourth fresh reviews: name-rule losses (moved out, above); reading the e-mail on the raw text took a glued
  name or address (now the e-mail grows only up to the hit before it); the grow was quadratic (now one sorted pass);
  a grown e-mail could be dropped by the unreadable-text filter (now the grow runs after the filter).
- Fifth review: "CT" in capitals (a street type and a state) cut "12 Oak St Dunmore CT" + ZIP short; now main's address
  pattern runs first and the capital types only add where it found nothing.
- Sixth review: the grow walked into a glued .mil address and showed its name ("jdoe@...mil_jane@..."); it no longer
  grows when it meets an "@". The walk cost grew with the name length; it is one .NET call now. Capital-only addresses
  could be High or Medium on capital text ("HOME OFFICE: REPLACE 2 HARD DRIVE"); now Low, or Medium with ", state ZIP".
- Known, not fixed: capital-only addresses can be cut short ("APT" is no unit; "CT" without a comma); main finds nothing
  there, so nothing is lost.
- Known, not fixed: street types in capitals add Low false positives on capital text ("REPLACE 2 HARD DRIVE"); on 1.3 MB
  of such text that is 40,000 Low hits and about 8x main's time (linear; the same cost as main on the same text in title case).

## Decisions

- Passport over SSN: a bare 9-digit number after "passport" with no SSN label is a Passport (the PAM approved it in chat,
  2026-11-24: "approve all"; docs/DECISIONS.md). The overlap order in PiiDetectors.psm1 $DetectorOrder is unchanged.
- Split of the name defects into their own pull request (the PAM, 2026-11-24, question dialog).
