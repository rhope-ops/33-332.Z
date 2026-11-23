# Ion e-signature DoD IDs: 25 layouts, before and after (2026-11-23)

Task: the coordinator session asked for the fix at the PAM's request (2026-11-23: "start the Ion fix now, it is a PII risk").
Evidence: archimedes_lever docs/reviews/2026-11-23-tool-retirement-plan.md section 1 ("Ion e-signature DoD IDs": 25 layouts,
3 give SSN, 7 give no hit). That source gives counts only. This session first built its own 25 layouts (S01-S25) from a
matrix of 1,110 texts; the source's 25 layouts then came from the tool retirement session (O01-OJ, below). Both sets are tests.

## Rule (src/PiiPatterns.psm1, Find-SignedIdHits)

- A CAC name in capitals, or in any case after "signed", "signature", "CN=" or "DN:" (within 40 characters before), then
  the 10 digits; or the 10 digits, then a CAC name in capitals (the ID on the line above the name).
- A line break is allowed inside or after a name part; a part after the first can have up to 3 words ("DOE.MARY ANN.").
- Between the name and the digits: spaces, tabs or no-break spaces, with at most one line break.
- Between any two digits: one space, tab, no-break space or line break; after a CAC name also a dash. No other digit may
  follow on the same line.
- The 10 digits alone after a signature word also count, but not a date (19xx or 20xx then a separator), and not a phone
  written 3-3-4 with spaces (it stays Phone; before this rule, "Signed copy, cell" and a spaced number gave DoD ID Low).
- Every such hit is DoD ID, Low (the PAM, 2026-11-21: e-signature DoD ID is Low, history not read).
- An unlabelled SSN match inside such a hit is dropped. A labelled SSN and a dashed SSN stay SSNs.

## The source's 25 layouts (tests/Signature.Tests.ps1, O01-OJ; same shapes, invented name and ID)

| Result | main b278ab0 | This PR |
|---|---|---|
| DoD ID Low | 10 (1, 2, 6, 7, 9, 10, 12, 13, 14, I) | 23 |
| SSN | 3 (A Medium, E Low, G Low) | 1 (A, Medium) |
| No hit | 12 (3, 4, 5, 8, 11, 15, B, C, D, F, H, J) | 1 (F) |

The source counted 9 DoD ID on its run (main 6cbc517): it had 14 ("O'BRIEN.SEAN.P.") as no hit; b278ab0 finds it.
By design: A ("DOE JOHN A" then a 3-2-4-1 number, no dots, no signature word) stays SSN Medium, so it is still reported;
F (10 spaced digits, no name, no signature word) is not reported, as 10 bare digits never are.

## This session's 25 layouts (tests/Signature.Tests.ps1; invented name DOE.JANE.Q. and an invented ID; expected: DoD ID Low, no SSN)

| Layout | Before (main b278ab0) | After |
|---|---|---|
| S01 base ("Digitally signed by" NAME ID, date line) | DoD ID Low | DoD ID Low |
| S02 CN= distinguished name | DoD ID Low | DoD ID Low |
| S03 no middle name | DoD ID Low | DoD ID Low |
| S04 suffix (JR.) | DoD ID Low | DoD ID Low |
| S05 hyphen and apostrophe | DoD ID Low | DoD ID Low |
| S06 digits 3-2-4-1 | DoD ID Low | DoD ID Low |
| S07 every digit apart | DoD ID Low | DoD ID Low |
| S08 digits 5-5 | DoD ID Low | DoD ID Low |
| S09 line break (LF) before the ID | no hit | DoD ID Low |
| S10 CRLF before the ID | no hit | DoD ID Low |
| S11 space before the ID | no hit | DoD ID Low |
| S12 tab before the ID | no hit | DoD ID Low |
| S13 no-break space before the ID | no hit | DoD ID Low |
| S14 line break inside the ID | no hit | DoD ID Low |
| S15 line break after a name dot | no hit | DoD ID Low |
| S16 line break inside a name | DoD ID Low | DoD ID Low |
| S17 title case after "signed" | no hit | DoD ID Low |
| S18 two-word first name | no hit | DoD ID Low |
| S19 no-break-space gaps in the ID | no hit | DoD ID Low |
| S20 "signed", ID 3-2-4-1, no name | no hit | DoD ID Low |
| S21 "signed", ID, no name | no hit | DoD ID Low |
| S22 file name with underscores | no hit | DoD ID Low |
| S23 space, then ID 3-2-4-1 | SSN Medium | DoD ID Low |
| S24 CRLF, then ID 3-2-4-1 | SSN Medium | DoD ID Low |
| S25 one digit, then nine | SSN Low | DoD ID Low |
| Total | 9 DoD ID, 3 SSN, 13 no hit | 25 DoD ID |

Matrix (1,110 texts, scratchpad only): before 45 DoD ID, 377 SSN, 688 no hit; after 1,040 DoD ID, 35 SSN, 35 no hit. The 70
left are by design: a title-case name with no signature word (60), or digits with no name and no signature word (10).

## Checks

- Recall: the R4e benchmark corpus (506 texts, seed 20261123, generator in the R4e review) gives the same Get-PiiSpan output
  before and after on all 506 texts: 0 spans lost, 0 confidences changed, 0 High or Medium hits on the 126 decoys.
- Decoys in the tests: a date after "signed", a signature with no ID, a phone after "signed", 11 digits, lower-case and
  title-case names with no signature word, a digit after the ID on the same line: no DoD ID.
- Masking: values with gaps, tabs, no-break spaces and line breaks show no full ID in the masked value, the snippet,
  Protect-PiiText or Protect-PiiName.
- Speed: 5 runs of 80,000 to 100,000 characters (dotted capitals, short lines, digit runs) under 4 s together. A first
  version let a name part run over every line (205 s); each part is now at most 40 letters with one line break.
- One old test changed: "Digitally signed by" 9 digits then 1 digit was 0 hits; it is now DoD ID Low.

## Known gaps (not fixed here)

- A labelled DoD ID with gaps ("EDIPI 123 45 6789 0") is still an SSN Medium: the label rule takes 10 digits in a row only.
- A title-case or lower-case CAC name with no signature word is not a signature (precision: "backup.config.").
- A line break that splits one name part twice, or the ID in two places, is not found.
