# Ion e-signature DoD IDs: 25 layouts, before and after (2026-11-23)

Task: the coordinator session asked for the fix at the PAM's request (2026-11-23: "start the Ion fix now, it is a PII risk").
Evidence: archimedes_lever docs/reviews/2026-11-23-tool-retirement-plan.md section 1 ("Ion e-signature DoD IDs": 25 layouts,
3 give SSN, 7 give no hit). That source gives counts only. This session first built its own 25 layouts (S01-S25) from a
matrix of 1,110 texts; the source's 25 layouts then came from the tool retirement session (O01-OJ, below). Both sets are tests.

## Rule (src/PiiDetectors.psm1 Find-SignedIdHits and Add-SignedIdHits)

Since the third review: main's SSN and DoD ID detectors are unchanged. Find-SignedIdHits gives candidates; Get-PiiSpan adds
a candidate after the overlap removal, only where no kept hit is. The one exception: SSN hits (not High, no dashes) wholly
inside a candidate give way when a signature word is just before the candidate or its digits touch the name's last dot.

- A CAC name in capitals, or in any case after "signed", "signature", "CN=" or "DN:" (within 40 characters before), then
  the 10 digits; or the 10 digits, then a CAC name in capitals (the ID on the line above the name).
- A line break is allowed inside or after a name part; a part after the first can have up to 3 words ("DOE.MARY ANN.").
- Between the name and the digits: spaces, tabs or no-break spaces, with at most one line break.
- Between any two digits: one space, tab, no-break space or line break; after a CAC name also a dash. No other digit may
  follow on the same line.
- The 10 digits alone on one line after a signature word also count (gaps: spaces only; no small word after them), but
  not a date (19xx or 20xx then a separator, or yyyymmdd) and not a phone (3-3-4 with any one separator).
- A phone shape (3-3-4) after white space is no candidate. Any hit the other detectors find (phone, e-mail, date, passport,
  card, dashed or labelled SSN) keeps its place: a candidate that overlaps it is not added.
- Every added hit is DoD ID, Low (the PAM, 2026-11-21: e-signature DoD ID is Low, history not read); High with a DoD ID label.

## The source's 25 layouts (tests/Signature.Tests.ps1, O01-OJ; same shapes, invented name and ID)

| Result | main b278ab0 | This PR |
|---|---|---|
| DoD ID Low | 10 (1, 2, 6, 7, 9, 10, 12, 13, 14, I) | 22 |
| SSN | 3 (A Medium, E Low, G Low) | 2 (A Medium; E Low, open with the PAM) |
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
| S23 space, then ID 3-2-4-1 | SSN Medium | SSN Medium (open with the PAM) |
| S24 CRLF, then ID 3-2-4-1 | SSN Medium | SSN Medium (open with the PAM) |
| S25 one digit, then nine | SSN Low | DoD ID Low |
| Total | 9 DoD ID, 3 SSN, 13 no hit | 23 DoD ID, 2 SSN (open with the PAM) |

Matrix (1,110 texts, scratchpad only): before 45 DoD ID, 377 SSN, 688 no hit; after 1,034 DoD ID, 39 SSN, 37 no hit. The 76
left are by design: a title-case name with no signature word (60), digits with no name and no signature word (10), or
no name and a line break inside the digits (6).

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

## Fresh review of 29fe0f5 (2026-11-23) and the fixes

The review found 3 MAJOR issues, all fixed and tested in tests/Signature.Tests.ps1 ("Review of PR 60"):
- A phone after a CAC name and white space, or across lines after a signature word, became DoD ID Low. Now a 3-3-4 phone
  shape is no ID after white space, and the ID alone after a signature word is on one line.
- A dashed SSN after a CAC name and a tab or a line break became DoD ID Low. Dashes now count only right after the name's
  last dot, and then with no line break.
- The SSN check looked at every signature ID for every SSN match (4,000 rows: 73 s). It now uses a lookup per character
  (4,000 rows: 3.2 s; main: 3.3 s).
Also: an ID with gaps after a label in a signature ("Signed, EDIPI 123 45 6789 0") was lost (no hit on main too); it is now
DoD ID High. An ID before "REF.NO." is no ID. The 1,110-text matrix now gives 1,034 DoD ID: 6 texts with no name and a line
break inside the ID give the same result as main.

## Second fresh review of 8329c97 (2026-11-24) and the fixes

4 MAJOR, all fixed and tested ("Review of PR 60" in tests/Signature.Tests.ps1); each case now gives main's result:
- A dashed phone right after a CAC name became DoD ID: the phone guard now covers the dash form too.
- A dashed SSN with a lone digit after it became DoD ID: a dashed SSN is never dropped now.
- An e-mail address with a CAC-like name and 10 digits after "signed" became DoD ID: no ID ends at "@".
- Slower than main: the any-case and the digits-alone scans run only when the text holds a signature word, and
  Find-SsnHits and Find-DodIdHits share one result per text. Best of 3 (main / this branch): 1.2 MB prose 323 / 239 ms;
  5,000 rows of 10 digits 264 / 199 ms; 5,000 rows of 5+5 digits 30 / 30 ms.
Minor, fixed: "signature" is not a signature word for the older SSN rule (Test-SsnPartOfLonger, as on main); an ID alone
after a signature word has space gaps only and no small word after it; yyyymmdd dates are no ID.

## Third fresh review of 878b867 (2026-11-24) and the redesign

4 MAJOR: main DoD ID hits lost (an "@" or a tab and a digit after the ID, a name part over 40 letters); new DoD ID hits that
took over e-mail, date of birth and passport hits; SSNs dropped in a numbered roster ("DOE.JANE.Q. 123 45 6789" then a line
"2 ROE.JOHN.A."); and slower runs. Three rounds of regex fixes each found new shapes, so the design changed (Rule, above):
the new rules can only add. A repo test proves it ("pass only adds": 3,584 texts with phones, SSNs, dates, passports, cards
and e-mails next to CAC names and signature words; a broken copy of the pass fails it with 108 losses). A run of 21,420
such texts in the scratchpad, main against this branch: every main hit kept, except 936 SSNs inside a signature ID that
became DoD ID Low (10 of them with a labelled name near, which goes from High to Low, as e-signature IDs never raise a name).
Open with the PAM: layouts S23, S24 and E (a CAC name, white space, then an SSN shape and one more digit, no signature word)
look the same as a roster row. They stay SSN, as on main, until the PAM decides.
Speed, each run in its own process, main then branch, 3 runs each: 1.2 MB prose 460-564 / 521-569 ms; 5,000 form rows
13.5-14.6 / 13.8-14.4 s; 2,000 signature blocks 1.7-1.8 / 3.1-3.2 s.

## Fourth fresh review of a397614 (2026-11-24) and the fixes

3 MAJOR, fixed: an added ID with a DoD ID label raised a nearby name to High (now an added ID never raises a name); in
unreadable text a swapped SSN was lost (no swap or add where the unreadable-text filter would drop the new Low hit); one
signature word made the any-case and digits-alone scans read the whole text (now only 400 characters after each word,
joined where they overlap, each match checked again on the whole text). Patterns start with a letter or digit (faster),
and a text with no digit skips the finder. The PAM set the budget (2026-11-24): signature-heavy text up to 2x main.
Speed, second call in each process, median of 5, main / branch: 1.2 MB prose with digits 433 / 494 ms (1.14x); 1.2 MB prose,
one signature word 472 / 428 ms; 5,000 form rows 10.6 / 10.9 s (1.03x); 2,000 signature blocks 1.32 / 2.03 s (1.53x).

## Known gaps (not fixed here)

- A labelled DoD ID with gaps ("EDIPI 123 45 6789 0") is still an SSN Medium: the label rule takes 10 digits in a row only.
- A title-case or lower-case CAC name with no signature word is not a signature (precision: "backup.config.").
- A line break that splits one name part twice, or the ID in two places, is not found.
