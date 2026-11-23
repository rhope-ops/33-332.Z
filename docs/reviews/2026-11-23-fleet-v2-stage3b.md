# Fleet v2 Stage 3B: evidence (2026-11-23, America/Chicago)

Plan: archimedes_lever docs/reviews/2026-11-23-fleet-v2-plan.md (branch claude/keen-cori-eiotd2, 99b8661; read by a
read-only clone because add_repo was denied). Coordinator: session_01PyXpdUat26YdmzaenEFPpj.

## Baseline gates (main 406653a, before any change)

| Gate | Exit | Count |
|---|---|---|
| `bash scripts/check.sh` | 0 | 29 passed, 2 skipped, of 31 |
| `python3 scripts/build_paste.py --check` | 0 | PASS, 2 files |
| `pwsh -NoProfile -File tests/Run-Tests.ps1` | 0 | 535 passed, 0 failed |
| `(cd ref && sha256sum -c SHA256SUMS)` | 0 | 2 OK |
| `bandit -r scripts/ -ll` (1.9.4) | 1 | 1 High: B602 shell=True, kit-locked scripts/bloat_check.py:241 (kit 11.0.0 fixes it) |

## PRs

| PR | Content | Lane |
|---|---|---|
| 1 | DEC 8 one bundle file; LICENSE, .editorconfig, .gitattributes, dependabot, .nojekyll, site.config.json, threat model, COMPLIANCE.md, build_paste_test.py | B |
| 2 | Splits <= 450; Protection negative PII-leak tests | B |
| 3 | index.html: Option A tokens, CSP, skip link, aria-live, theme-color, SVG files | B (DEC 5 approves the tokens) |
| 4 | Guard files: check.local.sh (ref/SHA256SUMS, build_paste_test, PSScriptAnalyzer hash), security.yml | C |

## Notes

- `.check` line `code-generated: paste/*.ps1` matches no file after DEC 8 (stale; `.check` is a guard file: left for the PAM).
- `.gitattributes`: `git add --renormalize .` gave 0 changes.

## Results

- PR 1 (#51), merged 9be100e: check.sh exit 0, 29 passed / 2 skipped on 9225ac6; fresh review: no MAJOR, 5 MINOR
  (COMPLIANCE label wording, site.config.json and README "checked" wording, a test for --check with the retired copy,
  the threat-model spoofing control): fixed in PR 2.
- PR 2 splits (wc -l before -> after): PiiDetectors 579 -> 215 + PiiPatterns 394; Report 565 -> 378 + ReportBook 212;
  PdfFonts 549 -> 319 + PdfShown 252; PdfText 503 -> 193 + PdfContent 331; Extraction.Tests 594 -> 248 +
  ReaderRounds.Tests 360. Largest hand-written file after: src/TextExtract.psm1 419. Each split keeps every $script:
  variable with the functions that use it (module mode = separate scopes; the bundle = one scope; build_paste.py
  check_names keeps names unique). Report rank in ReportBook: Get-ConfidenceRank (same map). New modules are in the 4
  load lists (build_paste.py, Run-Tests.ps1, Find-ChannelPII.ps1, ReaderPool.psm1) and the bundle test's unload list.
- PR 2 tests: Pester 535 -> 543 (+8 in Protection.Tests: no raw value in a hit, reason, file name, screen or workbook,
  plus a control that the check finds every raw value); build_paste_test 9 -> 10.
