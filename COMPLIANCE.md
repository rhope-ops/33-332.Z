# Compliance map

What this repo does for each rule it follows, and where the proof is. The rule source for privacy steps is AFI 33-332 with AFI33-332_DAFGM2026-01 (AGENTS.md "Source authority"; copies in `ref/`, pinned by `ref/SHA256SUMS`). The memorandum wins where the two differ.

## Privacy steps (AFI 33-332)

| Rule | Paragraph | What Ion does | Proof |
|---|---|---|---|
| Remove personal information on SharePoint when it is no longer needed | 7.1.1 | Finds PII in the files of one SharePoint site or Teams team, so the Administrator or site member can remove or move it | `tests/ChannelScan.Tests.ps1`, `tests/EndToEnd.Tests.ps1` |
| Safeguard personal information on shared drives (encrypted or password protected, need-to-know) | 7.1.2 | Scans one drive folder; a file with a password, or with a sensitivity label that encrypts it, is listed as NOT SCANNED with the reason (a label without encryption is read like any file) | Drive folder scan: `tests/DriveScan.Tests.ps1`; protected files: `tests/Protection.Tests.ps1` |
| Report a PII breach without delay | 3.1.2, as changed by AFI33-332_DAFGM2026-01 | The README, the site and the report give the breach advice; the text in `src/ReportExport.psm1` must match the README | `scripts/build_site_local.py` (tool-text check) |
| Protect SSNs | 6.4 | SSNs are found and shown masked | `tests/Detectors.Tests.ps1` |

## Project rules (AGENTS.md)

| Rule | Control | Proof |
|---|---|---|
| Data boundary | Invented test data only; PII-shaped samples are built from parts at run time; no tool writes data out of DoD365 | `scripts/check.sh --content` on each push |
| No reconnaissance | One site, team or folder per run; no listing of other sites, teams or users | Code review; `tests/FakePnP.ps1` fakes only the cmdlets Ion uses |
| Masked output | Every found value is masked where it is shown or saved | `tests/Detectors.Tests.ps1`, `tests/Report.Tests.ps1`, `tests/Protection.Tests.ps1` (no raw value in a hit, a reason, a name, the screen or the workbook) |
| Platform | Windows PowerShell 5.1, PnP.PowerShell 1.12.0, CAC sign-in, read-only, one pasted text, no execution-policy change | `tests/Lint.ps1` (5.1 syntax), the `ps51` workflow, `tests/Bundle.Tests.ps1` |
| Public repo | No unit names, real site paths, DoD IDs or network details | The content check (`.check`, profile dod) |

## Repo hygiene

| Item | File |
|---|---|
| License | `LICENSE` (MIT, as the kit) |
| Security policy | `SECURITY.md` |
| Threat model | `docs/threat-model.md` |
| Dependency updates (GitHub Actions) | `.github/dependabot.yml` |
| Line ends and text files | `.gitattributes`, `.editorconfig` |
| Rule source copies | `ref/SHA256SUMS` (checked by `scripts/check.local.sh`) |
| Static security scans | `.github/workflows/security.yml` (bandit, gitleaks) |
| Lint tool pin | `scripts/pssa-1.23.0.sha256` (PSScriptAnalyzer 1.23.0: SHA-256 of all 30 code files; `scripts/check.local.sh` checks them) |
