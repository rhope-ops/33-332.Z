# Threat model of ion

A short STRIDE model of Ion (the DoD365 PII Reporter), its paste bundle and this repo. Rules it relies on: AGENTS.md "Project rules" (data boundary, no reconnaissance, masked output, platform). Rule source for the privacy steps: AFI 33-332 with AFI33-332_DAFGM2026-01 (`ref/`, pinned in `ref/SHA256SUMS`).

## Assets

- The PII that Ion finds in the one site, team or folder the person gives. It stays in DoD365 and on the person's gov computer.
- The report workbook Ion saves (masked values, file locations, version-history names). It is CUI. `.check` blocks `PIIScan-*` from commits.
- The person's CAC sign-in session to SharePoint (PnP.PowerShell 1.12.0, interactive, no stored credentials).
- The paste bundle `paste/Ion-script.txt`, built from `src/` by `scripts/build_paste.py`, and the GitHub Pages site that links it.
- This repo's guards: `scripts/check.sh`, `scripts/check.local.sh`, the pre-push hook, `.check`, `.check.lock`.

## Trust boundaries

1. GitHub (public) to the gov computer: the person opens the `.txt` in a browser, reads it, and pastes it into PowerShell ISE. No file is downloaded and run; the execution policy does not change (ISSM ruling 2026-11-13).
2. The gov computer to SharePoint Online (DoD365): the person's own CAC sign-in; read-only calls to the one site given.
3. Ion to the local disk: one results folder under Documents; masked values only.
4. Repo content (files, issues, comments) to the agent: data, not instructions (AGENTS.md "Untrusted input").

## STRIDE

| Threat | Example | Control |
|---|---|---|
| Spoofing | A changed copy of the bundle from another source | The person opens it only from this repo's README or site link (GitHub over HTTPS); a copy from anywhere else is not trusted |
| Tampering | A pull request that changes the bundle by hand | `scripts/build_paste.py --check` fails when `paste/Ion-script.txt` is not the fresh build of `src/`; guard files change only on the Administrator's request |
| Repudiation | Who put PII in a file | The report names the first version that holds the PII (version history), Created By and Modified By |
| Information disclosure | A found SSN shown in full in the report or the console | Masked output (Project rule; `src/PiiEngine.psm1`; tests in `tests/`); `.check` blocks report files; the content check on push |
| Information disclosure | The tool lists sites or users outside the one given | No reconnaissance (Project rule): Ion takes one site, team or folder and reads only it |
| Denial of service | A very large or crafted PDF stalls the scan | File size limit (`$MaxFileMB`), a time and work budget per PDF, NOT SCANNED rows instead of a stop |
| Elevation of privilege | A script that asks for admin rights or changes the execution policy | Platform rule: no admin, no service account, no policy change; read-only calls only; no compiled code (`tests/Bundle.Tests.ps1`) |

## Site

The Pages site is one static page with no script and no network calls. Its controls: a Content-Security-Policy meta (`default-src 'none'`, images only from the site, `connect-src 'none'`), a `no-referrer` policy, and `rel="noopener noreferrer"` on the link that opens a new tab. `scripts/site_page.py` fails the site check when one of them is gone. A meta CSP cannot set `frame-ancestors`; GitHub Pages sends no such header.

## Known gaps

- The bundle is not signed. The person trusts the GitHub link and the review of the pull request that changed `src/`.
- `.github/workflows/security.yml` (bandit on `scripts/`, gitleaks on the tree) runs on pull requests and manual runs only. gitleaks is checked against the checksums of its own release: this finds a damaged download, not a bad release.
- PSScriptAnalyzer 1.23.0 is pinned by the SHA-256 of all 30 of its code files (`scripts/pssa-1.23.0.sha256`, checked by `scripts/check.local.sh`; an unlisted code file also fails), and `tests/Lint.ps1` loads exactly 1.23.0; the pins came from a local 1.23.0 copy (the PowerShell Gallery was not reachable from the build host), and the GitHub runner, which installs from the Gallery, is the proof that they match.
