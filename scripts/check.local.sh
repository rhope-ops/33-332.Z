#!/usr/bin/env bash
# This project's own tests, run last by scripts/check.sh (the pre-push hook and the "checks" pull-request Action).
# What it runs: the rule-source pins (ref/SHA256SUMS), the paste-bundle check (scripts/build_paste.py --check) and its
# self-test (scripts/build_paste_test.py), the PowerShell tests (tests/Run-Tests.ps1: invented data, fake PnP cmdlets, no
# network) and the lint (tests/Lint.ps1: parser and Windows PowerShell 5.1 syntax with PSScriptAnalyzer).
# Usage:   scripts/check.local.sh [--fast]   (--fast: the same; both are quick)
# Needs:   pwsh 7; PSScriptAnalyzer 1.23.0 for the lint (in GitHub Actions it is installed here, as the kit's workflow
#          file is locked; elsewhere a missing module is a FAIL, not a skip, so the lint is never left out silently).
#          The installed module's code files must match the SHA-256 pins below (fleet v2, 2026-11-23); a mismatch is a FAIL.
# Network: in GitHub Actions only, the PowerShell Gallery once, to install PSScriptAnalyzer. Else none.
# Writes:  none in the repo (the tests write report files to a temporary folder and delete them).
# Credentials: none.
set -u
cd "$(dirname "$0")/.."
if [ ! -f tests/Run-Tests.ps1 ]; then echo "SKIP  PowerShell tests: no tests/Run-Tests.ps1 yet"; exit 0; fi
if [ -n "${GITHUB_ACTIONS:-}" ]; then
  pwsh -NoProfile -Command "Install-Module PSScriptAnalyzer -RequiredVersion 1.23.0 -Scope CurrentUser -Force" || exit 1
fi
# A dev machine may keep PSScriptAnalyzer outside the default module path (as 90-161.Z's scripts/lint.sh does).
[ -d /opt/psmods ] && export PSModulePath="/opt/psmods:${PSModulePath:-}"
rc=0
# PSScriptAnalyzer 1.23.0: SHA-256 of its code files (pinned 2026-11-23 from a 1.23.0 copy; the GitHub runner installs
# it from the PowerShell Gallery, so a changed package fails here and prints the hash it found).
pwsh -NoProfile -Command '
  $pins = @{
    "PSScriptAnalyzer.psd1" = "72980ffab78dae5ee1f785232415e1ae9596b3121f7b10fe10bdfafd483d7f04"
    "PSScriptAnalyzer.psm1" = "af24e5ec4e3e570d8a5b83c51020d6f02f84f98417e6b9709f8619ff62e872f1"
    "Microsoft.Windows.PowerShell.ScriptAnalyzer.dll" = "98ac2fca382dc2a22c1c27baceb3a3ec1c03193b590600607df0590607add987"
    "Microsoft.Windows.PowerShell.ScriptAnalyzer.BuiltinRules.dll" = "07f0d3cdb7f8f53e98bbe54b63df3b84f52e4a48ed2571267485cf95a8fbeba1"
    "Microsoft.PowerShell.CrossCompatibility.dll" = "76577a2ff29055d5797a25328c6dbc3734ba9f019fe27333b01e5cdab3f3acb7"
  }
  $module = Get-Module -ListAvailable PSScriptAnalyzer | Where-Object { $_.Version -eq [version]"1.23.0" } | Select-Object -First 1
  if (-not $module) { Write-Host "FAIL  PSScriptAnalyzer 1.23.0 is not installed"; exit 1 }
  $bad = 0
  foreach ($name in $pins.Keys) {
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $module.ModuleBase $name)).Hash.ToLower()
    if ($hash -ne $pins[$name]) { Write-Host "FAIL  PSScriptAnalyzer $name sha256 $hash is not the pinned value"; $bad++ }
  }
  if ($bad) { exit 1 }
  Write-Host "PASS  PSScriptAnalyzer 1.23.0: 5 files match their SHA-256 pins"
' || rc=1
# The rule source copies (ref/) must match their pins.
(cd ref && sha256sum --check --quiet SHA256SUMS) && echo "PASS  ref/SHA256SUMS" || { echo "FAIL  ref/SHA256SUMS"; rc=1; }
# The paste bundle is a generated file: it must be the fresh build of src/ (scripts/build_paste.py).
python3 scripts/build_paste.py --check || rc=1
python3 scripts/build_paste_test.py 2>&1 | tail -1; [ "${PIPESTATUS[0]}" -eq 0 ] || rc=1
# Each runs once; its output is shown (without the posts sign-in notice) and its exit code kept.
out=$(pwsh -NoProfile -File tests/Run-Tests.ps1 2>&1) || rc=1
printf '%s\n' "$out" | grep -v '^Posts: '
out=$(pwsh -NoProfile -File tests/Lint.ps1 2>&1) || rc=1
printf '%s\n' "$out" | grep -v '^$'
exit "$rc"
