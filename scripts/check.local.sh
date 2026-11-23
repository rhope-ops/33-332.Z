#!/usr/bin/env bash
# This project's own tests, run last by scripts/check.sh (the pre-push hook and the "checks" pull-request Action).
# What it runs: the rule-source pins (ref/SHA256SUMS), the paste-bundle check (scripts/build_paste.py --check) and its
# self-test (scripts/build_paste_test.py), the PowerShell tests (tests/Run-Tests.ps1: invented data, fake PnP cmdlets, no
# network) and the lint (tests/Lint.ps1: parser and Windows PowerShell 5.1 syntax with PSScriptAnalyzer).
# Usage:   scripts/check.local.sh [--fast]   (--fast: the same; both are quick)
# Needs:   pwsh 7; PSScriptAnalyzer 1.23.0 for the lint (in GitHub Actions it is installed here, as the kit's workflow
#          file is locked; elsewhere a missing module is a FAIL, not a skip, so the lint is never left out silently).
#          The installed module's code files must match scripts/pssa-1.23.0.sha256 (fleet v2, 2026-11-23); a mismatch is a FAIL.
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
# PSScriptAnalyzer 1.23.0: the SHA-256 of every code file (.dll, .psd1, .psm1, .ps1xml, in every folder: pwsh 7 loads
# PSv7/, Windows PowerShell 5.1 the root) must match scripts/pssa-1.23.0.sha256, and no other code file may be there.
# Pinned 2026-11-23 from a 1.23.0 copy; the GitHub runner installs from the PowerShell Gallery, so a changed package
# fails here and prints the hash it found.
pwsh -NoProfile -Command '
  $pins = @{}
  foreach ($line in Get-Content -LiteralPath scripts/pssa-1.23.0.sha256) {
    if ($line -match "^([0-9a-f]{64})  (.+)$") { $pins[$Matches[2]] = $Matches[1] }
  }
  $module = Get-Module -ListAvailable PSScriptAnalyzer | Where-Object { $_.Version -eq [version]"1.23.0" } | Select-Object -First 1
  if (-not $module) { Write-Host "FAIL  PSScriptAnalyzer 1.23.0 is not installed"; exit 1 }
  $base = $module.ModuleBase
  $files = @(Get-ChildItem -LiteralPath $base -Recurse -File | Where-Object { $_.Extension -in ".dll", ".psd1", ".psm1", ".ps1xml" })
  $bad = 0
  foreach ($file in $files) {
    $name = $file.FullName.Substring($base.Length).TrimStart("/", "\").Replace("\", "/")
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $file.FullName).Hash.ToLower()
    if (-not $pins.ContainsKey($name)) { Write-Host "FAIL  PSScriptAnalyzer $name is not pinned (sha256 $hash)"; $bad++ }
    elseif ($hash -ne $pins[$name]) { Write-Host "FAIL  PSScriptAnalyzer $name sha256 $hash is not the pinned value"; $bad++ }
  }
  if ($files.Count -ne $pins.Count) { Write-Host ("FAIL  PSScriptAnalyzer: {0} code files, {1} pins" -f $files.Count, $pins.Count); $bad++ }
  if ($bad -or $pins.Count -eq 0) { exit 1 }
  Write-Host ("PASS  PSScriptAnalyzer 1.23.0: {0} code files match their SHA-256 pins" -f $files.Count)
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
