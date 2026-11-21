#!/usr/bin/env bash
# This project's own tests, run last by scripts/check.sh (the pre-push hook and the "checks" pull-request Action).
# What it runs: the paste-bundle check (scripts/build_paste.py --check), the PowerShell tests (tests/Run-Tests.ps1: invented data, fake PnP cmdlets, no network) and the lint
# (tests/Lint.ps1: parser and Windows PowerShell 5.1 syntax with PSScriptAnalyzer).
# Usage:   scripts/check.local.sh [--fast]   (--fast: the same; both are quick)
# Needs:   pwsh 7; PSScriptAnalyzer 1.23.0 for the lint (in GitHub Actions it is installed here, as the kit's workflow
#          file is locked; elsewhere a missing module is a FAIL, not a skip, so the lint is never left out silently).
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
# The paste bundle is a generated file: it must be the fresh build of src/ (scripts/build_paste.py).
python3 scripts/build_paste.py --check || rc=1
# Each runs once; its output is shown (without the posts sign-in notice) and its exit code kept.
out=$(pwsh -NoProfile -File tests/Run-Tests.ps1 2>&1) || rc=1
printf '%s\n' "$out" | grep -v '^Posts: '
out=$(pwsh -NoProfile -File tests/Lint.ps1 2>&1) || rc=1
printf '%s\n' "$out" | grep -v '^$'
exit "$rc"
