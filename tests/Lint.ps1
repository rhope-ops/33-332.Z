<#
Lint.ps1 - the parser check and PSScriptAnalyzer on every .ps1 and .psm1 file in the repo.

  Purpose:     Proves each file parses and uses only Windows PowerShell 5.1 syntax (the gov-device version).
  Inputs:      the .ps1/.psm1 files; PSScriptAnalyzerSettings.psd1 (rules, and why some are off).
  Outputs:     one line per finding, then 'LINT files=<n> findings=<n>'; exit code 0 when findings=0.
  How to run:  pwsh -NoProfile -File tests/Lint.ps1
  Run target:  PowerShell 7 with the PSScriptAnalyzer module (dev machine or CI only; not needed on the gov device).
  Network:     none.
  Writes:      none.
  Credentials: none.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module PSScriptAnalyzer -RequiredVersion 1.23.0   # the version scripts/check.local.sh pins
$settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
$files = Get-ChildItem -Path $root -Recurse -Include *.ps1, *.psm1
$all = foreach ($f in $files) {
    $tokens = $null; $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs)
    foreach ($x in $errs) { [pscustomobject]@{ RuleName = 'ParseError'; ScriptName = $f.Name; Line = $x.Extent.StartLineNumber; Message = $x.Message } }
    Invoke-ScriptAnalyzer -Path $f.FullName -Settings $settings
}
$all | Format-Table RuleName, ScriptName, Line, Message -AutoSize -Wrap | Out-String -Width 200
'LINT files=' + @($files).Count + ' findings=' + @($all).Count
if (@($all).Count -gt 0) { exit 1 }
