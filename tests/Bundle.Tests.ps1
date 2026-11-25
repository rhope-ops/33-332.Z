<#
Bundle.Tests.ps1 - runs the paste bundle (paste/Ion-script.txt) from start to end, as ISE would, against
the fake site.

  Purpose:     The text a person pastes works by itself: no module files, no script file of its own (the modules of
               the earlier test files are unloaded first). The Settings block is filled in as a person would edit it
               in ISE (plain value replacement); the text then runs dot-sourced, as ISE runs an untitled pane, and must
               leave the session's error setting and strict mode as they were.
  Inputs:      paste/Ion-script.txt; $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1, after EndToEnd.Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      a copy of the filled-in bundle and its report files in a new temporary folder, deleted at the end.
  Credentials: none.
#>

$bundlePath = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'paste') 'Ion-script.txt'
$bundleFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-bundle-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $bundleFolder | Out-Null
$library = '/sites/TeamA/Shared Documents'

# Fill in the Settings as a person would in ISE: the site and the report folder.
$text = Get-Content -LiteralPath $bundlePath -Raw
$filled = $text.Replace("`$SiteUrl = ''", "`$SiteUrl = '$($Sample.Site)'")
$filled = $filled.Replace("`$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Ion')",
    "`$OutputFolder = '$bundleFolder'")
Assert-Equal $true ($filled -ne $text) 'Settings filled in'
# No module file and no script file: the only Import-Module is the reader workers' line for the module files of a
# development run (src/), which the bundle never takes (its $IonWorkerInit holds the reader block's text).
$imports = @([regex]::Matches($filled, 'Import-Module[^\r\n]*|\$PSScriptRoot') | ForEach-Object { $_.Value })
Assert-Equal 'Import-Module -Global }|Import-Module $Init.Paths -Global -Force }' ($imports -join '|') 'Bundle loads no module or file'
Assert-Equal 1 ([regex]::Matches($filled, '(?m)^\$IonWorkerInit = @\{ Code = ').Count) 'Bundle workers load the reader block, not files'
# No compiled code (NIPRNet: no unapproved .exe, .dll or .jar): nothing compiled (Add-Type with code), no binary
# loaded from a file or bytes, nothing started; only assemblies of .NET Framework itself, by name.
$compiled = 'Add-Type\s+(?!-AssemblyName\b)|-TypeDefinition|-MemberDefinition|Assembly\]::Load|LoadFrom|LoadFile|Start-Process|\.(?:exe|dll|jar)\b'
Assert-Equal 0 ([regex]::Matches($filled, $compiled, 'IgnoreCase').Count) 'Bundle compiles and loads no binary code'

# Unload the modules the earlier test files imported, so the bundle can only use its own copies of the functions.
Get-Module PiiPatterns, PiiSignedIds, PiiDetectors, PiiMasking, TextExtract, XlsxWriter, Report, ReportBook, ScanMeter, ChannelScan | Remove-Module -Force
Assert-Equal $false ([bool](Get-Command Find-PiiInText -ErrorAction SilentlyContinue)) 'Modules unloaded before the bundle runs'

$runCopy = Join-Path $bundleFolder 'pasted.ps1'
Set-Content -LiteralPath $runCopy -Value $filled -Encoding UTF8
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
# ISE runs an untitled script pane in the session itself, as dot-sourcing does: run it that way, from a session with
# the default error setting and no strict mode, and check that the bundle puts both back.
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
. $runCopy 6>$null
$sessionErrorAction = $ErrorActionPreference
$strictStillOn = $false
try { $null = $notDefinedAnywhere } catch { $strictStillOn = $true }  # strict mode makes an unset variable an error
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'Continue' $sessionErrorAction 'Bundle put the session error setting back'
Assert-Equal $false $strictStillOn 'Bundle turned strict mode off at the end'
$results = @(Get-ResultFiles $bundleFolder)
Assert-Equal 1 $results.Count 'Bundle saved one workbook'
Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $results[0].FullName) 'Bundle found every planted type'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Bundle signed out'
Assert-Equal 'verified' $IonPasteCheck 'Paste check: the script file read back matches its hash'

# The clean-up also runs when the scan fails part way: the session's error setting comes back after the error.
Reset-FakePnP @{}
Set-FakePnPOption ListFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
. $runCopy 6>$null
$errorPathAction = $ErrorActionPreference
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'Continue' $errorPathAction 'Bundle put the session error setting back after a failed scan'
Assert-Equal $true (@(Get-ResultFiles $bundleFolder).Count -ge 2) 'Failed scan still saved its results'

# A sign-in that fails: the person sees why (no exception), is signed out, is offered another site (Enter exits),
# and the clean-up still puts strict mode and the session's error setting back.
Reset-FakePnP @{}
Set-FakePnPOption ConnectFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
$outerError = ''
$shown = ''
try { $shown = (. $runCopy 6>&1 | ForEach-Object { [string]$_ }) -join ' ' } catch { $outerError = $_.Exception.Message }
$outerAction = $ErrorActionPreference
$outerStrict = $false
try { $null = $notDefinedAnywhereEither } catch { $outerStrict = $true }  # strict mode makes an unset variable an error
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal '' $outerError 'A failed sign-in does not end the script with an error'
Assert-Equal $true ($shown -match 'sign-in refused') 'The failed sign-in is shown to the person'
Assert-Equal $true ($shown -match 'Build [0-9a-f]{8}\b') 'The start screen shows the build ID (the PAM, 2026-11-24)'
Assert-Equal 1 (Get-FakeCallCount 'Read-Host') 'Then another site is offered (Enter exits)'
Assert-Equal 'Continue' $outerAction 'Outer clean-up put the session error setting back'
Assert-Equal $false $outerStrict 'Outer clean-up turned strict mode off'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Outer error path signed out'
Assert-Equal $true ($shown -match 'Paste check: OK\.') 'Paste check: the person sees OK'

# ---- Paste check (fleet v2 plan P4c): a changed text stops before anything runs ----
# One character changed below the check's mark: the run stops with the STOPPED text, signs nothing in, saves
# nothing, and the clean-up still puts the session back.
$changedCopy = Join-Path $bundleFolder 'changed.ps1'
Set-Content -LiteralPath $changedCopy -Value $filled.Replace('# ==== Reader modules ====', '# ==== Reader module ====') -Encoding UTF8
Reset-FakePnP @{ "$library/General/notes.txt" = $Txt }
$before = @(Get-ResultFiles $bundleFolder).Count
Remove-Variable -Name IonWorkerInit   # set by the runs above, below the check's mark
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
$stopError = ''
try { . $changedCopy 6>$null } catch { $stopError = $_.Exception.Message }
$stopAction = $ErrorActionPreference
$stopStrict = $false
try { $null = $notDefinedAtTheStop } catch { $stopStrict = $true }
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal $true ($stopError -match '^Paste check: STOPPED\. Nothing ran\.') 'Paste check: a changed text stops with the STOPPED text'
Assert-Equal 'changed' $IonPasteCheck 'Paste check: the stop is for a changed text'
Assert-Equal $false (Test-FakePnPCall 'Connect-PnPOnline') 'Paste check: no sign-in after the stop'
Assert-Equal $before @(Get-ResultFiles $bundleFolder).Count 'Paste check: no workbook after the stop'
Assert-Equal $false (Test-Path -LiteralPath variable:IonWorkerInit) 'Paste check: no code below the check ran (no reader block)'
Assert-Equal 'Continue' $stopAction 'Paste check: the session error setting is back after the stop'
Assert-Equal $false $stopStrict 'Paste check: strict mode is off after the stop'

# The ISE script pane (no script file): the text comes from $psISE, with CRLF line ends and spaces at line ends as
# a browser copy can give; the check reads it and passes. A sign-in that fails keeps the run short.
$psISE = [pscustomobject]@{ CurrentFile = [pscustomobject]@{ Editor = [pscustomobject]@{ Text = ($filled.Replace("`n", "  `r`n")) } } }
Reset-FakePnP @{}
Set-FakePnPOption ConnectFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
try { . ([scriptblock]::Create($filled)) 6>$null } catch { $null = $_ }
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'verified' $IonPasteCheck 'Paste check: the ISE pane text (CRLF, spaces at line ends) matches its hash'
# The same pane with one character changed stops.
$psISE.CurrentFile.Editor.Text = $filled.Replace('# ==== Reader modules ====', '# ==== Reader module ====')
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
$stopError = ''
try { . ([scriptblock]::Create($filled)) 6>$null } catch { $stopError = $_.Exception.Message }
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'changed' $IonPasteCheck 'Paste check: a changed ISE pane text stops'
# No text to read back (no script file, no ISE): one line, and the scan goes on (a false stop must not block a paste).
Remove-Variable -Name psISE
Reset-FakePnP @{}
Set-FakePnPOption ConnectFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
$shown = ''
try { $shown = (. ([scriptblock]::Create($filled)) 6>&1 | ForEach-Object { [string]$_ }) -join ' ' } catch { $null = $_ }
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'not read' $IonPasteCheck 'Paste check: a text that cannot be read back is not checked'
Assert-Equal $true ($shown -match 'Paste check: not done\. .*sign-in refused') 'Paste check: not done is shown, and the run goes on'

# ---- A cut paste runs nothing (the try opens above the Settings) ----
# Each cut at 10/25/50/75/90/99 % does not parse; a cut above the check's mark that parses holds no statement.
$cutLines = $filled.Replace("`r`n", "`n").Split("`n")
$markLine = [array]::IndexOf($cutLines, ($cutLines | Where-Object { $_ -like '# ==== Paste check ====*' } | Select-Object -First 1))
$tokens = $null
$parseErrors = $null
$parsing = 0
foreach ($pct in 10, 25, 50, 75, 90, 99) {
    $cut = [int]($cutLines.Count * $pct / 100)
    $null = [System.Management.Automation.Language.Parser]::ParseInput(($cutLines[0..($cut - 1)] -join "`n"), [ref]$tokens, [ref]$parseErrors)
    if (-not $parseErrors) { $parsing++ }
}
Assert-Equal 0 $parsing 'Cut paste: no cut at 10/25/50/75/90/99 % parses'
$running = @()
for ($cut = 1; $cut -le $markLine; $cut++) {
    $ast = [System.Management.Automation.Language.Parser]::ParseInput(($cutLines[0..($cut - 1)] -join "`n"), [ref]$tokens, [ref]$parseErrors)
    if (-not $parseErrors -and $ast.EndBlock -and $ast.EndBlock.Statements.Count) { $running += $cut }
}
Assert-Equal '' ($running -join ',') 'Cut paste: no cut above the check runs a statement'

Remove-Item -LiteralPath $bundleFolder -Recurse -Force
