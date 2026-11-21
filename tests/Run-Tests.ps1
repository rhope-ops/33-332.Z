<#
Run-Tests.ps1 - runs every test of the Teams channel PII scan. No Pester (gov devices allow no extra modules).

  Purpose:     Imports the modules, loads the test kit and the fake PnP cmdlets, runs each test file in
               order, and prints one result line.
  Inputs:      the files in src/ and tests/.
  Outputs:     a FAIL line per failed test, then 'TESTS passed=<n> failed=<n>'; exit code 0 when all pass.
  How to run:  pwsh -NoProfile -File tests/Run-Tests.ps1
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none (PnP cmdlets are replaced by fakes).
  Writes:      report files in temporary folders, deleted by the tests that make them.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$sourceFolder = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
$readerModules = @('PiiDetectors.psm1', 'PiiMasking.psm1', 'OcrText.psm1', 'OleText.psm1', 'ContainerText.psm1',
    'PdfCore.psm1', 'PdfObjects.psm1', 'PdfFonts.psm1', 'PdfLayout.psm1', 'PdfText.psm1', 'TextExtract.psm1', 'ReaderPool.psm1')
foreach ($module in $readerModules + @('XlsxWriter.psm1', 'Report.psm1', 'ScanMeter.psm1', 'ScanQueue.psm1', 'Accountability.psm1', 'ChannelScan.psm1', 'DriveScan.psm1')) {
    Import-Module (Join-Path $sourceFolder $module) -Force
}

. (Join-Path $PSScriptRoot 'TestKit.ps1')
. (Join-Path $PSScriptRoot 'FakePnP.ps1')

# Order matters: Extraction builds $Docx and $Txt, which the later files use.
$testFiles = @('Detectors.Tests.ps1', 'Extraction.Tests.ps1', 'Protection.Tests.ps1', 'ChannelScan.Tests.ps1', 'DriveScan.Tests.ps1',
    'Report.Tests.ps1', 'EndToEnd.Tests.ps1', 'Bundle.Tests.ps1')
foreach ($testFile in $testFiles) {
    . (Join-Path $PSScriptRoot $testFile)
}

foreach ($failure in $script:Fail) { Write-Host "FAIL $failure" -ForegroundColor Red }
Write-Host ('TESTS passed={0} failed={1}' -f $script:Pass, $script:Fail.Count)
if ($script:Fail.Count -gt 0) { exit 1 }
