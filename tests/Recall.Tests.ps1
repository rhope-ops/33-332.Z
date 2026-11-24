<#
Recall.Tests.ps1 - tests the recall test (docs/reviews/2026-11-24-recall-test.txt) and measures Ion's local baseline.

  Purpose:     Loads the functions of the recall test from its paste text (one home: the gov device runs the same code),
               checks that its key is the same on every run and in range, that Make writes 12 files that the readers
               read with every planted value in their text, and that Score counts a known workbook right. Then it scans
               the made folder with Invoke-DriveScan (the reader block of paste/Ion-script.txt, as on the gov device),
               scores the workbook and prints one RECALL line.
  Inputs:      docs/reviews/2026-11-24-recall-test.txt; paste/Ion-script.txt (its reader block).
  Outputs:     passes and failures counted by Assert-Equal; one 'RECALL ...' line (counts only).
  How to run:  dot-sourced by tests/Run-Tests.ps1, last.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      the invented test files and the report in new temporary folders, deleted at the end.
  Credentials: none.
#>

$recallRoot = Split-Path $PSScriptRoot -Parent
$recallText = [System.IO.File]::ReadAllText((Join-Path (Join-Path $recallRoot 'docs') 'reviews/2026-11-24-recall-test.txt'))
$recallCut = $recallText.IndexOf('# ==== Questions')
Assert-Equal $true ($recallCut -gt 0) 'Recall: the paste text has its questions marker'
. ([scriptblock]::Create($recallText.Substring(0, $recallCut)))   # the functions only; the questions are not run

# The key: the same on every run (a number from every value, so Windows PowerShell 5.1 in the ps51 workflow proves it too).
function Get-RecallFingerprint {
    <#
    .SYNOPSIS
      A number from every file's lines, cells and planted values (a failure shows it, never a value).
    .PARAMETER Plan
      Output of Get-RecallPlan.
    .OUTPUTS
      [long] from 0 to 1000002.
    #>
    param([System.Collections.Generic.List[object]]$Plan)
    $print = 0L
    foreach ($file in $Plan) {
        $parts = @($file.Name, $file.Kind) + @($file.Lines) + @($file.Rows | ForEach-Object { @($_) -join '|' }) +
            @($file.Plants | ForEach-Object { $_.Type + '|' + $_.Form + '|' + $_.Key + '|' + $_.Raw })
        foreach ($ch in ($parts -join "`n").ToCharArray()) { $print = ($print * 31L + [int]$ch) % 1000003L }
    }
    return $print
}
$recallPlan = Get-RecallPlan
$recallPrint = Get-RecallFingerprint $recallPlan
Assert-Equal $recallPrint (Get-RecallFingerprint (Get-RecallPlan)) 'Recall: the key is the same on a second build'
Assert-Equal 367101 $recallPrint 'Recall: the key fingerprint is the frozen one (2026-11-24)'
$recallCounts = @($script:RecallTypes | ForEach-Object { $type = $_; '{0}={1}' -f $type, @($recallPlan | ForEach-Object { $_.Plants } | Where-Object { $_.Type -eq $type }).Count })
Assert-Equal 'SSN=9 DoD ID=10 Date of birth=7 Payment card=4 Passport=3 Phone=8 Personal e-mail=7 Home address=5' ($recallCounts -join ' ') 'Recall: 53 values over 8 types'
Assert-Equal 12 $recallPlan.Count 'Recall: 12 test files'
$recallAll = @($recallPlan | ForEach-Object { $_.Plants })
$badSsn = @($recallAll | Where-Object { $_.Type -eq 'SSN' } | Where-Object { $d = $_.Raw -replace '\D', ''; -not (Test-SsnValid $d.Substring(0, 3) $d.Substring(3, 2) $d.Substring(5)) })
Assert-Equal 0 $badSsn.Count 'Recall: every planted SSN is valid by the SSA rules'
Assert-Equal 0 @($recallAll | Where-Object { $_.Type -eq 'Payment card' -and -not (Test-Luhn ($_.Raw -replace '\D', '')) }).Count 'Recall: every planted card passes Luhn'

# Make: an empty folder only; every planted value is in the text the readers get.
$recallLetters = -join @(1..12 | ForEach-Object { [char](97 + (Get-Random -Maximum 26)) })
$recallDir = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-recall-' + $recallLetters)
$recallOut = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-recallout-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $recallDir -Force | Out-Null
try {
    Assert-Equal $true ((Invoke-RecallMake (Join-Path $recallDir 'missing')) -like '*does not exist*') 'Recall Make: a missing folder is refused'
    Assert-Equal $true ((Invoke-RecallMake $recallDir) -like 'Done: 12 test files with 53 planted values*') 'Recall Make: 12 files written'
    Assert-Equal $true ((Invoke-RecallMake $recallDir) -like '*not empty*') 'Recall Make: a folder with files is refused'
    $notRead = @(); $notInText = 0
    foreach ($file in $recallPlan) {
        $read = Get-TextFromBytes ([System.IO.File]::ReadAllBytes((Join-Path $recallDir $file.Name))) $file.Name
        if ($read.Status -ne 'OK') { $notRead += $file.Name; continue }
        $notInText += @($file.Plants | Where-Object { -not $read.Text.Contains($_.Raw) }).Count
    }
    Assert-Equal '' ($notRead -join ' ') 'Recall Make: the readers read every file'
    Assert-Equal 0 $notInText 'Recall Make: every planted value is in the text the readers get'

    # Score on a known workbook: each planted value once (masked as Ion masks it), a decoy hit and a hit in another file.
    $perfect = @{ Hits = (New-Object System.Collections.Generic.List[object]); NotScanned = @('decoys.txt', 'other.txt') }
    foreach ($file in $recallPlan) { foreach ($p in $file.Plants) { $perfect.Hits.Add(@{ Type = $p.Type; File = $file.Name; Masked = (Get-MaskedValue $p.Type $p.Raw) }) } }
    $perfect.Hits.Add(@{ Type = 'SSN'; File = 'decoys.txt'; Masked = '***-**-0000' })
    $perfect.Hits.Add(@{ Type = 'Phone'; File = 'other.txt'; Masked = '***-***-0000' })
    $known = Get-RecallScore $recallPlan $perfect
    $knownText = Format-RecallScore $known
    Assert-Equal 'All types 53 53 0 2 100%' (@($knownText | Where-Object { $_ -like 'All types*' })[0] -replace '\s+', ' ') 'Recall Score: every planted value found, 2 false hits'
    Assert-Equal '1 1' ('{0} {1}' -f $known.NotScanned, $known.OutsideHits) 'Recall Score: test files not scanned and hits outside the test are counted'
    Assert-Equal 0 @($knownText | Where-Object { $_ -match '\d{5,}|\S@' }).Count 'Recall Score: the lines hold no value (no long number, no e-mail name)'

    # A workbook that Excel saved keeps its text as shared strings: Score reads it too.
    $head = '<?xml version="1.0" encoding="UTF-8"?>'
    $rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $excelParts = [ordered]@{
        'xl/workbook.xml' = $head + '<workbook xmlns:r="' + $rel + '"><sheets><sheet name="Hits" sheetId="1" r:id="rId7"/><sheet name="Not scanned" sheetId="2" r:id="rId8"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = $head + '<Relationships><Relationship Id="rId7" Type="' + $rel + '/worksheet" Target="/xl/worksheets/sheet9.xml"/>' +
            '<Relationship Id="rId8" Type="' + $rel + '/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
        'xl/sharedStrings.xml' = $head + '<sst><si><t>Type</t></si><si><t>File</t></si><si><t>Masked value</t></si><si><r><t>Pass</t></r><r><t>port</t></r></si>' +
            '<si><t>memo.docx</t></si><si><t>***-**-0000</t></si></sst>'
        'xl/worksheets/sheet9.xml' = $head + '<worksheet><sheetData><row r="1"><c r="B1" t="s"><v>0</v></c><c r="E1" t="s"><v>1</v></c><c r="F1" t="s"><v>2</v></c></row>' +
            '<row r="2"><c r="B2" t="s"><v>3</v></c><c r="E2" t="s"><v>4</v></c><c r="F2" t="s"><v>5</v></c></row></sheetData></worksheet>'
        # Invented path: a name-shaped profile folder with an apostrophe and spaces (the reason must print none of it).
        'xl/worksheets/sheet2.xml' = $head + '<worksheet><sheetData><row r="1"><c r="C1" t="inlineStr"><is><t>File</t></is></c><c r="D1" t="inlineStr"><is><t>Reason</t></is></c>' +
            '<c r="E1" t="inlineStr"><is><t>Group</t></is></c></row><row r="2"><c r="C2" t="inlineStr"><is><t>RecallTest/roster.xlsx</t></is></c><c r="D2" t="inlineStr"><is><t>' +
            'Could not read the file (in use, damaged or protected?): Access to the path ''C:\Users\Sample.D''Arcy.Q\OneDrive - Test Org\RecallTest\roster.xlsx'' is denied.' +
            '</t></is></c><c r="E2" t="inlineStr"><is><t>Could not read (in use, damaged or protected)</t></is></c></row></sheetData></worksheet>'
    }
    $excelBook = Join-Path $recallDir 'saved.xlsx'
    [System.IO.File]::WriteAllBytes($excelBook, [byte[]](New-RecallZip $excelParts))
    $excelRead = Read-RecallWorkbook $excelBook
    $excelHit = $excelRead.Hits[0]
    Assert-Equal 'Passport|memo.docx|***-**-0000' ('{0}|{1}|{2}' -f $excelHit.Type, $excelHit.File, $excelHit.Masked) 'Recall Score: shared strings and a rich text cell are read'
    $excelCause = $excelRead.Causes[0]   # no @() around a list property: it fails in PowerShell
    Assert-Equal 'roster.xlsx|Could not read (in use, damaged or protected)|Could not read the file (in use, damaged or protected?): Access to the path * is denied.' `
        ('{0}|{1}|{2}' -f $excelCause.File, $excelCause.Group, $excelCause.Reason) 'Recall Score: the Not scanned tab gives file, group and a reason with no path'
    $causeLines = Format-RecallScore (Get-RecallScore $recallPlan $excelRead)
    Assert-Equal 1 @($causeLines | Where-Object { $_ -like '  roster.xlsx | Could not read*path * is denied.' }).Count 'Recall Score: the cause line of a test file not scanned is printed'
    Remove-Item -LiteralPath $excelBook -Force
    $notIon = Join-Path $recallDir 'other.xlsx'
    [System.IO.File]::WriteAllBytes($notIon, [byte[]](New-RecallZip ([ordered]@{ 'xl/workbook.xml' = $head + '<workbook><sheets><sheet name="Roster" sheetId="1" r:id="rId1"/></sheets></workbook>' })))
    $refused = $false; try { $null = Read-RecallWorkbook $notIon } catch { $refused = $_.Exception.Message -match 'not an Ion results workbook' }
    Assert-Equal $true $refused 'Recall Score: a workbook that is not Ion''s is refused, not scored as 0%'
    Remove-Item -LiteralPath $notIon -Force
    Assert-Equal $true ((Invoke-RecallMake 'relative-folder') -match 'full folder path') 'Recall Make: a relative path is refused'

    # The reason Score prints: allowed words only (review of PR 105). Each fixed reason of Ion prints unchanged.
    # The fixed reasons of Ion: each string in src/ (PowerShell parser) with 'check by hand' or a reason start; a variable
    # inside a string reads as 9.
    $reasonStart = '^(Could not read the file|The reader failed|Scanned PDF|Download failed|Download gave|Empty file|Larger than|Encrypted[ (]|Sensitivity label)'
    $fixedReasons = @(Get-ChildItem (Join-Path $recallRoot 'src') -Filter '*.ps*1' | ForEach-Object {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$null, [ref]$null)
            $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
                    $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true) | ForEach-Object {
                $text = [regex]::Replace([string]$_.Value, '\$\{?[\w:]+\}?', '9')
                if ($text -match 'check by hand' -or $text -match $reasonStart) { $text }
            } } | Sort-Object -Unique)
    Assert-Equal $true ($fixedReasons.Count -ge 40) 'Recall Score: the fixed reasons of Ion are found in src'
    Assert-Equal '' (@($fixedReasons | Where-Object { (Get-RecallSafeReason $_) -cne $_ }) -join ' || ') 'Recall Score: each fixed reason of Ion prints unchanged'
    # Invented paths and names in every form the reviews found (quoted, unquoted, relative, forward slash, nested quote).
    $leaky = @(
        "Access to the path 'C:\Users\Sample.D'Arcy.Q\OneDrive - Test Org\x.docx' is denied.",
        'The reader failed: Exception calling "ReadAllBytes": "Could not find file "C:\Users\Sample Person\x.docx"."',
        'Could not read the file (damaged or protected?): The network path \\host-one\Sample.Person\a b.docx was not found',
        'The reader failed: at Sample\Person line 4 for sample.person@example.test and 1234567890',
        "Access to the path '..\Users\Sample Person' is denied.", 'The reader failed: (/Users/Sample Person/x.docx)',
        'The reader failed: path=/Users/Sample Person/x.docx', 'Could not find "/Users/Sample Person/x.docx',
        'Could not find Users/Sample Person/x.docx', 'Could not find ./Sample Person/x.docx',
        "Access to the path 'Sample O'Neil\x.docx' is denied.", 'The file is locked for editing by Sample Person.',
        'Locked by Samplé.', 'Locked by Samplée Persön.', 'Locked by SampleException.')
    $leaked = @($leaky | ForEach-Object { Get-RecallSafeReason $_ } | Where-Object { $_ -match 'Samp|Pers|Neil|Arcy|host|Test Org|example|\d{5}' })
    Assert-Equal '' ($leaked -join ' || ') 'Recall Score: no word of a path or name shows'
    $inUse = 'Could not read the file (in use, damaged or protected?): The process cannot access the file '
    $inUseEnd = ' because it is being used by another process.'
    Assert-Equal ($inUse + '*' + $inUseEnd) (Get-RecallSafeReason ($inUse + "'C:\Users\Sample Person\OneDrive - Test Org\RecallTest\memo.docx'" + $inUseEnd)) `
        'Recall Score: an in-use error keeps its words and loses its path'
    Assert-Equal 163 (Get-RecallSafeReason ('file ' * 60)).Length 'Recall Score: a reason is cut at 160 characters'


    # The local baseline: Ion scans the made folder as on the gov device (the bundle's reader block in the workers).
    $ionText = [System.IO.File]::ReadAllText((Join-Path (Join-Path $recallRoot 'paste') 'Ion-script.txt'))
    $readerBlock = [regex]::Match($ionText, '(?s)\$IonReaderBlock = \{\r?\n(.*?)\r?\n\}\r?\n\$null = New-Module -Name IonReaders')
    Assert-Equal $true $readerBlock.Success 'Recall: the reader block is in paste/Ion-script.txt'
    Set-FakeSiteAnswer
    $recallBook = Invoke-DriveScan -Path $recallDir -OutputFolder $recallOut -MaxFileMB 25 -Init @{ Code = $readerBlock.Groups[1].Value } 6>$null
    $ionBook = Read-RecallWorkbook $recallBook
    Assert-Equal (@((Get-XlsxSheets $recallBook)['Hits']).Count - 1) $ionBook.Hits.Count 'Recall Score: every row of the Hits tab is read'
    $ionScore = Get-RecallScore $recallPlan $ionBook
    Assert-Equal 0 $ionScore.NotScanned 'Recall: Ion reads all 12 test files'
    $sums = @{ Found = 0; False = 0; Planted = 0 }
    foreach ($t in $ionScore.ByType.Values) { $sums.Found += $t.Found; $sums.False += $t.False; $sums.Planted += $t.Planted }
    Assert-Equal $ionBook.Hits.Count ($sums.Found + $sums.False) 'Recall Score: each hit is found or false, once'
    Write-Host ('RECALL found={0}/{1} false={2} (local baseline, docs/reviews/2026-11-24-recall-test.md)' -f $sums.Found, $sums.Planted, $sums.False)
}
finally {
    Remove-Item -LiteralPath $recallDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $recallOut -Recurse -Force -ErrorAction SilentlyContinue
}
