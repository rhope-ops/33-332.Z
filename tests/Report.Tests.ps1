<#
Report.Tests.ps1 - tests of Report.psm1 and XlsxWriter.psm1: the screen summary and the results workbook.

  Purpose:     The workbook goes to <site> as "Ion <site> <date time>", never overwrites, has the Summary tab first and one tab per PII
               type worst first, holds masked values only and no formula; the folder column is named for a team
               or a site; the screen shows hits per type.
  Inputs:      $Sample (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      report files in a new temporary folder, deleted at the end of this file.
  Credentials: none.
#>

$reportFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-report-' + [guid]::NewGuid())
$hit = @(Find-PiiInText "SSN $($Sample.Ssn)")[0]
$reportRows = @(
    (New-ReportRow 'File' 'General' 'Roster.docx' 'https://x/y' '2026-11-01' $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context),
    (New-ReportRow 'File' 'General' 'Bad.docx' '' '' 'Phone' 'Low' '***-***-0142' '=HYPERLINK("x")'),
    (New-ReportRow 'File' 'Ops' 'scan.pdf' '' '' 'NOT SCANNED' '' '' 'File type is not read by this script')
)

# ---- The workbook: Summary first, one tab per type (worst first), Low confidence, Not scanned ----
$dob = @(Find-PiiInText 'DOB: 03/14/1985')[0]
$reportRows += (New-ReportRow 'File' 'Ops' 'Recall.xlsx' '' '' $dob.Detector $dob.Confidence $dob.MaskedValue $dob.Context)
$header = [ordered]@{ Site = 'https://tenant.example.com/teams/TeamA'; Channels = 'General, Ops' }
$path = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site -FolderLabel 'Channel' 6>$null
Assert-Equal 'TeamA' (Split-Path (Split-Path $path -Parent) -Leaf) 'Saved in <site>, no day folder'
Assert-Equal $true ((Split-Path $path -Leaf) -match ('^Ion TeamA ' + (Get-Date -Format 'yyyy-MM-dd') + ' \d{4}( \(\d+\))?\.xlsx$')) 'Named Ion <site> <date time>.xlsx'
$book = Get-XlsxSheets $path
Assert-Equal 'Summary|SSN|Date of birth|Low confidence|Not scanned|Method' (@($book.Keys) -join '|') 'Tabs: Summary, types worst first, Low, Not scanned, Method last'
$hitHeader = 'Confidence|Channel|File|Masked value|Context|Last saved|Link|Found in|Created|Created by|Last saved by|PII first saved|PII first saved by|History note'
Assert-Equal $hitHeader ($book['SSN'][0] -join '|') 'Hit tab header, folder column named Channel for a team'
Assert-Equal 'High|General|Roster.docx|***-**-6789' (($book['SSN'][1] | Select-Object -First 4) -join '|') 'SSN tab row'
Assert-Equal 'Phone' $book['Low confidence'][1][0] 'Low hit in the Low confidence tab, with its type'
Assert-Equal '=HYPERLINK("x")' $book['Low confidence'][1][5] 'Formula-like text kept as plain text (an inline string is never a formula)'
Assert-Equal 'NOT SCANNED|Ops|scan.pdf' (($book['Not scanned'][1] | Select-Object -First 3) -join '|') 'Not scanned tab row'
$summary = ($book['Summary'] | ForEach-Object { $_ -join '|' }) -join "`n"
Assert-Equal $true ($summary -match '(?m)^Type not read: \.pdf\|1$') 'Summary: not-scanned files counted per reason'
# The link column is a clickable link: a sheet relationship, never a formula.
$zipStream = [System.IO.File]::OpenRead($path)
$zipBook = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $relsEntry = $zipBook.GetEntry('xl/worksheets/_rels/sheet2.xml.rels')
    $reader = New-Object System.IO.StreamReader($zipBook.GetEntry('xl/worksheets/sheet2.xml').Open()); $sheetXml = $reader.ReadToEnd(); $reader.Dispose()
    $rels = ''
    if ($relsEntry) { $reader = New-Object System.IO.StreamReader($relsEntry.Open()); $rels = $reader.ReadToEnd(); $reader.Dispose() }
}
finally { $zipBook.Dispose(); $zipStream.Dispose() }
Assert-Equal $true ($sheetXml -match '<hyperlink ref="G2" r:id="rId1"/>' -and $rels -match 'Target="https://x/y" TargetMode="External"') 'Link cell is a clickable link'
Assert-Equal $false ($sheetXml -match '<f>') 'No formula in a hit tab'
Assert-Equal $true ($summary -match 'Possible PII hits\|3') 'Summary: total hits'
$typeRows = $summary -match '(?m)^SSN\|1\|0\|0\|1$' -and $summary -match '(?m)^Date of birth\|1\|0\|0\|1$' -and $summary -match '(?m)^Phone\|0\|0\|1\|1$'
Assert-Equal $true $typeRows 'Summary: hits per type (High, Medium, Low, Total), worst first'
Assert-Equal $true ($summary -match 'potential or actual breach' -and $summary -match 'para 3\.1\.2') 'Summary gives the cited breach step'
$allCells = ($book.Values | ForEach-Object { $_ | ForEach-Object { $_ -join ' ' } }) -join ' '
Assert-Equal $false $allCells.Contains($Sample.Ssn) 'Workbook has no full SSN'
$again = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site 6>$null
Assert-Equal $true ($again -ne $path) 'A second run never overwrites the first'
Assert-Equal 'Folder' (Get-XlsxSheets $again)['SSN'][0][1] 'Folder column named Folder for a plain site'
$plain = @(Get-WorkbookSheets @() 0 ([ordered]@{}) 'Folder')
Assert-Equal 'Summary|Method' (@($plain | ForEach-Object { $_.Name }) -join '|') 'No hits: only the Summary and Method tabs'

# ---- The workbook writer: odd text and odd tab names ----
$oddPath = Join-Path $reportFolder 'odd.xlsx'
$long = 'x' * 40000
Export-Xlsx -Path $oddPath -Sheets @(
    (New-XlsxSheet 'A:B/C?D*E[F]G and a very long tab name past 31' @('Text') @(, @('bad' + [char]1 + 'char <&>')) @(20) -List),
    (New-XlsxSheet 'A:B/C?D*E[F]G and a very long tab name past 31' @('Text') @(, @($long)) @(20) -List))
$odd = Get-XlsxSheets $oddPath
Assert-Equal 'A-B-C-D-E-F-G and a very long t|A-B-C-D-E-F-G and a very lo (2)' (@($odd.Keys) -join '|') 'Tab names: bad characters replaced, cut to 31, made unique'
Assert-Equal 'badchar <&>' (@($odd.Values)[0][1][0]) 'Characters XML does not allow dropped; <&> kept as text'
Assert-Equal 32767 (@($odd.Values)[1][1][0]).Length 'A cell is cut at 32,767 characters'
$refused = $false
try { Export-Xlsx -Path $oddPath -Sheets @((New-XlsxSheet 'S' @('T') @() @(5))) } catch { $refused = $true }  # expected: the file exists
Assert-Equal $true $refused 'An existing workbook is never overwritten'
Assert-Equal $true ((Get-ResultPath $reportFolder 'https://tenant.example.com/teams/CON') -match '[\\/]_CON[\\/]') 'A Windows device name gets a leading _'
Assert-Equal $true ((Get-ResultPath $reportFolder 'https://tenant.example.com/teams/nul.x') -match '[\\/]_nul\.x[\\/]') 'A device name with an extension too'

# ---- The screen summary ----
$screenLines = @(Get-ScanSummaryText -Rows $reportRows -ItemsScanned 2 -Header $header -FolderLabel 'Channel')
$screen = ($screenLines | ForEach-Object { $_.Text }) -join "`n"
Assert-Equal $true ($screen -match 'POSSIBLE PII BY TYPE' -and $screen -match '\| SSN\s+\|\s+1 \|\s+0 \|\s+0 \|\s+1 \|') 'Screen: boxed table of hits per type (High, Medium, Low, Total)'
Assert-Equal 'Alert' ($screenLines | Where-Object { $_.Text -match '^\| SSN ' } | Select-Object -First 1).Kind 'Screen: a type with High hits is red'
Assert-Equal 'Text' ($screenLines | Where-Object { $_.Text -match '^\| Phone ' } | Select-Object -First 1).Kind 'Screen: a type with only Low hits is white'
Assert-Equal 'Accent' ($screenLines | Where-Object { $_.Text -match '^\| Type not read: \.pdf ' } | Select-Object -First 1).Kind 'Screen: not-scanned reasons in yellow'
Assert-Equal 'Site:' ($screenLines | Where-Object { $_.Text -like 'Site:*' } | ForEach-Object { $_.Text.Substring(0, $_.LabelLength).Trim() }) 'Screen: header label in the label colour'
Assert-Equal $true ($screen.IndexOf('SSN ') -lt $screen.IndexOf('Date of birth')) 'Screen: worst type first'
Assert-Equal $false $screen.Contains($Sample.Ssn) 'Screen: no full SSN'
$masked = Get-ResultPath $reportFolder ('https://tenant.example.com/teams/' + $Sample.SsnBare)
Assert-Equal $false $masked.Contains($Sample.SsnBare) 'A site name that looks like PII is masked in the folder name'

# ---- A run that did not reach its end (Ctrl+C runs 'finally' but not 'catch') ----
$cancelRows = New-Object System.Collections.Generic.List[object]
Add-IncompleteRow -Rows $cancelRows -Finished $false
Assert-Equal 'INCOMPLETE' $cancelRows[0].Detector 'Unfinished run gets an INCOMPLETE row'
Add-IncompleteRow -Rows $cancelRows -Finished $false
Assert-Equal 1 $cancelRows.Count 'Only one INCOMPLETE row'
$doneRows = New-Object System.Collections.Generic.List[object]
Add-IncompleteRow -Rows $doneRows -Finished $true
Assert-Equal 0 $doneRows.Count 'Finished run gets no INCOMPLETE row'

Remove-Item -LiteralPath $reportFolder -Recurse -Force

# ---- Screen polish (the PAM, 2026-11-21): a stopped scan says so first; tables fit the window; the confidence key ----
$stopRows = @((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped early'))
$stopText = @(Get-ScanSummaryText -Rows $stopRows -ItemsScanned 7 -Header ([ordered]@{ Site = 'x' }) -FolderLabel 'Folder')
Assert-Equal 'Alert|INCOMPLETE: the scan stopped early. Only the 7 files read are in these results; scan again.' `
    ($stopText[2].Kind + '|' + $stopText[2].Text) 'Stopped scan: a red line right under the Results heading'
$wide = @(& (Get-Module Report) { Format-BoxTable @('Worst', 'File') @(, @('SSN', ('a' * 200))) 60 })
Assert-Equal 60 ($wide | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum 'Table: no line wider than the window'
Assert-Equal $true ($wide[3] -match '\.\.\. \|$') 'Table: the cut file path ends in ...'
$keyText = (@(Get-ScanSummaryText -Rows $reportRows -ItemsScanned 2 -Header $header -FolderLabel 'Channel') | ForEach-Object { $_.Text }) -join "`n"
Assert-Equal $true ($keyText -match '(?m)^High = labelled') 'Summary: the key to High, Medium and Low is under the type table'
$narrow = @(& (Get-Module Report) { Format-BoxTable @('Reason', 'Files') @(, @(('r' * 115), 1)) 119 })
Assert-Equal $true (($narrow | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum -le 127) 'Table: a narrow last column is never widened'
Assert-Equal 0 @(& (Get-Module Report) { $script:ConfidenceKey } | Where-Object { $_.Length -gt 79 }).Count 'Summary: each confidence key line fits 80 characters'
# The AFI reference of the breach advice is its own line on screen and its own row on the Summary tab (the PAM, 2026-11-21).
$adviceShown = (Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site -FolderLabel 'Channel' 6>&1 |
    Where-Object { $_ -isnot [string] } | ForEach-Object { [string]$_ })
Assert-Equal $true (@($adviceShown | Where-Object { $_ -match '^Source: AFI 33-332 para 3\.1\.2' }).Count -eq 1) 'Advice: the AFI reference is its own line'
$sourceBook = Get-XlsxSheets (@(Get-ResultFiles $reportFolder) | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
Assert-Equal 1 @($sourceBook['Summary'] | Where-Object { $_[0] -eq 'Rule source' -and $_[1] -match '^AFI 33-332' }).Count 'Advice: the Summary tab has the reference as its own row'
# The Method tab explains the who-and-when columns and says the first-saved account is evidence, not proof (the PAM, 2026-11-21).
$methodText = (@($sourceBook['Method']) | ForEach-Object { $_ -join ' ' }) -join "`n"
$methodOk = $methodText -match 'PII first saved by' -and $methodText -match 'evidence, not proof' -and $methodText -match 'Version history'
Assert-Equal $true $methodOk 'Method tab: how the columns are worked out, and the limits'
# A long site name is cut so the workbook path stays under 260 characters (review of PR 41).
$longPath = & (Get-Module Report) { param($f) Get-ResultPath $f ('https://tenant.example.com/sites/' + ('a' * 150)) } $reportFolder
Assert-Equal $true ((Split-Path $longPath -Leaf).Length -lt 70 -and $longPath.Length -lt ($reportFolder.Length + 120)) 'Long site name: cut in the folder and the file name'
