<#
Report.Tests.ps1 - tests of Report.psm1 and XlsxWriter.psm1: the screen summary and the results workbook.

  Purpose:     The workbook goes to <site> as "Ion <site> <date time>", never overwrites, has the Summary tab first, then Hits and Files
               as Excel Tables (worst first), holds masked values only and no formula from scanned text; the folder column is named for a team
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

# ---- The workbook: Summary first, Hits and Files as Excel Tables, Not scanned (the PAM, 2026-11-24: as previewed) ----
$dob = @(Find-PiiInText 'DOB: 03/14/1985')[0]
$reportRows += (New-ReportRow 'File' 'Ops' 'Recall.xlsx' '' '' $dob.Detector $dob.Confidence $dob.MaskedValue $dob.Context)
$header = [ordered]@{ Site = 'https://tenant.example.com/teams/TeamA'; Channels = 'General, Ops' }
$path = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site -FolderLabel 'Channel' 6>$null
Assert-Equal 'TeamA' (Split-Path (Split-Path $path -Parent) -Leaf) 'Saved in <site>, no day folder'
Assert-Equal $true ((Split-Path $path -Leaf) -match ('^Ion TeamA ' + (Get-Date -Format 'yyyy-MM-dd') + ' \d{4}( \(\d+\))?\.xlsx$')) 'Named Ion <site> <date time>.xlsx'
$book = Get-XlsxSheets $path
Assert-Equal 'Summary|Hits|Files|Not scanned|Method' (@($book.Keys) -join '|') 'Tabs: Summary, Hits, Files, Not scanned, Method last'
$hitHeader = 'ID|Type|Confidence|Channel|File|Masked value|Context|Is it PII?|Action taken|Reported to Privacy Manager|Reviewed by|Review date|' +
    'Last saved|Last saved by|PII first saved|PII first saved by|Created|Created by|History note|Found in|Link'
Assert-Equal $hitHeader ($book['Hits'][0] -join '|') 'Hits header: the review columns right after the hit; folder column named Channel for a team'
Assert-Equal 'H001|SSN|High|General|Roster.docx|***-**-6789' (($book['Hits'][1] | Select-Object -First 6) -join '|') 'Hits row: ID, type, confidence, folder, file, value'
$order = ($book['Hits'][2][0], $book['Hits'][2][1], $book['Hits'][3][0], $book['Hits'][3][1], $book['Hits'][3][2]) -join '|'
Assert-Equal 'H002|Date of birth|H003|Phone|Low' $order 'Hits: worst first (High, then type), Low last, IDs in that order'
Assert-Equal '=HYPERLINK("x")' $book['Hits'][3][6] 'Formula-like text kept as plain text (an inline string is never a formula)'
Assert-Equal 'General|Roster.docx|1|1|0|0|SSN|0' (($book['Files'][1] | Select-Object -First 8) -join '|') 'Files tab: one row per file, counts, types, reviewed'
Assert-Equal 'NOT SCANNED|Ops|scan.pdf' (($book['Not scanned'][1] | Select-Object -First 3) -join '|') 'Not scanned tab row'
Assert-Equal 'Group|Type not read: .pdf' ($book['Not scanned'][0][4] + '|' + $book['Not scanned'][1][4]) 'Not scanned tab: a Group column, as the Summary counts it'
$summary = ($book['Summary'] | ForEach-Object { $_ -join '|' }) -join "`n"
Assert-Equal $true ($summary -match '(?m)^Type not read: \.pdf\|1$') 'Summary: not-scanned files counted per reason'
# The link column is a clickable link: a sheet relationship, never a formula.
$zipStream = [System.IO.File]::OpenRead($path)
$zipBook = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $relsEntry = $zipBook.GetEntry('xl/worksheets/_rels/sheet2.xml.rels')
    $reader = New-Object System.IO.StreamReader($zipBook.GetEntry('xl/worksheets/sheet2.xml').Open()); $sheetXml = $reader.ReadToEnd(); $reader.Dispose()
    $sheetCount = @($zipBook.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' }).Count
    $reader = New-Object System.IO.StreamReader($zipBook.GetEntry(('xl/worksheets/sheet{0}.xml' -f $sheetCount)).Open()); $methodXml = $reader.ReadToEnd(); $reader.Dispose()
    $rels = ''
    if ($relsEntry) { $reader = New-Object System.IO.StreamReader($relsEntry.Open()); $rels = $reader.ReadToEnd(); $reader.Dispose() }
    $reader = New-Object System.IO.StreamReader($zipBook.GetEntry('xl/tables/table1.xml').Open()); $tableXml = $reader.ReadToEnd(); $reader.Dispose()
    $reader = New-Object System.IO.StreamReader($zipBook.GetEntry('xl/worksheets/sheet1.xml').Open()); $summaryXml = $reader.ReadToEnd(); $reader.Dispose()
}
finally { $zipBook.Dispose(); $zipStream.Dispose() }
Assert-Equal $true ($sheetXml -match '<hyperlink ref="E2" r:id="rId1"/>' -and $rels -match 'Target="https://x/y" TargetMode="External"') 'The file name is a clickable link'
Assert-Equal $false ($sheetXml -match '<f>') 'No formula in the Hits tab'
Assert-Equal $true ($tableXml -match 'displayName="tblHits" ref="A1:U4"' -and $tableXml -match '<tableColumn id="8" name="Is it PII\?"/>' -and
    $rels -match 'Target="../tables/table1.xml"' -and $sheetXml -match '<tablePart r:id="rIdTable"/>' -and $sheetXml -notmatch '<autoFilter') 'Hits is an Excel Table'
Assert-Equal $true ($sheetXml -match '<c r="M2" s="4"><v>\d+(\.\d+)?</v></c>') 'Last saved is a date cell'
Assert-Equal $true ($sheetXml -match '<c r="C2" s="5"' -and $sheetXml -match '<c r="C4" s="7"') 'Confidence cells get the High and Low fill'
Assert-Equal $true ($sheetXml -match 'xSplit="5"' -and $sheetXml -match '&amp;CCUI' -and $sheetXml -match 'fitToWidth="1"') 'Hits: columns frozen; CUI header and footer; one page wide'
Assert-Equal $true ($sheetXml -match '<dataValidation type="date" operator="greaterThan"[^>]*sqref="L2:L4"><formula1>36526</formula1>') 'Review date: Excel accepts only a date'
Assert-Equal $true ($summary -match '(?m)^Start here: 1\|Open the tab Hits' -and $summary -match '(?m)^Start here: 4\|') 'Summary: the 4 Start here steps'
Assert-Equal $true ($summaryXml -match '<f>COUNTIFS\(tblHits\[Type\],\$A\d+,tblHits\[Confidence\],&quot;High&quot;\)</f><v>1</v>') 'Summary counts are formulas on the Hits table, with their value'
Assert-Equal $true ($summary -match 'Possible PII hits\|3') 'Summary: total hits'
$typeRows = $summary -match '(?m)^SSN\|1\|0\|0\|1\|0$' -and $summary -match '(?m)^Date of birth\|1\|0\|0\|1\|0$' -and $summary -match '(?m)^Phone\|0\|0\|1\|1\|0$' -and
    $summary -match '(?m)^Total\|2\|0\|1\|3\|0$'
Assert-Equal $true $typeRows 'Summary: hits per type (High, Medium, Low, Total, Reviewed), worst first, and a Total row'
Assert-Equal $true ($summary -match 'potential or actual breach' -and $summary -match 'para 3\.1\.2') 'Summary gives the cited breach step'
$allCells = ($book.Values | ForEach-Object { $_ | ForEach-Object { $_ -join ' ' } }) -join ' '
Assert-Equal $false $allCells.Contains($Sample.Ssn) 'Workbook has no full SSN'
$again = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site 6>$null
Assert-Equal $true ($again -ne $path) 'A second run never overwrites the first'
Assert-Equal 'Folder' (Get-XlsxSheets $again)['Hits'][0][3] 'Folder column named Folder for a plain site'
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
$wide = @(& (Get-Module ScanMeter) { Format-BoxTable @('Worst', 'File') @(, @('SSN', ('a' * 200))) 60 })
Assert-Equal 60 ($wide | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum 'Table: no line wider than the window'
Assert-Equal $true ($wide[3] -match '\.\.\. \|$') 'Table: the cut file path ends in ...'
$keyText = (@(Get-ScanSummaryText -Rows $reportRows -ItemsScanned 2 -Header $header -FolderLabel 'Channel') | ForEach-Object { $_.Text }) -join "`n"
Assert-Equal $true ($keyText -match '(?m)^High = labelled') 'Summary: the key to High, Medium and Low is under the type table'
$narrow = @(& (Get-Module ScanMeter) { Format-BoxTable @('Reason', 'Files') @(, @(('r' * 115), 1)) 119 })
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
$methodOk = $methodText -match 'PII first saved by' -and $methodText -match 'evidence, not proof' -and $methodText -match 'Version history' -and
    $methodText -match 'SSN Numbers the SSA never issues' -and $methodText -match 'DoD ID A 10-digit number, found only with' -and
    $methodText -match '10 digits in a row are not found'
Assert-Equal $true $methodOk 'Method tab: how the columns are worked out, and the limits'
# A long site name is cut so the workbook path stays under 260 characters (review of PR 41).
$longPath = & (Get-Module ReportBook) { param($f) Get-ResultPath $f ('https://tenant.example.com/sites/' + ('a' * 150)) } $reportFolder
Assert-Equal $true ((Split-Path $longPath -Leaf).Length -lt 70 -and $longPath.Length -lt ($reportFolder.Length + 120)) 'Long site name: cut in the folder and the file name'
# C6 (the PAM, 2026-11-21): the review columns have dropdowns; the Method tab wraps its text; By account counts per account.
$piiChoice = $sheetXml -match '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="H2:H\d+"><formula1>"Yes,No,Not sure"</formula1>'
Assert-Equal $true $piiChoice `
    'Review columns: "Is it PII?" has a Yes/No/Not sure dropdown'
Assert-Equal $true ($sheetXml -match 'sqref="I2:I\d+"><formula1>"Removed,Moved,Access limited,Left \(not PII\),Other"') 'Review columns: "Action taken" has its dropdown'
Assert-Equal $true ($methodXml -match '<c r="B2" s="3"') 'Method tab: the text column wraps (style 3)'
$accountHits = @(
    [pscustomobject]@{ AddedBy = 'a.one'; Channel = 'Ops'; Item = 'x.docx'; Confidence = 'High' },
    [pscustomobject]@{ AddedBy = 'a.one'; Channel = 'Ops'; Item = 'x.docx'; Confidence = 'Low' },
    [pscustomobject]@{ AddedBy = 'b.two'; Channel = 'Ops'; Item = 'y.docx'; Confidence = 'Medium' },
    [pscustomobject]@{ AddedBy = 'Not recorded'; Channel = 'Ops'; Item = 'z.docx'; Confidence = 'High' })
$accountRows = @(Get-AccountRows $accountHits)
Assert-Equal 'a.one 1 1 0 1|b.two 1 0 1 0' (($accountRows | ForEach-Object { $_ -join ' ' }) -join '|') 'By account: files once per account, hits per confidence, most first, Not recorded left out'
# The app's mark (gamma, Logo Studies) only on the start line next to "Ion" (the PAM, 2026-11-21).
$mark = [string][char]0x3B3
$bannerShown = (& (Get-Module Report) { Write-Banner 'Ion' -Mark } 6>&1 | ForEach-Object { [string]$_ }) -join ''
Assert-Equal ('==== ' + $mark + ' Ion ====') $bannerShown 'Start banner: the mark between the rules'
$bannerShown = (& (Get-Module Report) { Write-Banner 'Results' } 6>&1 | ForEach-Object { [string]$_ }) -join ''
Assert-Equal '==== Results ====' $bannerShown 'Other banners: no mark'
$tieHits = @([pscustomobject]@{ AddedBy = 'b.two'; Channel = 'Ops'; Item = 'y'; Confidence = 'Low' },
    [pscustomobject]@{ AddedBy = 'a.one'; Channel = 'Ops'; Item = 'x'; Confidence = 'Low' })
$tieRows = @(Get-AccountRows $tieHits)
Assert-Equal 'a.one|b.two' (($tieRows | ForEach-Object { $_[0] }) -join '|') 'By account: equal totals in name order'
# Not-scanned groups keep a sensitivity label, a password, the PDF time limit and a part not read apart (the PAM, 2026-11-21).
$groupOf = { param($reason) & (Get-Module Report) { param($r) Get-NotScannedGroup ([pscustomobject]@{ Detector = 'NOT SCANNED'; Item = 'a.pdf'; Context = $r }) } $reason }
Assert-Equal 'Sensitivity label with encryption' (& $groupOf 'Sensitivity label with encryption (Microsoft Purview): open it and check by hand') 'Groups: a sensitivity label'
Assert-Equal 'Encrypted (password or author limits)' (& $groupOf 'Encrypted (password-protected): check by hand') 'Groups: a password'
Assert-Equal 'PDF time limit reached (file may be fine)' (& $groupOf 'This PDF took longer than the time limit to read (2 minutes; the file may be fine): check by hand') 'Groups: the PDF time limit'
Assert-Equal 'PDF parts not read (file may open fine)' (& $groupOf 'Part of this PDF is not well formed (the file may still open fine): check by hand') 'Groups: a PDF part not read'
Assert-Equal 'Not a valid file (damaged or wrong name)' (& $groupOf 'Not a valid PDF (damaged, or another type with a .pdf name): check by hand') 'Groups: not a valid file'
Assert-Equal 'Could not read (in use, damaged or protected)' (& $groupOf 'Could not read the file (in use, damaged or protected?): x') 'Groups: a file that could not be read is not called invalid'
# An e-signature DoD ID (the only Low DoD ID) gets no version-history read; other hits do (the PAM, 2026-11-21).
$looked = New-Object System.Collections.Generic.List[string]
# A fake version-history lookup: records which hit types it was asked about.
$addedLookup = { param($hits) foreach ($h in $hits) { $looked.Add($h.Detector) }; @{ 'SSN|***-**-6789' = @{ AddedOn = 'x'; AddedBy = 'y'; Note = '' } } }
$addedInfo = @{ Added = $addedLookup.GetNewClosure() }
$mixed = @([pscustomobject]@{ Detector = 'DoD ID'; Confidence = 'Low'; MaskedValue = '******7890' },
    [pscustomobject]@{ Detector = 'SSN'; Confidence = 'High'; MaskedValue = '***-**-6789' })
$byHit = & (Get-Module ScanQueue) { param($i, $h) Get-AddedByHit $i $h } $addedInfo $mixed
$seen = ($looked -join ',') + '|' + $byHit['DoD ID|******7890'].AddedBy + '|' + $byHit['SSN|***-**-6789'].AddedBy
Assert-Equal 'SSN|Not recorded|y' $seen 'History: read for the SSN only; the e-signature DoD ID is not looked up'
$looked.Clear()
$null = & (Get-Module ScanQueue) { param($i, $h) Get-AddedByHit $i $h } $addedInfo @($mixed[0])
Assert-Equal 0 $looked.Count 'History: a file with only an e-signature DoD ID reads no versions'
# Review of PR 50: a labelled DoD ID with the same masked value keeps its history; no made-up account.
$sameId = @([pscustomobject]@{ Detector = 'DoD ID'; Confidence = 'High'; MaskedValue = '******7890' },
    [pscustomobject]@{ Detector = 'DoD ID'; Confidence = 'Low'; MaskedValue = '******7890' })
$idInfo = @{ Added = { param($hits) @{ 'DoD ID|******7890' = @{ AddedOn = 'x'; AddedBy = 'a.one'; Note = '' } } } }
$idByHit = & (Get-Module ScanQueue) { param($i, $h) Get-AddedByHit $i $h } $idInfo $sameId
Assert-Equal 'a.one' $idByHit['DoD ID|******7890'].AddedBy 'History: the labelled DoD ID keeps its result'
$notLooked = @([pscustomobject]@{ AddedBy = 'Not recorded'; Channel = 'Ops'; Item = 'sig.pdf'; Confidence = 'Low' })
Assert-Equal 0 @(Get-AccountRows $notLooked).Count 'By account: an e-signature DoD ID makes no account row'
