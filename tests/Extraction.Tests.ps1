<#
Extraction.Tests.ps1 - tests of TextExtract.psm1: text from Office files, text files and HTML.

  Purpose:     Each supported type gives its text; split Word runs join again; other types and damaged
               files are reported as NOT SCANNED.
  Inputs:      $Sample and New-ZipBytes (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal; $Docx and $Txt for the later test files.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- Word: an SSN split over two runs must join again ----
$ssnHead = $Sample.Ssn.Substring(0, 7)
$ssnTail = $Sample.Ssn.Substring(7)
$wordXml = "<w:document><w:body><w:p><w:r><w:t>SSN $ssnHead</w:t></w:r><w:r><w:t>$ssnTail</w:t></w:r></w:p>" +
           '<w:p><w:r><w:t>Next &amp; line</w:t></w:r></w:p></w:body></w:document>'
$Docx = New-ZipBytes @{ 'word/document.xml' = $wordXml }
$result = Get-TextFromBytes -Bytes $Docx -FileName 'Roster.docx'
Assert-Equal 'OK' $result.Status 'docx read'
Assert-Equal 'SSN/High/***-**-6789' (Get-HitSummary $result.Text) 'docx split runs joined'
Assert-Equal $true ($result.Text -match 'Next & line') 'docx XML entities decoded'

# ---- Excel: shared strings and cell values ----
$sheetXml = "<worksheet><sheetData><row><c t=`"s`"><v>0</v></c><c><v>$($Sample.DodId)</v></c></row></sheetData></worksheet>"
$xlsx = New-ZipBytes @{ 'xl/sharedStrings.xml' = '<sst><si><t>DoD ID</t></si></sst>'; 'xl/worksheets/sheet1.xml' = $sheetXml }
$result = Get-TextFromBytes -Bytes $xlsx -FileName 'Recall.xlsx'
Assert-Equal 'OK' $result.Status 'xlsx read'
Assert-Equal $true $result.Text.Contains($Sample.DodId) 'xlsx cell value read'
Assert-Equal "DoD ID | $($Sample.DodId)" $result.Text 'xlsx shared string back in its cell, next to its value'

# ---- Excel: no false positives from neighbouring cells, decimals, formulas or string indexes ----
$noise = '<worksheet><sheetData>' +
    '<row><c><v>123</v></c><c><v>45</v></c><c><v>6789</v></c></row>' +                            # three cells, not one SSN
    '<row><c><v>45231.416666666664</v></c><c><v>1.2345678901E+20</v></c></row>' +                  # a time; an exponent
    '<row><c><f>VLOOKUP(A1,Sheet2!A:B,2)*123456789</f><v>7</v></c></row>' +                        # formula text is dropped
    '<row>' + ((1..16 | ForEach-Object { '<c><v>4</v></c>' }) -join '') + '</row>' +                # 16 single digits, not a card
    '<row><c t="b"><v>1</v></c><c t="str"><v>OK</v></c><c t="inlineStr"><is><t>note</t></is></c></row>' +
    '</sheetData></worksheet>'
$noiseText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $noise }) -FileName 'Budget.xlsx').Text
Assert-Equal 0 @(Find-PiiInText $noiseText).Count 'Excel: no hit from cells, decimals, formulas or digit runs'
Assert-Equal $false $noiseText.Contains('VLOOKUP') 'Excel: formula text not read'
Assert-Equal $true ($noiseText -match '7' -and $noiseText -match 'OK \| note') 'Excel: formula results and text cells kept'
# Tags with a namespace prefix (some tools write <x:row>, <x:c>) and rich text in a shared string are read too.
$prefixed = "<x:worksheet><x:sheetData><x:row><x:c t=`"s`"><x:v>0</x:v></x:c><x:c><x:v>$($Sample.SsnBare)</x:v></x:c></x:row></x:sheetData></x:worksheet>"
$richShared = '<x:sst><x:si><x:r><x:t>S</x:t></x:r><x:r><x:rPr><x:b/></x:rPr><x:t>SN</x:t></x:r></x:si></x:sst>'
$prefixedText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $prefixed; 'xl/sharedStrings.xml' = $richShared }) -FileName 'p.xlsx').Text
Assert-Equal "SSN | $($Sample.SsnBare)" $prefixedText 'Excel: prefixed tags and rich-text shared strings read'
Assert-Equal 'SSN' (@(Find-PiiInText $prefixedText) | Select-Object -First 1).Detector 'Excel: SSN found in a file with prefixed tags'
$selfClosing = '<worksheet><sheetData><row><c r="A1"/><c r="B1"><v>5</v></c></row></sheetData></worksheet>'
Assert-Equal '5' (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $selfClosing }) -FileName 's.xlsx').Text 'Excel: an empty self-closing cell is skipped'
$realSsn = "<worksheet><sheetData><row><c t=`"inlineStr`"><is><t>SSN</t></is></c><c><v>$($Sample.SsnBare)</v></c></row></sheetData></worksheet>"
Assert-Equal 1 @(Find-PiiInText (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $realSsn }) -FileName 'a.xlsx').Text |
    Where-Object { $_.Detector -eq 'SSN' }).Count 'Excel: a real SSN in a number cell is still found'

# ---- Roster columns: a PII label only in a column header labels the values below it (the PAM, 2026-11-24) ----
function Get-RosterXlsxText {
    # The text of a one-sheet workbook; each row a list of 'B2=text' cells (digits only: a number cell, else an inline string).
    param([object[]]$Rows)
    $xml = foreach ($row in $Rows) {
        '<row>' + (-join @(foreach ($cell in $row) { $ref, $text = $cell -split '=', 2
            if ($text -match '^\d+$') { "<c r=`"$ref`"><v>$text</v></c>" } else { "<c r=`"$ref`" t=`"inlineStr`"><is><t>$text</t></is></c>" } })) + '</row>'
    }
    return (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = '<worksheet><sheetData>' + (-join $xml) + '</sheetData></worksheet>' }) -FileName 'r.xlsx').Text
}
$otherTen = '98765' + '43210'   # a 10-digit order number, no DoD ID
$roster = Get-RosterXlsxText @(@('A1=Person', 'B1=SSN', 'C1=DoD ID', 'D1=Date of birth'), @('A2=Person A', "B2=$($Sample.Ssn)", "C2=$($Sample.DodId)", 'D2=03/14/1985'))
$expected = "Person | SSN | DoD ID | Date of birth`nPerson A | SSN: $($Sample.Ssn) | DoD ID: $($Sample.DodId) | Date of birth: 03/14/1985"
Assert-Equal $expected $roster 'Roster: each value below a header cell gets its label'
Assert-Equal 'SSN/High/***-**-6789 DoD ID/High/******7890 Date of birth/High/[date masked]' ((Get-HitSummary $roster) -join ' ') 'Roster: SSN, DoD ID and DOB found by their header'
$gap = Get-RosterXlsxText @(@('A1=Person', 'B1=SSN', 'C1=DoD ID'), @('A2=Person B', "C2=$($Sample.DodId)"))
Assert-Equal "Person B | DoD ID: $($Sample.DodId)" ($gap -split "`n")[1] 'Roster: the column comes from the cell reference, not the place (B2 is empty)'
$variants = Get-RosterXlsxText @(@('A1=ssan:', 'B1=EDIPI #', 'C1=DOB.', 'D1=Passport No.'), @("A2=$($Sample.SsnBare)", "B2=$($Sample.DodId)", 'C2=1985-03-14', "D2=$($Sample.Passport)"))
$expected = "SSN: $($Sample.SsnBare) | DoD ID: $($Sample.DodId) | Date of birth: 1985-03-14 | Passport: $($Sample.Passport)"
Assert-Equal $expected ($variants -split "`n")[1] 'Roster: header cells in other forms (case, :, #, .)'
$numbered = Get-RosterXlsxText @(@('A1=No.', 'B1=SSN'), @('A2=1', "B2=$($Sample.Ssn)"), @('A3=2', "B3=$($Sample.Ssn)"))
Assert-Equal "No. | SSN`n1 | SSN: $($Sample.Ssn)`n2 | SSN: $($Sample.Ssn)" $numbered 'Roster: a row number is not moved (no spaces added)'
# A label never reaches a value in the next cell or the next row (the detectors read 40 characters before a value).
$reach = Get-RosterXlsxText @(@('A1=Ref', 'B1=DoD ID', 'C1=Order'), @('A2=(a note row, so that the header text is out of reach)'),
    @("A3=$otherTen", "B3=$($Sample.DodId)", "C3=$otherTen"), @("A4=$otherTen"))
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary $reach) 'Roster: the DoD ID label does not label the order numbers in the next cell and the next row'
$own = Get-RosterXlsxText @(@('A1=DoD ID', 'B1=Kind', 'C1=Value'), @("A2=$($Sample.DodId)", 'B2=SSN', "C2=$($Sample.SsnBare)"))
Assert-Equal 'DoD ID/High/******7890 SSN/High/***-**-6789' ((Get-HitSummary $own) -join ' ') 'Roster: a label of its own in a cell stays next to its value'
# Regression (an earlier attempt): "SSN | pending" is a header row, but only cells with a digit get a label, and no gap moves a label.
$perRow = Get-RosterXlsxText @(@('A1=SSN', 'B1=pending'), @('A2=DoD ID', "B2=$($Sample.DodId)"))
Assert-Equal "SSN | pending`nDoD ID | $($Sample.DodId)" $perRow 'Roster: a per-row label layout under "SSN | pending" keeps its text'
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary $perRow) 'Roster: the per-row DoD ID under "SSN | pending" is still found'
$late = Get-RosterXlsxText @(@('A1=Person'), @('A2=a'), @('A3=b'), @('A4=c'), @('A5=d'), @('A6=DoD ID'), @("A7=$($Sample.DodId)"))
Assert-Equal $false $late.Contains('DoD ID: ') 'Roster: a header below the first 5 rows is not taken'
# .csv and .tsv: fields by position, quotes and line ends kept; a file without a header row is not changed.
$csv = "Person,SSN,DoD ID,DOB`r`nPerson A,`"$($Sample.Ssn)`",$($Sample.DodId),1985-03-14`r`n"
$csvText = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($csv)) -FileName 'r.csv').Text
Assert-Equal "Person,SSN,DoD ID,DOB`r`nPerson A,SSN: `"$($Sample.Ssn)`",DoD ID: $($Sample.DodId),Date of birth: 1985-03-14`r`n" $csvText 'Roster csv: labels written in; quotes and line ends kept'
Assert-Equal 'SSN/High/***-**-6789 DoD ID/High/******7890 Date of birth/High/[date masked]' ((Get-HitSummary $csvText) -join ' ') 'Roster csv: values found by their header'
$tsv = "Name`tPassport`nPerson A`t$($Sample.Passport)"
Assert-Equal "Name`tPassport`nPerson A`tPassport: $($Sample.Passport)" (Get-TextFromBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($tsv)) -FileName 'r.tsv').Text 'Roster tsv: tab fields'
$csvReach = "Ref,DoD ID,Order`n(a note row, so that the header text is out of reach)`n$otherTen,$($Sample.DodId),$otherTen`n$otherTen`n"
$csvReachText = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($csvReach)) -FileName 'r.csv').Text
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary $csvReachText) 'Roster csv: the DoD ID label does not reach the next field or the next line'
$csvDate = "Name,Date of birth,Suspense`nPerson A,03/14/1985,5 Mar 27`n"   # a date with 3 digits: the fewest a DOB has
Assert-Equal 1 @(Find-PiiInText (Get-TextFromBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($csvDate)) -FileName 'd.csv').Text).Count 'Roster csv: the DOB label does not reach a suspense date'
$plainCsv = "Person,Code,Number`r`nPerson A,`"$($Sample.Ssn)`",$otherTen`r`n"
Assert-Equal $plainCsv (Get-TextFromBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($plainCsv)) -FileName 'p.csv').Text 'Roster csv: no header row, text not changed'

# ---- Word field codes are not read (a TOC bookmark number is no SSN) ----
$fieldXml = "<w:document><w:p><w:r><w:instrText xml:space=`"preserve`"> PAGEREF _Toc$($Sample.SsnBare) \h </w:instrText></w:r><w:r><w:t>Chapter 1</w:t></w:r></w:p></w:document>"
$fieldText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'word/document.xml' = $fieldXml }) -FileName 'Plan.docx').Text
Assert-Equal 'Chapter 1' $fieldText.Trim() 'docx field code text dropped, page text kept'
$selfClosingField = "<w:document><w:p><w:r><w:instrText/></w:r><w:r><w:t>SSN $($Sample.Ssn)</w:t></w:r><w:r><w:instrText> PAGE </w:instrText></w:r></w:p></w:document>"
$kept = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'word/document.xml' = $selfClosingField }) -FileName 'a.docx').Text
Assert-Equal 'SSN' (@(Find-PiiInText $kept) | Select-Object -First 1).Detector 'An empty field code does not hide the page text after it'

# ---- PowerPoint ----
$pptx = New-ZipBytes @{ 'ppt/slides/slide1.xml' = '<p:sld><a:p><a:r><a:t>DOB: 03/14/1985</a:t></a:r></a:p></p:sld>' }
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary (Get-TextFromBytes -Bytes $pptx -FileName 'Brief.pptx').Text) 'pptx read'

# ---- Text, HTML, and types that are not read ----
$Txt = [System.Text.Encoding]::UTF8.GetBytes("my cell $($Sample.Phone)")
Assert-Equal 'Phone/Medium/***-***-0142' (Get-HitSummary (Get-TextFromBytes -Bytes $Txt -FileName 'notes.txt').Text) 'txt read'
# ---- PDF: page text (compressed or not, with nested brackets), ToUnicode fonts, form fields, reasons ----
$pdf = New-PdfBytes "BT /F1 12 Tf 72 700 Td (SSN: $($Sample.Ssn)) Tj 0 -14 Td [(my cell \(home\) ) -300 ($($Sample.Phone))] TJ ET"
$pdfText = Get-TextFromBytes -Bytes $pdf -FileName 'Recall.pdf'
Assert-Equal 'OK' $pdfText.Status 'PDF: read'
Assert-Equal 'SSN' (@(Find-PiiInText $pdfText.Text) | Select-Object -First 1).Detector 'PDF: SSN found in compressed page text'
Assert-Equal $true ($pdfText.Text -match 'my cell \(home\)') 'PDF: escaped brackets undone'
$plainPdf = New-PdfBytes "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $plainPdf -FileName 'a.pdf').Text) | Select-Object -First 1).Detector 'PDF: uncompressed page text read'
# A Type0 font with 2-byte glyph codes: the text is only readable through its ToUnicode map.
$cmap = "begincmap 1 begincodespacerange <0000> <FFFF> endcodespacerange 1 beginbfrange <0020> <007E> <0020> endbfrange endcmap"
$glyphs = -join @("SSN: $($Sample.Ssn)".ToCharArray() | ForEach-Object { ([int]$_).ToString('X4') })
$unicodePdf = New-PdfBytes "BT /F1 12 Tf <$glyphs> Tj ET" -ToUnicode $cmap
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $unicodePdf -FileName 'u.pdf').Text) | Select-Object -First 1).Detector 'PDF: ToUnicode font decoded'
$formPdf = New-PdfBytes 'q Q' -Extra @("<< /FT /Tx /T (SSN) /V ($($Sample.Ssn)) >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $formPdf -FileName 'DD Form.pdf').Text) | Select-Object -First 1).Detector 'PDF: form field value found with its field name'
Assert-Equal 'No text in this PDF (a scanned image?): check by hand' (Get-TextFromBytes -Bytes (New-PdfBytes 'q 100 0 0 100 0 0 cm /Im1 Do Q') -FileName 's.pdf').Reason 'PDF: an image-only page named'
Assert-Equal $true ((Get-TextFromBytes -Bytes (New-PdfBytes 'BT (x) Tj ET' -Trailer '/Encrypt 9 0 R') -FileName 'e.pdf').Reason -match '^Encrypted PDF') 'PDF: encrypted named'
Assert-Equal $true ((Get-TextFromBytes -Bytes $Txt -FileName 'scan.pdf').Reason -match '^Not a valid PDF') 'PDF: not a PDF named'
$a85 = & (Get-Module PdfCore) { ConvertFrom-Ascii85 '9jqo^BlbD-BleB1DJ+*+F(f,q~>' }
Assert-Equal 'Man is distinguished' $a85 'PDF: ASCII85 filter decoded (a published test vector)'
$hexData = & (Get-Module PdfCore) { Read-PdfStream '/Filter /ASCIIHexDecode' '48 49>' }
Assert-Equal 'HI' $hexData 'PDF: ASCIIHex filter decoded'
# Object stream (PDF 1.5): a form field packed inside one is still read.
$packed = "9 0 << /T (SSN) /V ($($Sample.Ssn)) >>"
$objStm = New-PdfBytes 'q Q' -Extra @("<< /Type /ObjStm /N 1 /First 4 /Length $($packed.Length) >> stream`n$packed`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $objStm -FileName 'p.pdf').Text) | Select-Object -First 1).Detector 'PDF: object stream unpacked'
# A stream whose first line ends in CR only.
$latin = [System.Text.Encoding]::GetEncoding(28591)
$crPdf = $latin.GetBytes(($latin.GetString((New-PdfBytes "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain)) -replace "stream`n", "stream`r"))
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $crPdf -FileName 'cr.pdf').Text) | Select-Object -First 1).Detector 'PDF: stream line ending in CR read'
# An inline image and a Tj on a number: no error, the text after it still read.
$inline = Get-TextFromBytes -Bytes (New-PdfBytes "BI /W 1 /H 1 ID 7 Tj EI BT 7 Tj (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'i.pdf'
Assert-Equal 'OK|SSN' ($inline.Status + '|' + (@(Find-PiiInText $inline.Text) | Select-Object -First 1).Detector) 'PDF: inline image and Tj on a number skipped'
# Hostile input stays inside the time limit (lowered to 2 s here): unbalanced brackets, and headers without endobj.
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 2 }
$clock = [Diagnostics.Stopwatch]::StartNew()
$slow = Get-TextFromBytes -Bytes (New-PdfBytes ('BT ' + ('(a' * 20000) + ' Tj ET') -Plain) -FileName 'h.pdf'
$noEnd = Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes('%PDF-1.4 ' + ('1 0 obj << >> ' * 50000))) -FileName 'n.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15) 'PDF: hostile input ends inside the time limit'
Assert-Equal 'NOT SCANNED|NOT SCANNED' ($slow.Status + '|' + $noEnd.Status) 'PDF: hostile input reported, not read'
# Hostile font maps stay inside the limit: an unclosed bfchar block repeated, and many full ranges (looked up, not expanded).
$clock.Restart()
$badMap = Get-TextFromBytes -Bytes (New-PdfBytes 'BT /F1 12 Tf <0041> Tj ET' -ToUnicode ('beginbfchar ' * 100000)) -FileName 'm.pdf'
$ranges = ('1 beginbfrange <0000> <FFFE> <0000> endbfrange ' * 20)
$manyRanges = Get-TextFromBytes -Bytes (New-PdfBytes ("BT /F1 12 Tf <$glyphs> Tj ET") -ToUnicode $ranges) -FileName 'r.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15) 'PDF: hostile font maps end inside the time limit'
Assert-Equal $true ($badMap.Status -in @('OK', 'NOT SCANNED')) 'PDF: a damaged map gives a result, not a stall'
Assert-Equal 'SSN' (@(Find-PiiInText $manyRanges.Text) | Select-Object -First 1).Detector 'PDF: full ranges still decode'
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 30 }
# A long string is read; one over the string limit makes the PDF NOT SCANNED, never read in part without a word.
$longString = Get-TextFromBytes -Bytes (New-PdfBytes ('BT (' + ('\101' * 1100) + " $($Sample.Ssn)) Tj ET") -Plain) -FileName 'l.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $longString.Text) | Select-Object -First 1).Detector 'PDF: a string with 1,100 escapes read'
$tooLong = Get-TextFromBytes -Bytes (New-PdfBytes ('BT (hello) Tj (' + ('\101' * 6000) + " $($Sample.Ssn)) Tj ET") -Plain) -FileName 't.pdf'
Assert-Equal $true ($tooLong.Status -eq 'NOT SCANNED' -and $tooLong.Reason -match '^Part of this PDF could not be read') 'PDF: a string over the limit reported, not dropped'
$remarks = Get-TextFromBytes -Bytes (New-PdfBytes 'q Q' -Extra @("<< /T (Remarks) /V (\376\377" + ('\000a' * 1100) + ") >>") -Plain) -FileName 'f.pdf'
Assert-Equal $true ($remarks.Status -eq 'OK' -and $remarks.Text.Contains('a' * 1100)) 'PDF: a long UTF-16 form value read'
# Inline-image bytes cannot hide page text; a bracket in a comment is no unread string.
$imageGap = Get-TextFromBytes -Bytes (New-PdfBytes "BI /W 1 ID xx(yy EI BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET BI /W 1 ID zz)ww EI BT (After) Tj ET" -Plain) -FileName 'g.pdf'
Assert-Equal 'OK|SSN' ($imageGap.Status + '|' + (@(Find-PiiInText $imageGap.Text) | Select-Object -First 1).Detector) 'PDF: inline-image bytes do not hide page text'
$commented = Get-TextFromBytes -Bytes (New-PdfBytes "% draft (rev`nBT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'c.pdf'
Assert-Equal 'OK|SSN' ($commented.Status + '|' + (@(Find-PiiInText $commented.Text) | Select-Object -First 1).Detector) 'PDF: a bracket in a comment is fine'
# Text in a form XObject (a reusable page part drawn with Do) is read.
$xobject = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$formPdf2 = New-PdfBytes '/X1 Do' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Length $($xobject.Length) >> stream`n$xobject`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $formPdf2 -FileName 'x.pdf').Text) | Select-Object -First 1).Detector 'PDF: form XObject text read'
# A part not read is never silent: an attached file, or a stream over the unpack cap, gives a note.
$attached = Get-TextFromBytes -Bytes (New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /EmbeddedFile /Length 3 >> stream`nabc`nendstream")) -FileName 'a.pdf'
Assert-Equal $true ($attached.Status -eq 'OK' -and $attached.Note -match 'attached inside this PDF') 'PDF: attachments give a check-by-hand note'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 10 }
$capped = Get-TextFromBytes -Bytes (New-PdfBytes ("BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET " * 3000)) -FileName 'cap.pdf'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 64MB }
# The cut can end inside a string (then the whole PDF is Not scanned) or between tokens (a note): never silent.
Assert-Equal $true (($capped.Note + $capped.Reason) -match 'too large to unpack|could not be read') 'PDF: text past the unpack cap is never silent'
# A stream that inflates far past the cap is cut, not held whole in memory.
$bomb = & (Get-Module PdfCore) {
    $memory = New-Object System.IO.MemoryStream
    $deflate = New-Object System.IO.Compression.DeflateStream($memory, [System.IO.Compression.CompressionMode]::Compress, $true)
    $zeros = New-Object byte[] (1MB); for ($i = 0; $i -lt 40; $i++) { $deflate.Write($zeros, 0, $zeros.Length) }
    $deflate.Dispose()
    $script:PdfBudget = @{ Clock = $null; Inflated = 0L }
    (Expand-PdfStream ($script:Latin1.GetString($memory.ToArray()))).Length
}
Assert-Equal $true ($bomb -le 16MB + 65536) 'PDF: one stream inflates to at most 16 MB'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $Txt -FileName 'setup.ps1').Status 'script not read'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $Txt -FileName 'broken.docx').Status 'damaged docx not scanned'
# Bytes that are not a zip get a reason a person can act on, not .NET's "End of central directory not found".
$ole = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0)
Assert-Equal $true ((Get-TextFromBytes -Bytes $ole -FileName 'Recall.xlsx').Reason -match '^An old Office file') 'OLE bytes without encryption: an old format with a new name'
Assert-Equal 'Empty file (0 bytes)' (Get-TextFromBytes -Bytes ([byte[]]@()) -FileName 'New.xlsx').Reason 'Empty file named'
$page = [System.Text.Encoding]::ASCII.GetBytes('  <!DOCTYPE html><html>sign in</html>')
Assert-Equal $true ((Get-TextFromBytes -Bytes $page -FileName 'a.xlsx').Reason -match 'web page') 'A web page instead of the file named'
$bomPage = [byte[]](@(0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::ASCII.GetBytes('<html>sign in</html>'))
Assert-Equal $true ((Get-TextFromBytes -Bytes $bomPage -FileName 'a.xlsx').Reason -match 'web page') 'A web page with a byte order mark named'
Assert-Equal $true ((Get-TextFromBytes -Bytes $Txt -FileName 'broken.docx').Reason -match '^Not a valid Office file') 'Other bytes: not a valid Office file'
$html = (Convert-HtmlToText "<div><p>SSN $($Sample.Ssn)</p></div>here") -replace '\s+', ' '
Assert-Equal "SSN $($Sample.Ssn) here" $html.Trim() 'HTML tags removed'

# ==== Round 3 readers (the PAM, 2026-11-20): RTF, OpenDocument, Visio, old Office and Outlook, zip, e-mail, images ====
$ssnLine = "SSN: $($Sample.Ssn)"
$rtf = [System.Text.Encoding]::ASCII.GetBytes('{\rtf1\ansi{\fonttbl{\f0 Arial;}}{\*\generator Word}\f0 ' + $ssnLine + '\par caf\''e9 ' + [char]92 + 'u8364?}')
$rtfText = (Get-TextFromBytes -Bytes $rtf -FileName 'memo.rtf').Text
Assert-Equal 'SSN' (@(Find-PiiInText $rtfText) | Select-Object -First 1).Detector 'RTF: text read, tables dropped'
Assert-Equal $true ($rtfText.Contains([string][char]0xE9) -and $rtfText.Contains([string][char]0x20AC) -and -not $rtfText.Contains('Arial')) 'RTF: escapes decoded, font table dropped'
$odt = New-ZipBytes @{ 'content.xml' = "<office:document-content><text:p>$ssnLine</text:p></office:document-content>" }
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $odt -FileName 'a.odt').Text) | Select-Object -First 1).Detector 'ODT: text read'
$odsRow = '<table:table-row><table:table-cell><text:p>123</text:p></table:table-cell><table:table-cell><text:p>45</text:p></table:table-cell>' +
    '<table:table-cell><text:p>6789</text:p></table:table-cell></table:table-row>'
$ods = New-ZipBytes @{ 'content.xml' = $odsRow }
Assert-Equal 0 @(Find-PiiInText (Get-TextFromBytes -Bytes $ods -FileName 'a.ods').Text | Where-Object { $_.Detector -eq 'SSN' }).Count 'ODS: neighbouring cells never join into one SSN'
$vsdx = New-ZipBytes @{ 'visio/pages/page1.xml' = "<PageContents><Shapes><Shape><Text>$ssnLine</Text></Shape></Shapes></PageContents>" }
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $vsdx -FileName 'org.vsdx').Text) | Select-Object -First 1).Detector 'Visio: shape text read'

# Old Office and Outlook (OLE): text runs; encryption named; an .xls gets its note.
$oleHead = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1) + (New-Object byte[] 504)
$oleDoc = [byte[]]($oleHead + [System.Text.Encoding]::Unicode.GetBytes("Memo $ssnLine end") + (New-Object byte[] 64))
$docRead = Get-TextFromBytes -Bytes $oleDoc -FileName 'old.doc'
Assert-Equal 'OK|SSN|' ($docRead.Status + '|' + (@(Find-PiiInText $docRead.Text) | Select-Object -First 1).Detector + '|' + $docRead.Note) 'Old .doc: UTF-16 text read, no note'
Assert-Equal $true ((Get-TextFromBytes -Bytes $oleDoc -FileName 'old.xls').Note -match '^Old Excel file') 'Old .xls: note that numbers are not read'
$oleLocked = [byte[]]($oleHead + [System.Text.Encoding]::Unicode.GetBytes('EncryptedPackage') + (New-Object byte[] 64))
Assert-Equal $true ((Get-TextFromBytes -Bytes $oleLocked -FileName 'locked.doc').Reason -match '^Encrypted \(password') 'Old Office: encryption named'
$msgRead = Get-TextFromBytes -Bytes ([byte[]]($oleHead + [System.Text.Encoding]::ASCII.GetBytes("Subject: roster $ssnLine"))) -FileName 'mail.msg'
Assert-Equal 'SSN' (@(Find-PiiInText $msgRead.Text) | Select-Object -First 1).Detector 'Outlook .msg: 8-bit text read'

# Zip: each file read with its reader and named; a zip three deep is not opened; the entry limit gives a reason.
$innerZip = New-ZipBytes @{ 'deep.txt' = $ssnLine }
$middleZip = New-ZipBytes @{ 'inner.zip' = $innerZip }
$pack = New-ZipBytes @{ 'notes/a.txt' = $ssnLine; 'doc.odt' = $odt; 'middle.zip' = $middleZip }
$packRead = Get-TextFromBytes -Bytes $pack -FileName 'pack.zip'
$names = @($packRead.Parts | ForEach-Object { $_.Name + '=' + $_.Status }) | Sort-Object
Assert-Equal 'doc.odt=OK|middle.zip/inner.zip=NOT SCANNED|middle.zip=OK|notes/a.txt=OK' ($names -join '|') 'Zip: inner files read and named; a zip three deep not opened'
$odtPart = @($packRead.Parts | Where-Object { $_.Name -eq 'doc.odt' })[0]
Assert-Equal 'SSN' (@(Find-PiiInText $odtPart.Text) | Select-Object -First 1).Detector 'Zip: an ODT inside read'
& (Get-Module ContainerText) { $script:MaxArchiveEntries = 1 }
$limited = Get-TextFromBytes -Bytes $pack -FileName 'pack.zip'
& (Get-Module ContainerText) { $script:MaxArchiveEntries = 2000 }
Assert-Equal 1 @($limited.Parts | Where-Object { $_.Reason -match 'More than 1 files' }).Count 'Zip: past the file limit, the rest gets a reason'

# E-mail (.eml): headers (encoded words), a quoted-printable text body and a base64 attachment.
$attachment = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("my cell $($Sample.Phone)"))
$subject = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('Recall roster'))
$eml = "Subject: =?utf-8?B?$($subject)?=`r`nFrom: $($Sample.Email)`r`nContent-Type: multipart/mixed; boundary=`"XYZ`"`r`n`r`n" +
    "--XYZ`r`nContent-Type: text/plain; charset=utf-8`r`nContent-Transfer-Encoding: quoted-printable`r`n`r`n" +
    "SSN: $($Sample.Ssn.Substring(0, 3))=`r`n$($Sample.Ssn.Substring(3))`r`n" +
    "--XYZ`r`nContent-Type: text/plain; name=`"notes.txt`"`r`nContent-Disposition: attachment; filename=`"notes.txt`"`r`n" +
    "Content-Transfer-Encoding: base64`r`n`r`n$attachment`r`n--XYZ--`r`n"
$mail = Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($eml)) -FileName 'recall.eml'
$byName = @{}; foreach ($part in $mail.Parts) { $byName[$part.Name] = $part.Text }
Assert-Equal $true ($byName['(message headers)'] -match 'Subject: Recall roster') 'EML: encoded subject decoded'
Assert-Equal 'SSN' (@(Find-PiiInText $byName['(message text)']) | Select-Object -First 1).Detector 'EML: quoted-printable body read'
Assert-Equal 'Phone' (@(Find-PiiInText $byName['notes.txt']) | Select-Object -First 1).Detector 'EML: base64 attachment read'

# Images: without Windows OCR (this test machine) a reason; with OCR (a stand-in here) the text is read, also from a scanned PDF.
# Bytes that are no image: a reason either way (no OCR here; or, where Windows has OCR, the image cannot be opened).
$unreadable = '^(Image: Windows OCR is not available|Image: its text could not be read)'
Assert-Equal $true ((Get-TextFromBytes -Bytes ([byte[]](1, 2, 3)) -FileName 'scan.png').Reason -match $unreadable) 'Image: no OCR or no image gives a reason'
$scanPdf = New-PdfBytes 'q /Im1 Do Q' -Plain -Extra @("<< /Type /XObject /Subtype /Image /Filter /DCTDecode /Length 4 >> stream`nJPEG`nendstream")
Assert-Equal $true ((Get-TextFromBytes -Bytes $scanPdf -FileName 's.pdf').Reason -match ('^Scanned PDF: ' + $unreadable.Substring(1))) 'Scanned PDF: no OCR or a damaged page image gives a reason'
# Where Windows has OCR (the ps51 CI runner): a real image with an invented SSN drawn in it is read and the SSN found.
$ocrHere = (Get-TextFromBytes -Bytes ([byte[]](1, 2, 3)) -FileName 'x.png').Reason -notmatch 'not available'
if ($ocrHere -and $PSVersionTable.PSVersion.Major -eq 5) {
    Add-Type -AssemblyName System.Drawing   # tests only: draws the sample image (the bundle never loads it)
    # Types looked up by name: written as [System.Drawing...] they would be resolved, and fail, where it is missing.
    $drawing = @{}; foreach ($name in 'Graphics', 'Color', 'Brushes', 'Imaging.ImageFormat') { $drawing[$name] = ('System.Drawing.' + $name) -as [type] }
    $bitmap = New-Object System.Drawing.Bitmap 900, 120
    $graphics = $drawing['Graphics']::FromImage($bitmap)
    $graphics.Clear($drawing['Color']::White)
    $graphics.DrawString($ssnLine, (New-Object System.Drawing.Font 'Arial', 36), $drawing['Brushes']::Black, 10, 30)
    $png = New-Object System.IO.MemoryStream
    $bitmap.Save($png, $drawing['Imaging.ImageFormat']::Png); $graphics.Dispose(); $bitmap.Dispose()
    $ocrRead = Get-TextFromBytes -Bytes $png.ToArray() -FileName 'roster.png'
    Assert-Equal $true (@(Find-PiiInText $ocrRead.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'Image: Windows OCR reads an SSN drawn in a picture'
}
& (Get-Module OcrText) { param($line) $script:OcrOverride = [scriptblock]::Create("param(`$b) @{ Text = '$line'; Reason = ''; Note = '' }") } $ssnLine
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes ([byte[]](1, 2, 3)) -FileName 'scan.png').Text) | Select-Object -First 1).Detector 'Image: OCR text checked'
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $scanPdf -FileName 's.pdf').Text) | Select-Object -First 1).Detector 'Scanned PDF: page images read with OCR'
& (Get-Module OcrText) { $script:OcrOverride = $null }
