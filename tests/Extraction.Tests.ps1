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
Assert-Equal $true ((Get-TextFromBytes -Bytes $ole -FileName 'Recall.xlsx').Reason -match '^Encrypted, password-protected') 'OLE bytes: encrypted or labelled'
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
Assert-Equal $true ((Get-TextFromBytes -Bytes $oleLocked -FileName 'locked.doc').Reason -match '^Encrypted') 'Old Office: encryption named'
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
# ---- PDF review round 5 (2026-11-21): no silent loss, bounded time ----
# The word BI in page text is text, not an inline image (an image starts only at a BI operator).
$biText = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Clearance BI) Tj (ID $($Sample.Ssn)) Tj ( EI) Tj ET" -Plain) -FileName 'bi.pdf'
Assert-Equal 'OK|SSN' ($biText.Status + '|' + (@(Find-PiiInText ("SSN " + $biText.Text)) | Select-Object -First 1).Detector) 'PDF: the word BI does not hide text'
$biPage = Get-TextFromBytes -Bytes (New-PdfBytes ('BT /F1 12 Tf ' + ('(Status: BI complete for member) Tj 0 -14 Td ' * 3000) + 'ET') -Plain) -FileName 'bi2.pdf'
Assert-Equal 'OK' $biPage.Status 'PDF: a long page with the word BI is read in time'
# Comments and sticky notes, a value in an object of its own, and a content list one level deeper.
$annot = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /Annot /Subtype /Text /Contents (SSN $($Sample.Ssn)) >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $annot -FileName 'n.pdf').Text) | Select-Object -First 1).Detector 'PDF: a sticky-note comment read'
$valueRef = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V 8 0 R >>', "($($Sample.Ssn))")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $valueRef -FileName 'v.pdf').Text) | Select-Object -First 1).Detector 'PDF: a form value in an object of its own read'
$inner = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$nested = New-PdfBytes 'q Q' -Plain -PageContents '7 0 R' -Extra @('[8 0 R]', "<< /Length $($inner.Length) >> stream`n$inner`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $nested -FileName 'c.pdf').Text) | Select-Object -First 1).Detector 'PDF: /Contents through an array object read'
# LZW and RunLength streams decoded; any other filter gives a note.
$latin = [System.Text.Encoding]::GetEncoding(28591)
$lzwBits = New-Object System.Text.StringBuilder
foreach ($code in @(256) + @([System.Text.Encoding]::ASCII.GetBytes($inner) | ForEach-Object { [int]$_ }) + @(257)) { [void]$lzwBits.Append([Convert]::ToString($code, 2).PadLeft(9, '0')) }
while ($lzwBits.Length % 8) { [void]$lzwBits.Append('0') }
$lzw = -join @(for ($i = 0; $i -lt $lzwBits.Length; $i += 8) { [char][Convert]::ToInt32($lzwBits.ToString().Substring($i, 8), 2) })
$lzwPdf = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /LZWDecode /Length $($lzw.Length) >> stream`n$lzw`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $lzwPdf -FileName 'l.pdf').Text) | Select-Object -First 1).Detector 'PDF: LZW stream decoded'
$rle = [string][char]($inner.Length - 1) + $inner + [char]128
$rlePdf = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /RunLengthDecode /Length $($rle.Length) >> stream`n$rle`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes ($latin.GetBytes($latin.GetString($rlePdf))) -FileName 'r.pdf').Text) | Select-Object -First 1).Detector 'PDF: RunLength stream decoded'
$oddFilter = Get-TextFromBytes -Bytes (New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /JPXDecode /Length 3 >> stream`nabc`nendstream")) -FileName 'o.pdf'
Assert-Equal $true ($oddFilter.Note -match 'compression this reader cannot open') 'PDF: an unknown filter gives a note'
# A hostile object stream (a huge object list) ends at once with a note.
$clock.Restart()
$hugeList = New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /ObjStm /N 200000 /First 4 /Length 8 >> stream`n9 0 (x)`nendstream")
$hugeRead = Get-TextFromBytes -Bytes $hugeList -FileName 'h.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 10 -and $hugeRead.Note -match 'too large to unpack') 'PDF: a huge object list ends at once with a note'

# ---- PDF review round 6 (2026-11-21) ----
# LZW and RunLength output counts toward the unpack total: past it, a note (bounded memory).
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 10 }
$lzwCapped = Get-TextFromBytes -Bytes $lzwPdf -FileName 'l2.pdf'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 64MB }
Assert-Equal $true (($lzwCapped.Note + $lzwCapped.Reason) -match 'too large to unpack|could not be read') 'PDF: LZW output counts toward the unpack total'
# A composite font without a ToUnicode map: a reason or note, never glyph numbers read as text.
$type0 = New-PdfBytes 'BT /F1 12 Tf <001400150016> Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H >>'
$type0Read = Get-TextFromBytes -Bytes $type0 -FileName 't0.pdf'
Assert-Equal $true (($type0Read.Note + $type0Read.Reason) -match 'font in this PDF could not be read') 'PDF: a font without a map is never silent'
# An incremental update: the field defined again later with its value; the old and the new versions are both read.
$filled = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V () >>') -After "7 0 obj << /FT /Tx /T (SSN) /V ($($Sample.Ssn)) >> endobj`n"
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $filled -FileName 'u.pdf').Text) | Select-Object -First 1).Detector 'PDF: a value saved as an added edit is read'
# A multi-select value and a value in a text stream; a backslash before a line end inside a string.
$multi = New-PdfBytes 'q Q' -Plain -Extra @("<< /FT /Ch /T (SSN) /V [(none) ($($Sample.Ssn))] >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $multi -FileName 'm.pdf').Text) | Select-Object -First 1).Detector 'PDF: a multi-select value read'
$vStream = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V 8 0 R >>', "<< /Length 11 >> stream`n$($Sample.Ssn)`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $vStream -FileName 's2.pdf').Text) | Select-Object -First 1).Detector 'PDF: a value in a text stream read'
$wrapped = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (SSN: \`n$($Sample.Ssn)) Tj ET" -Plain) -FileName 'w.pdf'
Assert-Equal 'OK|SSN' ($wrapped.Status + '|' + (@(Find-PiiInText $wrapped.Text) | Select-Object -First 1).Detector) 'PDF: a backslash before a line end is a line continuation'

# ---- PDF review round 7 (2026-11-21) ----
# Text drawn one glyph at a time (Chrome): the glyphs join; a normal block keeps its word spaces.
# Each glyph moved by its real width at 12 pt, as a PDF writer does (Helvetica: S 8.0, N 8.7, digits 6.7, - 4.0, : and space 3.3).
$advance = @{ 'S' = 8.0; 'N' = 8.7; '-' = 4.0; ':' = 3.3; ' ' = 3.3 }
$glyphs1 = (("SSN: $($Sample.Ssn)").ToCharArray() | ForEach-Object { $w = 6.7; if ($advance.ContainsKey([string]$_)) { $w = $advance[[string]$_] }; "($_) Tj $w 0 Td" }) -join ' '
$glyphRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $glyphs1 ET" -Plain) -FileName 'g1.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $glyphRead.Text) | Select-Object -First 1).Detector 'PDF: text drawn glyph by glyph joins'
$words = Get-TextFromBytes -Bytes (New-PdfBytes 'BT /F1 12 Tf (Member) Tj 60 0 Td (roster) Tj ET' -Plain) -FileName 'w2.pdf'
Assert-Equal 'Member roster' $words.Text.Trim() 'PDF: words placed apart keep their space'
# XFA form data; an older page version kept by an incremental update; two comments in one dictionary.
$xfa = "<xfa:datasets><xfa:data><form><SSN>$($Sample.Ssn)</SSN></form></xfa:data></xfa:datasets>"
$xfaPdf = New-PdfBytes 'BT (Please wait) Tj ET' -Plain -Extra @("<< /Fields [] /XFA 8 0 R >>", "<< /Length $($xfa.Length) >> stream`n$xfa`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $xfaPdf -FileName 'x2.pdf').Text) | Select-Object -First 1).Detector 'PDF: XFA form data read'
$oldPage = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$edited = New-PdfBytes $oldPage -Plain -After "4 0 obj << /Length 22 >> stream`nBT (SSN: removed) Tj ET`nendstream endobj`n"
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $edited -FileName 'e2.pdf').Text) | Select-Object -First 1).Detector 'PDF: text of an older page version read'
$twoNotes = New-PdfBytes 'q Q' -Plain -Extra @("<< /Annots [<< /Subtype /Text /Contents (first note) >> << /Subtype /Text /Contents (SSN $($Sample.Ssn)) >>] >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $twoNotes -FileName 'n2.pdf').Text) | Select-Object -First 1).Detector 'PDF: every comment in one dictionary read'
# A glyph mapped to a private-use character: a space, and a note.
$pua = New-PdfBytes 'BT /F1 12 Tf <0041E0010042> Tj ET' -Plain -ToUnicode 'begincmap 1 beginbfrange <0000> <FFFF> <0000> endbfrange endcmap'
Assert-Equal $true ((Get-TextFromBytes -Bytes $pua -FileName 'p.pdf').Note -match 'Some characters in this PDF could not be read') 'PDF: private-use characters give a note'

# ---- PDF review round 8 (2026-11-21) ----
# Chrome splits a run at a kerning pair (its own space glyphs inside the strings): the pieces join again.
$kerned = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 16 Tf (SSN: 21) Tj 38.2 0 Td (1-45-671) Tj 66.7 0 Td (1) Tj ET" -Plain) -FileName 'k.pdf'
Assert-Equal ('SSN: 21' + '1-45-671' + '1') $kerned.Text.Trim() 'PDF: pieces split at kerning pairs join again'   # built from parts
# Word's bullet (Symbol font, U+F0B7) is a plain character: no needless note.
$bullet = New-PdfBytes 'BT /F1 12 Tf <F0B70020004D> Tj ET' -Plain -ToUnicode 'begincmap 1 beginbfrange <0000> <FFFF> <0000> endbfrange endcmap'
$bulletRead = Get-TextFromBytes -Bytes $bullet -FileName 'b.pdf'
Assert-Equal "|$([char]0xB7) M" ($bulletRead.Note + '|' + $bulletRead.Text.Trim()) 'PDF: a Word bullet reads as a plain character, no note'

# ---- PDF review round 9 (2026-11-21): gaps decided per text block ----
# Places text glyph by glyph as a PDF writer does: each move is the glyph's real Helvetica width at 12 pt plus the
# letter spacing (in points).
function Get-HelveticaWidth {
    # A character's Helvetica width in thousandths of an em (case-sensitive: S and s differ); 556 for the rest.
    param([string]$Char)
    switch -CaseSensitive ($Char) {
        'S' { 667 } 'N' { 722 } 'T' { 611 } 'M' { 833 } 'm' { 833 } 'i' { 222 } 's' { 500 } ':' { 278 } ' ' { 278 } '-' { 333 }
        default { 556 }
    }
}
function Get-GlyphRun {
    # A content-stream fragment that shows each character with its own Tj and moves by its width plus Spacing.
    param([string]$Text, [double]$Spacing)
    (($Text.ToCharArray() | ForEach-Object {
        $width = Get-HelveticaWidth ([string]$_)
        '({0}) Tj {1} 0 Td' -f $_, ($width * 12 / 1000 + $Spacing).ToString([Globalization.CultureInfo]::InvariantCulture)
    }) -join ' ')
}
# Chrome glyph by glyph, wide letters and its own space glyphs: no false spaces inside words or numbers.
$wide = Get-GlyphRun 'Name: Tom Mims' 0
$wideRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $wide ET" -Plain) -FileName 'wd.pdf'
Assert-Equal 'Name: Tom Mims' $wideRead.Text.Trim() 'PDF: wide glyphs placed one by one keep their words'
# Letter spacing 0.3 em (Chrome) with the writer's own spaces: the SSN stays whole.
$spaced = Get-GlyphRun "SSN: $($Sample.Ssn)" 3.6   # letter spacing 0.3 em
$spacedRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $spaced ET" -Plain) -FileName 'ls.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $spacedRead.Text) | Select-Object -First 1).Detector 'PDF: letter-spaced text keeps the SSN whole'
# Expanded character spacing in TJ (LibreOffice) with space characters: no space between letters.
$expanded = '[' + ((("SSN: $($Sample.Ssn)").ToCharArray() | ForEach-Object { "($_) -270" }) -join ' ') + '] TJ'
$expandedRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $expanded ET" -Plain) -FileName 'ex.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $expandedRead.Text) | Select-Object -First 1).Detector 'PDF: expanded TJ spacing keeps the SSN whole'
# A table row in one block with spaces: a wide jump between cells is still a break.
$cellsRead = Get-TextFromBytes -Bytes (New-PdfBytes 'BT /F1 12 Tf (Last name) Tj 120 0 Td (SSN) Tj ET' -Plain) -FileName 'tb.pdf'
Assert-Equal 'Last name SSN' $cellsRead.Text.Trim() 'PDF: a wide jump between cells is a break'

# ---- PDF review round 10 (2026-11-21): real glyph widths ----
# A composite font's /W widths (first [w1 w2 ...] and first last w), read whole; glyphs placed by them join.
$wFont = '<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H /ToUnicode 6 0 R /DescendantFonts [9 0 R] >>'
$wCid = '<< /Type /Font /Subtype /CIDFontType2 /DW 1000 /W [65 [500.5 600] 67 70 250] >>'
$wMap = 'begincmap 1 beginbfrange <0000> <FFFF> <0000> endbfrange endcmap'
$wRun = '<0041> Tj 6.006 0 Td <0042> Tj 7.2 0 Td <0043> Tj 3 0 Td <0044> Tj'
$wRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $wRun ET" -Plain -FontDict $wFont -ToUnicode $wMap -Extra @('<< >>', '<< >>', $wCid)) -FileName 'w.pdf'
Assert-Equal 'ABCD' $wRead.Text.Trim() 'PDF: glyphs placed by their /W widths join'
# A label and its value in one text object, a pad apart (reportlab textOut + moveCursor): they stay apart.
# "Sponsor SSN" is 60.58 wide in Helvetica 10 pt; the value starts a 5 pt pad after it.
$padded = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 10 Tf (Sponsor SSN) Tj 65.58 0 Td ($($Sample.Ssn)) Tj ET" -Plain) -FileName 'pad.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $padded.Text) | Select-Object -First 1).Detector 'PDF: a label and a value a pad apart stay apart'
# ASCII85: z (four zero bytes) is lower case only; a capital Z is data.
Assert-Equal 'hello' (& (Get-Module PdfCore) { ConvertFrom-Ascii85 'BOu!rDZ~>' }) 'PDF: ASCII85 keeps a capital Z (hello encodes as BOu!rDZ)'

# ---- PDF review round 11 (2026-11-21) ----
# pdfTeX style: word gaps are TJ adjustments, with one-letter words among them: the words stay apart.
$tex = "[(This)-333(is)-333(a)-333(note)-333(a)-333(b)-333(a)-333(c)-333(a)] TJ 0 -14 Td [(Contact)-333(SSN)-333($($Sample.Ssn))] TJ"
$texRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 10 Tf $tex ET" -Plain) -FileName 'tex.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $texRead.Text) | Select-Object -First 1).Detector 'PDF: TJ word gaps with one-letter words stay spaces'
# A Type3 font: widths in units of its /FontMatrix (0.01 here: 50 = 0.5 em).
$t3Font = '<< /Type /Font /Subtype /Type3 /FontMatrix [0.01 0 0 0.01 0 0] /FirstChar 32 /LastChar 122 /Widths [' + ((1..91 | ForEach-Object { '50' }) -join ' ') + '] >>'
# SSN is 15 wide (3 x 0.5 em x 10 pt); the value starts 19 after it: a 0.4 em gap.
$t3Run = '(SSN) Tj 19 0 Td (' + $Sample.Ssn + ') Tj'
$t3Read = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 10 Tf $t3Run ET" -Plain -FontDict $t3Font) -FileName 't3.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $t3Read.Text) | Select-Object -First 1).Detector 'PDF: Type3 widths use the font matrix'
# A huge string of spaces: bounded time and runs of spaces collapsed.
$clock.Restart()
$manySpaces = Get-TextFromBytes -Bytes (New-PdfBytes ('BT /F1 10 Tf (' + (' ' * 500000) + 'SSN ' + $Sample.Ssn + ') Tj ET')) -FileName 'sp.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 30 -and $manySpaces.Text.Length -lt 100) 'PDF: a huge run of spaces is read in bounded time and collapsed'

# ---- PDF review round 12 (2026-11-21) ----
# A letter-spaced heading above body text in one text object: letter spacing is measured per line, so the body's
# word gaps (0.33 em: "Tom" is 24 wide at 12 pt, the next word starts 28 after it) stay spaces.
$heading = Get-GlyphRun 'RECALL ROSTER' 3.6
$bodyRun = '(Tom) Tj 28 0 Td (Mims) Tj'
$headRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $heading 0 -14 Td $bodyRun ET" -Plain) -FileName 'hd.pdf'
Assert-Equal $true ($headRead.Text -match 'Tom Mims') 'PDF: a letter-spaced heading does not join the words below it'
# A /W list longer than the entry cap: read up to the cap, with a note, in bounded time.
$clock.Restart()
$bigCid = '<< /Type /Font /Subtype /CIDFontType2 /W [0 69999 500 70000 [' + ((1..2000 | ForEach-Object { '500' }) -join ' ') + ']] >>'
$bigW = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf <0041> Tj ET" -Plain -FontDict $wFont -ToUnicode $wMap -Extra @('<< >>', '<< >>', $bigCid)) -FileName 'bw.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 30 -and $bigW.Note -match 'could not be fully read') 'PDF: a /W list past the cap gives a note in bounded time'
# A rotated Type3 font matrix (a = 0): the unit is the length of its first column (0.01 here).
$t3Turned = $t3Font.Replace('/FontMatrix [0.01 0 0 0.01 0 0]', '/FontMatrix [0 0.01 -0.01 0 0 0]')
$t3TurnedRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 10 Tf $t3Run ET" -Plain -FontDict $t3Turned) -FileName 't3r.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $t3TurnedRead.Text) | Select-Object -First 1).Detector 'PDF: a rotated Type3 font matrix still gives its unit'

# ---- PDF review round 13 (2026-11-21) ----
# Many text state operators between text objects stay inside the time limit (lowered to 2 s here).
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 2 }
$clock.Restart()
$manyTc = Get-TextFromBytes -Bytes (New-PdfBytes (("1 Tc`n" * 300000) + "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET")) -FileName 'tc.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15 -and $manyTc.Status -in @('OK', 'NOT SCANNED')) 'PDF: many state operators between text objects end inside the time limit'
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 30 }
# A damaged string between text objects cannot hide the text after it: NOT SCANNED, never OK without it.
$damaged = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Roster) Tj ET (ab % ( ) BT (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'dm.pdf'
Assert-Equal $true ($damaged.Status -eq 'NOT SCANNED' -or @(Find-PiiInText $damaged.Text).Count -gt 0) 'PDF: a damaged string between text objects is reported or read past'
# Text shown outside BT/ET (against the spec) is still read.
$loose = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Roster) Tj ET 0 0 1 rg (SSN: $($Sample.Ssn)) Tj" -Plain) -FileName 'lo.pdf'
Assert-Equal $true (@(Find-PiiInText $loose.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'PDF: text shown outside a text object is read'

# ---- PDF review round 14 (2026-11-21) ----
# Text after an ET inside an open TJ array (damaged) is still read, as before.
$openArray = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Roster) Tj ET BT /F1 12 Tf [(Name) ET (SSN: $($Sample.Ssn))] TJ ET" -Plain) -FileName 'oa.pdf'
Assert-Equal $true (@(Find-PiiInText $openArray.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'PDF: text after an ET inside an open array is read'
# Short letter-spaced lines (3 letter gaps each) take the block's letter spacing: no spaces inside the words.
$shortRuns = (Get-GlyphRun 'SSN:' 3.6) + ' 0 -14 Td ' + (Get-GlyphRun '1234' 3.6)
$shortRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $shortRuns ET" -Plain) -FileName 'sr.pdf'
Assert-Equal $true ($shortRead.Text -match 'SSN:' -and $shortRead.Text -match '1234') 'PDF: short letter-spaced lines keep their words'

# ---- PDF review round 15 (2026-11-21) ----
# An unclosed [ takes in the rest of the stream: its strings are still read, with a note.
$unclosed = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Roster) Tj ET BT /F1 12 Tf [(x) TJ ET BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'uc.pdf'
Assert-Equal $true ((@(Find-PiiInText $unclosed.Text | ForEach-Object { $_.Detector }) -contains 'SSN') -and $unclosed.Note -match 'unclosed') 'PDF: text after an unclosed [ is read, with a note'
# Two initials and a word gap under a letter-spaced heading: the word gap stays a space.
$initials = (Get-GlyphRun 'ROSTER LIST' 3.6) + " 0 -14 Td (Y) Tj 10.272 0 Td (N) Tj 12.264 0 Td ($($Sample.Ssn)) Tj"
$initialsRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $initials ET" -Plain) -FileName 'in.pdf'
Assert-Equal $true (@(Find-PiiInText $initialsRead.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'PDF: a word gap after initials under a spaced heading stays a space'
# ---- PDF review round 16 (2026-11-21) ----
# An unclosed [ before a later, closed one: its text is read, with a note; its lines stay apart.
$twoArrays = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf [(SSN: $($Sample.Ssn)) TJ ET BT /F1 12 Tf [(ok)] TJ ET" -Plain) -FileName 'ta.pdf'
$twoFound = @(Find-PiiInText $twoArrays.Text | ForEach-Object { $_.Detector }) -contains 'SSN'
Assert-Equal $true ($twoFound -and $twoArrays.Note -match 'unclosed') 'PDF: an unclosed [ before a closed one is read, with a note'
$moves = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf [(Name) TJ 0 -14 Td (SSN: $($Sample.Ssn)) Tj 0 -14 Td (2024) Tj ET" -Plain) -FileName 'mv.pdf'
Assert-Equal $true (@(Find-PiiInText $moves.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'PDF: moves after an unclosed [ still break lines'
# ---- PDF review round 17 (2026-11-21) ----
# An empty [ closed by an operator: every string that follows is read by the bare TJ.
$emptyOpen = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf [ 0 Tc (SSN: $($Sample.Ssn)) ( end)] TJ ET" -Plain) -FileName 'eo.pdf'
Assert-Equal $true (@(Find-PiiInText $emptyOpen.Text | ForEach-Object { $_.Detector }) -contains 'SSN') 'PDF: strings after an empty unclosed [ are all read'
# A marked-content property list with true in an array: well formed, so no note and no extra text.
$property = Get-TextFromBytes -Bytes (New-PdfBytes "/P << /A [(lbl) true] >> BDC BT /F1 12 Tf (Body) Tj ET EMC" -Plain) -FileName 'pl.pdf'
Assert-Equal 'Body|' ($property.Text.Trim() + '|' + $property.Note) 'PDF: true inside an array is a value, not an operator'
# ---- PDF review round 18 (2026-11-21) ----
# Stray strings or arrays next to a show operator's own operand: every one is read, with a note.
foreach ($stray in @("[(SSN: $($Sample.Ssn))] (x) TJ", "(SSN: $($Sample.Ssn)) [(x)] TJ", "[(SSN: $($Sample.Ssn))] [(x)] TJ", "(SSN: $($Sample.Ssn)) (x) Tj")) {
    $strayRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $stray ET" -Plain) -FileName 'st.pdf'
    $strayFound = @(Find-PiiInText $strayRead.Text | ForEach-Object { $_.Detector }) -contains 'SSN'
    Assert-Equal $true ($strayFound -and $strayRead.Note -match 'stray') ('PDF: stray operands read, with a note: ' + $stray.Substring($stray.Length - 2))
}
# ---- PDF review round 19 (2026-11-21) ----
# A TJ whose array is its last operand is the usual case (no note), also with a stray number before it.
$usualTj = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf 5 [(Stray) -50 (Num)] TJ ET" -Plain) -FileName 'ut.pdf'
Assert-Equal '' $usualTj.Note 'PDF: a TJ with its array last is the usual case'

# ---- Review of PR 19 (2026-11-21): no silent misses in containers ----
# An encrypted zip entry (flag bit 0 in the central directory) is named, never read as clean.
$lockedZip = New-ZipBytes @{ 'secret.txt' = $ssnLine }
for ($z = 0; $z + 3 -lt $lockedZip.Length; $z++) {
    if ($lockedZip[$z] -eq 0x50 -and $lockedZip[$z + 1] -eq 0x4B -and $lockedZip[$z + 2] -eq 1 -and $lockedZip[$z + 3] -eq 2) { $lockedZip[$z + 8] = $lockedZip[$z + 8] -bor 1 }
}
$lockedPart = (Get-TextFromBytes -Bytes $lockedZip -FileName 'l.zip').Parts.ToArray()[0]
Assert-Equal $true ($lockedPart.Reason -match '^Encrypted') 'Zip: an encrypted entry is named, not read as clean'
# MIME: nested too deep gives a reason; multipart without a boundary and a part without headers are still read.
$deep = "SSN: $($Sample.Ssn)"
for ($d = 0; $d -lt 7; $d++) { $deep = "Content-Type: multipart/mixed; boundary=`"B$d`"`r`n`r`n--B$d`r`n$deep`r`n--B$d--" }
$deepParts = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($deep)) -FileName 'd.eml').Parts.ToArray()
Assert-Equal $true (@($deepParts | Where-Object { $_.Reason -match 'nested too deep' }).Count -eq 1) 'EML: a part nested too deep gives a reason'
$noBoundary = "Content-Type: multipart/mixed`r`n`r`nSSN: $($Sample.Ssn)`r`n"
$noBoundaryText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($noBoundary)) -FileName 'n.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $noBoundaryText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: multipart without a boundary is read as text'
$bare = "Content-Type: multipart/mixed; boundary=`"XX`"`r`n`r`n--XX`r`n`r`nSSN: $($Sample.Ssn)`r`n`r`nmore`r`n--XX--`r`n"
$bareText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($bare)) -FileName 'b.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $bareText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: a part without headers keeps its first lines'
# An encrypted OpenDocument file is named; an Outlook message with attachments gets a note.
$lockedOdt = New-ZipBytes @{ 'content.xml' = 'x'; 'META-INF/manifest.xml' = '<manifest:file-entry><manifest:encryption-data/></manifest:file-entry>' }
Assert-Equal $true ((Get-TextFromBytes -Bytes $lockedOdt -FileName 'l.odt').Reason -match '^Encrypted') 'ODT: a password-protected file is named'
$attachMark = [System.Text.Encoding]::Unicode.GetBytes('__attach_version1.0_#00000000')
$msgAttach = [byte[]]($oleHead + [System.Text.Encoding]::ASCII.GetBytes("Subject: roster $ssnLine") + $attachMark)
Assert-Equal $true ((Get-TextFromBytes -Bytes $msgAttach -FileName 'm.msg').Note -match 'attachments') 'Outlook .msg: attachments get a note'
# An e-mail with more parts than the limit: the rest gets a reason.
$manyParts = "Content-Type: multipart/mixed; boundary=`"P`"`r`n`r`n" + ((1..600 | ForEach-Object { "--P`r`nContent-Type: text/plain`r`n`r`nline $_`r`n" }) -join '') + "--P--`r`n"
$manyRead = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($manyParts)) -FileName 'p.eml').Parts.ToArray()
Assert-Equal 1 @($manyRead | Where-Object { $_.Reason -match '^More than 500 parts' }).Count 'EML: past the part limit, the rest gets a reason'
# Re-review of PR 19: a part's first line that looks like a header, with no blank line, keeps the lines after it.
$noBlank = "Content-Type: multipart/mixed; boundary=`"XX`"`r`n`r`n--XX`r`nNote: call me`r`nSSN $($Sample.Ssn)`r`n--XX--`r`n"
$noBlankText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($noBlank)) -FileName 'nb.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $noBlankText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: lines after a header-like first line are read'
# An encrypted entry whose name is not UTF-8 (byte 0x82, as 7-Zip writes some letters) is still named encrypted.
$oddName = New-ZipBytes @{ 'rXsume.txt' = $ssnLine }
for ($z = 0; $z + 3 -lt $oddName.Length; $z++) {
    if ($oddName[$z] -eq 0x50 -and $oddName[$z + 1] -eq 0x4B -and $oddName[$z + 3] -in @(2, 4) -and $oddName[$z + 2] -in @(1, 3)) {
        $oddName[$z + 6 + 2 * [int]($oddName[$z + 2] -eq 1)] = $oddName[$z + 6 + 2 * [int]($oddName[$z + 2] -eq 1)] -bor 1   # flags: bit 0
    }
}
for ($z = 0; $z + 1 -lt $oddName.Length; $z++) { if ($oddName[$z] -eq [byte][char]'X' -and $oddName[$z + 1] -eq [byte][char]'s') { $oddName[$z] = 0x82 } }
$oddPart = (Get-TextFromBytes -Bytes $oddName -FileName 'o.zip').Parts.ToArray()[0]
Assert-Equal $true ($oddPart.Reason -match '^Encrypted') 'Zip: an encrypted entry with a non-UTF-8 name is named encrypted'
# Third review of PR 19: past the entry limit an encrypted file still gives the "more files" reason.
$limitZip = New-ZipBytes ([ordered]@{ 'a.txt' = 'plain'; 'secret.txt' = $ssnLine })
$seen = 0
for ($z = 0; $z + 3 -lt $limitZip.Length; $z++) {
    if ($limitZip[$z] -eq 0x50 -and $limitZip[$z + 1] -eq 0x4B -and $limitZip[$z + 2] -eq 1 -and $limitZip[$z + 3] -eq 2) { if ((++$seen) -eq 2) { $limitZip[$z + 8] = $limitZip[$z + 8] -bor 1 } }
}
& (Get-Module ContainerText) { $script:MaxArchiveEntries = 1 }
$limitParts = (Get-TextFromBytes -Bytes $limitZip -FileName 'lim.zip').Parts.ToArray()
& (Get-Module ContainerText) { $script:MaxArchiveEntries = 2000 }
Assert-Equal 1 @($limitParts | Where-Object { $_.Reason -match '^More than 1 files' }).Count 'Zip: an encrypted file past the limit gives the more-files reason'
# A header given twice keeps both values; Content-Description is read as text.
$twice = "Comments: SSN $($Sample.Ssn)`r`nComments: second`r`nSubject: s`r`n`r`nbody`r`n"
$twiceText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($twice)) -FileName 't.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $twiceText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: a header given twice keeps both values'
$described = "Content-Type: multipart/mixed; boundary=`"D`"`r`n`r`n--D`r`nContent-Description: SSN $($Sample.Ssn)`r`nContent-Type: text/plain`r`n`r`nbody`r`n--D--`r`n"
$describedText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($described)) -FileName 'c.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $describedText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: Content-Description is read as text'
