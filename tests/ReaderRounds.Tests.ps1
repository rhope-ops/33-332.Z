<#
ReaderRounds.Tests.ps1 - tests of the readers from the review rounds 5 and later (PDF, fonts, glyph widths, e-mail).

  Purpose:     No silent loss and bounded time in the PDF reader (rounds 5-10), picture fonts, glyph names, real
               glyph widths, and the later reader fixes. Split from Extraction.Tests.ps1 (fleet v2 DEC 6, 2026-11-23:
               450 lines for hand-written code); it runs right after it and uses its variables.
  Inputs:      $Sample, New-PdfBytes and New-ZipBytes (TestKit.ps1); the variables of Extraction.Tests.ps1.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1, after Extraction.Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>
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
& (Get-Module PdfExtract) { $script:MaxInflatedBytes = 10 }
$lzwCapped = Get-TextFromBytes -Bytes $lzwPdf -FileName 'l2.pdf'
& (Get-Module PdfExtract) { $script:MaxInflatedBytes = 64MB }
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

# ---- Picture fonts (the Administrator, 2026-11-21): Wingdings and the like read as spaces, so their symbols give no hits ----
$digits = $Sample.Phone -replace '-', ' '
$wingdings = New-PdfBytes "BT /F1 12 Tf ($digits) Tj ET" -Plain -FontDict '<< /Type /Font /Subtype /TrueType /BaseFont /ABCDEF+Wingdings-Regular >>'
$wingRead = Get-TextFromBytes -Bytes $wingdings -FileName 'w.pdf'
Assert-Equal '' $wingRead.Text.Trim() 'PDF: a Wingdings run reads as spaces'
Assert-Equal 0 @(Find-PiiInText $wingRead.Text).Count 'PDF: a Wingdings run gives no hit'
$symbolPdf = New-PdfBytes "BT /F1 12 Tf ($digits) Tj ET" -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /SymbolMT >>'
Assert-Equal '' (Get-TextFromBytes -Bytes $symbolPdf -FileName 's.pdf').Text.Trim() 'PDF: a Symbol font run reads as spaces'
$plainPdf = New-PdfBytes "BT /F1 12 Tf ($digits) Tj ET" -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /SymbolicSans >>'
Assert-Equal $digits (Get-TextFromBytes -Bytes $plainPdf -FileName 'p.pdf').Text.Trim() 'PDF: a font name that only starts with Symbol is read'
$boldSymbol = New-PdfBytes "BT /F1 12 Tf ($digits) Tj ET" -Plain -FontDict '<< /Type /Font /Subtype /TrueType /BaseFont /ABCDEF+Symbol,Bold >>'
Assert-Equal '' (Get-TextFromBytes -Bytes $boldSymbol -FileName 'sb.pdf').Text.Trim() 'PDF: a bold Symbol font with a subset prefix reads as spaces'
$wideWing = New-PdfBytes 'BT /F1 12 Tf <00350035> Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type0 /BaseFont /Wingdings-Regular /Encoding /Identity-H /ToUnicode 6 0 R >>' `
    -ToUnicode 'begincmap 1 beginbfrange <0000> <FFFF> <0000> endbfrange endcmap'
Assert-Equal '' (Get-TextFromBytes -Bytes $wideWing -FileName 'ww.pdf').Text.Trim() 'PDF: a 2-byte Wingdings font reads as spaces'

# ---- A font's own glyph names (/Differences, the Administrator 2026-11-21): standard names read, made-up names never ----
$named = New-PdfBytes 'BT /F1 12 Tf (ABCD) Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+Garamond /Encoding << /Differences [65 /S /eacute /seven /a.sc] >> >>'
Assert-Equal ('S' + [char]0xE9 + '7a') (Get-TextFromBytes -Bytes $named -FileName 'n.pdf').Text.Trim() 'PDF: standard glyph names read (letter, accent, digit, suffix)'
$madeUp = New-PdfBytes 'BT /F1 12 Tf (ABC) Tj ( end) Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+X /Encoding 7 0 R >>' `
    -Extra @('<< /Type /Encoding /Differences [65 /g1 /g2 /g3] >>')
$madeUpRead = Get-TextFromBytes -Bytes $madeUp -FileName 'm.pdf'
Assert-Equal 'end|True' ($madeUpRead.Text.Trim() + '|' + ($madeUpRead.Note -match 'could not be read')) `
    'PDF: made-up glyph names (in an encoding object) read as spaces with a check-by-hand note; other codes as before'

# Review of PR 34: a picture font with glyph names still reads as spaces; real names this table lacks read as before.
$pictureNamed = New-PdfBytes "BT /F1 12 Tf ($digits) Tj ET" -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats /Encoding << /Differences [32 /space] >> >>'
Assert-Equal '' (Get-TextFromBytes -Bytes $pictureNamed -FileName 'pn.pdf').Text.Trim() 'PDF: a picture font with glyph names still reads as spaces'
$latin1 = [System.Text.Encoding]::GetEncoding(28591)
$signs = New-PdfBytes ('BT /F1 12 Tf (Acme' + $latin1.GetString([byte[]]@(169)) + ' 5' + $latin1.GetString([byte[]]@(177)) + '1 ' + $latin1.GetString([byte[]]@(161)) + ') Tj ET') -Plain `
    -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /X /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [161 /Aogonek 169 /copyright 177 /plusminus] >> >>'
$signsRead = Get-TextFromBytes -Bytes $signs -FileName 'sg.pdf'
Assert-Equal ('Acme' + [char]0xA9 + ' 5' + [char]0xB1 + '1 ' + [char]0x104 + '|') ($signsRead.Text.Trim() + '|' + $signsRead.Note) `
    'PDF: real glyph names the table lacks read as before, accents with ogonek read, no note'

$realNames = New-PdfBytes 'BT /F1 12 Tf (ABCDEF) Tj ET' -Plain `
    -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /X /Encoding << /Differences [65 /SF100000 /H18533 /u10000 /afii10017 /#4D /K 70 /g2] >> >>'
$realNamesRead = Get-TextFromBytes -Bytes $realNames -FileName 'rn.pdf'
Assert-Equal 'ABCDM|True' ($realNamesRead.Text.Trim() + '|' + ($realNamesRead.Note -match 'could not be read')) `
    'PDF: Adobe names with digits read as before; a #xx escape decoded; a later name for a code replaces the earlier one'

$caseNames = New-PdfBytes 'BT /F1 12 Tf (ABC) Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type1 /BaseFont /X /Encoding << /Differences [65 /h12 /u3 /uni00410042] >> >>'
Assert-Equal 'C' (Get-TextFromBytes -Bytes $caseNames -FileName 'cn.pdf').Text.Trim() `
    'PDF: lowercase made-up names (h12) and short u names (u3) read as spaces; a two-character uni name reads as before'

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
Assert-Equal 'hello' (& (Get-Module PdfExtract) { ConvertFrom-Ascii85 'BOu!rDZ~>' }) 'PDF: ASCII85 keeps a capital Z (hello encodes as BOu!rDZ)'

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
& (Get-Module PdfExtract) { $script:PdfSecondsLimit = 2 }
$clock.Restart()
$manyTc = Get-TextFromBytes -Bytes (New-PdfBytes (("1 Tc`n" * 300000) + "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET")) -FileName 'tc.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15 -and $manyTc.Status -in @('OK', 'NOT SCANNED')) 'PDF: many state operators between text objects end inside the time limit'
& (Get-Module PdfExtract) { $script:PdfSecondsLimit = 30 }
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
# Fourth review of PR 19: many repeated headers stay fast; conflicting Content headers give a reason; a second
# Content-Description is read.
$clock.Restart()
$repeated = ((1..20000 | ForEach-Object { 'X: ' + ('a' * 39) }) -join "`r`n") + "`r`n`r`nbody`r`n"
$null = Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($repeated)) -FileName 'r.eml'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 10) 'EML: 20,000 repeated headers are read in bounded time'
$conflicting = "Content-Type: text/plain`r`nContent-Type: multipart/mixed; boundary=`"B`"`r`n`r`n--B`r`nContent-Transfer-Encoding: base64`r`n`r`nU1NO`r`n--B--`r`n"
$conflictParts = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($conflicting)) -FileName 'cf.eml').Parts.ToArray()
Assert-Equal 1 @($conflictParts | Where-Object { $_.Reason -match 'conflicting Content headers' }).Count 'EML: conflicting Content headers give a reason'
$twoDesc = "Content-Description: first`r`nContent-Description: SSN $($Sample.Ssn)`r`n`r`nbody`r`n"
$twoDescText = ((Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($twoDesc)) -FileName 'd2.eml').Parts.ToArray() | ForEach-Object { $_.Text }) -join ' '
Assert-Equal $true (@(Find-PiiInText $twoDescText | ForEach-Object { $_.Detector }) -contains 'SSN') 'EML: a second Content-Description is read'
# The same Content-Type twice, written with other spacing or quotes, is no conflict.
$sameType = "Content-Type: text/plain; charset=utf-8`r`nContent-Type: text/plain;charset=`"utf-8`"`r`n`r`nbody`r`n"
$sameParts = (Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes($sameType)) -FileName 's.eml').Parts.ToArray()
Assert-Equal 0 @($sameParts | Where-Object { $_.Reason -match 'conflicting' }).Count 'EML: the same Content-Type in another spelling is no conflict'
