<#
PdfObjects.psm1 - the objects of a PDF and the text kept in its dictionaries, with .NET only (the PAM, 2026-11-20).

  Purpose:     Finds a PDF's objects (also in object streams and in incremental updates, older versions kept), decodes
               PDF strings, and reads the text kept in dictionaries: form-field values, comments, file info and XFA
               form data. A string it cannot read throws (Assert-PdfGap), so a PDF is never read in part without a word.
               PdfText.psm1 reads the page text on top of it.
  Inputs:      the PDF bytes as Latin-1 text.
  Outputs:     objects (Get-PdfObjects), decoded strings, dictionary and XFA text lines.
  How to run:  imported by Find-ChannelPII.ps1 (used by PdfText.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:ObjLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary). A literal holds at most
# 10,000 parts, so an unclosed bracket fails after a bounded scan. A "(" that no string token takes means a string
# was not read: the PDF is then NOT SCANNED, never read in part without a word (Assert-PdfGap).
$script:PdfStringPattern = '\((?>[^()\\]+|\\[\s\S]|\((?<d>)|\)(?<-d>)){0,10000}(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The message of the error a string not read throws (Get-PdfText turns it into "part of this PDF could not be read").
$script:PdfStringError = 'PDF string not read'

# The escapes of a literal string (\n, \053, a line break ...) and how each is undone, built once: building a
# PowerShell delegate per string cost about 450 microseconds (speed, 2026-11-21).
$script:EscapePattern = New-Object System.Text.RegularExpressions.Regex('\\(?:[0-7]{1,3}|\r?\n|.)')
$script:UndoEscape = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    $code = $m.Value.Substring(1)
    if ($code -match '^[0-7]+$') { return [string][char]([Convert]::ToInt32($code, 8) -band 0xFF) }
    switch -CaseSensitive ($code) { 'n' { "`n" } 'r' { "`r" } 't' { "`t" } 'b' { ' ' } 'f' { ' ' } default { $code.Trim("`r", "`n") } }
}

# An object stream (PDF 1.5) is unpacked only within these limits: its header of object numbers and offsets, and its
# object count (real files hold a few hundred objects per stream). More gives a "check by hand" note.
$script:MaxPackedHeader = 1MB

# Earlier versions of objects defined again further on (incremental updates): their text is still read.
$script:PdfOldObjects = New-Object System.Collections.Generic.List[object]
$script:MaxPackedObjects = 100000

# Dictionary keys whose string values are text a person wrote (Get-PdfDictText) -> the label put before the value:
# a form field's value (labelled with the field's name), a comment or note and its rich text, the file's info.
$script:PdfTextKeys = [ordered]@{ 'V' = ''; 'Contents' = 'Comment'; 'RC' = 'Comment'; 'Author' = 'Author'; 'Title' = 'Title'
    'Subject' = 'Subject'; 'Keywords' = 'Keywords' }

function Get-PdfStringPattern {
    # The pattern of one PDF string token (PdfText.psm1 builds its token pattern on it).
    return $script:PdfStringPattern
}

function Get-PdfStringError {
    # The message of the error a string not read throws.
    return $script:PdfStringError
}

function Stop-PdfHostile {
    # Ends the read of a PDF with more operands or array entries than any real one has: it is reported as NOT SCANNED
    # ("part could not be read"), so a hostile file cannot hold gigabytes (reviews r14, r18).
    throw (New-Object System.FormatException (Get-PdfStringError))
}

function Reset-PdfOldObjects {
    # Forgets the older object versions of the last PDF (Get-PdfText, before each PDF).
    $script:PdfOldObjects.Clear()
}

function Get-PdfOldObjects {
    # The older versions of objects defined again further on in this PDF (an incremental update).
    return , $script:PdfOldObjects.ToArray()
}

function Get-PdfObjects {
    <#
    .SYNOPSIS
      The objects of a PDF by number: each its dictionary text and its stream (inflated when FlateDecode; $null for
      other filters, such as images). Objects packed in object streams (PDF 1.5) are unpacked too.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary] number -> @{ Dict; Stream; Pos; Image (JPEG data or $null) }, in file order.
    #>
    param([string]$Text)
    $objects = New-Object System.Collections.Specialized.OrderedDictionary
    # Each object runs from its header to the first "endobj" after it, or to the next header when "endobj" is missing.
    # Two linear passes, so a damaged file without "endobj" cannot make the search quadratic.
    $heads = @((Get-PdfRegex '(\d+)\s+\d+\s+obj\b').Matches($Text))
    $ends = @((Get-PdfRegex '\bendobj\b').Matches($Text) | ForEach-Object { $_.Index })
    $e = 0
    for ($h = 0; $h -lt $heads.Count; $h++) {
        Assert-PdfTime
        $start = $heads[$h].Index + $heads[$h].Length
        $stop = $Text.Length; if ($h + 1 -lt $heads.Count) { $stop = $heads[$h + 1].Index }
        while ($e -lt $ends.Count -and $ends[$e] -lt $start) { $e++ }
        if ($e -lt $ends.Count -and $ends[$e] -lt $stop) { $stop = $ends[$e] }
        $body = $Text.Substring($start, $stop - $start)
        $stream = $null
        $parts = Split-PdfStreamBody $body
        $image = $null
        if ($parts) {
            $body = $parts.Dict; $stream = Read-PdfStream $body $parts.Data
            # A JPEG page image (a scanned page): kept as it is, for OCR when the PDF has no text (Get-ScanPagesText).
            if ($body -match '/Subtype\s*/Image\b' -and $body -match '/Filter\s*\[?\s*/DCTDecode\s*\]?') { $image = $parts.Data }
        }
        $number = $heads[$h].Groups[1].Value
        # A file saved with added edits (an incremental update) defines an object again further on: the later one is
        # the one in use; the earlier one's text is still in the file, so it is kept (Get-PdfDictText, Get-PdfLooseText).
        if ($objects.Contains($number)) { $script:PdfOldObjects.Add($objects[$number]) }
        $objects[$number] = @{ Dict = $body; Stream = $stream; Pos = $heads[$h].Index; Image = $image }
    }
    foreach ($packed in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/ObjStm' -and $_.Stream })) { Add-PackedObjects $objects $packed }
    return $objects
}

function Split-PdfStreamBody {
    <#
    .SYNOPSIS
      An object body with a stream, split into its dictionary and its data: from the "stream" keyword (then CR LF, LF
      or CR) to the last "endstream". Plain string searches, so a stream of any size is split in one linear pass.
      $null for an object without a stream; a stream without its end gives a note.
    .PARAMETER Body
      The object's text from its header to endobj.
    .OUTPUTS
      [hashtable] Dict, Data; or $null.
    #>
    param([string]$Body)
    $at = -1
    for ($from = 0; ($at = $Body.IndexOf('stream', $from, [StringComparison]::Ordinal)) -ge 0; $from = $at + 6) {
        $after = $at + 6
        $word = $at -eq 0 -or -not [char]::IsLetter($Body[$at - 1])   # not the end of "endstream"
        if ($word -and $after -lt $Body.Length -and ($Body[$after] -eq "`n" -or $Body[$after] -eq "`r")) { break }
    }
    if ($at -lt 0) { return $null }
    $start = $at + 6
    if ($Body[$start] -eq "`r") { $start++ }
    if ($start -lt $Body.Length -and $Body[$start] -eq "`n") { $start++ }
    $end = $Body.LastIndexOf('endstream', [StringComparison]::Ordinal)
    if ($end -lt $start) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return $null }   # cut short
    $stop = $end
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`n") { $stop-- }
    if ($stop -gt $start -and $Body[$stop - 1] -eq "`r") { $stop-- }
    return @{ Dict = $Body.Substring(0, $at); Data = $Body.Substring($start, $stop - $start) }
}

function Add-PackedObjects {
    <#
    .SYNOPSIS
      Adds the objects of one object stream (/Type /ObjStm: N pairs "number offset", then the objects from /First).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Packed
      One object stream of Get-PdfObjects.
    #>
    param($Objects, $Packed)
    $count = [int][regex]::Match($Packed.Dict, '/N\s+(\d+)').Groups[1].Value
    $first = [int][regex]::Match($Packed.Dict, '/First\s+(\d+)').Groups[1].Value
    if ($first -le 0 -or $first -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return }
    if ($first -gt $script:MaxPackedHeader -or $count -gt $script:MaxPackedObjects) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand'; return }
    $pairs = @($Packed.Stream.Substring(0, $first).Trim() -split '\s+')
    $number = 0
    for ($k = 0; $k -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        if (-not [int]::TryParse($pairs[$k], [ref]$number)) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; return }
    }
    for ($k = 0; $k -lt $count -and 2 * $k + 1 -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        $start = $first + [int]$pairs[2 * $k + 1]
        $end = $Packed.Stream.Length
        if (2 * $k + 3 -lt $pairs.Count) { $end = $first + [int]$pairs[2 * $k + 3] }
        if ($start -ge $end -or $end -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is not well formed (the file may still open fine): check by hand'; continue }
        $dict = $Packed.Stream.Substring($start, $end - $start)
        $number = $pairs[2 * $k]
        # The later definition in the file wins (an incremental update); the other one's text is kept too.
        if ($Objects.Contains($number) -and $Objects[$number].Pos -gt $Packed.Pos) { $script:PdfOldObjects.Add(@{ Dict = $dict; Stream = $null }); continue }
        if ($Objects.Contains($number)) { $script:PdfOldObjects.Add($Objects[$number]) }
        $Objects[$number] = @{ Dict = $dict; Stream = $null; Pos = $Packed.Pos }
    }
}

function ConvertFrom-PdfString {
    <#
    .SYNOPSIS
      The bytes of one PDF string token, as Latin-1 characters: (literal) with its escapes undone, or <hex>.
    .PARAMETER Token
      One PDF string token.
    #>
    param([string]$Token)
    if ($Token.StartsWith('<')) {
        $hex = $Token.Substring(1, $Token.Length - 2) -replace '\s', ''
        if ($hex.Length % 2) { $hex += '0' }
        return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
            if (($i % 200000) -eq 0) { Assert-PdfTime }
            [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
        })
    }
    $inner = $Token.Substring(1, $Token.Length - 2)
    if ($inner.IndexOf([char]92) -lt 0) { return $inner }   # no escapes (most strings): as it is (speed)
    return $script:EscapePattern.Replace($inner, $script:UndoEscape)
}

function ConvertFrom-PdfTextString {
    # A PDF text string (a form field's name or value): UTF-16 when it starts with the FE FF mark, else Latin-1.
    param([string]$Token)
    $raw = ConvertFrom-PdfString $Token
    if ($raw.Length -ge 2 -and [int]$raw[0] -eq 0xFE -and [int]$raw[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($script:ObjLatin1.GetBytes($raw.Substring(2)))
    }
    return $raw
}

function Assert-PdfGap {
    <#
    .SYNOPSIS
      Throws when the text between two tokens holds a "(": a string there was not read (too long, or damaged), so
      the PDF is reported NOT SCANNED rather than read in part without a word (fresh review, 2026-11-20).
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the gap starts.
    .PARAMETER To
      Where the gap ends.
    #>
    param([string]$Content, [int]$From, [int]$To)
    if ($To -gt $From -and $Content.IndexOf('(', $From, $To - $From) -ge 0) { throw (New-Object System.FormatException $script:PdfStringError) }
}

function Get-PdfFieldName {
    <#
    .SYNOPSIS
      A form field's name (/T) as the label of a value: the last name before the value in the dictionary (a
      dictionary can hold several fields inline), else the first; empty when there is none.
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .PARAMETER Before
      Where the value starts in the dictionary.
    #>
    param([string]$Dict, $NameRegex, [int]$Before)
    $names = @($NameRegex.Matches($Dict))
    if ($names.Count -eq 0) { return '' }
    $pick = $names[0]
    foreach ($name in $names) { if ($name.Index -lt $Before) { $pick = $name } }
    return (ConvertFrom-PdfTextString $pick.Groups[1].Value)
}

function Format-PdfDictLine {
    # One line of dictionary text: "label: text", or the text alone.
    param([string]$Label, [string]$Text)
    if ($Label) { return ($Label + ': ' + $Text) }
    return $Text
}

function Get-PdfDictText {
    <#
    .SYNOPSIS
      Text kept in dictionaries, not on a page: form-field values ("name: value", so labels stay with values),
      comments and sticky notes (/Contents of an annotation, /RC rich text) and the file's info (author, title,
      subject, keywords), in current and older object versions; every value of a key, also several in one
      dictionary. A literal value that is not read throws (Get-PdfText: "part of this PDF could not be read").
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] one line per value.
    #>
    param($Objects)
    $nameRegex = Get-PdfRegex ('/T\s*(' + $script:PdfStringPattern + ')')
    $checks = foreach ($key in $script:PdfTextKeys.Keys) {
        @{ Key = $key; Label = $script:PdfTextKeys[$key]; Value = (Get-PdfRegex ('/' + $key + '\s*(' + $script:PdfStringPattern + ')'))
           Ref = (Get-PdfRegex ('/' + $key + '\s+(\d+)\s+\d+\s+R')); Open = ('/' + $key + '\s*\(') }
    }
    # Older versions too. ToArray(), not @(): @() around an empty generic list fails ("Argument types do not match").
    $all = @($Objects.Values) + $script:PdfOldObjects.ToArray()
    foreach ($obj in $all) {
        Assert-PdfTime
        foreach ($check in $checks) {
            if ($obj.Dict.IndexOf('/' + $check.Key) -ge 0) { Get-PdfKeyLines $obj.Dict $check $Objects $nameRegex }
        }
    }
}

function Get-PdfKeyLines {
    <#
    .SYNOPSIS
      The lines of one key in one dictionary: each string value, a list of strings (a multi-select field), or a
      value held in an object of its own (a string, or a text stream for /V and /RC).
    .PARAMETER Dict
      The dictionary text.
    .PARAMETER Check
      The key, its label and its regexes (Get-PdfDictText).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER NameRegex
      The timed regex of a /T string.
    .OUTPUTS
      [string] lines.
    #>
    param([string]$Dict, [hashtable]$Check, $Objects, $NameRegex)
    $values = @($Check.Value.Matches($Dict))
    $literals = @($values | Where-Object { $_.Groups[1].Value.StartsWith('(') }).Count
    if ([regex]::Matches($Dict, $Check.Open).Count -gt $literals) { throw (New-Object System.FormatException $script:PdfStringError) }   # not read
    foreach ($value in $values) {
        $text = ConvertFrom-PdfTextString $value.Groups[1].Value
        if ($Check.Key -eq 'RC') { $text = $text -replace '<[^>]+>', ' ' }   # rich text: XHTML tags dropped
        $label = $Check.Label
        if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $value.Index }
        Format-PdfDictLine $label $text
    }
    if ($values.Count -gt 0) { return }
    $list = [regex]::Match($Dict, '/' + $Check.Key + '\s*\[([^\]]*)\]')
    if ($list.Success) {
        $items = @(foreach ($item in (Get-PdfRegex $script:PdfStringPattern).Matches($list.Groups[1].Value)) { ConvertFrom-PdfTextString $item.Value })
        $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $list.Index }
        if ($items.Count -gt 0) { Format-PdfDictLine $label ($items -join ', ') }
        return
    }
    $ref = $Check.Ref.Match($Dict)
    if (-not $ref.Success -or -not $Objects.Contains($ref.Groups[1].Value)) { return }
    $held = $Objects[$ref.Groups[1].Value]
    $target = $held.Dict.Trim()
    $label = $Check.Label; if ($Check.Key -eq 'V') { $label = Get-PdfFieldName $Dict $NameRegex $ref.Index }
    if ($target.StartsWith('(') -or ($target.StartsWith('<') -and -not $target.StartsWith('<<'))) { Format-PdfDictLine $label (ConvertFrom-PdfTextString $target) }
    elseif ($Check.Key -in @('V', 'RC') -and $held.Stream) { Format-PdfDictLine $label $held.Stream }   # a text stream
}

function Get-PdfXfaText {
    <#
    .SYNOPSIS
      The data of an XFA form (/AcroForm /XFA: one stream, or an array of name and stream pairs): XML, read as
      "element: value" lines, so a field name stays with its value (a dynamic XFA form shows only "Please wait..."
      on its page; its data is here).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] lines.
    #>
    param($Objects)
    $pair = Get-PdfRegex '<([\w:.-]+)\b[^<>]*>([^<]+)</\1>'
    foreach ($obj in @($Objects.Values)) {
        $xfa = [regex]::Match($obj.Dict, '/XFA\s*(\[[^\]]*\]|\d+\s+\d+\s+R)')
        if (-not $xfa.Success) { continue }
        foreach ($ref in [regex]::Matches($xfa.Groups[1].Value, '(\d+)\s+\d+\s+R')) {
            Assert-PdfTime
            $part = $Objects[$ref.Groups[1].Value]
            if (-not $part -or -not $part.Stream) { continue }   # a stream not decoded has its note already
            $lines = $pair.Replace($part.Stream, "`n`$1: `$2`n") -replace '<[^>]+>', ' '
            [System.Net.WebUtility]::HtmlDecode($lines)
        }
    }
}

Export-ModuleMember -Function Get-PdfObjects, ConvertFrom-PdfString, ConvertFrom-PdfTextString, Assert-PdfGap, Get-PdfDictText, Get-PdfXfaText, Get-PdfStringPattern,
Get-PdfStringError, Stop-PdfHostile, Reset-PdfOldObjects, Get-PdfOldObjects
