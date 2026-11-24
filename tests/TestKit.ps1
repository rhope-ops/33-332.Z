<#
TestKit.ps1 - shared test tools: the assertion, the invented sample values and an Office file builder.

  Purpose:     Dot-sourced by Run-Tests.ps1 before the test files.
  Inputs:      none.
  Outputs:     $script:Pass and $script:Fail (counted by Assert-Equal); the $Sample table.
  How to run:  not on its own; run tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.

  Test data:   all invented. PII-shaped samples are built from parts when the tests run (the PAM,
               2026-11-20), so this public repo holds no complete SSN-, DoD ID- or phone-shaped string.
               Card numbers are the published test numbers of the card networks.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = '$Sample is read by the test files.')]
param()

$script:Pass = 0
$script:Fail = New-Object System.Collections.Generic.List[string]

# Invented sample values, assembled from parts.
$Sample = @{
    Ssn        = '123' + '-45-' + '6789'     # valid by the SSA rules, invented
    SsnBare    = '123' + '45' + '6789'
    SsnArea000 = '000' + '-12-' + '3456'
    SsnArea666 = '666' + '-12-' + '3456'
    SsnArea9xx = '912' + '-12-' + '3456'
    SsnGroup00 = '123' + '-00-' + '3456'
    SsnSerial0 = '123' + '-45-' + '0000'
    SsnMixed   = '123' + '-45 ' + '6789'
    DodId      = '12345' + '67890'
    Phone      = '555-201' + '-0142'
    PhoneParen = '(555) 201' + '-0142'
    Passport   = 'X12' + '34567'
    Email      = 'pat.sample' + '@example.com'
    EmailQuote = "o'connor.jane" + '@example.org'
    EmailIdn   = 'john.q' + '@ex' + [char]0xE4 + 'mple.com'
    EmailAccent = 'jos' + [char]0xE9 + '.garc' + [char]0xED + 'a' + '@example.net'
    GovEmail   = 'first.last' + '@agency.gov'
    MilEmail   = 'first.last' + '@unit.af.' + 'mil'
    Street     = '42 Elm' + ' Street'
    CityZip    = 'Springfield, IL ' + '62704'
    Site       = 'https://tenant.example.com/sites/TeamA'
}

function Assert-Equal {
    <#
    .SYNOPSIS
      Compares two values as text (case-sensitive) and counts a pass or records a failure.
    .PARAMETER Expected
      The expected value.
    .PARAMETER Actual
      The value the code returned.
    .PARAMETER Name
      The test name, shown on a failure.
    .OUTPUTS
      None.
    #>
    param($Expected, $Actual, [string]$Name)
    if ("$Expected" -ceq "$Actual") { $script:Pass++ }
    else { $script:Fail.Add("$Name : expected [$Expected] got [$Actual]") }
}

function Get-HitSummary {
    # Find-PiiInText results as 'Detector/Confidence/MaskedValue' strings; always an array.
    param([string]$Text)
    return , @(Find-PiiInText $Text | ForEach-Object { '{0}/{1}/{2}' -f $_.Detector, $_.Confidence, $_.MaskedValue })
}

function New-ZipBytes {
    <#
    .SYNOPSIS
      Builds a minimal Office file or zip archive in memory.
    .PARAMETER Entries
      Part name -> XML text, or file bytes.
    .OUTPUTS
      [byte[]]
    #>
    param([hashtable]$Entries)
    Add-Type -AssemblyName System.IO.Compression
    $memory = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($name in $Entries.Keys) {
        $stream = $zip.CreateEntry($name).Open()
        # File bytes (a function's byte[] can arrive as object[]: PowerShell unrolls arrays on return).
        if ($Entries[$name] -is [array]) { $data = [byte[]]$Entries[$name]; $stream.Write($data, 0, $data.Length); $stream.Dispose(); continue }
        $writer = New-Object System.IO.StreamWriter($stream)
        $writer.Write($Entries[$name])
        $writer.Dispose()
    }
    $zip.Dispose()
    return $memory.ToArray()
}

function New-PdfBytes {
    <#
    .SYNOPSIS
      Builds a minimal one-page PDF in memory: a content stream (zlib-compressed unless -Plain), font F1 (with a
      ToUnicode map when one is given) and any extra objects (form fields, an encryption entry).
    .PARAMETER Content
      The page content stream.
    .PARAMETER ToUnicode
      A ToUnicode CMap for font F1, or empty.
    .PARAMETER Extra
      More object bodies, numbered from 7.
    .PARAMETER Trailer
      Extra trailer entries (for example an /Encrypt entry).
    .PARAMETER Plain
      Leave the content stream uncompressed.
    .PARAMETER PageContents
      The page's /Contents value (default 4 0 R, the content stream).
    .PARAMETER FontDict
      Font F1's dictionary instead of the default (a Type1 font, or a Type0 font with the ToUnicode map).
    .PARAMETER After
      Raw text added after the last object (for example an object defined again, as an incremental update does).
    .OUTPUTS
      [byte[]]
    #>
    param([string]$Content, [string]$ToUnicode = '', [string[]]$Extra = @(), [string]$Trailer = '', [switch]$Plain, [string]$PageContents = '4 0 R',
          [string]$FontDict = '', [string]$After = '')
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $stream = $latin1.GetBytes($Content)
    $filter = ''
    if (-not $Plain) {
        $memory = New-Object System.IO.MemoryStream
        $memory.WriteByte(0x78); $memory.WriteByte(0x9C)   # zlib header, as PDF writers put it
        $deflate = New-Object System.IO.Compression.DeflateStream($memory, [System.IO.Compression.CompressionMode]::Compress, $true)
        $deflate.Write($stream, 0, $stream.Length); $deflate.Dispose()
        $stream = $memory.ToArray(); $filter = ' /Filter /FlateDecode'
    }
    $font = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    if ($ToUnicode) { $font = '<< /Type /Font /Subtype /Type0 /BaseFont /X /ToUnicode 6 0 R >>' }
    if ($FontDict) { $font = $FontDict }
    $text = "%PDF-1.4`n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj`n2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj`n" +
        "3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents $PageContents >> endobj`n" +
        "4 0 obj << /Length $($stream.Length)$filter >> stream`n" + $latin1.GetString($stream) + "`nendstream endobj`n5 0 obj $font endobj`n" +
        "6 0 obj << /Length $($ToUnicode.Length) >> stream`n$ToUnicode`nendstream endobj`n"
    for ($i = 0; $i -lt $Extra.Count; $i++) { $text += "$($i + 7) 0 obj $($Extra[$i]) endobj`n" }
    $text += $After + "trailer << /Root 1 0 R $Trailer >>`n%%EOF"
    return $latin1.GetBytes($text)
}

function Get-XlsxSheets {
    <#
    .SYNOPSIS
      Reads a workbook back for the tests: tab name -> rows, each row an array of cell texts (inline strings
      and numbers). Only the parts XlsxWriter.psm1 writes are read.
    .PARAMETER Path
      The .xlsx file.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    $stream = [System.IO.File]::OpenRead($Path)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        # Reads one part of the zip as text.
        $read = { param($name) $r = New-Object System.IO.StreamReader($zip.GetEntry($name).Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
        $sheets = [ordered]@{}
        $names = @([regex]::Matches((& $read 'xl/workbook.xml'), '<sheet name="([^"]*)"') | ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value) })
        for ($i = 0; $i -lt $names.Count; $i++) {
            $rows = foreach ($row in [regex]::Matches((& $read ('xl/worksheets/sheet' + ($i + 1) + '.xml')), '(?s)<row[^>]*>(.*?)</row>')) {
                # Each cell at its column (an empty cell is not written), a formula cell as its value.
                $cells = New-Object System.Collections.Generic.List[string]
                foreach ($c in [regex]::Matches($row.Groups[1].Value, '(?s)<c r="([A-Z]+)\d+"[^>]*>(?:<is><t[^>]*>(.*?)</t></is>|(?:<f>.*?</f>)?<v>(.*?)</v>)</c>')) {
                    $column = 0; foreach ($ch in $c.Groups[1].Value.ToCharArray()) { $column = $column * 26 + ([int]$ch - 64) }
                    while ($cells.Count -lt $column - 1) { $cells.Add('') }
                    $cells.Add([System.Net.WebUtility]::HtmlDecode($c.Groups[2].Value + $c.Groups[3].Value))
                }
                , $cells.ToArray()
            }
            $sheets[$names[$i]] = @($rows)
        }
        return $sheets
    }
    finally { $zip.Dispose(); $stream.Dispose() }
}

function Get-ResultFiles {
    # Every results workbook under a folder (the script saves them in <site> folders as "Ion <site> <date time>.xlsx").
    param([string]$Folder)
    return @(Get-ChildItem -LiteralPath $Folder -Recurse -Filter 'Ion *.xlsx' -ErrorAction SilentlyContinue)
}

function Get-WorkbookTypes {
    <#
    .SYNOPSIS
      The PII types and note words in a workbook, sorted, one each: the Type column of the Hits tab and the Status
      column of the Not scanned tab.
    .PARAMETER Path
      The .xlsx file.
    .OUTPUTS
      [string] the names, joined with spaces.
    #>
    param([string]$Path)
    $book = Get-XlsxSheets $Path
    $types = foreach ($tab in $book.Keys) {
        if ($tab -eq 'Hits') { $book[$tab] | Select-Object -Skip 1 | ForEach-Object { $_[1] } }
        elseif ($tab -eq 'Not scanned') { $book[$tab] | Select-Object -Skip 1 | ForEach-Object { $_[0] } }
    }
    return (@($types | Sort-Object -Unique) -join ' ')
}

