<#
ScanMeter.psm1 - the progress bars and the per-folder lines of a scan: counts only, so the screen shows no PII.

  Purpose:     Lets the person follow a scan: files done of the total, read, not scanned, possible PII,
               time used and time left (the PAM, 2026-11-20). No found value is shown; channel and file
               names show only masked PII.
  Inputs:      the report row list and the counts the scan functions of ChannelScan.psm1 give.
  Outputs:     a meter (hashtable), status text, percent, a per-file progress bar and one coloured line per folder,
               counts in fixed widths (the PAM, 2026-11-20); the boxed screen tables of the summary, cut to the window.
  Depends on:  PiiMasking.psm1 (Protect-PiiName), Report.psm1 (Write-Ui), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none of its own: it reads a download stream that ChannelScan.psm1 opens.
  Writes:      none (screen only).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# The per-file bar is shown only from this size on: a small file is done in a moment, and each bar update costs time in
# Windows PowerShell 5.1 (the author's choice, 2026-11-20: faster scans of many small files).
$script:FileBarBytes = 1MB

# The per-file bar shows a speed and a time left only after this many seconds (before that the speed jumps about).
$script:RateAfterSeconds = 1

# A streamed download is read this much at a time (64 KB); the bar is updated at most this often (8 a second): smaller
# steps, so the bar moves smoothly (the PAM, 2026-11-21: 256 KB and 4 a second looked choppy).
$script:ChunkBytes = 64KB
$script:BarRefreshMs = 125

# Screen tables fit the window: under 60 characters or unknown = 120; the last column keeps 12 (the author's choice, 2026-11-21).
$script:DefaultScreenWidth = 120
$script:MinScreenWidth = 60
$script:MinLastColumn = 12

# A coloured line of live counts is printed under the progress bar at most this often (and after each folder).
$script:LiveLineSeconds = 10

# Width of each count on the live lines: 5 characters (up to 99,999), so the columns do not move (the PAM, 2026-11-20).
$script:CountWidth = 5

# Length of the grey rule after each folder's lines.
$script:RuleWidth = 100

# The name column of the folder rows: as wide as the longest folder name, at most this; the total row says "Total".
$script:MaxNameWidth = 40
$script:TotalName = 'Total'

# Short PII type names for the live line, in the PAM's order (Report.psm1 $SeverityOrder).
$script:ShortTypeNames = [ordered]@{ 'SSN' = 'SSN'; 'Date of birth' = 'DOB'; 'Passport' = 'Passport'; 'Payment card' = 'Card'
    'DoD ID' = 'DoD ID'; 'Home address' = 'Address'; 'Name' = 'Name'; 'Phone' = 'Phone'; 'Personal e-mail' = 'E-mail' }

function New-ScanMeter {
    <#
    .SYNOPSIS
      Starts a meter for one scan step: items done, read, not scanned, possible PII, and the time.
    .PARAMETER Total
      How many items the step has (0 when not known up front).
    .OUTPUTS
      [hashtable] the meter (Update-ScanMeter and Get-ScanMeterText use it).
    #>
    param([int]$Total)
    # Clock: anything with an Elapsed [TimeSpan] (a running stopwatch here; a fixed time in the tests).
    @{ Total = $Total; Done = 0; Read = 0; Skipped = 0; Hits = 0; HitItems = 0; Types = @{}; LastLine = [TimeSpan]::Zero; NameWidth = 0
       Clock = [Diagnostics.Stopwatch]::StartNew() }
}

function Update-ScanMeter {
    <#
    .SYNOPSIS
      Counts one item into the meter from the report rows it added.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Rows
      The report row list.
    .PARAMETER FirstRow
      The row count before the item was scanned.
    .PARAMETER ReadCount
      The items whose text was read (the scan function's return value).
    #>
    param([hashtable]$Meter, $Rows, [int]$FirstRow, [int]$ReadCount)
    $new = @(for ($r = $FirstRow; $r -lt $Rows.Count; $r++) { $Rows[$r] })
    $hits = @($new | Where-Object { $_.Confidence })   # a row with a confidence is a possible PII hit
    $Meter.Done++
    $Meter.Read += $ReadCount
    $Meter.Skipped += @($new | Where-Object { $_.Detector -eq 'NOT SCANNED' -or $_.Detector -eq 'BLOCKED' }).Count
    $Meter.Hits += $hits.Count
    $Meter.HitItems += @($hits | Select-Object -ExpandProperty Item -Unique).Count   # files with a hit
    foreach ($hit in $hits) { $Meter.Types[$hit.Detector] = 1 + [int]$Meter.Types[$hit.Detector] }
}

function Format-Duration {
    # A time as minutes:seconds; the minutes go past 59 (65 minutes = 65:00), so a long scan never wraps to 05:00.
    param([TimeSpan]$Time)
    return '{0:00}:{1:00}' -f [int][Math]::Floor($Time.TotalMinutes), $Time.Seconds
}

function Format-Count {
    # A count right-aligned in $CountWidth characters, so the columns of the live lines stay in place (the PAM, 2026-11-20).
    param([long]$Value)
    return ([string]$Value).PadLeft($script:CountWidth)
}

function Get-TimeLeftText {
    # ", about mm:ss left" from the pace so far; empty when the total is not known or the step is done.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0 -or $Meter.Done -le 0 -or $Meter.Done -ge $Meter.Total) { return '' }
    $elapsed = $Meter.Clock.Elapsed
    return ', about ' + (Format-Duration ([TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done))))) + ' left'
}

function Get-ScanMeterText {
    <#
    .SYNOPSIS
      The run so far as one line for the scan bar (its middle line; the PAM, 2026-11-21): files done of the total with
      the percent, read, not scanned, possible PII, the time elapsed and left, then the hits per type.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .OUTPUTS
      [string] for example "12 of 40 files (30%) | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 elapsed, about 01:48 left".
    #>
    param([hashtable]$Meter, [string]$Noun)
    $totalNoun = $Noun + 's'; if ([Math]::Max($Meter.Total, $Meter.Done) -eq 1) { $totalNoun = $Noun }   # "1 of 1 file"
    $text = '{0:N0} {1}' -f $Meter.Done, $totalNoun
    if ($Meter.Total -gt 0) { $text = '{0:N0} of {1:N0} {2} ({3}%)' -f $Meter.Done, $Meter.Total, $totalNoun, (Get-MeterPercent $Meter) }
    $hitNoun = $Noun + 's'; if ($Meter.HitItems -eq 1) { $hitNoun = $Noun }   # "in 1 file", "in 2 files"
    $text += ' | read {0:N0} | not scanned {1:N0} | possible PII {2:N0} in {3:N0} {4} | {5} elapsed{6}' -f
        $Meter.Read, $Meter.Skipped, $Meter.Hits, $Meter.HitItems, $hitNoun, (Format-Duration $Meter.Clock.Elapsed), (Get-TimeLeftText $Meter)
    $types = Get-TypeCountText $Meter
    if ($types) { $text += ' | ' + $types }
    return $text
}

function Get-TypeCountText {
    # The hits per type so far, worst type first, as "SSN 2, DOB 1"; empty when there is none.
    param([hashtable]$Meter)
    if (-not $Meter.ContainsKey('Types') -or $Meter.Types.Count -eq 0) { return '' }
    $parts = foreach ($type in $script:ShortTypeNames.Keys) {
        if ($Meter.Types[$type]) { '{0} {1}' -f $script:ShortTypeNames[$type], $Meter.Types[$type] }
    }
    return (@($parts) -join ', ')
}

function Write-CountPart {
    # One part of a folder or total row: " | " (dark grey), the label (grey), the count (fixed width) in its colour.
    param([string]$Label, [long]$Value, [string]$Kind)
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Label + ' ') 'Muted' -NoNewline
    Write-Ui (Format-Count $Value) $Kind -NoNewline
}

function Get-CountKind {
    # The colour of a count: White when 0; else Yellow (not scanned: check by hand) or Red (possible PII).
    param([long]$Value, [string]$WhenSome)
    if ($Value -gt 0) { return $WhenSome }
    return 'Text'
}

function Get-NameWidth {
    # The width of the name column: the longest (shown) folder name, at least "Total", at most $MaxNameWidth.
    param([string[]]$Names)
    $width = $script:TotalName.Length
    foreach ($name in $Names) { $width = [Math]::Max($width, $name.Length) }
    return [Math]::Min($width, $script:MaxNameWidth)
}

function Format-Name {
    # A name padded (or cut, ending in "...") to the name column's width, so every row lines up.
    param([string]$Name, [int]$Width)
    if ($Name.Length -gt $Width) { return $Name.Substring(0, $Width - 3) + '...' }
    return $Name.PadRight($Width)
}

function Get-RowWidth {
    # The name column's width of a meter (a meter made in a test may have none).
    param([hashtable]$Meter)
    if ($Meter.ContainsKey('NameWidth') -and $Meter.NameWidth -gt 0) { return [int]$Meter.NameWidth }
    return $script:TotalName.Length
}

function Write-MeterLine {
    <#
    .SYNOPSIS
      Prints the running total as one row in the same columns as the folder rows: "Total", files done of the total
      and percent, read, not scanned, possible PII, time used and left, then the PII per type. At most every
      $LiveLineSeconds, unless -Force. Counts only: no file name and no value.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .PARAMETER Force
      Print now (after a folder).
    #>
    param([hashtable]$Meter, [string]$Noun, [switch]$Force)
    $elapsed = $Meter.Clock.Elapsed
    if (-not $Force -and ($elapsed - $Meter.LastLine).TotalSeconds -lt $script:LiveLineSeconds) { return }
    $Meter.LastLine = $elapsed
    Write-Ui ('  ' + (Format-Name $script:TotalName (Get-RowWidth $Meter))) 'Text' -NoNewline
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Noun + 's ') 'Muted' -NoNewline
    Write-Ui ('{0}/{1} {2,3}%' -f (Format-Count $Meter.Done), (Format-Count $Meter.Total), (Get-MeterPercent $Meter)) 'Text' -NoNewline
    Write-CountPart 'read' $Meter.Read 'Text'
    Write-CountPart 'not scanned' $Meter.Skipped (Get-CountKind $Meter.Skipped 'Accent')
    Write-CountPart 'possible PII' $Meter.Hits (Get-CountKind $Meter.Hits 'Alert')
    # No time here: the scan bar shows the time elapsed and left (the PAM, 2026-11-21).
    $types = Get-TypeCountText $Meter
    if ($types) { Write-Ui ' | ' 'Rule' -NoNewline; Write-Ui $types 'Alert' -NoNewline }
    Write-Ui ''
}

function Get-MeterPercent {
    # The meter's percent done (0-100); 0 when the total is not known.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0) { return 0 }
    return [int][Math]::Floor([Math]::Min(100.0, 100.0 * $Meter.Done / $Meter.Total))   # down: 100% only when all are done
}

function Test-FileBarSize {
    # True when a file is large enough for the per-file bar (smaller files are done in a moment).
    param([long]$Size)
    return $Size -ge $script:FileBarBytes
}

function Write-FileStep {
    <#
    .SYNOPSIS
      The second progress bar: the download of the file in work, by byte (MB done of the size, speed, time left), so a
      large file does not look hung (the PAM, 2026-11-21: the bar must move). Reading happens in the workers, after it.
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER Step
      What is being done now, for example 'Downloading'.
    .PARAMETER Done
      The bytes done so far; -1 when the download method gives no byte count (the bar then says so and does not move).
    .PARAMETER Elapsed
      The time the step has taken so far.
    #>
    param([long]$Size, [string]$Step, [long]$Done = -1, [TimeSpan]$Elapsed = [TimeSpan]::Zero)
    if (-not (Test-FileBarSize $Size)) { return }
    $sizeMB = $Size / 1MB
    if ($Done -lt 0) {
        $status = '{0:N1} MB - SharePoint gives no byte count for this file: please wait' -f $sizeMB
        $percent = 0
    }
    else {
        $percent = [int][Math]::Floor([Math]::Min(100.0, 100.0 * $Done / [Math]::Max(1, $Size)))
        $status = '{0:N1} of {1:N1} MB ({2}%)' -f ($Done / 1MB), $sizeMB, $percent
        $seconds = $Elapsed.TotalSeconds
        if ($Done -gt 0 -and $seconds -ge $script:RateAfterSeconds) {
            $rate = $Done / $seconds
            $status += ' | {0:N1} MB/s' -f ($rate / 1MB)
            if ($Done -lt $Size) { $status += ' | about {0} left' -f (Format-Duration ([TimeSpan]::FromSeconds(($Size - $Done) / $rate))) }
        }
    }
    # No -ParentId: Windows PowerShell ISE indents a child bar; the PAM wants both bars aligned (2026-11-20).
    Write-Progress -Id 2 -Activity ((Get-IonMark) + ' ' + $Step) -Status $status -PercentComplete $percent
}

function Copy-StreamWithProgress {
    <#
    .SYNOPSIS
      Reads a stream (a download) to its end into memory, a chunk at a time, and moves the per-file bar at most every
      $BarRefreshMs (each bar update costs time in Windows PowerShell 5.1) and once at the end.
    .PARAMETER Stream
      The stream to read.
    .PARAMETER Size
      The expected size (the bar's 100%, and the first buffer size); 0 when not known.
    .PARAMETER Step
      The bar's step text, for example 'Downloading'.
    .OUTPUTS
      [byte[]]
    #>
    param([System.IO.Stream]$Stream, [long]$Size, [string]$Step)
    $memory = New-Object System.IO.MemoryStream([int][Math]::Min([long][int]::MaxValue, [Math]::Max(0L, $Size)))
    try {
        $buffer = New-Object byte[] $script:ChunkBytes
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $last = 0L
        Write-FileStep $Size $Step 0 $clock.Elapsed
        while (($count = $Stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $memory.Write($buffer, 0, $count)
            if ($clock.ElapsedMilliseconds - $last -lt $script:BarRefreshMs) { continue }
            $last = $clock.ElapsedMilliseconds
            Write-FileStep $Size $Step $memory.Length $clock.Elapsed
        }
        Write-FileStep $Size $Step $memory.Length $clock.Elapsed
        # The leading comma keeps the byte array whole: PowerShell would otherwise unroll it, byte by byte (slow).
        return , $memory.ToArray()
    }
    finally { $memory.Dispose() }
}

function Format-BoxTable {
    <#
    .SYNOPSIS
      A text table with ASCII borders (no other characters: Windows PowerShell 5.1 and ISE show them all alike).
      Text columns are left-aligned, number columns right-aligned.
    .PARAMETER Columns
      The header texts.
    .PARAMETER Rows
      The rows, each an array of cells.
    .PARAMETER MaxWidth
      The widest a line may be (0: no limit); the last column is cut to fit.
    .OUTPUTS
      [string[]] the border line, the header, the border, one line per row, the border.
    #>
    param([string[]]$Columns, [object[]]$Rows, [int]$MaxWidth = 0)
    $widths = @(for ($i = 0; $i -lt $Columns.Count; $i++) {
        [Math]::Max($Columns[$i].Length, (@($Rows | ForEach-Object { ([string]$_[$i]).Length }) + 0 | Measure-Object -Maximum).Maximum)
    })
    # Too wide for the screen: the last column (a file path) is cut, with '...', so no line wraps (the PAM, 2026-11-21).
    $total = ($widths | Measure-Object -Sum).Sum + 3 * $widths.Count + 1
    if ($MaxWidth -gt 0 -and $total -gt $MaxWidth) { $widths[-1] = [Math]::Max([Math]::Min($script:MinLastColumn, $widths[-1]), $widths[-1] - ($total - $MaxWidth)) }
    $border = '+' + ((@($widths) | ForEach-Object { '-' * ($_ + 2) }) -join '+') + '+'
    # One table line: each cell padded to its column width, numbers to the right.
    $line = {
        param($cells)
        $parts = for ($i = 0; $i -lt $Columns.Count; $i++) {
            $cell = [string]$cells[$i]
            if ($cell.Length -gt $widths[$i]) { $cell = $cell.Substring(0, $widths[$i] - 3) + '...' }
            if ($cells[$i] -is [int]) { ' ' + $cell.PadLeft($widths[$i]) + ' ' } else { ' ' + $cell.PadRight($widths[$i]) + ' ' }
        }
        '|' + ($parts -join '|') + '|'
    }
    $border
    & $line $Columns
    $border
    foreach ($row in $Rows) { & $line $row }
    $border
}

function Get-ScreenWidth {
    # The width of the console or ISE pane in characters, less one (a full line would wrap); $DefaultScreenWidth when unknown.
    try { $width = [int]$Host.UI.RawUI.WindowSize.Width } catch { $width = 0 }   # no window size: the default below
    if ($width -lt $script:MinScreenWidth) { $width = $script:DefaultScreenWidth }
    return $width - 1
}

function Get-TableLines {
    <#
    .SYNOPSIS
      A boxed table as screen lines: borders grey, the header in the label colour, each row in its own colour.
    .PARAMETER Columns
      The header texts.
    .PARAMETER Rows
      The rows, each an array of cells.
    .PARAMETER Kinds
      One Write-Ui colour per row.
    .OUTPUTS
      [pscustomobject] Text, Kind, LabelLength (0), one per line.
    #>
    param([string[]]$Columns, [object[]]$Rows, [string[]]$Kinds)
    $table = @(Format-BoxTable $Columns $Rows (Get-ScreenWidth))
    for ($i = 0; $i -lt $table.Count; $i++) {
        $kind = 'Rule'
        if ($i -eq 1) { $kind = 'Label' } elseif ($i -ge 3 -and $i -lt $table.Count - 1) { $kind = $Kinds[$i - 3] }
        [pscustomobject]@{ Text = $table[$i]; Kind = $kind; LabelLength = 0 }
    }
}

function Write-ChannelLine {
    <#
    .SYNOPSIS
      Prints one row when a folder (a channel) is done: its (masked) name first in teal, in a fixed-width column,
      then its counts in fixed widths with grey labels (the PAM, 2026-11-20). Not scanned in yellow and possible
      PII in red when there are some; other numbers white.
    .PARAMETER Channel
      The channel name (masked here).
    .PARAMETER Before
      A copy of the meter before the channel (Done, Read, Skipped, Hits).
    .PARAMETER Meter
      The meter after the channel.
    .PARAMETER Noun
      What an item is, for example 'file'.
    #>
    param([string]$Channel, [hashtable]$Before, [hashtable]$Meter, [string]$Noun)
    $hits = $Meter.Hits - $Before.Hits
    $skipped = $Meter.Skipped - $Before.Skipped
    Write-Ui ('  ' + (Format-Name (Protect-PiiName $Channel) (Get-RowWidth $Meter))) 'Folder' -NoNewline
    Write-Ui ' | ' 'Rule' -NoNewline
    Write-Ui ($Noun + 's ') 'Muted' -NoNewline
    # Padded to the width of the total row's "done/total percent", so the next columns line up.
    Write-Ui ((Format-Count ($Meter.Done - $Before.Done)).PadRight($script:CountWidth * 2 + 6)) 'Text' -NoNewline
    Write-CountPart 'read' ($Meter.Read - $Before.Read) 'Text'
    Write-CountPart 'not scanned' $skipped (Get-CountKind $skipped 'Accent')
    Write-CountPart 'possible PII' $hits (Get-CountKind $hits 'Alert')
    Write-Ui ''
}

function Write-RuleLine {
    # A grey line that ends one folder's block on screen.
    Write-Ui ('  ' + ('-' * $script:RuleWidth)) 'Rule'
}

Export-ModuleMember -Function Format-BoxTable, Get-TableLines, Get-ScreenWidth,
    Format-Duration, New-ScanMeter, Update-ScanMeter, Get-ScanMeterText, Get-MeterPercent, Get-NameWidth, Write-MeterLine, Test-FileBarSize,
    Write-FileStep, Copy-StreamWithProgress, Write-ChannelLine, Write-RuleLine
