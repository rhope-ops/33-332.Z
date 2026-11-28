<#
ScanQueue.psm1 - progress meters, box tables, reader-pool queue feeding, and shared SharePoint/drive scan loop.

  Purpose:     Displays live progress bars and masked folder lines, feeds ReaderPool workers, collects results,
               populates accountability columns, and runs the Scan now? confirmation prompt.
  Inputs:      file lists, ReaderPool workers, target info.
  Outputs:     live progress UI, queued scan summary, report rows.
  Depends on:  PiiEngine.psm1, TextExtract.psm1, ReaderPool.psm1, ReportExport.psm1, Accountability.psm1, imported first.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none itself (callers supply SharePoint getters when scanning online).
  Writes:      none.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# ==== ScanMeter.psm1 ====
# The per-file bar is shown only from this size on: a small file is done in a moment, and each bar update costs time in
# Windows PowerShell 5.1 (the author's choice, 2026-11-20: faster scans of many small files).
$script:FileBarBytes = 1MB

# The per-file bar shows a speed and a time left only after this many seconds (before that the speed jumps about).
$script:RateAfterSeconds = 1

# A streamed download is read this much at a time (64 KB); the bar is updated at most this often (8 a second): smaller
# steps, so the bar moves smoothly (the Administrator, 2026-11-21: 256 KB and 4 a second looked choppy).
$script:ChunkBytes = 64KB
$script:BarRefreshMs = 125

# Screen tables fit the window: under 60 characters or unknown = 120; the last column keeps 12 (the author's choice, 2026-11-21).
$script:DefaultScreenWidth = 120
$script:MinScreenWidth = 60
$script:MinLastColumn = 12

# A coloured line of live counts is printed under the progress bar at most this often (and after each folder).
$script:LiveLineSeconds = 10

# Width of each count on the live lines: 5 characters (up to 99,999), so the columns do not move (the Administrator, 2026-11-20).
$script:CountWidth = 5

# Length of the grey rule after each folder's lines.
$script:RuleWidth = 100

# The name column of the folder rows: as wide as the longest folder name, at most this; the total row says "Total".
$script:MaxNameWidth = 40
$script:TotalName = 'Total'

# Short PII type names for the live line, in the Administrator's order (ReportExport.psm1 $SeverityOrder).
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
    # A count right-aligned in $CountWidth characters, so the columns of the live lines stay in place (the Administrator, 2026-11-20).
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
      The run so far as one line for the scan bar (its middle line; the Administrator, 2026-11-21): files done of the total with
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
    Write-Ui ('{0} / {1} {2,3}%' -f (Format-Count $Meter.Done), (Format-Count $Meter.Total), (Get-MeterPercent $Meter)) 'Text' -NoNewline
    Write-CountPart 'read' $Meter.Read 'Text'
    Write-CountPart 'not scanned' $Meter.Skipped (Get-CountKind $Meter.Skipped 'Accent')
    Write-CountPart 'possible PII' $Meter.Hits (Get-CountKind $Meter.Hits 'Alert')
    # No time here: the scan bar shows the time elapsed and left (the Administrator, 2026-11-21).
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
      large file does not look hung (the Administrator, 2026-11-21: the bar must move). Reading happens in the workers, after it.
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER Step
      What is being done now, for example 'Downloading...'.
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
    # No -ParentId: Windows PowerShell ISE indents a child bar; the Administrator wants both bars aligned (2026-11-20).
    Write-Progress -Id 2 -Activity $Step -Status $status -PercentComplete $percent
}

function Write-DownloadsStep {
    <#
    .SYNOPSIS
      The second progress bar while several files download at once (ChannelScan.psm1, the Administrator 2026-11-24): one bar for
      all files in flight, "3 files: 12.4 of 30.1 MB (41%) | 2.1 MB/s". No bar while the total is under 1 MB, as for one file.
    .PARAMETER Files
      The number of files downloading now.
    .PARAMETER Size
      The bytes of this burst: the files in flight and the ones that finished since the first of them started.
    .PARAMETER Done
      The bytes done of Size.
    .PARAMETER Elapsed
      The time since the burst started.
    .PARAMETER Step
      The bar's step text, for example 'Downloading..'.
    #>
    param([int]$Files, [long]$Size, [long]$Done, [TimeSpan]$Elapsed, [string]$Step)
    if (-not (Test-FileBarSize $Size)) { return }
    $noun = 'files'; if ($Files -eq 1) { $noun = 'file' }
    $percent = [int][Math]::Floor([Math]::Min(100.0, 100.0 * $Done / [Math]::Max(1, $Size)))
    $status = '{0} {1}: {2:N1} of {3:N1} MB ({4}%)' -f $Files, $noun, ($Done / 1MB), ($Size / 1MB), $percent
    if ($Done -gt 0 -and $Elapsed.TotalSeconds -ge $script:RateAfterSeconds) { $status += ' | {0:N1} MB/s' -f ($Done / $Elapsed.TotalSeconds / 1MB) }
    Write-Progress -Id 2 -Activity $Step -Status $status -PercentComplete $percent
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
      The bar's step text, for example 'Downloading...'.
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
    # Too wide for the screen: the last column (a file path) is cut, with '...', so no line wraps (the Administrator, 2026-11-21).
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
      then its counts in fixed widths with grey labels (the Administrator, 2026-11-20). Not scanned in yellow and possible
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
    # Padded to the width of the total row's "done / total percent", so the next columns line up.
    Write-Ui ((Format-Count ($Meter.Done - $Before.Done)).PadRight($script:CountWidth * 2 + 8)) 'Text' -NoNewline
    Write-CountPart 'read' ($Meter.Read - $Before.Read) 'Text'
    Write-CountPart 'not scanned' $skipped (Get-CountKind $skipped 'Accent')
    Write-CountPart 'possible PII' $hits (Get-CountKind $hits 'Alert')
    Write-Ui ''
}

function Write-RuleLine {
    # A grey line that ends one folder's block on screen.
    Write-Ui ('  ' + ('-' * $script:RuleWidth)) 'Rule'
}

# ==== ScanQueue.psm1 ====
# The "PII first saved" cells of a hit whose version history is not read (an e-signature DoD ID): "Not recorded", the
# word the By account tab leaves out, so no account is made up (review of PR 50); the History note says why.
$script:NotLookedUp = 'Not recorded'
$script:NotLookedUpNote = 'Not looked up: a DoD ID in an e-signature; its version history is not read'

function New-ScanQueue {
    <#
    .SYNOPSIS
      A queue for one scan: the reader pool, the report rows, the meter and what each file in work is.
    .PARAMETER Init
      How the workers load the reader code (New-ReaderPool).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .OUTPUTS
      [hashtable] the queue.
    #>
    param([hashtable]$Init, $Rows, [hashtable]$Meter)
    return @{ Pool = (New-ReaderPool $Init); Rows = $Rows; Meter = $Meter; Info = @{}; Results = @{}; Waiting = @{}; NextId = 0
        Seconds = 0.0; Activity = '' }
}

function Submit-ScanFile {
    <#
    .SYNOPSIS
      Hands one file to the readers; adds the rows of the files that finished meanwhile.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Info
      @{ Channel; Item; Link; Modified } as the report shows the file (already masked).
    .PARAMETER Name
      The file name (its extension decides how it is read).
    .PARAMETER Bytes
      The file bytes (SharePoint), or $null.
    .PARAMETER Path
      The file path when Bytes is $null (drives).
    #>
    param([hashtable]$Queue, [hashtable]$Info, [string]$Name, [byte[]]$Bytes, [string]$Path)
    $Queue.NextId++
    $id = $Queue.NextId
    $Queue.Info[$id] = $Info
    $job = @{ Id = $id; Name = $Name; Bytes = $Bytes; Path = $Path }
    # Keeps the scan bar moving while every worker is busy.
    $onWait = { param($busy) Write-WaitBar $Queue $busy }
    foreach ($result in @(Add-PoolJob $Queue.Pool $job $onWait)) { Resolve-ScanResult $Queue $result }
}

function Resolve-ScanResult {
    <#
    .SYNOPSIS
      Adds the rows of one finished file (and of the files with the same content that waited for it) and counts
      them on the meter.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Result
      A result of Invoke-PooledRead.
    #>
    param([hashtable]$Queue, [hashtable]$Result)
    if ($Result.ContainsKey('Seconds')) { $Queue.Seconds += [double]$Result.Seconds }
    $source = $Result
    if ($Result.ContainsKey('Duplicate') -and $Result.Duplicate) {
        if (-not $Queue.Results.ContainsKey($Result.Original)) {
            if (-not $Queue.Waiting.ContainsKey($Result.Original)) { $Queue.Waiting[$Result.Original] = New-Object System.Collections.Generic.List[int] }
            $Queue.Waiting[$Result.Original].Add($Result.Id)
            return
        }
        $source = $Queue.Results[$Result.Original]
    }
    else { $Queue.Results[$Result.Id] = $Result }
    Add-QueuedRows $Queue $Result.Id $source
    if ($Queue.Waiting.ContainsKey($Result.Id)) {
        foreach ($waiting in $Queue.Waiting[$Result.Id]) { Add-QueuedRows $Queue $waiting $Result }
        $Queue.Waiting.Remove($Result.Id)
    }
}

function Add-QueuedRows {
    <#
    .SYNOPSIS
      The rows of one file from a result (its own, or the one of a file with the same content), then the meter
      and the live line.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Id
      The file's job number.
    .PARAMETER Result
      The result whose hits the file gets.
    #>
    param([hashtable]$Queue, [int]$Id, [hashtable]$Result)
    $info = $Queue.Info[$Id]
    $Queue.Info.Remove($Id)
    $rows = $Queue.Rows
    $first = $rows.Count
    $read = 0
    if ($Result.Status -ne 'OK') {
        $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Reason)))
    }
    else {
        $read = 1
        # A file read in part (a PDF over a cap, damaged, or with attachments): a "check by hand" row next to its hits.
        if ($Result.Note) { $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified 'NOT SCANNED' '' '' (Protect-PiiText $Result.Note))) }
        foreach ($hit in $Result.Hits) {
            $rows.Add((New-ReportRow 'File' $info.Channel $info.Item $info.Link $info.Modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
        }
        foreach ($part in $Result.Parts) { Add-PartRows $part $info.Channel $info.Item $info.Link $info.Modified $rows }
        $nameHit = $info.ContainsKey('NameRows') -and $info.NameRows.Count -gt 0
        if ($Result.Hits.Count -gt 0 -or $rows.Count -gt $first -or $nameHit) { Add-WhoToRows $info $Result @($rows | Select-Object -Skip $first) }
    }
    Update-ScanMeter $Queue.Meter $rows $first $read
    Write-MeterLine $Queue.Meter 'file'
}

function Invoke-GroupScan {
    <#
    .SYNOPSIS
      Scans the groups (folders) one after the other: each file is started (its name checked, then handed to the
      readers or given a NOT SCANNED row), and the folder's line is written when all its files have their rows.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Groups
      The groups, each with Name and Files.
    .PARAMETER StartFile
      { param($Queue, $Group, $File) } that starts one file; returns true when the file went to the readers.
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    #>
    param([hashtable]$Queue, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile)
    $meter = $Queue.Meter
    for ($c = 0; $c -lt $Groups.Count; $c++) {
        $group = $Groups[$c]
        $before = $meter.Clone()
        $Queue.Activity = 'Folder {0} of {1}: {2}' -f ($c + 1), $Groups.Count, (Protect-PiiName $group.Name)
        # A scan with its own way to start a folder's files (ChannelScan.psm1: parallel downloads in a planned order).
        $files = $group.Files
        if ($Queue.ContainsKey('GroupScan') -and $Queue.GroupScan) { & $Queue.GroupScan $Queue $group; $files = @() }
        foreach ($file in $files) {
            # The file in work, with PII masked (the Administrator, 2026-11-20: a slow file is then visible).
            Write-ScanBar $Queue ('File: ' + (& $ShownFile $group $file))
            $firstRow = $Queue.Rows.Count
            if (-not (& $StartFile $Queue $group $file)) {
                Update-ScanMeter $meter $Queue.Rows $firstRow 0
                Write-MeterLine $meter 'file'   # the running total, at most every $LiveLineSeconds
            }
        }
        Wait-ScanQueue $Queue   # the folder's line counts all its files
        Write-ChannelLine $group.Name $before $meter 'file'
    }
    Write-Progress -Id 1 -Activity 'Scan' -Completed
}

function Add-WhoToRows {
    <#
    .SYNOPSIS
      The accountability columns of a file's rows (the Administrator, 2026-11-21): from its SharePoint fields and version
      history (Info.Who and Info.Added), or from the drive file (Info.DrivePath), plus the names an Office file stores.
    .PARAMETER Info
      The file's queue info.
    .PARAMETER Result
      Its reader result (Hits, Props).
    .PARAMETER FileRows
      The rows of this file.
    #>
    param([hashtable]$Info, [hashtable]$Result, [object[]]$FileRows)
    $props = $null; if ($Result.ContainsKey('Props')) { $props = $Result.Props }
    if ($Info.ContainsKey('DrivePath')) {
        $who = Get-DriveWho $Info.DrivePath $props
        if ($Info.ContainsKey('DriveNote')) {   # a OneDrive file not synced yet (DriveScan.psm1): its name-hit rows say so too
            $who.HistoryNote = $Info.DriveNote
            if ($Info.ContainsKey('NameRows')) { foreach ($row in $Info.NameRows) { $row.HistoryNote = $Info.DriveNote } }
        }
    }
    elseif ($Info.ContainsKey('Who')) {
        $who = $Info.Who.Clone()
        if ($Info.ContainsKey('Added') -and $Result.Hits.Count -gt 0) { $who.ByHit = Get-AddedByHit $Info $Result.Hits }
        # Plan step 5: who has access, for a web file with a hit in its text, a part (a file in a zip or an e-mail) or its name.
        # A loop, not '@()': '@()' of a New-Object list throws "Argument types do not match" (probe 6).
        $nameRows = @(); if ($Info.ContainsKey('NameRows')) { $nameRows = @(foreach ($row in $Info.NameRows) { $row }) }
        if ($Info.ContainsKey('Access') -and @(@($FileRows) + $nameRows | Where-Object { $_ -and (Get-ConfidenceRank $_.Confidence) -gt 0 }).Count -gt 0) {
            $access = & $Info.Access
            $who.OwnPermissions = $access.OwnPermissions; $who.AccessEntries = $access.AccessEntries
            foreach ($row in $nameRows) { $row.OwnPermissions = $access.OwnPermissions; $row.AccessEntries = $access.AccessEntries }
        }
    }
    else { return }
    Set-RowAccountability $FileRows $who
}

function Get-AddedByHit {
    <#
    .SYNOPSIS
      Who first saved each hit, from the version history, except for a DoD ID in an e-signature (the only Low DoD ID):
      its history is not read, so a file with only such hits costs no version reads (the Administrator, 2026-11-21).
    .PARAMETER Info
      The file's queue info (Added: the version-history lookup).
    .PARAMETER Hits
      The file's hits.
    .OUTPUTS
      [hashtable] "type|masked value" -> @{ AddedOn; AddedBy; Note }.
    #>
    param([hashtable]$Info, [object[]]$Hits)
    # True for an e-signature DoD ID (the only Low DoD ID).
    $signature = { param($hit) $hit.Detector -eq 'DoD ID' -and $hit.Confidence -eq 'Low' }
    $looked = @($Hits | Where-Object { -not (& $signature $_) })
    $byHit = @{}
    if ($looked.Count -gt 0) { $byHit = & $Info.Added $looked }
    foreach ($hit in @($Hits | Where-Object { & $signature $_ })) {
        $key = [string]$hit.Detector + '|' + [string]$hit.MaskedValue
        # The same masked value as a labelled DoD ID of this file keeps that one's history (review of PR 50).
        if ($byHit.ContainsKey($key)) { continue }
        $byHit[$key] = @{ AddedOn = $script:NotLookedUp; AddedBy = $script:NotLookedUp; Note = $script:NotLookedUpNote }
    }
    return $byHit
}

function Wait-ScanQueue {
    <#
    .SYNOPSIS
      Waits until every file handed in has its rows (the end of a folder, or of the scan).
    .PARAMETER Queue
      Output of New-ScanQueue.
    #>
    param([hashtable]$Queue)
    # Keeps the scan bar moving while the last files are read.
    $onWait = { param($busy) Write-WaitBar $Queue $busy }
    while ($Queue.Pool.Jobs.Count -gt 0) {
        foreach ($result in @(Receive-PoolResult $Queue.Pool -Wait -OnWait $onWait)) { Resolve-ScanResult $Queue $result }
    }
}

function Write-WaitBar {
    <#
    .SYNOPSIS
      The scan bar while every worker is busy or a folder waits for its last files: the time keeps moving and the bar
      says what it waits for, so the scan does not look hung (the Administrator, 2026-11-21).
    .PARAMETER Queue
      Output of New-ScanQueue.
    .PARAMETER Busy
      The number of files being read now.
    #>
    param([hashtable]$Queue, [int]$Busy)
    if (-not $Queue.Activity) { return }
    $noun = 'files'; if ($Busy -eq 1) { $noun = 'file' }
    Write-ScanBar $Queue ('Reading {0} {1} (a large or long file takes longer)...' -f $Busy, $noun)
}

function Write-ScanBar {
    <#
    .SYNOPSIS
      The scan bar: the folder, the run so far, then the file in work (CurrentOperation). ISE shows CurrentOperation
      after the status, and a line break in the status does not show (the Administrator's layout test, 2026-11-21); layout
      test 2 (docs/reviews/2026-11-21-ise-progress-test.md) decides how the file gets a line of its own.
    .PARAMETER Queue
      Output of New-ScanQueue (Activity: the folder line; Meter).
    .PARAMETER Third
      The third line: the file in work, or what the scan waits for.
    #>
    param([hashtable]$Queue, [string]$Third)
    Write-Progress -Id 1 -Activity $Queue.Activity -Status (Get-ScanMeterText $Queue.Meter 'file') -PercentComplete (Get-MeterPercent $Queue.Meter) `
        -CurrentOperation $Third
}

function Close-ScanQueue {
    <#
    .SYNOPSIS
      Closes the readers (also after a cancel: workers still running are stopped). Returns the reading time of all
      workers together, for the header.
    .PARAMETER Queue
      Output of New-ScanQueue.
    .OUTPUTS
      [TimeSpan]
    #>
    param([hashtable]$Queue)
    Close-ReaderPool $Queue.Pool
    return [TimeSpan]::FromSeconds($Queue.Seconds)
}

function Get-QueueWorkers {
    # The number of files read at the same time now (for the drive scan's header).
    param([hashtable]$Queue)
    return $Queue.Pool.Target
}

function Add-FileStartRows {
    <#
    .SYNOPSIS
      The rows every file gets before it is read, for both scans: a hit row for PII in its name or folder, then a NOT
      SCANNED row when its type is not read or it is over the size limit.
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Info
      @{ Channel; Item; Link; Modified } as the report shows the file (already masked); NameRows is added (its name-hit rows).
    .PARAMETER RawPath
      The folder and file path as written (checked for PII; never shown).
    .PARAMETER Name
      The file name (its extension decides whether it is read).
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER MaxFileMB
      The size limit.
    .OUTPUTS
      [bool] true when the file is to be read; false when it got its NOT SCANNED row.
    #>
    param($Rows, [hashtable]$Info, [string]$RawPath, [string]$Name, [long]$Size, [int]$MaxFileMB)
    $nameRows = New-Object System.Collections.Generic.List[object]
    foreach ($hit in @(Find-PiiInText $RawPath)) {
        $row = New-ReportRow 'File name' $Info.Channel $Info.Item $Info.Link $Info.Modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'
        $Rows.Add($row); $nameRows.Add($row)
    }
    $Info.NameRows = $nameRows   # the file's name-hit rows: they get its permission cells after the read (plan step 5)
    $reason = ''
    if (-not (Test-SupportedFile $Name)) { $reason = 'File type is not read by this script: check by hand' }
    elseif ($Size -gt ($MaxFileMB * 1MB)) { $reason = "Larger than $MaxFileMB MB: check by hand" }
    if ($reason) { $Rows.Add((New-ReportRow 'File' $Info.Channel $Info.Item $Info.Link $Info.Modified 'NOT SCANNED' '' '' $reason)) }
    return (-not $reason)
}

function Confirm-ScanStart {
    <#
    .SYNOPSIS
      Before a scan: how many files and MB were found, and how many will be listed but not read (type or size); then
      "Scan now? [Y/n]" (Enter = yes; the Administrator, 2026-11-21: a guard against a scope larger than meant).
    .PARAMETER Groups
      The groups, each with Files.
    .PARAMETER Facts
      { param($File) } that gives @(name, size in bytes) of one file.
    .PARAMETER MaxFileMB
      The size limit.
    .PARAMETER NoAsk
      Show the counts but do not ask: the person already said yes to the whole list (several targets in one run).
    .OUTPUTS
      [bool] true to scan.
    #>
    param([object[]]$Groups, [scriptblock]$Facts, [int]$MaxFileMB, [switch]$NoAsk)
    $count = 0; $bytes = 0L; $skipped = 0
    foreach ($group in $Groups) {
        foreach ($file in $group.Files) {
            $name, $size = & $Facts $file
            $count++; $bytes += [long]$size
            if (-not (Test-SupportedFile $name) -or [long]$size -gt $MaxFileMB * 1MB) { $skipped++ }
        }
    }
    if ($count -eq 0) { return $true }   # nothing to read: no question
    # A count with its noun: "1 file", "2 files".
    $plural = { param($n, $noun) if ($n -eq 1) { "1 $noun" } else { '{0:N0} {1}s' -f $n, $noun } }
    Write-Ui ('Found {0} ({1:N1} MB) in {2}; {3:N0} will be listed but not read (file type, or over {4} MB).' -f
        (& $plural $count 'file'), ($bytes / 1MB), (& $plural $Groups.Count 'folder'), $skipped, $MaxFileMB) 'Text'
    if ($NoAsk) { return $true }
    $answer = Read-Host 'Scan now? [Y/n]'
    return -not ($answer -and $answer.Trim() -match '^(n|no)$')
}

function Invoke-ScanRun {
    <#
    .SYNOPSIS
      One scan of the groups, for both scans: the "Scanning N files" banner, the meter, the readers, every group, then
      the rule and the total line. The run's meter, reading time and worker count go in Run, also when the scan
      stops part way (the caller's 'finally' reads them).
    .PARAMETER Init
      How the reader workers load their code (New-ReaderPool).
    .PARAMETER Rows
      The list that receives the rows.
    .PARAMETER Groups
      The groups, each with Name and Files.
    .PARAMETER StartFile
      { param($Queue, $Group, $File) } that starts one file (Invoke-GroupScan).
    .PARAMETER ShownFile
      { param($Group, $File) } that gives the file's masked path for the progress bar.
    .PARAMETER Run
      Receives Meter, Reading ([TimeSpan]) and Workers; Declined = $true when the person answered no to "Scan now?".
      AutoStart = $true (set by the caller): the counts show, but "Scan now?" is not asked. GroupScan (optional, set by
      the caller): { param($Queue, $Group) } that starts all files of one folder in place of the per-file loop.
    .PARAMETER Facts
      { param($File) } that gives @(name, size in bytes) of one file, for the question before the scan.
    .PARAMETER MaxFileMB
      The size limit (files over it are listed, not read).
    #>
    param([hashtable]$Init, $Rows, [object[]]$Groups, [scriptblock]$StartFile, [scriptblock]$ShownFile, [hashtable]$Run,
          [scriptblock]$Facts, [int]$MaxFileMB)
    $total = 0
    foreach ($group in $Groups) { $total += $group.Files.Count }
    $Run.Meter = New-ScanMeter 0; $Run.Reading = [TimeSpan]::Zero; $Run.Workers = 0; $Run.Declined = $false
    # Declined until a yes: a cancel (Ctrl+C) at the question saves nothing either, as nothing was read (review of PR 37).
    if ($Facts) { $Run.Declined = $true; if (-not (Confirm-ScanStart $Groups $Facts $MaxFileMB -NoAsk:($Run.ContainsKey('AutoStart') -and $Run.AutoStart))) { return }; $Run.Declined = $false }
    Write-Banner ('Scanning {0} files in {1} folders' -f $total, $Groups.Count)
    $meter = New-ScanMeter $total
    $meter.NameWidth = Get-NameWidth @($Groups | ForEach-Object { Protect-PiiName $_.Name })
    $Run.Meter = $meter; $Run.Reading = [TimeSpan]::Zero; $Run.Workers = 0
    $queue = New-ScanQueue $Init $Rows $meter
    if ($Run.ContainsKey('GroupScan')) { $queue.GroupScan = $Run.GroupScan }   # { param($Queue, $Group) }: see Invoke-GroupScan
    try { Invoke-GroupScan $queue $Groups $StartFile $ShownFile; $Run.Workers = Get-QueueWorkers $queue }
    finally { $Run.Reading = Close-ScanQueue $queue }
    # The folder rows form one table; the total closes it, under a rule.
    Write-RuleLine
    Write-MeterLine $meter 'file' -Force
}

Export-ModuleMember -Function Format-BoxTable, Get-TableLines, Get-ScreenWidth, Format-Duration, New-ScanMeter, Update-ScanMeter, Get-ScanMeterText, Get-MeterPercent, Get-NameWidth,
    Write-MeterLine, Test-FileBarSize, Write-FileStep, Write-DownloadsStep, Copy-StreamWithProgress, Write-ChannelLine, Write-RuleLine, New-ScanQueue, Submit-ScanFile, Invoke-GroupScan,
    Wait-ScanQueue, Close-ScanQueue, Get-QueueWorkers, Add-FileStartRows, Invoke-ScanRun, Write-ScanBar
