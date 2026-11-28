<#
ReaderPool.psm1 - reads several files at once in PowerShell runspaces, and each identical file only once per run.

  Purpose:     Speed (the Administrator, 2026-11-21): files are read and checked for PII by 1 to 8 workers at a time (Normal speed: at most half
               the processors; Fast: one per processor; the Administrator, 2026-11-25), while the
               main thread lists, downloads (SharePoint) and reports. The number of workers follows the load: it
               grows while the throughput grows and shrinks when it falls. A file whose content (SHA-256) was
               already read in this run is not read again: it gets the same hits. Nothing is kept after the run.
  Inputs:      jobs: a file's bytes (SharePoint) or its local or network path (drives), and its name.
  Outputs:     results: Status, Reason, Note, Hits, Parts, Hash, Duplicate, Seconds, Size.
  How to run:  imported by Find-ChannelPII.ps1; each worker loads the same reader code (the readers above this
               module and this module) once.
  Depends on:  PiiEngine.psm1 (Find-PiiInText) and TextExtract.psm1 (Get-TextFromBytes), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7). Runspaces are part of PowerShell: no module, no
               compiled code, no other process.
  Network:     none here (a network-drive path is opened read-only with the person's Windows logon).
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# Workers: at most this many (and at most one per processor); the first target is 2.
$script:MaxWorkers = 8
$script:FirstWorkers = 2
# The run's speed (the Administrator, 2026-11-25: a Fast run made the fan loud): 'Normal' uses at most half the processors for
# readers; 'Fast' one per processor. In both, workers run at below-normal Windows priority, so other programs come first.
$script:ScanSpeed = 'Normal'

# The processors of this computer (tests set another count, so a Fast test holds on a runner with 1 processor).
$script:Processors = [Environment]::ProcessorCount
# The target is reviewed after every this many finished files: up one when the MB per second rose by more than
# $GrowthStep, down one when it fell by more than $DropStep.
$script:ReviewEvery = 6
$script:GrowthStep = 0.05
$script:DropStep = 0.15
# The longest one file may take in a worker (seconds). Above the PDF reader's own limit (120 s, PdfExtract.psm1), so that
# limit, with its own reason, fires first. A worker over it is stopped without waiting, and the scan goes on (the Administrator's
# gov run, 2026-11-23: a SharePoint scan hung after 50 minutes and Ctrl+C did not stop it).
$script:FileSecondsLimit = 180
$script:FileSlowReason = 'Reading took longer than 3 minutes (the file may be fine): check by hand'
$script:LastInit = $null      # the reader set-up of the last New-ReaderPool (Invoke-LimitedRead uses it)
$script:LimitedRunspace = $null
# How long the main thread waits for a worker between screen updates (milliseconds).
$script:WaitSliceMs = 250
# Paths longer than this are opened with the \\?\ prefix (Windows' classic limit is 260 characters).
$script:LongPathAt = 240

# The reader modules, in load order, for a development run (the bundle gives the reader block's text instead).
$script:ReaderModules = @('PiiEngine', 'ContainerText', 'PdfExtract', 'TextExtract', 'ReaderPool')

# The scripts each runspace runs (a worker per file; Invoke-LimitedRead per old version): load the reader code, then read.
$script:ReaderLoader = @'
[System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal
if (-not $global:IonReaderReady) {
    if ($Init.ContainsKey('Ast')) {
        New-Module -Name IonReaders -ScriptBlock ($Init.Ast.GetScriptBlock()) | Import-Module -Global
    }
    elseif ($Init.ContainsKey('Paths')) { Import-Module $Init.Paths -Global -Force }
    else { throw 'Worker reader set-up is missing Ast and Paths.' }
    $global:IonReaderReady = $true
}
'@
$script:PoolWorkerScript = "param(`$Init, `$Job, `$Seen)`n" + $script:ReaderLoader + "`nInvoke-PooledRead `$Job `$Seen"
# The script that reads one old SharePoint version in its own runspace (Invoke-LimitedRead), so a stuck read cannot
# hang the main thread (review of PR 58). It loads the reader code once per runspace, as a worker does.
$script:LimitedReadScript = "param(`$Init, `$Bytes, `$Name)`n" + $script:ReaderLoader + "`nGet-TextFromBytes -Bytes `$Bytes -FileName `$Name"

function Read-LocalFileBytes {
    <#
    .SYNOPSIS
      The bytes of one local or network file, opened read-only and shared (others may keep it open and write):
      the file is never locked, changed or touched otherwise. Long paths get the \\?\ prefix.
    .PARAMETER Path
      The full path (X:\... or \\server\share\...).
    .OUTPUTS
      [byte[]]
    #>
    param([string]$Path)
    $open = $Path
    if ($Path.Length -gt $script:LongPathAt -and -not $Path.StartsWith('\\?\')) {
        if ($Path.StartsWith('\\')) { $open = '\\?\UNC\' + $Path.Substring(2) } else { $open = '\\?\' + $Path }
    }
    $stream = New-Object System.IO.FileStream($open, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]'ReadWrite, Delete')
    try {
        $memory = New-Object System.IO.MemoryStream
        try { $stream.CopyTo($memory); return , $memory.ToArray() } finally { $memory.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Get-ContentHash {
    <#
    .SYNOPSIS
      The SHA-256 of a file's bytes as hex text, with the FIPS-validated Windows provider (DoD computers enforce
      FIPS mode); empty when it is not available, and then the file is simply read.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    try {
        $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
        try { return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') } finally { $sha.Dispose() }
    }
    catch { return '' }
}

function Invoke-PooledRead {
    <#
    .SYNOPSIS
      Reads one file and checks its text for PII (runs in a worker). A file whose content another job of this run
      already claimed is not read again: Duplicate, with the Id of that job.
    .PARAMETER Job
      @{ Id; Name; Bytes (or $null); Path (when Bytes is $null) }.
    .PARAMETER Seen
      A synchronized hashtable shared by all workers: content hash and extension -> the Id of the job that reads it.
    .OUTPUTS
      [hashtable] Id, Status, Reason, Note, Hits, Parts, Hash, Duplicate, Original, Seconds, Size, Props.
    #>
    param([hashtable]$Job, [hashtable]$Seen)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $result = @{ Id = $Job.Id; Status = 'OK'; Reason = ''; Note = ''; Hits = @(); Parts = @(); Hash = ''; Duplicate = $false
        Original = $null; Seconds = 0.0; Size = 0; Props = $null }
    try {
        $bytes = $Job.Bytes
        if ($null -eq $bytes) { $bytes = Read-LocalFileBytes $Job.Path }
        $result.Size = $bytes.Length
        # Same content and same extension (the extension decides the reader): read once.
        $hash = Get-ContentHash $bytes
        if ($hash) { $result.Hash = $hash + '|' + [System.IO.Path]::GetExtension($Job.Name).ToLowerInvariant() }
        if ($result.Hash -and $null -ne $Seen) {
            $claimed = $false
            [System.Threading.Monitor]::Enter($Seen.SyncRoot)
            try { if ($Seen.ContainsKey($result.Hash)) { $result.Original = $Seen[$result.Hash] } else { $Seen[$result.Hash] = $Job.Id; $claimed = $true } }
            finally { [System.Threading.Monitor]::Exit($Seen.SyncRoot) }
            if (-not $claimed) { $result.Duplicate = $true; return $result }
        }
        $text = Get-TextFromBytes -Bytes $bytes -FileName $Job.Name
        $result.Status = $text.Status; $result.Reason = $text.Reason; $result.Note = $text.Note
        if ($text.Status -eq 'OK') { $result.Hits = @(Find-PiiInText $text.Text) }
        if ($result.Hits.Count -gt 0) { $result.Props = Get-OfficeProps $bytes }   # who made and changed it (Office files)
        # No @() around the list itself: @() around an empty generic list property fails in PowerShell.
        if ($null -ne $text.Parts -and $text.Parts.Count -gt 0) { $result.Parts = $text.Parts.ToArray() }
    }
    catch {
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file (in use, damaged or protected?): ' + $_.Exception.Message
    }
    finally { $result.Seconds = $clock.Elapsed.TotalSeconds }
    return $result
}

function Set-ScanSpeed {
    <#
    .SYNOPSIS
      Sets the run's speed: 'Normal' (half the processors for readers) or 'Fast' (one per processor).
    .PARAMETER Speed
      'Normal' or 'Fast'.
    #>
    param([ValidateSet('Normal', 'Fast')][string]$Speed)
    $script:ScanSpeed = $Speed
}

function Get-ScanSpeed {
    # The run's speed: 'Normal' or 'Fast'.
    return $script:ScanSpeed
}

function Get-ReaderLimit {
    <#
    .SYNOPSIS
      The most readers at one time: Fast = one per processor; Normal = half the processors (rounded down, at least 1);
      never more than $MaxWorkers.
    .PARAMETER Speed
      'Normal' or 'Fast'.
    .PARAMETER Processors
      The number of processors.
    .OUTPUTS
      [int]
    #>
    param([string]$Speed, [int]$Processors)
    $limit = $Processors
    if ($Speed -ne 'Fast') { $limit = [Math]::Floor($Processors / 2) }
    return [int][Math]::Max(1, [Math]::Min($script:MaxWorkers, $limit))
}

function New-ReaderPool {
    <#
    .SYNOPSIS
      Opens the worker pool: up to Get-ReaderLimit runspaces (the run's speed), $FirstWorkers busy at first.
    .PARAMETER Init
      How a worker loads the reader code: @{ Ast = $block.Ast } (the paste bundle) or @{ Paths = <module files> }; when
      empty, the files of the reader modules loaded now (a development run).
    .OUTPUTS
      [hashtable] the pool state for Add-PoolJob, Receive-PoolResult and Close-ReaderPool.
    #>
    param([hashtable]$Init)
    if (-not $Init) { $Init = @{ Paths = @(Get-Module $script:ReaderModules | ForEach-Object { $_.Path }) } }
    $script:LastInit = $Init
    $max = Get-ReaderLimit $script:ScanSpeed $script:Processors
    $pool = [runspacefactory]::CreateRunspacePool(1, $max)
    $pool.Open()
    return @{ Pool = $pool; Init = $Init; Jobs = (New-Object System.Collections.Generic.List[object]); Max = $max
        Target = [Math]::Min($script:FirstWorkers, $max); Seen = [hashtable]::Synchronized(@{}); Done = 0; Stuck = 0
        WindowBytes = 0L; WindowClock = [Diagnostics.Stopwatch]::StartNew(); LastRate = 0.0 }
}

function Add-PoolJob {
    <#
    .SYNOPSIS
      Hands one file to a worker; first waits (and returns the finished results) while the target number of
      workers is busy.
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Job
      @{ Id; Name; Bytes or Path } (see Invoke-PooledRead).
    .PARAMETER OnWait
      { param($busy) } called before each wait slice (the screen shows the wait); optional.
    .OUTPUTS
      The results that finished while waiting (Receive-PoolResult).
    #>
    param([hashtable]$State, [hashtable]$Job, [scriptblock]$OnWait)
    while ($State.Jobs.Count -ge $State.Target) { Receive-PoolResult $State -Wait -OnWait $OnWait }
    $shell = [powershell]::Create()
    $shell.RunspacePool = $State.Pool
    [void]$shell.AddScript($script:PoolWorkerScript).AddArgument($State.Init).AddArgument($Job).AddArgument($State.Seen)
    $State.Jobs.Add(@{ Shell = $shell; Handle = $shell.BeginInvoke(); Id = $Job.Id; Clock = [Diagnostics.Stopwatch]::StartNew() })
}

function Receive-PoolResult {
    <#
    .SYNOPSIS
      The results of the workers that have finished (with -Wait: waits until at least one has, a slice at a time).
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Wait
      Wait for at least one result when none is ready (and jobs are running).
    .PARAMETER OnWait
      { param($busy) } called before each wait slice with the number of busy workers; optional.
    .OUTPUTS
      [hashtable] results of Invoke-PooledRead.
    #>
    param([hashtable]$State, [switch]$Wait, [scriptblock]$OnWait)
    $slow = 0
    while ($true) {
        $ready = @($State.Jobs | Where-Object { $_.Handle.IsCompleted })
        # A worker over the time limit: stopped without waiting; its file gets a NOT SCANNED result now.
        foreach ($job in @($State.Jobs | Where-Object { -not $_.Handle.IsCompleted -and $_.Clock.Elapsed.TotalSeconds -gt $script:FileSecondsLimit })) {
            Stop-PoolJob $State $job
            $slow++
            $result = @{ Id = $job.Id; Status = 'NOT SCANNED'; Reason = $script:FileSlowReason }
            Update-PoolTarget $State $result
            $result
        }
        if ($ready.Count -gt 0 -or $slow -gt 0 -or -not $Wait -or $State.Jobs.Count -eq 0) { break }
        if ($OnWait) { & $OnWait $State.Jobs.Count }
        [void]$State.Jobs[0].Handle.AsyncWaitHandle.WaitOne($script:WaitSliceMs)
    }
    foreach ($job in $ready) {
        [void]$State.Jobs.Remove($job)
        try {
            $output = @($job.Shell.EndInvoke($job.Handle))
            $result = $output | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
            if (-not $result) { $result = @{ Id = $job.Id; Status = 'NOT SCANNED'; Reason = 'The reader gave no result: check by hand' } }
        }
        catch { $result = @{ Id = $job.Id; Status = 'NOT SCANNED'; Reason = 'The reader failed: ' + $_.Exception.Message } }
        finally { $job.Shell.Dispose() }
        Update-PoolTarget $State $result
        $result
    }
}

function Update-PoolTarget {
    <#
    .SYNOPSIS
      Counts one finished file and, every $ReviewEvery files, moves the number of busy workers by one: up while the
      MB per second grows, down when it falls or a file could not be read.
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Result
      The finished result.
    #>
    param([hashtable]$State, [hashtable]$Result)
    $State.Done++
    if ($Result.ContainsKey('Size')) { $State.WindowBytes += [long]$Result.Size }
    $failed = $Result.Status -ne 'OK' -and $Result.Reason -match '^Could not read the file|^The reader'
    if ($failed -and $State.Target -gt 1) { $State.Target--; return }
    if ($State.Done % $script:ReviewEvery -ne 0) { return }
    $rate = $State.WindowBytes / [Math]::Max(0.001, $State.WindowClock.Elapsed.TotalSeconds)
    if ($rate -gt $State.LastRate * (1 + $script:GrowthStep) -and $State.Target -lt $State.Max) { $State.Target++ }
    elseif ($rate -lt $State.LastRate * (1 - $script:DropStep) -and $State.Target -gt 1) { $State.Target-- }
    $State.LastRate = $rate
    $State.WindowBytes = 0L
    $State.WindowClock.Restart()
}

function Stop-PoolJob {
    <#
    .SYNOPSIS
      Stops one worker without waiting for it: a worker stuck inside a .NET call does not stop, and a wait for it would
      hang the scan (and Ctrl+C). The pool gets one more worker in its place; the stuck one ends with the pool.
    .PARAMETER State
      The pool of New-ReaderPool.
    .PARAMETER Job
      The job (from $State.Jobs).
    #>
    param([hashtable]$State, [hashtable]$Job)
    [void]$State.Jobs.Remove($Job)
    try { [void]$Job.Shell.BeginStop($null, $null) } catch { Write-Verbose 'A worker had already stopped.' }   # nothing left to stop
    $State.Stuck++
    [void]$State.Pool.SetMaxRunspaces($State.Max + $State.Stuck)
}

function Invoke-LimitedRead {
    <#
    .SYNOPSIS
      Get-TextFromBytes in a runspace of its own, with the per-file time limit ($FileSecondsLimit): for reads on the
      main thread (old SharePoint versions). The wait is in short slices, so Ctrl+C still stops the scan. Over the
      limit the runspace is stopped without waiting and left; the next read gets a new one.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER Name
      The file name (for its reader).
    .OUTPUTS
      Output of Get-TextFromBytes. Throws [TimeoutException] over the limit.
    #>
    param([byte[]]$Bytes, [string]$Name)
    if (-not $script:LastInit) { return Get-TextFromBytes -Bytes $Bytes -FileName $Name }   # no scan set-up: read here
    if (-not $script:LimitedRunspace) { $script:LimitedRunspace = [runspacefactory]::CreateRunspace(); $script:LimitedRunspace.Open() }
    $shell = [powershell]::Create()
    $shell.Runspace = $script:LimitedRunspace
    [void]$shell.AddScript($script:LimitedReadScript).AddArgument($script:LastInit).AddArgument($Bytes).AddArgument($Name)
    $handle = $shell.BeginInvoke()
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while (-not $handle.AsyncWaitHandle.WaitOne($script:WaitSliceMs)) {
        if ($clock.Elapsed.TotalSeconds -gt $script:FileSecondsLimit) {
            try { [void]$shell.BeginStop($null, $null) } catch { Write-Verbose 'The read had already stopped.' }   # nothing to stop
            $script:LimitedRunspace = $null
            throw (New-Object System.TimeoutException $script:FileSlowReason)
        }
    }
    try { return (@($shell.EndInvoke($handle)) | Select-Object -Last 1) } finally { $shell.Dispose() }
}

function Close-ReaderPool {
    <#
    .SYNOPSIS
      Stops the workers that still run (after a cancel or an error) and closes the pool, without waiting for a stuck
      worker (Stop-PoolJob): a cancel always ends.
    .PARAMETER State
      The pool of New-ReaderPool.
    #>
    param([hashtable]$State)
    $script:LastInit = $null   # no scan now: a later old-version read is done on the calling thread
    $limited = $script:LimitedRunspace
    $script:LimitedRunspace = $null
    # A Runspace closes with CloseAsync (BeginClose is a RunspacePool method: review of PR 58); it never stops the close.
    if ($limited) { try { $limited.CloseAsync() } catch { Write-Verbose 'The old-version reader was already closed.' } }
    foreach ($job in $State.Jobs.ToArray()) {
        if ($job.Handle.IsCompleted) { $job.Shell.Dispose(); [void]$State.Jobs.Remove($job) }
        else { Stop-PoolJob $State $job }
    }
    if ($State.Stuck -gt 0) { [void]$State.Pool.BeginClose($null, $null); return }   # it closes when the stuck workers end
    $State.Pool.Close()
    $State.Pool.Dispose()
}

Export-ModuleMember -Function Read-LocalFileBytes, Get-ContentHash, Invoke-PooledRead, New-ReaderPool, Add-PoolJob, Receive-PoolResult, Close-ReaderPool,
    Invoke-LimitedRead, Set-ScanSpeed, Get-ScanSpeed, Get-ReaderLimit
