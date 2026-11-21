<#
Accountability.psm1 - who made and changed a file with PII, and who put the PII in and when.

  Purpose:     The PAM, 2026-11-21: for accountability and to date a possible incident, each file with a hit shows
               who last modified it, when it was made and by whom, and who added the PII and when, by e-mail
               address (the report is CUI and stays in DoD365). SharePoint and Teams: the file's fields and its
               version history (the first version that holds each hit). Drives: the file's owner (a Windows logon
               name, which Windows looks up from the owner's ID; no search of the directory: "No reconnaissance"), its times, and the names an Office file stores.
  Inputs:      the file's list item, path or reader result, and its hits.
  Outputs:     @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; ByHit } for Report.psm1 Set-RowAccountability.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  TextExtract.psm1 (Get-TextFromBytes), PiiDetectors.psm1 (Find-PiiInText) and, for SharePoint,
               PnP.PowerShell 1.12.0 (Get-PnPFileVersion, Get-PnPContext; read-only), imported first.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the signed-in site (version history of files with a hit).
  Writes:      none. Old versions are read in memory only.
  Credentials: none (the PnP sign-in of the scan).
#>

Set-StrictMode -Version 2.0

# Old versions read per file with a hit, oldest first (a file with more is dated from the oldest ones read).
$script:MaxVersionsRead = 30
$script:NoHistory = 'Not recorded (no version history)'
# Put-backs named per hit in "PII added by" (the rest are counted).
$script:MaxPutBacks = 3
$script:MaybeEarlier = ' (an older version could not be read: it may have been added earlier)'
$script:PartHistory = "Not recorded (not in the old versions read; check the file's version history by hand)"
$script:DriveNoHistory = 'Not recorded (a drive keeps no file history)'

function Get-PersonText {
    <#
    .SYNOPSIS
      A person field as text: the e-mail address of a SharePoint user value, else its name, else the text itself.
    .PARAMETER Value
      A field value (FieldUserValue, User or text).
    .OUTPUTS
      [string]
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    foreach ($name in @('Email', 'LookupValue', 'Title')) {
        $property = $Value.PSObject.Properties[$name]
        if ($property -and $property.Value) { return [string]$property.Value }
    }
    return [string]$Value
}

function Get-DateText {
    # A date as yyyy-MM-dd HH:mm (the report's form); empty when missing.
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Get-SharePointWho {
    <#
    .SYNOPSIS
      Who made and last changed a SharePoint file, from its list item (Author, Created, Editor).
    .PARAMETER File
      The list item.
    .OUTPUTS
      [hashtable] ModifiedBy, Created, CreatedBy.
    #>
    param($File)
    return @{ ModifiedBy = (Get-PersonText $File['Editor']); Created = (Get-DateText $File['Created']); CreatedBy = (Get-PersonText $File['Author']) }
}

function Get-HitKey {
    # The same hit in two versions: its type and masked value.
    param($Hit)
    return ([string]$Hit.Detector + '|' + [string]$Hit.MaskedValue)
}

function Complete-CurrentVersion {
    <#
    .SYNOPSIS
      The current version's part: a hit missing from the newest old version (all of them read) was put back now; a
      hit in no old version came with the current one (with no old version: the file's maker) - only when every old
      version was read; else who added it is not known, and the report says so (review of PR 20).
    .PARAMETER ByHit
      Hit key -> @{ AddedOn; AddedBy }.
    .PARAMETER Back
      Hit key -> put-back texts.
    .PARAMETER Present
      Hit key -> in the newest old version read.
    .PARAMETER Keys
      The hit keys.
    .PARAMETER Who
      Get-SharePointWho of the file, with Modified.
    .PARAMETER Read
      @{ AllRead; Complete; Versions }: how much of the history was read.
    #>
    param([hashtable]$ByHit, [hashtable]$Back, [hashtable]$Present, [object[]]$Keys, [hashtable]$Who, [hashtable]$Read)
    if ($Read.AllRead) {
        foreach ($key in @($Keys | Where-Object { $ByHit.ContainsKey($_) -and $Present.ContainsKey($_) -and $Present[$_] -eq $false })) {
            Add-PutBack $Back $key $Who.ModifiedBy $Who.Modified
        }
    }
    foreach ($key in @($Keys | Where-Object { -not $ByHit.ContainsKey($_) })) {
        if (-not $Read.Complete) { $ByHit[$key] = @{ AddedOn = $script:PartHistory; AddedBy = $script:PartHistory } }
        elseif ($Read.Versions -eq 0) { $ByHit[$key] = @{ AddedOn = $Who.Created; AddedBy = $Who.CreatedBy } }
        else { $ByHit[$key] = @{ AddedOn = $Who.Modified; AddedBy = $Who.ModifiedBy } }
    }
}

function Get-OldVersions {
    <#
    .SYNOPSIS
      A file's old versions, oldest first (at most $MaxVersionsRead), with who saved each (read-only); $null when
      the history cannot be read.
    .PARAMETER Ref
      The file's server-relative path.
    .OUTPUTS
      [hashtable] Versions, All (how many there are), Context; or $null.
    #>
    param([string]$Ref)
    try {
        $all = @(Get-PnPFileVersion -Url $Ref -ErrorAction Stop | Sort-Object Created)
        $versions = @($all | Select-Object -First $script:MaxVersionsRead)
        $context = Get-PnPContext
        foreach ($version in $versions) { $context.Load($version.CreatedBy) }
        if ($versions.Count -gt 0) { $context.ExecuteQuery() }
        return @{ Versions = $versions; All = $all.Count; Context = $context }
    }
    catch { return $null }
}

function Read-VersionText {
    <#
    .SYNOPSIS
      The text of one old version of a file (read-only, in memory), or $null when it cannot be read or read in full.
    .PARAMETER Version
      A FileVersion of Get-PnPFileVersion.
    .PARAMETER Context
      The PnP client context.
    .PARAMETER Name
      The file name (for its reader).
    .OUTPUTS
      Output of Get-TextFromBytes, or $null.
    #>
    param($Version, $Context, [string]$Name)
    try {
        $stream = $Version.OpenBinaryStream(); $Context.ExecuteQuery()
        $memory = New-Object System.IO.MemoryStream
        try { $stream.Value.CopyTo($memory); $bytes = $memory.ToArray() } finally { $memory.Dispose() }
        $text = Get-TextFromBytes -Bytes $bytes -FileName $Name
    }
    catch { return $null }
    if ($text.Status -ne 'OK') { return $null }
    return $text
}

function Add-PutBackText {
    <#
    .SYNOPSIS
      Adds the put-backs to "PII added by": "first adder; put back by X on D" (at most $MaxPutBacks named, the rest
      counted), and notes what the history read cannot show.
    .PARAMETER ByHit
      Hit key -> @{ AddedOn; AddedBy }.
    .PARAMETER Back
      Hit key -> the put-back texts (Add-PutBack).
    .PARAMETER Unsure
      Hit keys with a version not read after they were first seen: a removal and put-back there cannot be named.
    .PARAMETER AllRead
      Whether every old version was read (else later ones were not checked).
    #>
    param([hashtable]$ByHit, [hashtable]$Back, [hashtable]$Unsure, [bool]$AllRead)
    foreach ($key in $Back.Keys) {
        $list = $Back[$key]
        $text = @($list | Select-Object -First $script:MaxPutBacks) -join '; '
        $more = $list.Count - $script:MaxPutBacks
        if ($more -eq 1) { $text += '; and 1 more time' } elseif ($more -gt 1) { $text += ('; and {0} more times' -f $more) }
        $ByHit[$key].AddedBy = $ByHit[$key].AddedBy + '; ' + $text
    }
    # Say what the history read cannot show, so "PII added by" never reads as complete when it is not (review).
    foreach ($key in @($ByHit.Keys | Where-Object { $ByHit[$_].AddedBy -ne $script:PartHistory })) {
        if ($Unsure.ContainsKey($key)) { $ByHit[$key].AddedBy += '; it may have been removed and put back (a version could not be read)' }
        if (-not $AllRead) { $ByHit[$key].AddedBy += "; versions after the $script:MaxVersionsRead read were not checked" }
    }
}

function Add-PutBack {
    # Notes that a hit was put back (in a version after one without it): who saved that version and when.
    param([hashtable]$Back, [string]$Key, [string]$By, [string]$On)
    if (-not $Back.ContainsKey($Key)) { $Back[$Key] = New-Object System.Collections.Generic.List[string] }
    $Back[$Key].Add(('put back by {0} on {1}' -f $By, $On))
}

function Find-PiiAddedVersion {
    <#
    .SYNOPSIS
      For each hit of a SharePoint file, the first version that holds it (who saved it and when), and each time it
      was put back after a version without it (the PAM, 2026-11-21). A hit
      in no older version came with the current one (the last modifier). Read-only; versions stay in memory.
    .PARAMETER Ref
      The file's server-relative path.
    .PARAMETER Name
      The file name (for its reader).
    .PARAMETER Hits
      The current hits (Detector, MaskedValue).
    .PARAMETER Who
      Get-SharePointWho of the file, with Modified (the current version's time).
    .OUTPUTS
      [hashtable] hit key -> @{ AddedOn; AddedBy }.
    #>
    param([string]$Ref, [string]$Name, [object[]]$Hits, [hashtable]$Who)
    $byHit = @{}
    $open = @($Hits | ForEach-Object { Get-HitKey $_ } | Sort-Object -Unique)
    $history = Get-OldVersions $Ref
    if ($null -eq $history) {
        foreach ($key in $open) { $byHit[$key] = @{ AddedOn = $script:NoHistory; AddedBy = $script:NoHistory } }
        return $byHit
    }
    $versions = $history.Versions; $context = $history.Context
    $complete = $history.All -le $script:MaxVersionsRead   # every older version is read (else a later one may add a hit)
    $allRead = $complete
    $readFailed = $false   # a version before this one could not be read (the cap only leaves newer ones unread)
    $keys = $open
    $present = @{}   # per hit: in the last version read ($true/$false), or not known after a version not read ($null)
    $back = @{}      # per hit: when it was put back after a version without it (the PAM, 2026-11-21)
    $unsure = @{}    # per hit: a version not read after the hit was seen: a removal and put-back there cannot be named
    foreach ($version in $versions) {   # every version read (not only to the first with each hit): put-backs count too
        $text = Read-VersionText $version $context $Name
        if ($null -eq $text) {   # not read: later versions can neither be first nor a put-back for sure
            $complete = $false; $readFailed = $true
            foreach ($key in $keys) {
                if ($byHit.ContainsKey($key)) { $unsure[$key] = $true }   # seen before: it may be removed and put back in this one
                $present[$key] = $null
            }
            continue
        }
        $found = @(Find-PiiInText $text.Text | ForEach-Object { Get-HitKey $_ })
        $on = Get-DateText $version.Created; $by = Get-PersonText $version.CreatedBy
        foreach ($key in $keys) {
            $has = $found -contains $key
            if ($has -and -not $byHit.ContainsKey($key)) {
                # After a version that could not be read, this may not be the first one with the hit (review of PR 20).
                if (-not $readFailed) { $byHit[$key] = @{ AddedOn = $on; AddedBy = $by } }
                else { $byHit[$key] = @{ AddedOn = 'by ' + $on; AddedBy = $by + $script:MaybeEarlier } }
            }
            elseif ($has -and $present.ContainsKey($key) -and $present[$key] -eq $false) { Add-PutBack $back $key $by $on }
            $present[$key] = $has
        }
    }
    Complete-CurrentVersion $byHit $back $present $keys $Who @{ AllRead = $allRead; Complete = $complete; Versions = $versions.Count }
    Add-PutBackText $byHit $back $unsure $allRead
    return $byHit
}

function Get-DriveWho {
    <#
    .SYNOPSIS
      Who made and changed a file on a drive: its owner (usually its maker; a Windows logon name), its creation
      time, and the names an Office file stores (creator, last modified by). Windows does not record who last
      changed other files.
    .PARAMETER Path
      The file path.
    .PARAMETER Props
      The Office names of the reader result (Creator, LastModifiedBy), or $null.
    .OUTPUTS
      [hashtable] ModifiedBy, Created, CreatedBy, AddedOn, AddedBy.
    #>
    param([string]$Path, [hashtable]$Props)
    $owner = ''
    try { $owner = (Get-Acl -LiteralPath $Path -ErrorAction Stop).Owner } catch { $owner = '' }   # no owner where the system keeps none
    $created = ''
    try { $created = (Get-Item -LiteralPath $Path -ErrorAction Stop).CreationTime.ToString('yyyy-MM-dd HH:mm') } catch { $created = '' }
    $modifiedBy = 'Not recorded (Windows keeps no last editor)'
    $createdBy = $owner
    if ($Props) {
        if ($Props.LastModifiedBy) { $modifiedBy = $Props.LastModifiedBy + ' (as saved in the file)' }
        if (-not $createdBy -and $Props.Creator) { $createdBy = $Props.Creator + ' (as saved in the file)' }
    }
    return @{ ModifiedBy = $modifiedBy; Created = $created; CreatedBy = $createdBy; AddedOn = $script:DriveNoHistory; AddedBy = $script:DriveNoHistory }
}

Export-ModuleMember -Function Get-PersonText, Get-SharePointWho, Find-PiiAddedVersion, Get-DriveWho
