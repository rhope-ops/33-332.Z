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

function Find-PiiAddedVersion {
    <#
    .SYNOPSIS
      For each hit of a SharePoint file, the first version that holds it: who saved that version and when. A hit
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
    try {
        $all = @(Get-PnPFileVersion -Url $Ref -ErrorAction Stop | Sort-Object Created)
        $versions = @($all | Select-Object -First $script:MaxVersionsRead)
        $context = Get-PnPContext
        foreach ($version in $versions) { $context.Load($version.CreatedBy) }
        if ($versions.Count -gt 0) { $context.ExecuteQuery() }
    }
    catch { $versions = $null }
    if ($null -eq $versions) {
        foreach ($key in $open) { $byHit[$key] = @{ AddedOn = $script:NoHistory; AddedBy = $script:NoHistory } }
        return $byHit
    }
    $complete = $all.Count -le $script:MaxVersionsRead   # every older version is read (else a later one may add a hit)
    foreach ($version in $versions) {
        if ($open.Count -eq 0) { break }
        try {
            $stream = $version.OpenBinaryStream(); $context.ExecuteQuery()
            $memory = New-Object System.IO.MemoryStream
            try { $stream.Value.CopyTo($memory); $bytes = $memory.ToArray() } finally { $memory.Dispose() }
            $text = Get-TextFromBytes -Bytes $bytes -FileName $Name
        }
        catch { $complete = $false; continue }   # a version that cannot be read: a later one can only date the hit if it is in
        if ($text.Status -ne 'OK') { $complete = $false; continue }
        $found = @(Find-PiiInText $text.Text | ForEach-Object { Get-HitKey $_ })
        foreach ($key in @($open | Where-Object { $found -contains $_ })) {
            $byHit[$key] = @{ AddedOn = (Get-DateText $version.Created); AddedBy = (Get-PersonText $version.CreatedBy) }
        }
        $open = @($open | Where-Object { -not $byHit.ContainsKey($_) })
    }
    # In no older version: the current version added it (with no older version: the file's maker) - but only when
    # every older version was read; else who added it is not known, and the report says so (review of PR 20).
    foreach ($key in $open) {
        if (-not $complete) { $byHit[$key] = @{ AddedOn = $script:PartHistory; AddedBy = $script:PartHistory } }
        elseif ($versions.Count -eq 0) { $byHit[$key] = @{ AddedOn = $Who.Created; AddedBy = $Who.CreatedBy } }
        else { $byHit[$key] = @{ AddedOn = $Who.Modified; AddedBy = $Who.ModifiedBy } }
    }
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
