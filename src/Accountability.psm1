<#
Accountability.psm1 - who made and changed a file with PII, and who put the PII in and when.

  Purpose:     The PAM, 2026-11-21: for accountability and to date a possible incident, each file with a hit shows
               who last modified it, when it was made and by whom, and who added the PII and when, by e-mail
               address (the report is CUI and stays in DoD365). SharePoint and Teams: the file's fields and its
               version history (the first version that holds each hit). Drives: the file's owner (a Windows logon
               name, which Windows looks up from the owner's ID; no search of the directory: "No reconnaissance"), its times, and the names an Office file stores.
  Inputs:      the file's list item, path or reader result, and its hits.
  Outputs:     @{ ModifiedBy; Created; CreatedBy; AddedOn; AddedBy; HistoryNote; ByHit } for Report.psm1 Set-RowAccountability.
               "PII first saved" and "PII first saved by" hold only a date and a person (or "Not recorded"); every caveat goes in
               "History note" (the PAM, 2026-11-21: the columns must be easy to understand).
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
$script:NotRecorded = 'Not recorded'
$script:NoHistory = 'The file has no version history'
# Put-backs named per hit in its history note (the rest are counted).
$script:MaxPutBacks = 3
$script:MaybeEarlier = 'An older version could not be read: it may have been added earlier'
$script:PartHistory = "Not in the old versions read: check the file's version history by hand"
$script:DriveNoHistory = 'A drive keeps no file history'
# The two fallbacks of "PII first saved by" say so in the History note (review of PR 41).
$script:NoOlderVersion = "No older version: the file's creator is named"
$script:OnlyCurrentVersion = 'In no older version: the current version is named'

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
      Hit key -> @{ AddedOn; AddedBy; Note }.
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
        if (-not $Read.Complete) { $ByHit[$key] = New-AddedEntry $script:NotRecorded $script:NotRecorded $script:PartHistory }
        elseif ($Read.Versions -eq 0) { $ByHit[$key] = New-AddedEntry $Who.Created $Who.CreatedBy $script:NoOlderVersion }
        else { $ByHit[$key] = New-AddedEntry $Who.Modified $Who.ModifiedBy $script:OnlyCurrentVersion }
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
      Adds the put-backs to the history note: "Put back by X on D" (at most $MaxPutBacks named, the rest counted),
      and notes what the history read cannot show.
    .PARAMETER ByHit
      Hit key -> @{ AddedOn; AddedBy; Note }.
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
        Add-HistoryNote $ByHit[$key] $text
    }
    # Say what the history read cannot show, so "PII first saved by" never reads as complete when it is not (review).
    foreach ($key in @($ByHit.Keys | Where-Object { $ByHit[$_].AddedBy -ne $script:NotRecorded })) {
        if ($Unsure.ContainsKey($key)) { Add-HistoryNote $ByHit[$key] 'It may have been removed and put back (a version could not be read)' }
        if (-not $AllRead) { Add-HistoryNote $ByHit[$key] "Versions after the $script:MaxVersionsRead oldest were not checked" }
    }
}

function Add-PutBack {
    # Notes that a hit was put back (in a version after one without it): who saved that version and when.
    param([hashtable]$Back, [string]$Key, [string]$By, [string]$On)
    if (-not $Back.ContainsKey($Key)) { $Back[$Key] = New-Object System.Collections.Generic.List[string] }
    $Back[$Key].Add(('Put back by {0} on {1}' -f $By, $On))
}

function New-AddedEntry {
    # Who added one hit and when, with an optional history note.
    param([string]$On, [string]$By, [string]$Note = '')
    return @{ AddedOn = $On; AddedBy = $By; Note = $Note }
}

function Add-HistoryNote {
    # Adds one sentence to a hit's history note ("; " between sentences).
    param([hashtable]$Entry, [string]$Text)
    if ($Entry.Note) { $Entry.Note += '; ' + $Text } else { $Entry.Note = $Text }
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
      [hashtable] hit key -> @{ AddedOn; AddedBy; Note }.
    #>
    param([string]$Ref, [string]$Name, [object[]]$Hits, [hashtable]$Who)
    $byHit = @{}
    $open = @($Hits | ForEach-Object { Get-HitKey $_ } | Sort-Object -Unique)
    $history = Get-OldVersions $Ref
    if ($null -eq $history) {
        foreach ($key in $open) { $byHit[$key] = New-AddedEntry $script:NotRecorded $script:NotRecorded $script:NoHistory }
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
                if (-not $readFailed) { $byHit[$key] = New-AddedEntry $on $by }
                else { $byHit[$key] = New-AddedEntry $on $by $script:MaybeEarlier }
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
      [hashtable] ModifiedBy, Created, CreatedBy, AddedOn, AddedBy, HistoryNote.
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
    return @{ ModifiedBy = $modifiedBy; Created = $created; CreatedBy = $createdBy; AddedOn = $script:NotRecorded; AddedBy = $script:NotRecorded
        HistoryNote = $script:DriveNoHistory }
}

function Get-AccountRows {
    <#
    .SYNOPSIS
      The By account tab (C6, the PAM, 2026-11-21): one row per "PII first saved by" account, most hits first, with how
      many files and how many High, Medium and Low hits it first saved. Each file counts once per account; "Not
      recorded" and empty cells are left out. The same limits as the column apply (see the Method tab).
    .PARAMETER Hits
      The hit rows (High, Medium, Low).
    .OUTPUTS
      [object[]] rows: account, files, High, Medium, Low.
    #>
    param([object[]]$Hits)
    $byAccount = @{}
    foreach ($hit in $Hits) {
        $account = ([string]$hit.AddedBy).Trim()
        if (-not $account -or $account -eq $script:NotRecorded) { continue }
        if (-not $byAccount.ContainsKey($account)) { $byAccount[$account] = @{ Files = @{}; High = 0; Medium = 0; Low = 0 } }
        $entry = $byAccount[$account]
        $entry.Files[$hit.Channel + '/' + $hit.Item] = $true
        $entry[$hit.Confidence]++
    }
    $byTotal = @{ Expression = { $_.Value.High + $_.Value.Medium + $_.Value.Low }; Descending = $true }
    $byAccount.GetEnumerator() | Sort-Object $byTotal, Key | ForEach-Object {   # most hits first; equal totals by name (stable)
        , @($_.Key, [int]$_.Value.Files.Count, [int]$_.Value.High, [int]$_.Value.Medium, [int]$_.Value.Low)
    }
}

function Get-MethodRows {
    <#
    .SYNOPSIS
      The rows of the workbook's Method tab: how Ion finds possible PII, how sure each confidence is for each type (as
      src/PiiDetectors.psm1 decides it), and how it fills the who-and-when columns, with what that evidence can and
      cannot show (the PAM, 2026-11-21: the method on its own page; "PII first saved by" is evidence, not proof).
    .OUTPUTS
      [object[]] rows of two cells (topic, text).
    #>
    , @('How Ion finds possible PII', ('Ion reads the text of each file in memory and looks for patterns. A match is a lead, not proof: ' +
        'open each file and check it. The confidence (High, Medium, Low) depends on the type, as below.'))
    , @('SSN', ('Numbers the SSA never issues (000, 666 or 9xx first; 00 in the middle; 0000 last) are skipped. High: "SSN", "SSAN", ' +
        '"social security" or "soc sec" just before it. Medium: 3-2-4 digits with dashes or spaces. Low: 9 digits in a row. ' +
        'Without a label, 9 digits that go on into a longer number (a digit one space away, ' +
        'or a CAC name such as LAST.FIRST.M. just before) are not an SSN.'))
    , @('DoD ID', ('A 10-digit number, found only with a label just before it ("DoD ID", "DODID", "EDIPI", "CAC", "DoD #", "DoD number"), ' +
        'High. At the end of a CAC name (LAST.FIRST.M. then the 10 digits, as digital signatures show it): Low, ' +
        'and its version history is not read ("Not looked up").'))
    , @('Date of birth', ('A date written as 1/2/1990, 01-02-90, 1990-01-02, 2 Jan 1990, 02JAN1990 or January 2, 1990, found only with a label ' +
        'just before it ("DOB", "D.O.B.", "date of birth", "birth date", "birthdate", "born"): always High. Other date forms are not found.'))
    , @('Payment card', ('Medium: 13 to 19 digits (the length each card network uses), starting with that network''s digits, and the whole ' +
        'number passes the card check digit (Luhn). ' +
        'High: the same with a card word ("card", "Visa", "Amex", "GTC", "GPC" ...) before it.'))
    , @('Passport', 'Found only with "passport" just before it: always Medium.')
    , @('Phone', ('A US number written with separators: (###) ###-####, ###-###-####, ###.###.#### or ### ### ####; 10 digits in a row ' +
        'are not found; a +1 in front is allowed. Low: no label (office numbers are common). Medium: "cell", "mobile", "home", "personal", ' +
        '"my number" or "my phone" just before it.'))
    , @('Personal e-mail', 'Low: an address that does not end in .gov or .mil. Medium: a personal mail service (Gmail, Yahoo and the like).')
    , @('Home address', 'Low: number, street name and street type. Medium: it ends with a state and ZIP code. High: "home address" or similar before it.')
    , @('Name', 'Low: a name after "Name:", "Member:", "Spouse:" and the like. High: the same, with an SSN, DoD ID, date of birth, passport or card number close by.')
    , @('Noise', ('A Low match is dropped when, of the visible characters within 40 on each side (at least 20 of them), 30% or more are ' +
        'unreadable signs (symbols, a font read wrong). High and Medium matches are always kept.'))
    , @('')
    , @('Created, Created by', 'From SharePoint: when the file was first saved to the site, and by which account. On a drive: the file time and its owner.')
    , @('Last saved, Last saved by', 'The last save of the whole file, by anyone, for any reason. It says nothing about the PII; it shows who works on the file now.')
    , @('PII first saved, PII first saved by', (('Ion reads the file''s SharePoint version history, oldest first (at most {0} versions), and finds ' +
        'the first version whose text holds the same type and masked value. These columns name the account that saved that version, and when. ' +
        'If every older version was read and none holds the value, they name the current version (the same as Last saved). If the file has no ' +
        'older versions at all (version history off or never used), they name the file''s creator. If not every version could be read and none ' +
        'read holds the value, they say "Not recorded"; if a later version holds it, that version is named with a note that it may be earlier.') -f $script:MaxVersionsRead))
    , @('What that shows', ('SharePoint records which account saved each version. So these columns show which account first saved a version ' +
        'that holds the value: strong evidence, not proof of who typed it.'))
    , @('Why it can be wrong', ('The person who saves or uploads a file is not always the one who wrote the text: it may be pasted from an e-mail, ' +
        'copied from another file, or uploaded by OneDrive. When several people edit at once, SharePoint names one of them. Old versions may have ' +
        'been deleted. Two different values can look the same once masked (for example the same last 4 digits). A file copied in starts a new ' +
        'history. A drive keeps no history at all.'))
    , @('History note', ('Says how sure the two columns are: no version history; not in the versions read; a version that could not be read; ' +
        'a value removed and put back; more versions than were read; "No older version: the file''s creator is named"; "In no older ' +
        'version: the current version is named"; or, on a drive, "A drive keeps no file history".'))
    , @('Before you name a person', ('Open the file in SharePoint > Version history. Open the version Ion names and the one before it, and check ' +
        'that the full value first appears in that version.'))
    , @('By account tab', ('How many files and hits each "PII first saved by" account first saved, most first: a list of whom to contact first. ' +
        'It has the same limits as that column: evidence, not proof.'))
    , @('')
    , @('Review columns', ('Reviewed by, Review date, Is it PII?, Action taken and Reported to Privacy Manager are for you to fill in as you ' +
        'check each hit. Ion leaves them empty. The last three have a dropdown list.'))
}

Export-ModuleMember -Function Get-PersonText, Get-SharePointWho, Find-PiiAddedVersion, Get-DriveWho, Get-AccountRows, Get-MethodRows
