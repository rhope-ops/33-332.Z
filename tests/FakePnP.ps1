<#
FakePnP.ps1 - stand-ins for the PnP.PowerShell cmdlets the scan calls, so the tests need no network.

  Purpose:     Defines global functions with the names of the real PnP cmdlets. Module code that calls
               those names reaches these fakes. $global:FakePnP holds the fake site and switches.
  Inputs:      $global:FakePnP (set by Reset-FakePnP, changed by the tests).
  Outputs:     fake site data; $global:FakePnP.Calls lists each cmdlet called.
  How to run:  dot-sourced by Run-Tests.ps1; not on its own.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none (the fake sign-ins accept nothing and store nothing).

  Why global:  module code finds a cmdlet by name in the global scope, so the fakes are global functions,
               and the state they share ($global:FakePnP) is global too. Tests change it only through
               Set-FakePnPOption and read it through Test-FakePnPCall.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Fakes of global cmdlets share global state.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Read by the fakes.')]
param()

function Reset-FakePnP {
    <#
    .SYNOPSIS
      Sets up the fake site: library '/sites/TeamA/Shared Documents', channels General and Ops, the
      given files, file download through -AsMemoryStream.
    .PARAMETER Files
      Server-relative file path -> file bytes.
    .OUTPUTS
      None (sets $global:FakePnP).
    #>
    param([hashtable]$Files)
    $global:FakePnP = @{
        Files       = $Files
        Folders     = @('General', 'Ops', 'Forms')
        Libraries   = @('Shared Documents')                       # the visible document libraries (Get-PnPList)
        ListFails   = $false
        ConnectFails = $false                                     # the web login to a site fails (an error before the scan)
        OnlineAnswer = ''                                         # the answer to "Download and read them?" (Enter = no)
        SiteFails   = $false                                      # Get-PnPSite fails (the site's group cannot be read)
        TypedSite   = New-Object System.Collections.Generic.Queue[string]   # the fake Read-Host's answers, in order; none = Enter
        NoAccess    = $false                                      # the site refuses the signed-in person (403)
        StopInFolder2 = $false                                    # the scan stops (as at Ctrl+C) when it reaches folder 2
        SiteGroupId = [guid]'11111111-2222-3333-4444-555555555555'
        LookupSite  = ''                                          # the tenant root's answer for a Teams link's group ('' = none; 'fail' = refused)
        OpenAnswer  = ''                                          # the answer to "Open the workbook now?" ('' = Enter = yes)
        ScanAnswer  = ''                                          # the answer to "Scan now?" ('' = Enter = yes)
        Calls       = New-Object System.Collections.Generic.List[string]
        Progress    = New-Object System.Collections.Generic.List[string]
    }
}

function Set-FakePnPOption {
    # Changes one switch of the fake site: ListFails, ConnectFails, SiteFails or NoAccess.
    param([ValidateSet('ListFails', 'ConnectFails', 'SiteFails', 'NoAccess', 'StopInFolder2')][string]$Name, [bool]$Value)
    $global:FakePnP[$Name] = $Value
}

function Set-FakeLookupSite {
    # Sets the tenant root's answer to the group-site lookup of a Teams link ('' = no site; 'fail' = refused).
    param([string]$Site)
    $global:FakePnP.LookupSite = $Site
}

function Set-FakeFolders {
    # Sets the top-level folders of the fake library (Get-PnPFolderItem).
    param([string[]]$Names)
    $global:FakePnP.Folders = $Names
}

function Set-FakeSiteAnswer {
    <#
    .SYNOPSIS
      Sets what Read-Host returns, answer by answer, and the group ID of the connected site.
    .PARAMETER TypedSite
      The answers of the fake Read-Host, in order; after the last, it returns '' (Enter). 'cancel' = the prompt fails.
    .PARAMETER SiteGroupId
      The group ID of the connected site (Get-PnPSite).
    .PARAMETER ScanAnswer
      The answer to "Scan now?" ('' = Enter = yes; 'cancel' = the prompt fails).
    .OUTPUTS
      None (changes $global:FakePnP).
    #>
    param([string[]]$TypedSite = @(), [guid]$SiteGroupId = '11111111-2222-3333-4444-555555555555', [string]$ScanAnswer = '')
    $global:FakePnP.TypedSite.Clear()
    foreach ($answer in $TypedSite) { if ($answer) { $global:FakePnP.TypedSite.Enqueue($answer) } }
    $global:FakePnP.SiteGroupId = $SiteGroupId
    $global:FakePnP.ScanAnswer = $ScanAnswer
}

function Get-FakePnPConnection {
    # Every site the fake Connect-PnPOnline was called with since the last Reset-FakePnP, in order.
    return @($global:FakePnP.Calls | Where-Object { $_ -like 'Connect *' } | ForEach-Object { $_.Substring(8) })
}

function Get-FakeProgressText {
    # Everything the progress bar showed since the last Reset-FakePnP, as one string.
    return ($global:FakePnP.Progress -join ' | ')
}

function Get-FakeCallCount {
    # How many times the named fake call was made since the last Reset-FakePnP.
    param([string]$Name)
    return @($global:FakePnP.Calls | Where-Object { $_ -eq $Name }).Count
}

function Test-FakePnPCall {
    # True when the named fake cmdlet was called since the last Reset-FakePnP.
    param([string]$Name)
    return $global:FakePnP.Calls.Contains($Name)
}

# Read the site: library, its root folder, its channel folders and its items.
function global:Get-PnPList {
    <#
    .SYNOPSIS
      Fake: the site's lists: its libraries (Set-FakeLibraries), a hidden library and SharePoint's Style Library.
    #>
    [CmdletBinding()] param()
    $global:FakePnP.Calls.Add('Get-PnPList')
    foreach ($name in $global:FakePnP.Libraries) {
        $title = $name; if ($name -eq 'Shared Documents') { $title = 'Documents' }
        [pscustomobject]@{ Title = $title; BaseTemplate = 101; Hidden = $false; Url = '/sites/TeamA/' + $name }
    }
    [pscustomobject]@{ Title = 'Style Library'; BaseTemplate = 101; Hidden = $false; Url = '/sites/TeamA/Style Library' }
    [pscustomobject]@{ Title = 'Hidden Files'; BaseTemplate = 101; Hidden = $true; Url = '/sites/TeamA/Hidden Files' }
    [pscustomobject]@{ Title = 'Site Pages'; BaseTemplate = 119; Hidden = $false; Url = '/sites/TeamA/SitePages' }
}

function Set-FakeLibraries {
    # Sets the visible document libraries of the fake site, by URL name (default: Shared Documents).
    param([string[]]$Names)
    $global:FakePnP.Libraries = $Names
}

function global:Get-PnPProperty {
    # Fake: the library's root folder.
    [CmdletBinding()] param($ClientObject, $Property)
    [pscustomobject]@{ ServerRelativeUrl = $ClientObject.Url }
}

function global:Get-PnPFolderItem {
    # Fake: the top-level folders of the library.
    [CmdletBinding()] param($FolderSiteRelativeUrl, $ItemType)
    $global:FakePnP.Folders | ForEach-Object { [pscustomobject]@{ Name = $_ } }
}

function global:Get-PnPListItem {
    <#
    .SYNOPSIS
      Fake: one folder item plus one item per fake file; throws when ListFails is set.
    .PARAMETER List
      Not used.
    .PARAMETER PageSize
      Not used.
    .PARAMETER Fields
      Not used (every field is given).
    #>
    [CmdletBinding()] param($List, $PageSize, $Fields)
    if ($global:FakePnP.ListFails) { throw 'FAKE: list read refused' }
    if ($List.Title -eq 'Preservation Hold Library') { throw 'Access denied.' }   # members often cannot open it
    @{ FileLeafRef = 'General'; FileRef = '/sites/TeamA/Shared Documents/General'; FSObjType = 1 }
    foreach ($path in $global:FakePnP.Files.Keys) {
        @{ FileLeafRef = ($path -split '/')[-1]; FileRef = $path; FSObjType = 0
           File_x0020_Size = [string]$global:FakePnP.Files[$path].Length; Modified = [datetime]'2026-11-01'
           UniqueId = [guid]'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' }
    }
}

function global:Get-PnPFile {
    # Fake: the file bytes as a memory stream (a later PnP version's -AsMemoryStream).
    [CmdletBinding()] param($Url, [switch]$AsMemoryStream)
    New-Object System.IO.MemoryStream(, $global:FakePnP.Files[$Url])
}

function Get-FakeFileStream {
    # Fake of the streamed download (ChannelScan.psm1 OpenDirectStream): the file bytes, or with -Short two bytes of a web page.
    param([string]$Url, [switch]$Short)
    if ($Short) { return New-Object System.IO.MemoryStream(, [byte[]]@(60, 104)) }
    New-Object System.IO.MemoryStream(, $global:FakePnP.Files[$Url])
}

function Remove-FakeMemoryStreamSwitch {
    # Redefines the fake Get-PnPFile without -AsMemoryStream, as in PnP versions that lack it.
    function global:Get-PnPFile { [CmdletBinding()] param($Url) throw 'FAKE: should not be called' }
}

function global:Get-PnPContext {
    <#
    .SYNOPSIS
      Fake client context for the OpenBinaryStream path: Web.GetFileByServerRelativeUrl(url)
      .OpenBinaryStream() gives an object whose Value is the file stream; ExecuteQuery does nothing.
    .OUTPUTS
      [pscustomobject]
    #>
    [CmdletBinding()] param()
    $global:FakePnP.Calls.Add('Get-PnPContext')
    $web = [pscustomobject]@{}
    $web | Add-Member -MemberType ScriptMethod -Name GetFileByServerRelativeUrl -Value {
        param($url)
        $file = [pscustomobject]@{ Url = $url }
        $file | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value {
            [pscustomobject]@{ Value = (New-Object System.IO.MemoryStream(, $global:FakePnP.Files[$this.Url])) }
        }
        return $file
    }
    $context = [pscustomobject]@{ Web = $web }
    $context | Add-Member -MemberType ScriptMethod -Name ExecuteQuery -Value { }
    return $context
}

# The progress bar is on screen too: the fake records its text so tests can check it is masked.
function global:Write-Progress {
    # Fake: records the activity and status text instead of drawing a bar.
    [CmdletBinding()] param($Activity, $Status, $PercentComplete, [switch]$Completed, $Id, $ParentId, $CurrentOperation)
    $global:FakePnP.Progress.Add(('{0} {1} {2}' -f $Activity, $Status, $CurrentOperation))
    if ($global:FakePnP.StopInFolder2 -and [string]$Activity -like '*Folder 2 of*') { throw 'FAKE: the scan was stopped' }
    if ($Completed) { $global:FakePnP.Calls.Add('Progress done ' + $Id) }
    elseif ($null -ne $ParentId) { $global:FakePnP.Calls.Add('Progress ' + $Id + ' under ' + $ParentId) }
}

# Sign-in, sign-out and the team behind the site.
function global:Connect-PnPOnline {
    # Fake: the web login works unless ConnectFails is set.
    [CmdletBinding()] param($Url, [switch]$UseWebLogin)
    $global:FakePnP.Calls.Add('Connect ' + $Url)
    if ($PSBoundParameters['WarningAction'] -eq 'SilentlyContinue') { $global:FakePnP.Calls.Add('Quiet sign-in ' + $Url) }
    if ($UseWebLogin -and $global:FakePnP.ConnectFails) { throw 'FAKE: sign-in refused' }
}

function global:Invoke-PnPSPRestMethod {
    # Fake: the tenant root's group-site lookup (GetSiteStatus): LookupSite, nothing, or a refusal.
    [CmdletBinding()] param($Method, $Url, $Content)
    $global:FakePnP.Calls.Add('Lookup ' + $Url + ' ' + $Content.groupId)
    if ($global:FakePnP.LookupSite -eq 'fail') { throw 'FAKE: lookup refused' }
    if (-not $global:FakePnP.LookupSite) { return $null }
    [pscustomobject]@{ SiteUrl = $global:FakePnP.LookupSite; SiteStatus = 2 }
}

function global:Disconnect-PnPOnline {
    # Fake: nothing to close; records the call.
    [CmdletBinding()] param()
    $global:FakePnP.Calls.Add('Disconnect-PnPOnline')
}

function global:Get-PnPWeb {
    # Fake: the site's web; refused (403) when NoAccess is set.
    [CmdletBinding()] param()
    if ($global:FakePnP.NoAccess) { throw 'The remote server returned an error: (403) Forbidden.' }
    [pscustomobject]@{ Title = 'Team A' }
}

function global:Get-PnPSite {
    # Fake: the site belongs to a team (a Microsoft 365 group).
    [CmdletBinding()] param($Includes)
    if ($global:FakePnP.SiteFails) { throw 'FAKE: site refused' }
    [pscustomobject]@{ GroupId = $global:FakePnP.SiteGroupId }
}

function global:Read-Host {
    <#
    .SYNOPSIS
      Fake: the address the person would type (queued answers), or the answer to the open-the-workbook question.
    .PARAMETER Prompt
      The question.
    #>
    [CmdletBinding()] param($Prompt)
    # The open-the-workbook question has its own answer (Enter = open), so the queued addresses stay in order.
    if ([string]$Prompt -like 'Open the workbook*') { $global:FakePnP.Calls.Add('Open question'); return $global:FakePnP.OpenAnswer }
    if ([string]$Prompt -like 'Download and read them*') { $global:FakePnP.Calls.Add('Online question'); return $global:FakePnP.OnlineAnswer }
    if ([string]$Prompt -like 'Scan now*') {
        $global:FakePnP.Calls.Add('Scan question')
        if ($global:FakePnP.ScanAnswer -eq 'cancel') { throw 'FAKE: prompt cancelled' }   # as a cancel at the question
        return $global:FakePnP.ScanAnswer
    }
    $global:FakePnP.Calls.Add('Read-Host')
    # A real cancel stops the whole pipeline (the test run too); an error leaves through the same 'finally'.
    if ($global:FakePnP.TypedSite.Count -eq 0) { return '' }
    $answer = $global:FakePnP.TypedSite.Dequeue()
    if ($answer -eq 'cancel') { throw 'FAKE: prompt cancelled' }
    return $answer
}

function global:Invoke-Item {
    # Fake: records that the workbook would open in Excel, instead of opening it.
    [CmdletBinding()] param($LiteralPath)
    $global:FakePnP.Calls.Add('Open ' + $LiteralPath)
}
