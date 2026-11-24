<#
PiiDetectors.psm1 - finds text that looks like PII and returns where it is.

  Purpose:     Get-PiiSpan runs every detector (PiiPatterns.psm1, and the labelled-name detector here), removes
               overlaps, raises a name near an identifier, drops Low hits in unreadable text, and returns each hit
               with its position and raw value.
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     hit objects WITH the raw value. Only PiiMasking.psm1 reads them; it masks every value
               before anything is shown or saved (the PAM, 2026-11-20: all values masked).
  Depends on:  PiiPatterns.psm1 (the one-type detectors, New-PiiHit) and
               PiiSignedIds.psm1 (Add-SignedIdHits, Test-GarbledNear), imported first.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings (named values; each says where it comes from)
# =====================================================================================

# A labelled name within this many characters of an identifier is raised to High. DoD privacy
# guidance treats a name together with an identifier (SSN, DoD ID, date of birth) as sensitive PII.
$script:NameToIdentifierChars = 80

# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last. One exception: Remove-SsnUnderPassport.
$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone',
    'Personal e-mail', 'Home address', 'Name')

# Types that make a nearby name sensitive (see $NameToIdentifierChars).
$script:IdentifierTypes = @('SSN', 'DoD ID', 'Date of birth', 'Passport', 'Payment card')

# Confidence labels as numbers, for sorting (High first).
$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# =====================================================================================
# Small helper and the labelled-name detector (it calls Get-PiiSpan, so it stays here)
# =====================================================================================

function Get-ConfidenceRank {
    # Returns 3, 2 or 1 for High, Medium or Low (0 for anything else), for sorting.
    param([string]$Confidence)
    if ($script:ConfidenceRank.ContainsKey($Confidence)) { return $script:ConfidenceRank[$Confidence] }
    return 0
}

function Find-NameHits {
    <#
    .SYNOPSIS
      Names after a label (Name:, Full name:, Member:, Employee:, Spouse:, Dependent:, NOK:,
      Next of kin:, POC:). Added at the PAM's request, 2026-11-20.
    .DESCRIPTION
      Low by itself: a name alone is PII but rarely sensitive. Get-PiiSpan raises it to High when an
      identifier is near ($NameToIdentifierChars). A name without a label is not found: that would
      need a list of real names. Military "LAST, FIRST MI" form is found.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit) for the name only, not its label.
    #>
    param([string]$Text)
    $label = '(?:full\s+name|member\s+name|name|member|employee|spouse|dependent|next\s+of\s+kin|NOK|POC)\s*[:=-]\s*'
    # Labels of other data often follow a name on the same line; a name stops before them.
    $stopWords = '(?!(?:SSN|SSAN|DOB|DoD|EDIPI|CAC|Phone|Cell|Mobile|Home|Address|Email|E-mail|Rank|Grade|Unit|Date|Born)\b)'
    $name = "[A-Z][A-Za-z'\-]+(?:,?\s+$stopWords[A-Z](?:[A-Za-z'\-]+|\.))+"
    foreach ($m in [regex]::Matches($Text, '(?i:' + $label + ')(' + $name + ')')) {
        New-PiiHit 'Name' 'Low' $m.Groups[1].Index $m.Groups[1].Value
    }
}

# =====================================================================================
# All detectors together
# =====================================================================================

function Get-AllHits {
    <#
    .SYNOPSIS
      Runs every detector on the text and returns all hits, overlaps included.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    Find-SsnHits $Text
    Find-DodIdHits $Text
    Find-CardHits $Text
    Find-DobHits $Text
    Find-PassportHits $Text
    Find-PhoneHits $Text
    Find-PersonalEmailHits $Text
    Find-HomeAddressHits $Text
    Find-NameHits $Text
}

function Remove-OverlappingHits {
    <#
    .SYNOPSIS
      Keeps one hit per stretch of text: the type earlier in $DetectorOrder wins, then the higher
      confidence. Example: the 9 digits inside a 16-digit card number are not also an SSN.
    .DESCRIPTION
      The kept hits never overlap, so kept in order of their start they also end in order. A new hit overlaps a
      kept one only if it overlaps the kept hit that starts last before the new hit ends: one binary search per
      hit instead of a check against every kept hit. (Each kept hit is inserted in place, a fast memory move.)
    .PARAMETER Hits
      All hits from Get-AllHits.
    .OUTPUTS
      The kept hits (a list, in priority order).
    #>
    param([object[]]$Hits)
    $sorted = $Hits | Where-Object { $_ } | Sort-Object `
        @{ Expression = { [array]::IndexOf($script:DetectorOrder, $_.Detector) } },
        @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true },
        Index
    $kept = New-Object System.Collections.Generic.List[object]
    $starts = New-Object System.Collections.Generic.List[int]   # starts of the kept hits, ascending
    $ends = New-Object System.Collections.Generic.List[int]     # their ends, in the same order (also ascending)
    foreach ($hit in $sorted) {
        $end = $hit.Index + $hit.Length
        # The first kept hit that starts at or after this hit's end (BinarySearch gives its place as ~n if absent).
        $after = $starts.BinarySearch($end)
        if ($after -lt 0) { $after = -bnot $after }
        while ($after -gt 0 -and $starts[$after - 1] -ge $end) { $after-- }
        if ($after -gt 0 -and $ends[$after - 1] -gt $hit.Index) { continue }   # overlaps the kept hit before it
        $starts.Insert($after, $hit.Index)
        $ends.Insert($after, $end)
        $kept.Add($hit)
    }
    return , $kept
}

function Remove-SsnUnderPassport {
    <#
    .SYNOPSIS
      Drops an SSN Low hit (9 bare digits, no SSN label) that covers the same characters as a Passport hit, so the
      Passport hit is kept (the PAM, 2026-11-24: a 9-digit value right after a Passport label is a Passport).
    .DESCRIPTION
      The one exception to $DetectorOrder. A Passport hit exists only with a Passport label in the label window
      (Find-PassportHits). An SSN with an SSN label (High) or with dashes or spaces (Medium) stays an SSN.
    .PARAMETER Hits
      All hits from Get-AllHits.
    .OUTPUTS
      The hits without those SSN hits.
    #>
    param([object[]]$Hits)
    $passportPlaces = @{}
    foreach ($hit in $Hits) {
        if ($hit -and $hit.Detector -eq 'Passport') { $passportPlaces['{0}+{1}' -f $hit.Index, $hit.Length] = $true }
    }
    if ($passportPlaces.Count -eq 0) { return $Hits }
    return @($Hits | Where-Object {
        -not ($_ -and $_.Detector -eq 'SSN' -and $_.Confidence -eq 'Low' -and $passportPlaces.ContainsKey(('{0}+{1}' -f $_.Index, $_.Length)))
    })
}

function Set-NameNearIdentifierHigh {
    <#
    .SYNOPSIS
      Raises a labelled name to High when an identifier hit is within $NameToIdentifierChars.
    .PARAMETER Hits
      The kept hits (changed in place).
    .OUTPUTS
      None.
    #>
    param([object[]]$Hits)
    foreach ($nameHit in @($Hits | Where-Object { $_.Detector -eq 'Name' })) {
        foreach ($other in $Hits) {
            $near = [Math]::Abs($other.Index - $nameHit.Index) -le $script:NameToIdentifierChars
            # An e-signature DoD ID (Low, or added by Add-SignedIdHits) does not raise a name: the name would be surer than the
            # ID (reviews of PR 50 and PR 60).
            $signatureId = $other.Detector -eq 'DoD ID' -and ($other.Confidence -eq 'Low' -or $other.PSObject.Properties['Signed'])
            if ($near -and -not $signatureId -and $script:IdentifierTypes -contains $other.Detector) { $nameHit.Confidence = 'High'; break }
        }
    }
}

function Get-PiiSpan {
    <#
    .SYNOPSIS
      The one entry point of this module: every PII hit in a text, without overlaps, in text order.
      Hits hold raw values: pass them only to PiiMasking.psm1.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), sorted by position.
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    # Underscores join words in file names ("CAC_1234..."), which defeats word-boundary checks.
    # Seen as spaces by the detectors only; the length is the same, so positions still fit the text.
    $searchText = $Text.Replace('_', ' ')
    $kept = Remove-OverlappingHits @(Remove-SsnUnderPassport @(Get-AllHits $searchText))
    Add-SignedIdHits $searchText $kept
    Set-NameNearIdentifierHigh $kept.ToArray()
    # A Low hit inside unreadable text (symbols, a font read wrong) is noise: it is dropped. High and Medium hits stay.
    $kept = @($kept | Where-Object { $_.Confidence -ne 'Low' -or -not (Test-GarbledNear $searchText $_.Index $_.Length) })
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

Export-ModuleMember -Function Get-PiiSpan, Get-ConfidenceRank
