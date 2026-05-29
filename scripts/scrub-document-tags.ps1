<#
.SYNOPSIS
    Scrub every document library for files missing their Owner / Property
    tags, suggest matches from filenames, optionally apply them.

.DESCRIPTION
    For each row in the 8 property-linked libraries + CAHP Entity Documents:
      1. Read the current Owner and Property tags (multi-value Lookup).
      2. If either is empty, look for matching properties / owners whose
         Title appears in the filename (case-insensitive substring match).
      3. Score matches by overlap length so 'Oakwood Apartments' beats 'O'.
      4. Print a report of every orphan with its top match suggestions.

    Defaults to DRY RUN. Pass -Execute to write the suggested tags back
    to SharePoint (only when a HIGH-confidence single match is found).
    Use -Interactive to be prompted per file before tagging.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID.

.PARAMETER Execute
    Pass to actually write the tags. Without it, runs as dry-run only.

.PARAMETER Interactive
    Prompt per file before writing each tag, even if -Execute is set.

.PARAMETER MinScoreToAutoTag
    Minimum match score (filename overlap length) required to auto-tag
    when -Execute is set without -Interactive. Default 5 — anything
    shorter is too risky for an unattended tag.

.EXAMPLE
    # Dry run — see what's orphaned and what would be tagged
    .\scrub-document-tags.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"

.EXAMPLE
    # Apply only high-confidence matches automatically
    .\scrub-document-tags.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59" `
      -Execute

.EXAMPLE
    # Apply with per-file prompts
    .\scrub-document-tags.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59" `
      -Execute -Interactive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [switch]$Execute,
    [switch]$Interactive,
    [int]$MinScoreToAutoTag = 5,
    # When set, for any file we successfully Property-tag, also look up the
    # property's primary direct owner via Ownership Structure and set Owner
    # too. Off by default to keep dry-run output predictable.
    [switch]$DeriveOwnerFromProperty
)

# Known filename → property-title aliases. Filenames in your libraries
# reference street addresses or legacy entity names that don't always match
# the canonical Property Title in SharePoint. Add to this map as you discover
# more aliases.
$FilenameAliases = @{
    "1200 college pointe" = "Fusion Pointe"
    "1200 cp"             = "Fusion Pointe"
    "144 w henry"         = "City View"
    "144 whenry"          = "City View"
    "arlington 16"        = "Arlington Townes"
    "hampton 101"         = "Hampton Avenue"
    "hampton ave"         = "Hampton Avenue"
    "greenwood gardens"   = "Greensboro Greenwood Gardens"
    "iv fund global ii"   = "IV Fund Global II LLC"
    "iv fund global"      = "IV Fund Global SFR Portfolio"
    "iv 3"                = "IV 3 LLC"
    "iv 4"                = "IV 4 LLC"
    "iv 5"                = "IV 5 LLC"
    "iv spb 3"            = "IV SPB 3 LLC"
    "iv spb 4"            = "IV SPB 4 LLC"
    "iv spb 5"            = "IV SPB 5 LLC"
    "iv spb 6"            = "IV SPB 6 LLC"
    "iv spb 7"            = "IV SPB 7 LLC"
    "iv spb ii"           = "IV SPB II LLC"
    "iv spb,"             = "IV SPB LLC"
    "iv spb llc"          = "IV SPB LLC"
    "iv mobile 5"         = "IV Mobile 5 LLC"
}

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Document tag scrub" -ForegroundColor Cyan
Write-Host "  Mode: $(if ($Execute) { 'EXECUTE (will write tags)' } else { 'DRY RUN (no changes)' })" -ForegroundColor $(if ($Execute) { 'Yellow' } else { 'Green' })
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}

# Pre-fetch the master lists for matching. Sort properties + owners by
# Title length descending so longer (more specific) matches win.
Write-Host "-> Fetching Properties + Owners for filename matching..." -ForegroundColor White
$properties = Get-PnPListItem -List "Properties Registry" -PageSize 500 -Fields "ID","Title" |
    ForEach-Object { [PSCustomObject]@{ Id = $_.Id; Title = [string]$_["Title"] } } |
    Where-Object { $_.Title } |
    Sort-Object { $_.Title.Length } -Descending
$owners = Get-PnPListItem -List "Owners" -PageSize 500 -Fields "ID","Title" |
    ForEach-Object { [PSCustomObject]@{ Id = $_.Id; Title = [string]$_["Title"] } } |
    Where-Object { $_.Title } |
    Sort-Object { $_.Title.Length } -Descending
Write-Host "   $($properties.Count) properties, $($owners.Count) owners" -ForegroundColor Green

# Lookup tables for fast id→title and title→id
$propertyByTitle = @{}
foreach ($p in $properties) { $propertyByTitle[$p.Title.ToLower()] = $p }

# If -DeriveOwnerFromProperty, pre-fetch Ownership Structure rows so we can
# resolve property → primary direct owner ID for each property-tagged file.
$primaryOwnerByPropertyId = @{}
if ($DeriveOwnerFromProperty) {
    Write-Host "-> Building property → primary-owner map from Ownership Structure..." -ForegroundColor White
    $ownership = Get-PnPListItem -List "Ownership Structure" -PageSize 500 `
        -Fields "ID","LinkedPropertyLookupId","OwnerLookupId","OwnershipPercent"
    $byProp = @{}
    foreach ($row in $ownership) {
        $propId = [string]$row["LinkedPropertyLookupId"]
        $ownId  = [string]$row["OwnerLookupId"]
        if (-not $propId -or -not $ownId) { continue }
        $pct = $row["OwnershipPercent"]
        if (-not $byProp.ContainsKey($propId)) { $byProp[$propId] = @() }
        $byProp[$propId] += [PSCustomObject]@{ OwnerId = $ownId; Percent = [double]($pct ?? 0) }
    }
    foreach ($pid in $byProp.Keys) {
        $primaryOwnerByPropertyId[$pid] = ($byProp[$pid] | Sort-Object Percent -Descending | Select-Object -First 1).OwnerId
    }
    Write-Host "   primary owner resolved for $($primaryOwnerByPropertyId.Count) propert$(if ($primaryOwnerByPropertyId.Count -eq 1) {'y'} else {'ies'})" -ForegroundColor Green
}
Write-Host ""

# Match helper — scores by length of the matched Title token. Returns top 3.
# Also honors the $FilenameAliases map so e.g. "144 W Henry" → "City View"
# even though "City View" doesn't appear in the filename literally.
function Find-Matches {
    param(
        [string]$Filename,
        [object[]]$Candidates,
        [hashtable]$Aliases = @{},
        [hashtable]$CandidatesByTitle = @{}
    )
    $hits = @()
    $low = $Filename.ToLower()

    # First check aliases — these win regardless of length since they're hand-curated
    foreach ($aliasKey in $Aliases.Keys) {
        if ($low.Contains($aliasKey)) {
            $targetTitle = $Aliases[$aliasKey]
            $match = $CandidatesByTitle[$targetTitle.ToLower()]
            if ($match) {
                $hits += [PSCustomObject]@{
                    Id    = $match.Id
                    Title = $match.Title
                    Score = $aliasKey.Length + 100   # bias aliases above substring matches
                    ViaAlias = $aliasKey
                }
            }
        }
    }

    foreach ($c in $Candidates) {
        $t = $c.Title.ToLower()
        # Skip super-short titles (1-2 chars) to avoid junk matches
        if ($t.Length -lt 3) { continue }
        if ($low.Contains($t)) {
            $hits += [PSCustomObject]@{
                Id    = $c.Id
                Title = $c.Title
                Score = $t.Length
                ViaAlias = $null
            }
        }
    }
    # Dedupe by Id, keeping the highest score
    $hits = $hits | Group-Object Id | ForEach-Object { $_.Group | Sort-Object Score -Descending | Select-Object -First 1 }
    $hits | Sort-Object Score -Descending | Select-Object -First 3
}

# Libraries to scrub (mirror PROPERTY_LINKED_LIBRARIES + CAHP Entity Documents)
$libraries = @(
    "AMI Certification Letters",
    "DOR Correspondence",
    "DOR Submittal Packages",
    "Land Use Restriction Agreements",
    "Operating Agreements",
    "Org Charts",
    "Property Deeds",
    "Supporting Documentation",
    "CAHP Entity Documents"
)

$totalScanned = 0
$totalOrphans = 0
$totalTagged = 0
$totalSkipped = 0
$totalAmbiguous = 0

foreach ($lib in $libraries) {
    Write-Host "===== $lib =====" -ForegroundColor Cyan
    $items = Get-PnPListItem -List $lib -PageSize 500 -Fields "ID","FileLeafRef","Owner","Property" -ErrorAction SilentlyContinue
    if (-not $items) {
        Write-Host "   (library not found or empty)" -ForegroundColor DarkGray
        continue
    }

    foreach ($it in $items) {
        $totalScanned++
        $filename = [string]$it["FileLeafRef"]
        if (-not $filename) { continue }

        $hasOwner = $it["Owner"] -and ($it["Owner"].Count -gt 0)
        $hasProperty = $it["Property"] -and ($it["Property"].Count -gt 0)
        if ($hasOwner -and $hasProperty) { continue }
        $totalOrphans++

        # Look for property + owner matches in the filename (aliases applied for property)
        $propMatches = if (-not $hasProperty) {
            Find-Matches -Filename $filename -Candidates $properties -Aliases $FilenameAliases -CandidatesByTitle $propertyByTitle
        } else { @() }
        $ownerMatches = if (-not $hasOwner) { Find-Matches -Filename $filename -Candidates $owners } else { @() }

        Write-Host "  $filename" -ForegroundColor White
        if (-not $hasProperty) {
            if ($propMatches.Count -eq 0) {
                Write-Host "     Property: (no match found)" -ForegroundColor DarkGray
            } else {
                $bestP = $propMatches[0]
                $tag = if ($bestP.Score -ge $MinScoreToAutoTag) { 'AUTO' } else { 'low-confidence' }
                Write-Host ("     Property: top match = {0} (id {1}, score {2}, {3})" -f $bestP.Title, $bestP.Id, $bestP.Score, $tag) -ForegroundColor $(if ($bestP.Score -ge $MinScoreToAutoTag) { 'Green' } else { 'Yellow' })
                if ($propMatches.Count -gt 1) {
                    $alts = ($propMatches | Select-Object -Skip 1 | ForEach-Object { "$($_.Title)(#$($_.Id),s$($_.Score))" }) -join ", "
                    Write-Host "       other candidates: $alts" -ForegroundColor DarkGray
                }
            }
        }
        if (-not $hasOwner) {
            if ($ownerMatches.Count -eq 0) {
                Write-Host "     Owner: (no match found)" -ForegroundColor DarkGray
            } else {
                $bestO = $ownerMatches[0]
                $tag = if ($bestO.Score -ge $MinScoreToAutoTag) { 'AUTO' } else { 'low-confidence' }
                Write-Host ("     Owner: top match = {0} (id {1}, score {2}, {3})" -f $bestO.Title, $bestO.Id, $bestO.Score, $tag) -ForegroundColor $(if ($bestO.Score -ge $MinScoreToAutoTag) { 'Green' } else { 'Yellow' })
                if ($ownerMatches.Count -gt 1) {
                    $alts = ($ownerMatches | Select-Object -Skip 1 | ForEach-Object { "$($_.Title)(#$($_.Id),s$($_.Score))" }) -join ", "
                    Write-Host "       other candidates: $alts" -ForegroundColor DarkGray
                }
            }
        }

        # Execute path — write the tags if a confident single match exists
        if ($Execute) {
            $patchValues = @{}
            $propWrite = $null
            $ownerWrite = $null

            if (-not $hasProperty -and $propMatches.Count -ge 1 -and $propMatches[0].Score -ge $MinScoreToAutoTag) {
                # Skip when multiple matches are tied at the top score (ambiguous)
                if ($propMatches.Count -gt 1 -and $propMatches[1].Score -eq $propMatches[0].Score) {
                    Write-Host "       skipping Property tag — ambiguous (tied top score)" -ForegroundColor Yellow
                    $totalAmbiguous++
                } else {
                    $propWrite = $propMatches[0]
                }
            }
            if (-not $hasOwner -and $ownerMatches.Count -ge 1 -and $ownerMatches[0].Score -ge $MinScoreToAutoTag) {
                if ($ownerMatches.Count -gt 1 -and $ownerMatches[1].Score -eq $ownerMatches[0].Score) {
                    Write-Host "       skipping Owner tag — ambiguous (tied top score)" -ForegroundColor Yellow
                    $totalAmbiguous++
                } else {
                    $ownerWrite = $ownerMatches[0]
                }
            }

            # Derive Owner from Property's primary direct owner when asked.
            # Works whether the Property is already set on the file OR we're
            # about to set it via $propWrite.
            if ($DeriveOwnerFromProperty -and -not $hasOwner -and -not $ownerWrite) {
                $effectivePropId = if ($propWrite) { [string]$propWrite.Id } elseif ($hasProperty) { [string]$it["Property"][0].LookupId } else { $null }
                if ($effectivePropId -and $primaryOwnerByPropertyId.ContainsKey($effectivePropId)) {
                    $derivedOwnerId = $primaryOwnerByPropertyId[$effectivePropId]
                    $derivedOwner = $owners | Where-Object { [string]$_.Id -eq $derivedOwnerId } | Select-Object -First 1
                    if ($derivedOwner) {
                        $ownerWrite = $derivedOwner
                        Write-Host ("     Owner: derived from property's chain → {0}" -f $derivedOwner.Title) -ForegroundColor Cyan
                    }
                }
            }

            if (-not $propWrite -and -not $ownerWrite) {
                $totalSkipped++
                continue
            }

            if ($Interactive) {
                $msg = "Apply"
                if ($propWrite)  { $msg += " Property='$($propWrite.Title)'" }
                if ($ownerWrite) { $msg += " Owner='$($ownerWrite.Title)'" }
                $msg += "? (y/n)"
                $resp = Read-Host $msg
                if ($resp -notmatch '^y') {
                    Write-Host "       skipped by user" -ForegroundColor Yellow
                    $totalSkipped++
                    continue
                }
            }

            try {
                if ($propWrite)  { $patchValues["Property"] = @($propWrite.Id) }
                if ($ownerWrite) { $patchValues["Owner"]    = @($ownerWrite.Id) }
                Set-PnPListItem -List $lib -Identity $it.Id -Values $patchValues | Out-Null
                $totalTagged++
                Write-Host "       -> tagged" -ForegroundColor Green
            } catch {
                Write-Host "       ! failed to tag: $_" -ForegroundColor Red
            }
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Files scanned:         $totalScanned" -ForegroundColor White
Write-Host "  Files missing a tag:   $totalOrphans" -ForegroundColor $(if ($totalOrphans -gt 0) { 'Yellow' } else { 'White' })
if ($Execute) {
    Write-Host "  Files auto-tagged:     $totalTagged" -ForegroundColor Green
    Write-Host "  Files skipped:         $totalSkipped (no confident match or user declined)" -ForegroundColor White
    Write-Host "  Ambiguous skips:       $totalAmbiguous (multiple tied top matches)" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "DRY RUN — no changes were made." -ForegroundColor Green
    Write-Host "To apply the AUTO-flagged matches, re-run with -Execute:" -ForegroundColor Yellow
    Write-Host "  .\scrub-document-tags.ps1 -SiteUrl '$SiteUrl' -ClientId '$ClientId' -Execute" -ForegroundColor Yellow
    Write-Host "Or per-file prompts: add -Interactive" -ForegroundColor Yellow
}
Write-Host ""

Disconnect-PnPOnline
