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
    [int]$MinScoreToAutoTag = 5
)

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
Write-Host ""

# Match helper — scores by length of the matched Title token. Returns top 3.
function Find-Matches {
    param([string]$Filename, [object[]]$Candidates)
    $hits = @()
    $low = $Filename.ToLower()
    foreach ($c in $Candidates) {
        $t = $c.Title.ToLower()
        # Skip super-short titles (1-2 chars) to avoid junk matches
        if ($t.Length -lt 3) { continue }
        if ($low.Contains($t)) {
            $hits += [PSCustomObject]@{
                Id = $c.Id
                Title = $c.Title
                Score = $t.Length
            }
        }
    }
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

        # Look for property + owner matches in the filename
        $propMatches = if (-not $hasProperty) { Find-Matches -Filename $filename -Candidates $properties } else { @() }
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
