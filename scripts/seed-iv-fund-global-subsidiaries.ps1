<#
.SYNOPSIS
    One-shot script: add 11 LLC subsidiaries wholly-owned by IV Fund Global LLC.

.DESCRIPTION
    Creates an Owners row for each of the 11 subsidiary LLCs (if not already
    present) and an Ownership Structure row tying it to IV Fund Global LLC
    as Sole Member at 100%.

    The Ownership row schema (recap):
      OwnerLookupId       = IV Fund Global LLC  (the holder of the stake)
      ParentOwnerLookupId = the subsidiary LLC  (the entity being held)
      RelationshipType    = 'Sole Member'
      OwnershipPercent    = 100

    Idempotent: existing Owners rows by Title are reused, and existing
    Ownership rows that match (parent + child) are skipped.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.PARAMETER ParentEntityTitle
    The Title of the parent Owner row that holds 100% of each subsidiary.
    Defaults to "IV Fund Global LLC". Override if the parent is named
    differently in your Owners list (e.g., "IV Fund Global" without LLC).

.PARAMETER OwnerState
    State of formation for the new LLCs. Defaults to "SC". Change before
    running if the LLCs are formed elsewhere.

.EXAMPLE
    .\seed-iv-fund-global-subsidiaries.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [string]$ParentEntityTitle = "IV Fund Global LLC",
    [string]$OwnerState = "SC"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Seed IV Fund Global subsidiaries (11 wholly-owned LLCs)" -ForegroundColor Cyan
Write-Host "  Site:   $SiteUrl" -ForegroundColor Cyan
Write-Host "  Parent: $ParentEntityTitle" -ForegroundColor Cyan
Write-Host "  State:  $OwnerState (formation state for the new LLCs)" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}
Write-Host "  Connected" -ForegroundColor Green
Write-Host ""

# Verify both target lists exist
$ownersList = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if (-not $ownersList) {
    Write-Error "'Owners' list not found. Run provision-owners.ps1 first."
    exit 1
}
$ownershipList = Get-PnPList -Identity "Ownership Structure" -ErrorAction SilentlyContinue
if (-not $ownershipList) {
    Write-Error "'Ownership Structure' list not found. Run provision-owners.ps1 first."
    exit 1
}

# Find the parent entity (IV Fund Global LLC). We use a paged GetPnPListItem
# + client-side Title match because the field isn't guaranteed indexed.
Write-Host "-> Locating parent '$ParentEntityTitle' in Owners..." -ForegroundColor White
$allOwners = Get-PnPListItem -List "Owners" -PageSize 500 -Fields "ID","Title"
$parent = $allOwners | Where-Object { $_["Title"] -eq $ParentEntityTitle } | Select-Object -First 1
if (-not $parent) {
    Write-Error "'$ParentEntityTitle' not found in Owners list. Create it first via the app (Owners -> New Owner) or pass -ParentEntityTitle with the exact title."
    exit 1
}
$parentId = $parent.Id
Write-Host "   Found '$ParentEntityTitle' (ID: $parentId)" -ForegroundColor Green
Write-Host ""

# Build a Title -> ID map for fast existing-Owner lookup
$ownerIdByTitle = @{}
foreach ($o in $allOwners) {
    $t = [string]$o["Title"]
    if ($t -and -not $ownerIdByTitle.ContainsKey($t)) {
        $ownerIdByTitle[$t] = $o.Id
    }
}

# Subsidiaries to seed
$subsidiaries = @(
    "IV 3 LLC",
    "IV 4 LLC",
    "IV 5 LLC",
    "IV Mobile 5 LLC",
    "IV SPB LLC",
    "IV SPB II LLC",
    "IV SPB 3 LLC",
    "IV SPB 4 LLC",
    "IV SPB 5 LLC",
    "IV SPB 6 LLC",
    "IV SPB 7 LLC"
)

# Pre-fetch existing Ownership rows to dedupe idempotently. We only need rows
# that point at the parent in either direction (OwnerLookupId=parent or
# ParentOwnerLookupId=parent) — but for safety we just check all of them
# against the (OwnerLookupId=$parentId AND ParentOwnerLookupId=$subId) tuple.
Write-Host "-> Pre-fetching existing Ownership Structure rows..." -ForegroundColor White
$existingOwnership = Get-PnPListItem -List "Ownership Structure" -PageSize 500 -Fields "ID","OwnerLookupId","ParentOwnerLookupId"
$existingPairs = @{}
foreach ($r in $existingOwnership) {
    $ownerId  = [string]$r["OwnerLookupId"]
    $parentId2 = [string]$r["ParentOwnerLookupId"]
    if ($ownerId -and $parentId2) {
        $key = "$ownerId|$parentId2"
        $existingPairs[$key] = $r.Id
    }
}
Write-Host "   Indexed $($existingPairs.Count) existing Ownership row(s)" -ForegroundColor Green
Write-Host ""

$createdOwners = 0
$reusedOwners = 0
$createdOwnership = 0
$skippedOwnership = 0

foreach ($name in $subsidiaries) {
    Write-Host "-> $name" -ForegroundColor White

    # Step 1: ensure the subsidiary exists in Owners
    $subId = $null
    if ($ownerIdByTitle.ContainsKey($name)) {
        $subId = $ownerIdByTitle[$name]
        Write-Host "     Owner row exists (ID: $subId)" -ForegroundColor DarkGray
        $reusedOwners++
    } else {
        try {
            $newOwner = Add-PnPListItem -List "Owners" -Values @{
                Title         = $name
                OwnerType     = "LLC"
                OwnerState    = $OwnerState
                IsCAHPEntity  = $false
                OwnerNotes    = "Wholly-owned subsidiary of $ParentEntityTitle. Seeded $(Get-Date -Format 'yyyy-MM-dd')."
            }
            $subId = $newOwner.Id
            $ownerIdByTitle[$name] = $subId
            Write-Host "     + Created Owner row (ID: $subId)" -ForegroundColor Cyan
            $createdOwners++
        } catch {
            Write-Host "     ! Failed to create Owner '$name': $_" -ForegroundColor Red
            continue
        }
    }

    # Step 2: ensure the Ownership relationship exists
    $pairKey = "$parentId|$subId"
    if ($existingPairs.ContainsKey($pairKey)) {
        Write-Host "     Ownership row already links '$ParentEntityTitle' -> '$name'; skipping" -ForegroundColor DarkGray
        $skippedOwnership++
        continue
    }

    try {
        Add-PnPListItem -List "Ownership Structure" -Values @{
            Title               = $name
            OwnerLookupId       = $parentId
            ParentOwnerLookupId = $subId
            RelationshipType    = "Sole Member"
            OwnershipPercent    = 100
            EffectiveDate       = (Get-Date).ToString("yyyy-MM-ddT00:00:00Z")
            SourceDocument      = "Seeded via seed-iv-fund-global-subsidiaries.ps1"
        } | Out-Null
        Write-Host "     + Created Ownership row ($ParentEntityTitle -> $name, 100% Sole Member)" -ForegroundColor Cyan
        $createdOwnership++
        $existingPairs[$pairKey] = $true
    } catch {
        Write-Host "     ! Failed to create Ownership row for '$name': $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Owners created   : $createdOwners" -ForegroundColor White
Write-Host "  Owners reused    : $reusedOwners" -ForegroundColor White
Write-Host "  Ownership added  : $createdOwnership" -ForegroundColor White
Write-Host "  Ownership skipped: $skippedOwnership" -ForegroundColor White
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Open '$ParentEntityTitle' on the Owners page." -ForegroundColor Yellow
Write-Host "     Property Holdings section should show the new sub-entities" -ForegroundColor Yellow
Write-Host "     under the 'Ownership Tree' section." -ForegroundColor Yellow
Write-Host "  3. For each subsidiary, click into its Owner page and use the" -ForegroundColor Yellow
Write-Host "     '+ Add Property' button to create the SFR / property records" -ForegroundColor Yellow
Write-Host "     it directly owns." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
