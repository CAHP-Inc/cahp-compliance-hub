<#
.SYNOPSIS
    Seed 11 Property records — each named for an LLC directly owned by
    IV Fund Global, LLC — into the Properties Registry, with an Ownership
    row tying each one to IV Fund Global, LLC as Sole Member at 100%.

.DESCRIPTION
    Each LLC named in the script becomes a row in the Properties Registry
    (Title and LegalEntity both set to the LLC name) plus a single
    Ownership Structure row:

      OwnerLookupId             = IV Fund Global, LLC
      LinkedPropertyLookupId    = the new property
      RelationshipType          = 'Sole Member'
      OwnershipPercent          = 100

    No CAHP SC LLC co-member is added — these are directly-owned by the
    parent entity (matches the 'Direct Ownership' toggle in the in-app
    New Property wizard).

    Idempotent: a property is reused if one already exists with the same
    Title; an ownership row is skipped if (OwnerLookupId, LinkedProperty)
    already exists.

    Defaults each property to cahpState=SC, PropertyStatus=Pending,
    cahpVerificationStatus='Inherited - Unverified'. Address, county, unit
    count, AMI program, DOR account ID, and EIN are left blank — fill those
    in via the in-app property detail page afterward.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.PARAMETER ParentEntityTitle
    Exact Title of the Owner row that holds 100% of each new property.
    Defaults to "IV Fund Global, LLC". Override if the parent is named
    differently in your Owners list.

.PARAMETER State
    cahpState for the new properties. Defaults to "SC".

.EXAMPLE
    .\seed-iv-fund-global-properties.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [string]$ParentEntityTitle = "IV Fund Global, LLC",
    [ValidateSet("SC","NC")] [string]$State = "SC"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Seed IV Fund Global properties (11 directly-owned LLCs)" -ForegroundColor Cyan
Write-Host "  Site:   $SiteUrl" -ForegroundColor Cyan
Write-Host "  Parent: $ParentEntityTitle" -ForegroundColor Cyan
Write-Host "  State:  $State" -ForegroundColor Cyan
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

# Verify all three target lists exist
foreach ($listName in @("Owners","Properties Registry","Ownership Structure")) {
    if (-not (Get-PnPList -Identity $listName -ErrorAction SilentlyContinue)) {
        Write-Error "'$listName' list not found. Re-run the base provisioning first."
        exit 1
    }
}

# Find the parent entity (IV Fund Global, LLC). Title isn't guaranteed indexed,
# so we page and match in PowerShell.
Write-Host "-> Locating parent '$ParentEntityTitle' in Owners..." -ForegroundColor White
$allOwners = Get-PnPListItem -List "Owners" -PageSize 500 -Fields "ID","Title"
$parent = $allOwners | Where-Object { $_["Title"] -eq $ParentEntityTitle } | Select-Object -First 1
if (-not $parent) {
    Write-Error "'$ParentEntityTitle' not found in Owners. Pass -ParentEntityTitle with the exact title to override."
    exit 1
}
$parentId = $parent.Id
Write-Host "   Found '$ParentEntityTitle' (ID: $parentId)" -ForegroundColor Green
Write-Host ""

# Pre-index existing Property titles to avoid duplicates
Write-Host "-> Pre-fetching existing Properties..." -ForegroundColor White
$allProperties = Get-PnPListItem -List "Properties Registry" -PageSize 500 -Fields "ID","Title"
$propertyIdByTitle = @{}
foreach ($p in $allProperties) {
    $t = [string]$p["Title"]
    if ($t -and -not $propertyIdByTitle.ContainsKey($t)) {
        $propertyIdByTitle[$t] = $p.Id
    }
}
Write-Host "   Indexed $($propertyIdByTitle.Count) existing Property row(s)" -ForegroundColor Green

# Pre-index existing (owner -> property) ownership pairs so we don't double-create
Write-Host "-> Pre-fetching existing Ownership Structure rows..." -ForegroundColor White
$existingOwnership = Get-PnPListItem -List "Ownership Structure" -PageSize 500 -Fields "ID","OwnerLookupId","LinkedPropertyLookupId"
$existingPairs = @{}
foreach ($r in $existingOwnership) {
    $oid = [string]$r["OwnerLookupId"]
    $pid = [string]$r["LinkedPropertyLookupId"]
    if ($oid -and $pid) {
        $existingPairs["$oid|$pid"] = $r.Id
    }
}
Write-Host "   Indexed $($existingPairs.Count) existing Owner->Property pair(s)" -ForegroundColor Green
Write-Host ""

# Property records to create. Title and LegalEntity are both set to the LLC name.
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

$createdProperties = 0
$reusedProperties = 0
$createdOwnership = 0
$skippedOwnership = 0

foreach ($name in $subsidiaries) {
    Write-Host "-> $name" -ForegroundColor White

    # Step 1: ensure the Property exists
    $propId = $null
    if ($propertyIdByTitle.ContainsKey($name)) {
        $propId = $propertyIdByTitle[$name]
        Write-Host "     Property exists (ID: $propId)" -ForegroundColor DarkGray
        $reusedProperties++
    } else {
        try {
            $newProperty = Add-PnPListItem -List "Properties Registry" -Values @{
                Title                  = $name
                LegalEntity            = $name
                cahpState              = $State
                PropertyStatus         = "Pending"
                cahpVerificationStatus = "Inherited - Unverified"
                DateAddedToCAHP        = (Get-Date).ToString("yyyy-MM-ddT00:00:00Z")
                PropertyNotes          = "Directly owned by $ParentEntityTitle. Seeded $(Get-Date -Format 'yyyy-MM-dd') — fill in address, county, units, AMI, DOR Account ID, and EIN from the property detail page."
            }
            $propId = $newProperty.Id
            $propertyIdByTitle[$name] = $propId
            Write-Host "     + Created Property (ID: $propId)" -ForegroundColor Cyan
            $createdProperties++
        } catch {
            Write-Host "     ! Failed to create Property '$name': $_" -ForegroundColor Red
            continue
        }
    }

    # Step 2: ensure the Ownership row exists
    $pairKey = "$parentId|$propId"
    if ($existingPairs.ContainsKey($pairKey)) {
        Write-Host "     Ownership row already links '$ParentEntityTitle' -> '$name'; skipping" -ForegroundColor DarkGray
        $skippedOwnership++
        continue
    }

    try {
        Add-PnPListItem -List "Ownership Structure" -Values @{
            Title                  = $name
            OwnerLookupId          = $parentId
            LinkedPropertyLookupId = $propId
            RelationshipType       = "Sole Member"
            OwnershipPercent       = 100
            EffectiveDate          = (Get-Date).ToString("yyyy-MM-ddT00:00:00Z")
            SourceDocument         = "Seeded via seed-iv-fund-global-properties.ps1"
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
Write-Host "  Properties created : $createdProperties" -ForegroundColor White
Write-Host "  Properties reused  : $reusedProperties" -ForegroundColor White
Write-Host "  Ownership added    : $createdOwnership" -ForegroundColor White
Write-Host "  Ownership skipped  : $skippedOwnership" -ForegroundColor White
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Open the Properties page. The 11 new entries should appear" -ForegroundColor Yellow
Write-Host "     nested under '$ParentEntityTitle' as sub-entities (or as that" -ForegroundColor Yellow
Write-Host "     group's direct holdings if IV Fund Global itself has no LLC parent)." -ForegroundColor Yellow
Write-Host "  3. Click each new property to fill in address, county, units," -ForegroundColor Yellow
Write-Host "     AMI program, DOR Account ID, EIN, and add tax map IDs." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
