<#
.SYNOPSIS
    PR-09a: Provisions Owners SharePoint list + adds Owner lookup columns to Ownership Structure.

.DESCRIPTION
    Splits the existing flat Ownership Structure into a proper entity master + junction table:
      - NEW LIST: Owners (entity master — one row per VanRock Holdings, CAHP SC LLC, individual, etc.)
      - EXISTING LIST UPDATED: Ownership Structure gets two new lookup columns:
          - Owner (Lookup → Owners) — the entity holding the ownership interest in this row
          - ParentOwner (Lookup → Owners) — if this entity is itself a member of another entity (recursion)

    After this script runs, the Owners list can be populated via the app's new Owners module
    (ships in the PR-09a code deploy). The existing Title-based Ownership Structure rows remain
    but will need OwnerLookupId set manually (or recreated) — there should be few/none currently.

    Idempotent — re-running is safe.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-owners.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"

.NOTES
    Authored by Brandy Turner / NewShire Property Management
    Part of: CAHP Compliance Hub — Phase 1 strict closure, PR-09a
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  PR-09a: Owners list + Ownership Structure lookups" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}
Write-Host "  ✓ Connected" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Part 1: Create Owners list
# =============================================================================

$ListTitle = "Owners"
$ownersList = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $ownersList) {
    Write-Host "→ Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/Owners" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Master list of entities holding ownership interest in CAHP properties — individuals, LLCs, and nonprofits. The single source of truth for entity names, types, and contact info. Ownership Structure references this list via lookup." `
        -EnableVersioning $true | Out-Null
    Write-Host "  ✓ Owners list created" -ForegroundColor Green
} else {
    Write-Host "→ Owners list already exists; verifying columns" -ForegroundColor Yellow
}

$ownersCols = @(
    @{ Display = "Owner Type"; Internal = "OwnerType"; Type = "Choice"; Choices = @("Individual", "LLC", "Nonprofit"); InView = $true }
    @{ Display = "State"; Internal = "OwnerState"; Type = "Text"; InView = $true }
    @{ Display = "Tax ID"; Internal = "TaxID"; Type = "Text"; InView = $false }
    @{ Display = "Contact Email"; Internal = "ContactEmail"; Type = "Text"; InView = $true }
    @{ Display = "Notes"; Internal = "OwnerNotes"; Type = "Note"; InView = $false }
)

foreach ($col in $ownersCols) {
    $existing = Get-PnPField -List $ListTitle -Identity $col.Internal -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  → $($col.Display) exists, skipping" -ForegroundColor DarkGray
        continue
    }
    try {
        $params = @{
            List         = $ListTitle
            DisplayName  = $col.Display
            InternalName = $col.Internal
            Type         = $col.Type
        }
        if ($col.InView) { $params.AddToDefaultView = $true }
        if ($col.Type -eq "Choice") { $params.Choices = $col.Choices }
        Add-PnPField @params -ErrorAction Stop | Out-Null
        Write-Host "  + $($col.Display) [$($col.Type)]" -ForegroundColor Cyan
    } catch {
        Write-Host "  ! Failed: $($col.Display) — $_" -ForegroundColor Red
    }
}

# =============================================================================
# Part 2: Add Owner + ParentOwner lookup columns to Ownership Structure
# =============================================================================

Write-Host ""
Write-Host "→ Adding lookup columns to existing 'Ownership Structure' list..." -ForegroundColor White

$ownersListReloaded = Get-PnPList -Identity "Owners" -ErrorAction Stop
$ownersListId = $ownersListReloaded.Id.ToString()

# Owner lookup — the entity whose ownership stake this row represents
if (-not (Get-PnPField -List "Ownership Structure" -Identity "Owner" -ErrorAction SilentlyContinue)) {
    $ownerFieldXml = @"
<Field
  Type='Lookup'
  DisplayName='Owner'
  Name='Owner'
  List='{$ownersListId}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
    Add-PnPFieldFromXml -List "Ownership Structure" -FieldXml $ownerFieldXml | Out-Null

    # Add to default view
    $view = Get-PnPView -List "Ownership Structure" -Identity "All Items" -ErrorAction SilentlyContinue
    if ($view) {
        $viewFields = $view.ViewFields
        if ($viewFields -notcontains "Owner") {
            Set-PnPView -List "Ownership Structure" -Identity "All Items" -Fields ($viewFields + "Owner") | Out-Null
        }
    }
    Write-Host "  + Owner [Lookup → Owners] on Ownership Structure" -ForegroundColor Cyan
} else {
    Write-Host "  → Owner column already exists, skipping" -ForegroundColor DarkGray
}

# ParentOwner lookup — for recursive entity-owns-entity relationships
if (-not (Get-PnPField -List "Ownership Structure" -Identity "ParentOwner" -ErrorAction SilentlyContinue)) {
    $parentOwnerFieldXml = @"
<Field
  Type='Lookup'
  DisplayName='Parent Owner'
  Name='ParentOwner'
  List='{$ownersListId}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
    Add-PnPFieldFromXml -List "Ownership Structure" -FieldXml $parentOwnerFieldXml | Out-Null
    Write-Host "  + Parent Owner [Lookup → Owners] on Ownership Structure" -ForegroundColor Cyan
} else {
    Write-Host "  → Parent Owner column already exists, skipping" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT: deploy PR-09a app code." -ForegroundColor Yellow
Write-Host "  Any existing Ownership Structure rows are 'legacy' — their Owner lookup is empty." -ForegroundColor Yellow
Write-Host "  The new Owners module will let you recreate them properly. Few-to-zero rows currently." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
