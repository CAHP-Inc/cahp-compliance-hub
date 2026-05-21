<#
.SYNOPSIS
    Provisions the Contacts SharePoint list + adds the Owner Contact lookup
    column to Properties Registry.

.DESCRIPTION
    Adds a Contacts list for the people we communicate with about properties
    (owners, attorneys, vendors, etc.), separate from the entity-level Owners
    list. The Properties Registry gets a new lookup column pointing here so
    each property can have a designated owner-side point of contact.

    Idempotent — re-running is safe.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-contacts.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"
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
Write-Host "  Contacts list + Properties.PropertyOwnerContact lookup" -ForegroundColor Cyan
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
# Part 1: Create Contacts list
# =============================================================================

$ListTitle = "Contacts"
$contactsList = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $contactsList) {
    Write-Host "→ Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/Contacts" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "People we communicate with about properties — owners, attorneys, vendors, property managers. Distinct from Owners (which are entity records). Used to populate the assignee picker and surface what's waiting on each person." `
        -EnableVersioning $true | Out-Null
    Write-Host "  ✓ Contacts list created" -ForegroundColor Green
} else {
    Write-Host "→ Contacts list already exists; verifying columns" -ForegroundColor Yellow
}

# Contact columns. Title is the display name; the rest are added below.
$contactCols = @(
    @{ Display = "Contact Email"; Internal = "ContactEmail"; Type = "Text"; InView = $true }
    @{ Display = "Contact Phone"; Internal = "ContactPhone"; Type = "Text"; InView = $true }
    @{ Display = "Contact Role";  Internal = "ContactRole";  Type = "Choice";
       Choices = @("Property Owner", "Sponsor", "Attorney", "Accountant", "Property Manager", "Vendor", "Lender", "Other");
       InView = $true }
    @{ Display = "Contact Notes"; Internal = "ContactNotes"; Type = "Note";   InView = $false }
)

foreach ($col in $contactCols) {
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
        if ($col.InView)            { $params.AddToDefaultView = $true }
        if ($col.Type -eq "Choice") { $params.Choices = $col.Choices }
        Add-PnPField @params -ErrorAction Stop | Out-Null
        Write-Host "  + $($col.Display) [$($col.Type)]" -ForegroundColor Cyan
    } catch {
        Write-Host "  ! Failed: $($col.Display) — $_" -ForegroundColor Red
    }
}

# Lookup → Owners (optional linkage between Contact and the entity they represent)
$ownersList = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if (-not $ownersList) {
    Write-Host "  ! 'Owners' list not found — run provision-owners.ps1 first" -ForegroundColor Red
} else {
    if (-not (Get-PnPField -List $ListTitle -Identity "ContactOwner" -ErrorAction SilentlyContinue)) {
        $ownersListId = $ownersList.Id.ToString()
        $contactOwnerXml = @"
<Field
  Type='Lookup'
  DisplayName='Contact Owner'
  Name='ContactOwner'
  List='{$ownersListId}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $contactOwnerXml | Out-Null
        $view = Get-PnPView -List $ListTitle -Identity "All Items" -ErrorAction SilentlyContinue
        if ($view) {
            $viewFields = $view.ViewFields
            if ($viewFields -notcontains "ContactOwner") {
                Set-PnPView -List $ListTitle -Identity "All Items" -Fields ($viewFields + "ContactOwner") | Out-Null
            }
        }
        Write-Host "  + Contact Owner [Lookup → Owners]" -ForegroundColor Cyan
    } else {
        Write-Host "  → Contact Owner column already exists, skipping" -ForegroundColor DarkGray
    }
}

# =============================================================================
# Part 2: Add PropertyOwnerContact lookup to Properties Registry
# =============================================================================

Write-Host ""
Write-Host "→ Adding PropertyOwnerContact lookup to Properties Registry..." -ForegroundColor White

$contactsListReloaded = Get-PnPList -Identity "Contacts" -ErrorAction Stop
$contactsListId = $contactsListReloaded.Id.ToString()

if (-not (Get-PnPField -List "Properties Registry" -Identity "PropertyOwnerContact" -ErrorAction SilentlyContinue)) {
    $propOwnerContactXml = @"
<Field
  Type='Lookup'
  DisplayName='Owner Contact'
  Name='PropertyOwnerContact'
  List='{$contactsListId}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
    Add-PnPFieldFromXml -List "Properties Registry" -FieldXml $propOwnerContactXml | Out-Null
    $view = Get-PnPView -List "Properties Registry" -Identity "All Items" -ErrorAction SilentlyContinue
    if ($view) {
        $viewFields = $view.ViewFields
        if ($viewFields -notcontains "PropertyOwnerContact") {
            Set-PnPView -List "Properties Registry" -Identity "All Items" -Fields ($viewFields + "PropertyOwnerContact") | Out-Null
        }
    }
    Write-Host "  + Owner Contact [Lookup → Contacts] on Properties Registry" -ForegroundColor Cyan
} else {
    Write-Host "  → Owner Contact column already exists, skipping" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Open the app, go to Contacts in the sidebar, add some contacts." -ForegroundColor Yellow
Write-Host "  2. On each Property's Overview tab, click Edit and pick the Owner Contact." -ForegroundColor Yellow
Write-Host "  3. Contacts whose Contact Owner == an Owner entity now show up under that owner's" -ForegroundColor Yellow
Write-Host "     'Waiting on this owner' filter on the Owner detail page." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
