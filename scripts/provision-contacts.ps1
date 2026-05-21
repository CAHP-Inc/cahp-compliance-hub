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

# =============================================================================
# Part 3: Create Contact Owner Links junction list (many-to-many)
# =============================================================================

Write-Host ""
Write-Host "→ Creating 'Contact Owner Links' junction list..." -ForegroundColor White

$JunctionTitle = "Contact Owner Links"
$junctionList = Get-PnPList -Identity $JunctionTitle -ErrorAction SilentlyContinue
if (-not $junctionList) {
    New-PnPList -Title $JunctionTitle -Template GenericList -Url "lists/ContactOwnerLinks" | Out-Null
    Set-PnPList -Identity $JunctionTitle `
        -Description "Junction list — one row per (Contact, Owner) association. Lets a single Contact represent multiple Owner entities." `
        -EnableVersioning $true | Out-Null
    Write-Host "  ✓ Junction list created" -ForegroundColor Green
} else {
    Write-Host "  → Junction list already exists; verifying columns" -ForegroundColor Yellow
}

# Lookup → Contacts
$contactsListReloaded2 = Get-PnPList -Identity "Contacts" -ErrorAction Stop
$contactsId2 = $contactsListReloaded2.Id.ToString()
if (-not (Get-PnPField -List $JunctionTitle -Identity "Contact" -ErrorAction SilentlyContinue)) {
    $contactXml = @"
<Field
  Type='Lookup'
  DisplayName='Contact'
  Name='Contact'
  List='{$contactsId2}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
    Add-PnPFieldFromXml -List $JunctionTitle -FieldXml $contactXml | Out-Null
    Write-Host "  + Contact [Lookup → Contacts]" -ForegroundColor Cyan
} else {
    Write-Host "  → Contact column already exists, skipping" -ForegroundColor DarkGray
}

# Lookup → Owners
$ownersListForJunction = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if ($ownersListForJunction) {
    $ownersListId2 = $ownersListForJunction.Id.ToString()
    if (-not (Get-PnPField -List $JunctionTitle -Identity "Owner" -ErrorAction SilentlyContinue)) {
        $ownerXml = @"
<Field
  Type='Lookup'
  DisplayName='Owner'
  Name='Owner'
  List='{$ownersListId2}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
        Add-PnPFieldFromXml -List $JunctionTitle -FieldXml $ownerXml | Out-Null
        Write-Host "  + Owner [Lookup → Owners]" -ForegroundColor Cyan
    } else {
        Write-Host "  → Owner column already exists, skipping" -ForegroundColor DarkGray
    }
} else {
    Write-Host "  ! Owners list missing — run provision-owners.ps1 first" -ForegroundColor Red
}

# Add Contact + Owner to default view
$junctionView = Get-PnPView -List $JunctionTitle -Identity "All Items" -ErrorAction SilentlyContinue
if ($junctionView) {
    $viewFields = $junctionView.ViewFields
    $newFields = $viewFields
    if ($newFields -notcontains "Contact") { $newFields = $newFields + "Contact" }
    if ($newFields -notcontains "Owner") { $newFields = $newFields + "Owner" }
    if ($newFields.Count -ne $viewFields.Count) {
        Set-PnPView -List $JunctionTitle -Identity "All Items" -Fields $newFields | Out-Null
    }
}

# =============================================================================
# Part 4: One-time migration — copy any existing single Contact.ContactOwner
# values into the junction list so nothing gets orphaned.
# =============================================================================

Write-Host ""
Write-Host "→ Migrating existing single Contact.ContactOwner values into the junction..." -ForegroundColor White

$existingContacts = Get-PnPListItem -List "Contacts" -PageSize 500 -Fields "ID","ContactOwner" -ErrorAction SilentlyContinue
$existingJunctionRows = Get-PnPListItem -List $JunctionTitle -PageSize 500 -Fields "ID","Contact","Owner" -ErrorAction SilentlyContinue
$migratedCount = 0
$skippedCount = 0

foreach ($contact in ($existingContacts ?? @())) {
    $contactOwner = $contact["ContactOwner"]
    if (-not $contactOwner) { continue }
    $ownerId = $null
    if ($contactOwner.LookupId) {
        $ownerId = $contactOwner.LookupId
    } elseif ($contactOwner -is [int]) {
        $ownerId = $contactOwner
    }
    if (-not $ownerId) { continue }
    $contactId = $contact.Id

    # Skip if a junction row already exists for this (contact, owner)
    $alreadyLinked = $false
    foreach ($row in ($existingJunctionRows ?? @())) {
        $rowContact = $row["Contact"]
        $rowOwner = $row["Owner"]
        if ($rowContact -and $rowOwner -and $rowContact.LookupId -eq $contactId -and $rowOwner.LookupId -eq $ownerId) {
            $alreadyLinked = $true
            break
        }
    }
    if ($alreadyLinked) {
        $skippedCount++
        continue
    }
    Add-PnPListItem -List $JunctionTitle -Values @{
        Title = "Contact $contactId <-> Owner $ownerId"
        Contact = $contactId
        Owner = $ownerId
    } | Out-Null
    $migratedCount++
}

Write-Host "  ✓ Migration: $migratedCount new junction row(s) created, $skippedCount already linked" -ForegroundColor Green

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Open the app, go to Contacts in the sidebar, add or edit a contact." -ForegroundColor Yellow
Write-Host "  2. Each contact can now be linked to multiple Owner entities at once." -ForegroundColor Yellow
Write-Host "  3. On each Property's Overview tab, click Edit and pick the Owner Contact." -ForegroundColor Yellow
Write-Host "  4. Contacts linked to an Owner show up under that owner's 'Waiting on this owner' filter." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
