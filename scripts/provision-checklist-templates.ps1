<#
.SYNOPSIS
    Provisions the Checklist Templates SharePoint list and seeds the
    hardcoded DOR defaults if the list is empty.

.DESCRIPTION
    Moves the filing checklist templates from browser localStorage to a
    shared SharePoint list so teammates see the same configuration.

    Schema:
      - Title              (Text)    — checklist item title
      - TemplateCategory   (Text)    — maps to ItemCategory
      - TemplateScope      (Choice)  — cahp / owner / property
      - TemplateState      (Choice)  — SC / NC, blank = applies to all
      - TemplateLibrary    (Text)    — optional SharePoint library override
      - TemplateNotes      (Note)
      - TemplateSortOrder  (Number)  — preserves display order

    On first run, seeds the 12-item DOR_FILING_CHECKLIST default set so the
    app has something to work with immediately. Subsequent runs detect any
    existing rows and skip the seed.

    Idempotent — safe to re-run.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-checklist-templates.ps1 `
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
Write-Host "  Checklist Templates list provisioning" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
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

# =============================================================================
# Part 1: Create the Checklist Templates list
# =============================================================================

$ListTitle = "Checklist Templates"
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "-> Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/ChecklistTemplates" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Filing Checklist Templates used by the Property Wizard + Filing Checklist Generator. One row per template item. Edit in the app at Settings -> Checklist Templates." `
        -EnableVersioning $true | Out-Null
    Write-Host "  List created" -ForegroundColor Green
} else {
    Write-Host "-> List already exists; verifying columns" -ForegroundColor Yellow
}

# Columns
$columns = @(
    @{ Display = "Category";    Internal = "TemplateCategory";  Type = "Text";   InView = $true }
    @{ Display = "Scope";       Internal = "TemplateScope";     Type = "Choice"; InView = $true; Choices = @("cahp","owner","property") }
    @{ Display = "State";       Internal = "TemplateState";     Type = "Choice"; InView = $true; Choices = @("SC","NC") }
    @{ Display = "Library";     Internal = "TemplateLibrary";   Type = "Text";   InView = $false }
    @{ Display = "Notes";       Internal = "TemplateNotes";     Type = "Note";   InView = $false }
    @{ Display = "Sort Order";  Internal = "TemplateSortOrder"; Type = "Number"; InView = $true }
)

foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListTitle -Identity $col.Internal -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  -> $($col.Display) exists, skipping" -ForegroundColor DarkGray
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
        Write-Host "  ! Failed: $($col.Display) -- $_" -ForegroundColor Red
    }
}

# =============================================================================
# Part 2: Seed default DOR templates if the list is empty
# =============================================================================

$existingItems = Get-PnPListItem -List $ListTitle -PageSize 500 -Fields "ID","Title" -ErrorAction SilentlyContinue
$itemCount = ($existingItems | Measure-Object).Count

if ($itemCount -gt 0) {
    Write-Host ""
    Write-Host "  -> List has $itemCount existing item(s); skipping default seed" -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "-> Seeding the 12 DOR default templates..." -ForegroundColor White

    $defaults = @(
        @{ Title="CAHP Operating Agreement (Non Profit OA)";              Category="Operating Agreement";          Scope="cahp";     Notes="CAHP SC LLC Operating Agreement. Lives at the CAHP entity level - should be reusable across all filings." }
        @{ Title="CAHP 501(c)(3) Determination Letter";                   Category="501(c)(3) Determination";       Scope="cahp";     Notes="IRS determination letter for Carolina Affordable Housing Project (501(c)(3) status)." }
        @{ Title="CAHP EIN Confirmation";                                  Category="EIN Confirmation";              Scope="cahp";     Notes="IRS EIN letter for the nonprofit." }
        @{ Title="CAHP Articles of Incorporation";                         Category="Articles of Incorporation";     Scope="cahp";     Notes="Nonprofit Articles of Incorporation." }
        @{ Title="CAHP Certificate of Existence (COE)";                    Category="Certificate of Existence";      Scope="cahp";     Notes="State-issued Certificate of Existence / Good Standing for the nonprofit." }
        @{ Title="Entity Certification Letter (Cert of Authorization)";    Category="Certificate of Authorization";  Scope="owner";    Notes="State-issued Cert of Authorization for the property-owning LLC." }
        @{ Title="Entity EIN Confirmation";                                Category="EIN Confirmation";              Scope="owner";    Notes="IRS EIN letter for the property-owning LLC." }
        @{ Title="Entity Operating Agreement";                             Category="Operating Agreement";           Scope="owner";    Notes="Operating Agreement for the property-owning LLC." }
        @{ Title="Entity Articles of Organization";                        Category="Articles of Incorporation";     Scope="owner";    Notes="Articles of Organization for the property-owning LLC." }
        @{ Title="Property Deed(s)";                                       Category="Deed";                          Scope="property"; Notes="Recorded property deed(s). Multiple parcels = multiple deeds." }
        @{ Title="Rent Roll (current year)";                               Category="Rent Roll";                     Scope="property"; Notes="Current-year rent roll showing tenant income qualification." }
        @{ Title="IRS Determination Letter (property-specific, if applicable)"; Category="Determination Letter";      Scope="property"; Notes="Property-specific IRS determination, if one exists for this filing." }
    )

    $sortOrder = 0
    foreach ($d in $defaults) {
        Add-PnPListItem -List $ListTitle -Values @{
            Title             = $d.Title
            TemplateCategory  = $d.Category
            TemplateScope     = $d.Scope
            TemplateNotes     = $d.Notes
            TemplateSortOrder = $sortOrder
        } | Out-Null
        $sortOrder++
    }

    Write-Host "  Seeded $sortOrder default template row(s)" -ForegroundColor Green
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Go to Settings -> Checklist Templates." -ForegroundColor Yellow
Write-Host "  3. Confirm the list looks right. If you had customizations saved" -ForegroundColor Yellow
Write-Host "     in this browser from the old version, click 'Import from this" -ForegroundColor Yellow
Write-Host "     browser' to push them up to SharePoint." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
