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
    Write-Host "  -> To replace existing rows with the latest SC defaults, go to" -ForegroundColor DarkGray
    Write-Host "     Settings -> Checklist Templates and click 'Load defaults' then 'Save'." -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "-> Seeding the SC PT-401-O default templates..." -ForegroundColor White

    # Mirrors DOR_FILING_CHECKLIST in src/lib/filing-checklist.ts. Keep both in sync.
    $defaults = @(
        # CAHP nonprofit corporation docs
        @{ Title="CAHP 501(c)(3) Determination Letter";                          Category="501(c)(3) Determination";   Scope="cahp";     State="SC"; Notes="IRS exempt determination letter for Carolina Affordable Housing Project." }
        @{ Title="CAHP Bylaws";                                                   Category="Bylaws";                    Scope="cahp";     State="SC"; Notes="Bylaws of the nonprofit housing corporation." }
        @{ Title="CAHP Articles of Incorporation (Stamped SC)";                  Category="Articles of Incorporation"; Scope="cahp";     State="SC"; Notes="Stamped SC Articles of Incorporation for the nonprofit corporation." }
        # CAHP SC LLC (wholly-owned affiliate)
        @{ Title="CAHP SC LLC Operating Agreement (LLC <-> sole member)";        Category="Operating Agreement";       Scope="cahp";     State="SC"; Notes="Operating agreement between CAHP SC LLC and its sole member (the nonprofit). Demonstrates wholly-owned-affiliate relationship." }
        @{ Title="CAHP SC LLC Articles of Organization (Stamped SC)";            Category="Articles of Incorporation"; Scope="cahp";     State="SC"; Notes="Stamped SC Articles of Organization for the wholly-owned LLC affiliate." }
        # Property-owning entity docs
        @{ Title="Entity Partnership Agreement / Operating Agreement";           Category="Partnership Agreement";     Scope="owner";    State="SC"; Notes="Partnership agreement, or Operating Agreement if the property-owning entity is an LLC." }
        @{ Title="Entity Articles of Organization (Stamped SC)";                 Category="Articles of Incorporation"; Scope="owner";    State="SC"; Notes="Stamped SC Articles of Organization for the property-owning entity." }
        @{ Title="Organizational Structure Chart";                                Category="Org Chart";                 Scope="owner";    State="SC"; Notes="Org chart showing the property-owning entity, its members, and the chain up to the CAHP nonprofit. The app can export this from the property detail page." }
        # Per-filing property docs
        @{ Title="PT-401-O Exemption Application (completed)";                   Category="Exemption Application";     Scope="property"; State="SC"; Notes="Completed SC PT-401-O exemption application form." }
        @{ Title="Recorded Property Deed(s)";                                     Category="Deed";                      Scope="property"; State="SC"; Notes="Recorded property deed(s) from the county. Multiple parcels = multiple deeds." }
        @{ Title="Rent Roll (current year)";                                      Category="Rent Roll";                 Scope="property"; State="SC"; Notes="Current-year rent roll. Required either as standalone or paired with the recorded restrictive covenants below." }
        @{ Title="Recorded Restrictive Covenants (SC State Housing)";            Category="Restrictive Covenants";     Scope="property"; State="SC"; Notes="Recorded restrictive covenants filed with SC State Housing. Pair with the rent roll above (at least one of the two is required)." }
        @{ Title="Most Recent SC State Housing Compliance Certificate";          Category="Compliance Certificate";    Scope="property"; State="SC"; Notes="Optional but recommended when one exists - supports the rent-roll/covenants documentation." }
        # Conditional re-filing item
        @{ Title="Reassignment of Interest Sign-Off (if re-filing)";             Category="Reassignment of Interest";  Scope="property"; State="SC"; Notes="Only required when the property was previously filed under a different nonprofit and is being re-assigned. Skip otherwise." }
    )

    $sortOrder = 0
    foreach ($d in $defaults) {
        Add-PnPListItem -List $ListTitle -Values @{
            Title             = $d.Title
            TemplateCategory  = $d.Category
            TemplateScope     = $d.Scope
            TemplateState     = $d.State
            TemplateNotes     = $d.Notes
            TemplateSortOrder = $sortOrder
        } | Out-Null
        $sortOrder++
    }

    Write-Host "  Seeded $sortOrder default template row(s) (all tagged SC)" -ForegroundColor Green
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
