<#
.SYNOPSIS
    Provisions schema changes for DOR Correspondence multi-property + general comms.

.DESCRIPTION
    Two changes to the existing DOR Correspondence Log:
      1. Adds a CorrChannel Choice column (Letter / Email / Phone / Meeting / Other)
         so the app can distinguish formal letters from ad-hoc calls/emails.
      2. Creates a new Correspondence Property Links junction list with
         Corr + Property lookups, so a single DOR comm can reference many
         properties (or none at all for general inquiries).

    On first run, migrates any existing PropertyLookupId values from the
    Correspondence rows into the new junction list as a starting point.

    Idempotent — safe to re-run.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-correspondence-channel-links.ps1 `
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
Write-Host "  DOR Correspondence: CorrChannel column + property junction" -ForegroundColor Cyan
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

# Pre-flight checks
$corrList = Get-PnPList -Identity "DOR Correspondence Log" -ErrorAction SilentlyContinue
if (-not $corrList) {
    Write-Error "'DOR Correspondence Log' list not found. Aborting."
    exit 1
}
$propsList = Get-PnPList -Identity "Properties Registry" -ErrorAction SilentlyContinue
if (-not $propsList) {
    Write-Error "'Properties Registry' list not found. Aborting."
    exit 1
}

$corrListId  = $corrList.Id.ToString()
$propsListId = $propsList.Id.ToString()

# =============================================================================
# Part 1: Add CorrChannel column to DOR Correspondence Log
# =============================================================================

if (-not (Get-PnPField -List "DOR Correspondence Log" -Identity "CorrChannel" -ErrorAction SilentlyContinue)) {
    Write-Host "-> Adding CorrChannel choice column to 'DOR Correspondence Log'..." -ForegroundColor White
    Add-PnPField -List "DOR Correspondence Log" `
        -DisplayName "Channel" `
        -InternalName "CorrChannel" `
        -Type Choice `
        -Choices "Letter","Email","Phone","Meeting","Other" `
        -AddToDefaultView | Out-Null
    Write-Host "  + CorrChannel column added (default values: Letter for legacy rows)" -ForegroundColor Cyan
} else {
    Write-Host "-> CorrChannel column already exists, skipping" -ForegroundColor DarkGray
}

# =============================================================================
# Part 2: Create Correspondence Property Links junction list
# =============================================================================

$JunctionTitle = "Correspondence Property Links"
$junctionList = Get-PnPList -Identity $JunctionTitle -ErrorAction SilentlyContinue
if (-not $junctionList) {
    Write-Host "-> Creating list '$JunctionTitle'..." -ForegroundColor White
    New-PnPList -Title $JunctionTitle -Template GenericList -Url "lists/CorrespondencePropertyLinks" | Out-Null
    Set-PnPList -Identity $JunctionTitle `
        -Description "Junction list. One row per (DOR Correspondence, Property) linkage. Lets a single comm span multiple properties or none." `
        -EnableVersioning $true | Out-Null
    Write-Host "  Junction list created" -ForegroundColor Green
} else {
    Write-Host "-> Junction list already exists; verifying columns" -ForegroundColor Yellow
}

if (-not (Get-PnPField -List $JunctionTitle -Identity "Corr" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Corr' Name='Corr' List='{$corrListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $JunctionTitle -FieldXml $xml | Out-Null
    Write-Host "  + Corr [Lookup -> DOR Correspondence Log]" -ForegroundColor Cyan
}
if (-not (Get-PnPField -List $JunctionTitle -Identity "Property" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Property' Name='Property' List='{$propsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $JunctionTitle -FieldXml $xml | Out-Null
    Write-Host "  + Property [Lookup -> Properties Registry]" -ForegroundColor Cyan
}

# Add Corr + Property to default view
$junctionView = Get-PnPView -List $JunctionTitle -Identity "All Items" -ErrorAction SilentlyContinue
if ($junctionView) {
    $viewFields = $junctionView.ViewFields
    $newFields = $viewFields
    if ($newFields -notcontains "Corr") { $newFields = $newFields + "Corr" }
    if ($newFields -notcontains "Property") { $newFields = $newFields + "Property" }
    if ($newFields.Count -ne $viewFields.Count) {
        Set-PnPView -List $JunctionTitle -Identity "All Items" -Fields $newFields | Out-Null
    }
}

# =============================================================================
# Part 3: Migrate existing single PropertyLookupId values into the junction
# =============================================================================

Write-Host ""
Write-Host "-> Migrating existing Correspondence.PropertyLookupId values..." -ForegroundColor White

$existingCorr = Get-PnPListItem -List "DOR Correspondence Log" -PageSize 500 -Fields "ID","Property" -ErrorAction SilentlyContinue
$existingLinks = Get-PnPListItem -List $JunctionTitle -PageSize 500 -Fields "ID","Corr","Property" -ErrorAction SilentlyContinue

$migrated = 0
$skipped = 0

foreach ($c in ($existingCorr ?? @())) {
    $cId = $c.Id
    $propRef = $c["Property"]
    if (-not ($propRef -and $propRef.LookupId)) { continue }
    $pId = $propRef.LookupId

    $alreadyLinked = $false
    foreach ($row in ($existingLinks ?? @())) {
        $rowCorr = $row["Corr"]
        $rowProp = $row["Property"]
        if ($rowCorr -and $rowProp -and $rowCorr.LookupId -eq $cId -and $rowProp.LookupId -eq $pId) {
            $alreadyLinked = $true
            break
        }
    }
    if ($alreadyLinked) {
        $skipped++
        continue
    }
    Add-PnPListItem -List $JunctionTitle -Values @{
        Title = "Corr $cId <-> Property $pId"
        Corr = $cId
        Property = $pId
    } | Out-Null
    $migrated++
}

Write-Host "  Migration: $migrated new junction row(s) created; $skipped already linked" -ForegroundColor Green

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. The DOR Correspondence page now has two buttons:" -ForegroundColor Yellow
Write-Host "     - Log Letter: formal letters that cascade to Outstanding Items" -ForegroundColor Yellow
Write-Host "     - Log Comm: phone calls, emails, meetings (notes-only)" -ForegroundColor Yellow
Write-Host "  3. Both modals support multi-property select or no property at all." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
