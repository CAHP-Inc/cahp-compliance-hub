<#
.SYNOPSIS
    Add 'Owner' and 'Property' lookup columns to every document library the
    app uses, so uploads can tag their files and the EntityDocumentsSection
    can filter by entity.

.DESCRIPTION
    The original provision-sharepoint.ps1 only added a 'Property ID' TEXT
    column to libraries, not the lookup columns the app actually reads/writes
    ('OwnerLookupId' / 'PropertyLookupId'). Result: uploads landed in the
    library with no entity tag and never appeared on owner/property pages.

    This script adds two real Lookup columns to each library that's missing
    them:
      - Owner       → Owners list (Title)
      - Property    → Properties Registry (Title)

    The app's upload code already tolerates several internal-name variants,
    but the canonical InternalName here is 'Owner' / 'Property'. SharePoint
    surfaces these as 'OwnerLookupId' / 'PropertyLookupId' on read via Graph,
    which is what the app expects.

    Idempotent — re-running is safe.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-document-library-columns.ps1 `
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
Write-Host "  Add Owner + Property lookup columns to all libraries" -ForegroundColor Cyan
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

# Look up the source lists for the Lookup columns
$ownersList = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if (-not $ownersList) {
    Write-Error "'Owners' list not found. Run provision-owners.ps1 first."
    exit 1
}
$propertiesList = Get-PnPList -Identity "Properties Registry" -ErrorAction SilentlyContinue
if (-not $propertiesList) {
    Write-Error "'Properties Registry' list not found. Run provision-sharepoint.ps1 first."
    exit 1
}

# Libraries the app reads — must match PROPERTY_LINKED_LIBRARIES +
# CAHP_ENTITY_LIBRARY in src/components/UploadDocumentModal.tsx.
$targetLibraries = @(
    "Operating Agreements",
    "LURAs",
    "AMI Certifications",
    "Rent Rolls",
    "Property Deeds",
    "Org Charts",
    "Supporting Documentation",
    "DOR Submittal Packages",
    "CAHP Entity Documents"
)

$ownerColAdded = 0
$ownerColExisted = 0
$propColAdded = 0
$propColExisted = 0
$libsTouched = 0
$libsMissing = 0

foreach ($lib in $targetLibraries) {
    Write-Host "-> $lib" -ForegroundColor White

    $list = Get-PnPList -Identity $lib -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "   ! Library not found; skipping" -ForegroundColor Red
        $libsMissing++
        continue
    }
    $libsTouched++

    # Owner lookup column (multi-value so one doc can be tagged to multiple owners,
    # e.g. an Assignment of LLC Interest letter that applies to both the assignor
    # entity and the assignee entity)
    $ownerField = Get-PnPField -List $lib -Identity "Owner" -ErrorAction SilentlyContinue
    $needsRecreate = $false
    if ($ownerField) {
        # If it's not already a multi-value lookup, drop + recreate
        if ($ownerField.TypeAsString -ne 'LookupMulti') {
            Write-Host "   Owner column exists as single-value — recreating as multi-value" -ForegroundColor Yellow
            try {
                Remove-PnPField -List $lib -Identity "Owner" -Force | Out-Null
                $needsRecreate = $true
            } catch {
                Write-Host "   ! Could not drop the existing single-value Owner column: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "   Owner multi-value lookup column already exists" -ForegroundColor DarkGray
            $ownerColExisted++
        }
    } else {
        $needsRecreate = $true
    }
    if ($needsRecreate) {
        try {
            $schemaXml = @"
<Field Type="LookupMulti" Mult="TRUE" DisplayName="Owner" Name="Owner" StaticName="Owner"
       List="{$($ownersList.Id)}" ShowField="Title" />
"@
            Add-PnPFieldFromXml -List $lib -FieldXml $schemaXml | Out-Null
            Write-Host "   + Added Owner multi-value lookup column" -ForegroundColor Cyan
            $ownerColAdded++
        } catch {
            Write-Host "   ! Failed to add Owner lookup: $_" -ForegroundColor Red
        }
    }

    # Property lookup column — skip for the CAHP Entity library since CAHP docs
    # aren't property-scoped (they're shared across every filing).
    if ($lib -eq "CAHP Entity Documents") {
        Write-Host "   (Property column intentionally not added — CAHP docs are shared)" -ForegroundColor DarkGray
        continue
    }

    $propField = Get-PnPField -List $lib -Identity "Property" -ErrorAction SilentlyContinue
    $needsRecreate = $false
    if ($propField) {
        if ($propField.TypeAsString -ne 'LookupMulti') {
            Write-Host "   Property column exists as single-value — recreating as multi-value" -ForegroundColor Yellow
            try {
                Remove-PnPField -List $lib -Identity "Property" -Force | Out-Null
                $needsRecreate = $true
            } catch {
                Write-Host "   ! Could not drop the existing single-value Property column: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "   Property multi-value lookup column already exists" -ForegroundColor DarkGray
            $propColExisted++
        }
    } else {
        $needsRecreate = $true
    }
    if ($needsRecreate) {
        try {
            $schemaXml = @"
<Field Type="LookupMulti" Mult="TRUE" DisplayName="Property" Name="Property" StaticName="Property"
       List="{$($propertiesList.Id)}" ShowField="Title" />
"@
            Add-PnPFieldFromXml -List $lib -FieldXml $schemaXml | Out-Null
            Write-Host "   + Added Property multi-value lookup column" -ForegroundColor Cyan
            $propColAdded++
        } catch {
            Write-Host "   ! Failed to add Property lookup: $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Libraries touched   : $libsTouched" -ForegroundColor White
Write-Host "  Libraries missing   : $libsMissing" -ForegroundColor $(if ($libsMissing -gt 0) { 'Red' } else { 'White' })
Write-Host "  Owner columns added : $ownerColAdded (already existed on $ownerColExisted)" -ForegroundColor White
Write-Host "  Property columns added: $propColAdded (already existed on $propColExisted)" -ForegroundColor White
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app." -ForegroundColor Yellow
Write-Host "  2. Try uploading a document to an owner page. It should now" -ForegroundColor Yellow
Write-Host "     appear immediately on the Owner Documents section." -ForegroundColor Yellow
Write-Host "  3. For any existing uploads that 'disappeared into the void':" -ForegroundColor Yellow
Write-Host "     open the library directly in SharePoint, locate the file," -ForegroundColor Yellow
Write-Host "     and set the new 'Owner' (or 'Property') column on each row." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
