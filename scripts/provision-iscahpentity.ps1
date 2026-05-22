<#
.SYNOPSIS
    Adds the IsCAHPEntity Yes/No column to the Owners list. Used by the
    org chart to highlight the exemption chain on property pages.

.DESCRIPTION
    Check this on every entity in the CAHP family (the parent 501(c)(3)
    nonprofit and its wholly-owned subsidiaries like CAHP SC LLC, CAHP NC LLC).

    The org chart will:
      - Render CAHP-flagged entities with a gold border + "CAHP" badge.
      - Render any LLC that has a CAHP-flagged entity as a direct member
        with a gold ring + "Exemption Source" badge -- those entities'
        formation docs accompany the DOR filing for their subsidiaries.

    Existing entities with "CAHP" or "Carolina Affordable Housing Project"
    in the title are auto-detected as a fallback, so the org chart still
    works before you flip the explicit flag on each row.

    Idempotent.

.EXAMPLE
    .\provision-iscahpentity.ps1 `
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
Write-Host "  Owners: add IsCAHPEntity column" -ForegroundColor Cyan
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

$owners = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if (-not $owners) {
    Write-Error "'Owners' list not found. Run provision-owners.ps1 first."
    exit 1
}

if (Get-PnPField -List "Owners" -Identity "IsCAHPEntity" -ErrorAction SilentlyContinue) {
    Write-Host "-> IsCAHPEntity column already exists, skipping" -ForegroundColor DarkGray
} else {
    Add-PnPField -List "Owners" `
        -DisplayName "CAHP Entity" `
        -InternalName "IsCAHPEntity" `
        -Type Boolean `
        -AddToDefaultView | Out-Null
    Write-Host "  + IsCAHPEntity column added (Yes/No)" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Open the Owners list (or each Owner detail page) and check" -ForegroundColor Yellow
Write-Host "     'CAHP Entity = Yes' on every CAHP-family entity:" -ForegroundColor Yellow
Write-Host "        - Carolina Affordable Housing Project Inc" -ForegroundColor Yellow
Write-Host "        - CAHP SC LLC" -ForegroundColor Yellow
Write-Host "        - CAHP NC LLC (if it exists as an Owner)" -ForegroundColor Yellow
Write-Host "        - Any other wholly-owned CAHP subsidiary" -ForegroundColor Yellow
Write-Host "  3. Open any property's Org Chart tab. CAHP-flagged entities get" -ForegroundColor Yellow
Write-Host "     a gold border + 'CAHP' badge. LLCs with a CAHP entity as a" -ForegroundColor Yellow
Write-Host "     direct member get a gold ring + 'Exemption Source' badge --" -ForegroundColor Yellow
Write-Host "     those entities' docs accompany the DOR filing." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
