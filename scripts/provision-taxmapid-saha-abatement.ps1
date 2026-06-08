<#
.SYNOPSIS
    Adds the PriorSAHAAbatement (Yes/No) and PriorSAHANotes (text) columns to the
    Tax Map IDs list, so a parcel can be tagged as previously approved for
    property tax abatement under SAHA.

.DESCRIPTION
    PriorSAHAAbatement - Yes/No flag. Surfaces a "SAHA" badge on the parcel in the
                         property's Tax Map IDs section and on the submittal detail
                         page (filing context).
    PriorSAHANotes     - Optional free text (year approved, reference #, context).

    Idempotent.

.EXAMPLE
    .\provision-taxmapid-saha-abatement.ps1 `
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
Write-Host "  Tax Map IDs: add prior SAHA abatement columns" -ForegroundColor Cyan
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

$list = Get-PnPList -Identity "Tax Map IDs" -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Error "'Tax Map IDs' list not found."
    exit 1
}

if (Get-PnPField -List "Tax Map IDs" -Identity "PriorSAHAAbatement" -ErrorAction SilentlyContinue) {
    Write-Host "-> PriorSAHAAbatement column already exists, skipping" -ForegroundColor DarkGray
} else {
    Add-PnPField -List "Tax Map IDs" `
        -DisplayName "Prior SAHA Abatement" `
        -InternalName "PriorSAHAAbatement" `
        -Type Boolean `
        -AddToDefaultView | Out-Null
    Write-Host "  + PriorSAHAAbatement column added (Yes/No)" -ForegroundColor Cyan
}

if (Get-PnPField -List "Tax Map IDs" -Identity "PriorSAHANotes" -ErrorAction SilentlyContinue) {
    Write-Host "-> PriorSAHANotes column already exists, skipping" -ForegroundColor DarkGray
} else {
    Add-PnPField -List "Tax Map IDs" `
        -DisplayName "SAHA Abatement Notes" `
        -InternalName "PriorSAHANotes" `
        -Type Note | Out-Null
    Write-Host "  + PriorSAHANotes column added (multi-line text)" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. On a property's Tax Map IDs section, Edit a parcel and tick" -ForegroundColor Yellow
Write-Host "     'Previously approved for abatement under SAHA'. A gold SAHA badge" -ForegroundColor Yellow
Write-Host "     then shows on the parcel and on its submittal detail page." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
