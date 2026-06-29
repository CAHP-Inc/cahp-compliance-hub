<#
.SYNOPSIS
    Adds a BillTaxMapID lookup column to the Billing Tracker so an annual "% of
    Annual Savings" invoice can optionally be scoped to a single Tax Map ID
    (parcel) — useful when a property/portfolio has more than one TMID and each
    parcel has its own tax bills.

.DESCRIPTION
    BillTaxMapID - Lookup -> Tax Map IDs. Optional. Graph surfaces it as
                   BillTaxMapIDLookupId. When set, the % invoice is tied to that
                   parcel and de-dupes per (property, year, TMID), so each TMID
                   can be billed separately for the same year. When left blank
                   the invoice is property-level (the default).

    Idempotent - re-running skips the column if it already exists.

.EXAMPLE
    .\provision-billing-taxmapid-link.ps1 `
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
Write-Host "  Billing Tracker: add BillTaxMapID lookup (-> Tax Map IDs)" -ForegroundColor Cyan
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

$BillingTitle = "Billing Tracker"
if (-not (Get-PnPList -Identity $BillingTitle -ErrorAction SilentlyContinue)) {
    Write-Error "'$BillingTitle' list not found."
    exit 1
}

$taxmaps = Get-PnPList -Identity "Tax Map IDs" -ErrorAction SilentlyContinue
if (-not $taxmaps) {
    Write-Error "'Tax Map IDs' list not found (needed as the BillTaxMapID lookup target)."
    exit 1
}
$taxmapsListId = $taxmaps.Id.ToString()

if (Get-PnPField -List $BillingTitle -Identity "BillTaxMapID" -ErrorAction SilentlyContinue) {
    Write-Host "-> BillTaxMapID column already exists, skipping" -ForegroundColor DarkGray
} else {
    $xml = @"
<Field Type='Lookup' DisplayName='Bill Tax Map ID' Name='BillTaxMapID' List='{$taxmapsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $BillingTitle -FieldXml $xml | Out-Null
    Write-Host "  + BillTaxMapID column added (Lookup -> Tax Map IDs)" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT: hard-refresh the app (Ctrl+Shift+R). The annual % of savings" -ForegroundColor Yellow
Write-Host "      panel now shows an optional Tax Map ID selector for multi-TMID" -ForegroundColor Yellow
Write-Host "      properties." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
