<#
.SYNOPSIS
    Adds Last Full Tax Bill + Most Recent Tax Bill currency columns to the Billing
    Tracker so the annual CAHP "% of Annual Savings" fee can be computed from the
    two tax bills: savings = LastFullTaxBill - MostRecentTaxBill, fee = savings x %.

.DESCRIPTION
    The recurring % of savings invoice records the abatement basis as two figures:
      LastFullTaxBill    - Currency. The full (pre-abatement) property tax bill.
      MostRecentTaxBill  - Currency. The most recent (abated) tax bill.
    The savings (BillApprovedAbatement) is their difference, and the CAHP fee is
    that difference x CAHPFeePercent / 100.

    REQUIRED before the annual % billing UI can save — Graph rejects writes to a
    column that doesn't exist.

    Idempotent - re-running skips columns that already exist.

.EXAMPLE
    .\provision-billing-tax-bill-fields.ps1 `
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
Write-Host "  Billing Tracker: add tax-bill columns (annual % of savings)" -ForegroundColor Cyan
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

$columns = @(
    @{ Internal = "LastFullTaxBill";   Display = "Last Full Tax Bill" },
    @{ Internal = "MostRecentTaxBill"; Display = "Most Recent Tax Bill" }
)

foreach ($c in $columns) {
    if (Get-PnPField -List $BillingTitle -Identity $c.Internal -ErrorAction SilentlyContinue) {
        Write-Host "-> $($c.Internal) already exists, skipping" -ForegroundColor DarkGray
    } else {
        Add-PnPField -List $BillingTitle `
            -DisplayName $c.Display `
            -InternalName $c.Internal `
            -Type Currency | Out-Null
        Write-Host "  + $($c.Internal) column added (Currency)" -ForegroundColor Cyan
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Property > Billing tab, or Billing > To Invoice (roll forward):" -ForegroundColor Yellow
Write-Host "     enter the last full + most recent tax bills to bill the CAHP %," -ForegroundColor Yellow
Write-Host "     or Mark N/A for a year you are not claiming." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
