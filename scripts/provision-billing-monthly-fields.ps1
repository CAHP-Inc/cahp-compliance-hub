<#
.SYNOPSIS
    Adds the two inputs the Monthly Billing view needs to prorate the annual CAHP
    fee into a monthly amount, on the Billing Tracker list:
      PreviouslyAbated - Boolean. TRUE = abatement was already in effect, bill all
                         12 months. FALSE = newly abated this year, prorate from
                         BillStartDate.
      BillStartDate    - DateTime. The date monthly billing starts (used to
                         prorate the first year: remaining months, +1 if the start
                         is before mid-month).

.DESCRIPTION
    The annual % of savings fee is already stored on each Percent-of-Savings
    Billing row (AmountBilled = savings x CAHPFeePercent). These two columns let
    the Monthly Billing tab compute Months and Monthly = annual / Months, exactly
    like the CAHP Bill Backs spreadsheet.

    Idempotent - re-running skips columns that already exist.

.EXAMPLE
    .\provision-billing-monthly-fields.ps1 `
      -SiteUrl "https://newshirepmcom.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "7f310acf-12b1-4ba9-a113-c027614268b9"
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
Write-Host "  Billing Tracker: add monthly-proration columns" -ForegroundColor Cyan
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
    @{ Internal = "PreviouslyAbated"; Display = "Previously Abated"; Type = "Boolean" },
    @{ Internal = "BillStartDate";    Display = "Bill Start Date";   Type = "DateTime" }
)

foreach ($c in $columns) {
    if (Get-PnPField -List $BillingTitle -Identity $c.Internal -ErrorAction SilentlyContinue) {
        Write-Host "-> $($c.Internal) already exists, skipping" -ForegroundColor DarkGray
    } else {
        Add-PnPField -List $BillingTitle `
            -DisplayName $c.Display `
            -InternalName $c.Internal `
            -Type $c.Type | Out-Null
        Write-Host "  + $($c.Internal) column added ($($c.Type))" -ForegroundColor Cyan
    }
}

Write-Host ""
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "  NEXT: hard-refresh the app; open Billing > Monthly Billing." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
