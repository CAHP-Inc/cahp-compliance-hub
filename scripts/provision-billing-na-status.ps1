<#
.SYNOPSIS
    Adds an 'N/A' choice to the Billing Tracker Status (BillingStatus) column so a
    property's initial filing fee can be marked "not charged".

.DESCRIPTION
    The hub records "we are not charging this property an initial filing fee" as a
    $0 Filing Fee billing row with BillingStatus = 'N/A'. That row still satisfies
    the filing-fee dedupe, so the property drops out of the to-invoice queue while
    leaving an auditable record. This script appends 'N/A' to the existing Status
    choices (preserving all current values).

    Idempotent - re-running skips the choice if it already exists.

    NOTE: Graph writes to a SharePoint Choice field generally accept off-list
    values, so the app can already write 'N/A' without this. Running it just makes
    'N/A' a first-class, filterable choice in SharePoint views.

.EXAMPLE
    .\provision-billing-na-status.ps1 `
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
Write-Host "  Billing Tracker: add 'N/A' Status choice" -ForegroundColor Cyan
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
$billing = Get-PnPList -Identity $BillingTitle -ErrorAction SilentlyContinue
if (-not $billing) {
    Write-Error "'$BillingTitle' list not found."
    exit 1
}

$field = Get-PnPField -List $BillingTitle -Identity "BillingStatus" -ErrorAction SilentlyContinue
if (-not $field) {
    Write-Error "'BillingStatus' (Status) column not found on '$BillingTitle'."
    exit 1
}

# Read the current choices straight from the field schema (reliable), then append.
[xml]$schema = $field.SchemaXml
$current = @()
if ($schema.Field.CHOICES) { $current = @($schema.Field.CHOICES.CHOICE) }

if ($current -contains "N/A") {
    Write-Host "-> 'N/A' is already a Status choice, skipping" -ForegroundColor DarkGray
} else {
    $updated = $current + "N/A"
    Set-PnPField -List $BillingTitle -Identity "BillingStatus" -Values @{ Choices = $updated } | Out-Null
    Write-Host "  + 'N/A' added to Status choices (kept: $($current -join ', '))" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Billing > To Invoice: 'Mark N/A' on a filing-fee row, or" -ForegroundColor Yellow
Write-Host "     a Property's Billing tab: 'Mark filing fee N/A'." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
