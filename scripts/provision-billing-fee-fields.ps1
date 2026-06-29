<#
.SYNOPSIS
    Adds the CAHP Filing Fee, Billing Type, and Bill Submittal lookup columns to
    the Billing Tracker list so the two kinds of CAHP fee invoice can be
    generated from an approved submittal.

.DESCRIPTION
    CAHP fees are billed as two separate invoices:
      - Filing Fee         : a flat, one-time charge per property (Initial filing).
      - Percent of Savings : a % of the DOR-approved tax savings, billed per tax
                             year (Initial year + each Annual thereafter).
    Both are revenue CAHP collects from the owner; there is no owner disbursement.

    BillingType     - Choice ('Filing Fee' | 'Percent of Savings'). Distinguishes
                      the two invoice kinds so the UI and reporting can separate
                      one-time fees from recurring contingency fees.

    CAHPFilingFee   - Currency. The flat one-time filing-fee amount (used on
                      'Filing Fee' invoices). The 'Percent of Savings' invoice
                      amount is BillApprovedAbatement * CAHPFeePercent / 100.

    BillSubmittal   - Lookup -> Submittals Tracker. Links each invoice back to
                      the approved filing it bills against. Mirrors
                      DisbSubmittal on the Disbursements list. Graph surfaces
                      this as BillSubmittalLookupId. Used to detect whether a
                      submittal already has its % of savings invoice and to drive
                      the Generate Invoice actions on approved submittals.

    Idempotent - re-running skips columns that already exist.

.EXAMPLE
    .\provision-billing-fee-fields.ps1 `
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
Write-Host "  Billing Tracker: add filing-fee + submittal-link columns" -ForegroundColor Cyan
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

$submittals = Get-PnPList -Identity "Submittals Tracker" -ErrorAction SilentlyContinue
if (-not $submittals) {
    Write-Error "'Submittals Tracker' list not found (needed as the BillSubmittal lookup target)."
    exit 1
}
$submittalsListId = $submittals.Id.ToString()

# 1. BillingType (Choice) -----------------------------------------------------
if (Get-PnPField -List $BillingTitle -Identity "BillingType" -ErrorAction SilentlyContinue) {
    Write-Host "-> BillingType column already exists, skipping" -ForegroundColor DarkGray
} else {
    Add-PnPField -List $BillingTitle `
        -DisplayName "Billing Type" `
        -InternalName "BillingType" `
        -Type Choice `
        -Choices @("Filing Fee", "Percent of Savings") | Out-Null
    Write-Host "  + BillingType column added (Choice)" -ForegroundColor Cyan
}

# 2. CAHPFilingFee (Currency) -------------------------------------------------
if (Get-PnPField -List $BillingTitle -Identity "CAHPFilingFee" -ErrorAction SilentlyContinue) {
    Write-Host "-> CAHPFilingFee column already exists, skipping" -ForegroundColor DarkGray
} else {
    Add-PnPField -List $BillingTitle `
        -DisplayName "CAHP Filing Fee" `
        -InternalName "CAHPFilingFee" `
        -Type Currency | Out-Null
    Write-Host "  + CAHPFilingFee column added (Currency)" -ForegroundColor Cyan
}

# 3. BillSubmittal (Lookup -> Submittals Tracker) -----------------------------
#    Graph exposes this as BillSubmittalLookupId.
if (Get-PnPField -List $BillingTitle -Identity "BillSubmittal" -ErrorAction SilentlyContinue) {
    Write-Host "-> BillSubmittal column already exists, skipping" -ForegroundColor DarkGray
} else {
    $xml = @"
<Field Type='Lookup' DisplayName='Bill Submittal' Name='BillSubmittal' List='{$submittalsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $BillingTitle -FieldXml $xml | Out-Null
    Write-Host "  + BillSubmittal column added (Lookup -> Submittals Tracker)" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Approve a submittal -> the Approval modal now captures a flat" -ForegroundColor Yellow
Write-Host "     filing fee + a % of tax savings, and creates the invoice +" -ForegroundColor Yellow
Write-Host "     owner disbursement automatically." -ForegroundColor Yellow
Write-Host "  3. For submittals already marked Approved without an invoice, open" -ForegroundColor Yellow
Write-Host "     the submittal and click 'Generate Invoice'." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
