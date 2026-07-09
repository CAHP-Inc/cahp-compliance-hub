<#
.SYNOPSIS
    Adds a RefundStatus choice column to the Billing Tracker so each recorded
    abatement can track the previously-paid-tax refund: whether a refund is
    Needed, has been Requested, has been Approved & Sent, or No Request Needed.

.EXAMPLE
    .\provision-billing-refund-status.ps1 `
      -SiteUrl "https://newshirepmcom.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "7f310acf-12b1-4ba9-a113-c027614268b9"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId
)
$ErrorActionPreference = "Stop"

if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
    $cleanPnP = "C:\Users\brand\PnPModules\PnP.PowerShell\3.2.0\PnP.PowerShell.psd1"
    if (Test-Path $cleanPnP) { Import-Module $cleanPnP -Force }
    elseif (Get-Module -ListAvailable -Name PnP.PowerShell) { Import-Module PnP.PowerShell }
    else { Write-Error "PnP.PowerShell not available. Dot-source Start-CleanSession.ps1 first."; exit 1 }
}

Write-Host "`n  Billing Tracker: add RefundStatus choice column`n" -ForegroundColor Cyan
try { Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop }
catch { Write-Error "Connection failed: $_"; exit 1 }
Write-Host "  Connected" -ForegroundColor Green

$BillingTitle = "Billing Tracker"
if (-not (Get-PnPList -Identity $BillingTitle -ErrorAction SilentlyContinue)) { Write-Error "'$BillingTitle' not found."; exit 1 }

$choices = @("Needed", "Requested", "Approved & Sent", "No Request Needed")
if (Get-PnPField -List $BillingTitle -Identity "RefundStatus" -ErrorAction SilentlyContinue) {
    Write-Host "-> RefundStatus already exists, ensuring choices" -ForegroundColor DarkGray
    Set-PnPField -List $BillingTitle -Identity "RefundStatus" -Values @{ Choices = $choices } | Out-Null
} else {
    Add-PnPField -List $BillingTitle -DisplayName "Refund Status" -InternalName "RefundStatus" `
        -Type Choice -Choices $choices | Out-Null
    Write-Host "  + RefundStatus column added (Choice)" -ForegroundColor Cyan
}

Write-Host "`n  Provisioning complete.`n" -ForegroundColor Green
Disconnect-PnPOnline
