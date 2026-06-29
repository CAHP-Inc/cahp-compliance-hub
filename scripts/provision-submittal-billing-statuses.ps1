<#
.SYNOPSIS
    Adds the 'Invoiced' and 'Paid' choices to the Submittals Tracker
    SubmittalStatus column so the billing lifecycle can write them.

.DESCRIPTION
    The submittal lifecycle now extends past Approved:
        Approved -> Invoiced (a CAHP fee invoice was generated)
                 -> Paid     (all of the submittal's invoices are marked Paid)
    Both are set automatically by the app. SharePoint choice columns reject
    values that aren't in their allowed list, so without these two choices the
    status write fails silently and the submittal appears stuck on Approved.

    This script reads the column's CURRENT choices and appends only the missing
    ones, so it won't disturb any other choices already on the live list.

    Idempotent.

.EXAMPLE
    .\provision-submittal-billing-statuses.ps1 `
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
Write-Host "  Submittals Tracker: add Invoiced / Paid status choices" -ForegroundColor Cyan
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

$ListTitle = "Submittals Tracker"
$FieldName = "SubmittalStatus"

$field = Get-PnPField -List $ListTitle -Identity $FieldName -ErrorAction SilentlyContinue
if (-not $field) {
    Write-Error "Column '$FieldName' not found on '$ListTitle'."
    exit 1
}

# Current choices (string array)
$current = @($field.Choices)
Write-Host "  Current choices: $($current -join ', ')" -ForegroundColor DarkGray

$wanted = @("Invoiced", "Paid")
$toAdd = $wanted | Where-Object { $current -notcontains $_ }

if (-not $toAdd -or $toAdd.Count -eq 0) {
    Write-Host "-> Invoiced and Paid already present, nothing to do" -ForegroundColor DarkGray
} else {
    $updated = $current + $toAdd
    Set-PnPField -List $ListTitle -Identity $FieldName -Values @{ Choices = $updated } | Out-Null
    Write-Host "  + Added: $($toAdd -join ', ')" -ForegroundColor Cyan
    Write-Host "  Choices now: $($updated -join ', ')" -ForegroundColor Green
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Disconnect-PnPOnline
