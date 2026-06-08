<#
.SYNOPSIS
    Adds the DateResponded and DateLetterReceived date columns to the Submittals
    Tracker list. These anchor the DOR clocks surfaced on the Submittals page and
    My Day ("DOR Deadlines").

.DESCRIPTION
    DateResponded     - Date we sent our response back to DOR. Anchors the ~12-week
                        "Expect DOR response" clock when a submittal moves to
                        "Responded - Awaiting DOR".
    DateLetterReceived - Date we received a DOR RFI letter. Anchors the 30-day
                        "Respond to DOR" clock when a submittal moves to
                        "Letter Received - Action Needed".

    Both feed Next Action / Next Action Due automatically at the status transition,
    and can be backfilled on existing submittals (Edit -> set date -> Save
    recomputes the clock for submittals still in the matching state).

    Idempotent.

.EXAMPLE
    .\provision-submittal-dor-clock-fields.ps1 `
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
Write-Host "  Submittals Tracker: add DOR clock date columns" -ForegroundColor Cyan
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

$list = Get-PnPList -Identity "Submittals Tracker" -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Error "'Submittals Tracker' list not found."
    exit 1
}

$columns = @(
    @{ Internal = "DateResponded";      Display = "Date Responded" }
    @{ Internal = "DateLetterReceived"; Display = "Date RFI Received" }
)

foreach ($col in $columns) {
    if (Get-PnPField -List "Submittals Tracker" -Identity $col.Internal -ErrorAction SilentlyContinue) {
        Write-Host "-> $($col.Internal) column already exists, skipping" -ForegroundColor DarkGray
    } else {
        Add-PnPField -List "Submittals Tracker" `
            -DisplayName $col.Display `
            -InternalName $col.Internal `
            -Type DateTime | Out-Null
        Write-Host "  + $($col.Internal) column added (Date)" -ForegroundColor Cyan
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. New transitions now capture the relevant date and set the clock:" -ForegroundColor Yellow
Write-Host "       - Mark Responded   -> Date Responded   -> DOR response +12 weeks" -ForegroundColor Yellow
Write-Host "       - Letter Received  -> Date RFI Received -> our response +30 days" -ForegroundColor Yellow
Write-Host "  3. Backfill: open a submittal, Edit, set the date, Save -- the clock" -ForegroundColor Yellow
Write-Host "     recomputes for submittals still in the matching state." -ForegroundColor Yellow
Write-Host "  4. Both surface in 'DOR Deadlines' on the Submittals page + My Day." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
