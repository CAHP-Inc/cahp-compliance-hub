<#
.SYNOPSIS
    Provisions the "Submittal Reviews" list — a weekly status journal per
    submittal, recorded until the submittal is Approved.

.DESCRIPTION
    Each row is one weekly review of a submittal, capturing the status at that
    time plus a progress note, the planned next action, and an expected
    resolution date. Author + timestamp are auto-captured by SharePoint.

    Columns:
      ReviewSubmittal    - Lookup -> Submittals Tracker (Graph: ReviewSubmittalLookupId)
      ReviewStatus       - Text   (submittal status snapshot at review time)
      ReviewNote         - Note   (weekly progress note)
      ReviewNextAction   - Note   (planned next action)
      ReviewNextActionETA- DateTime (expected resolution date)

    Idempotent — re-running detects and skips the existing list/columns.

.EXAMPLE
    .\provision-submittal-reviews.ps1 `
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
Write-Host "  Submittal Reviews list provisioning" -ForegroundColor Cyan
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

$ListTitle = "Submittal Reviews"

$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/SubmittalReviews" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Weekly status journal per submittal (status + note + next action/ETA), recorded until Approved." `
        -EnableVersioning $true | Out-Null
    Write-Host "  + List created" -ForegroundColor Green
} else {
    Write-Host "-> List already exists; verifying columns" -ForegroundColor Yellow
}

function Add-SimpleField($internal, $display, $type) {
    if (Get-PnPField -List $ListTitle -Identity $internal -ErrorAction SilentlyContinue) {
        Write-Host "  -> $internal exists, skipping" -ForegroundColor DarkGray
    } else {
        Add-PnPField -List $ListTitle -DisplayName $display -InternalName $internal -Type $type -AddToDefaultView | Out-Null
        Write-Host "  + $internal [$type]" -ForegroundColor Cyan
    }
}

Add-SimpleField "ReviewStatus"        "Review Status"      Text
Add-SimpleField "ReviewNote"          "Review Note"        Note
Add-SimpleField "ReviewNextAction"    "Next Action"        Note
Add-SimpleField "ReviewNextActionETA" "Next Action ETA"    DateTime

# ReviewSubmittal lookup -> Submittals Tracker (Graph: ReviewSubmittalLookupId)
if (Get-PnPField -List $ListTitle -Identity "ReviewSubmittal" -ErrorAction SilentlyContinue) {
    Write-Host "  -> ReviewSubmittal exists, skipping" -ForegroundColor DarkGray
} else {
    $submittals = Get-PnPList -Identity "Submittals Tracker" -ErrorAction Stop
    $listId = $submittals.Id.ToString()
    $fieldXml = @"
<Field Type='Lookup' DisplayName='Review Submittal' Name='ReviewSubmittal' List='{$listId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $ListTitle -FieldXml $fieldXml | Out-Null
    $view = Get-PnPView -List $ListTitle -Identity "All Items"
    if ($view.ViewFields -notcontains "ReviewSubmittal") {
        Set-PnPView -List $ListTitle -Identity "All Items" -Fields ($view.ViewFields + "ReviewSubmittal") | Out-Null
    }
    Write-Host "  + ReviewSubmittal [Lookup -> Submittals Tracker]" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Disconnect-PnPOnline
