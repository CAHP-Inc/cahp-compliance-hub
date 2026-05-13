<#
.SYNOPSIS
    Provisions the AuditLog SharePoint list for PR-07.

.DESCRIPTION
    Creates the AuditLog list with all 7 columns needed to capture every CRUD operation
    across the CAHP Compliance Hub data layer.

    Idempotent — re-running is safe; existing list and columns are detected and skipped.

    REQUIRES: "Allow public client flows" toggled ON on the Azure AD app at the moment
    you run this script. Toggle it OFF afterward so the SPA sign-in flow stays clean.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID (same value used in the React app).

.EXAMPLE
    .\provision-auditlog.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "eeb92696-399a-4394-858c-ee73de0e94c6"

.NOTES
    Authored by Brandy Turner / NewShire Property Management
    Part of: CAHP Compliance Hub — Phase 1, PR-07
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  PR-07: AuditLog list provisioning" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Host ""
    Write-Host "Connection failed. Most common cause:" -ForegroundColor Red
    Write-Host "  'Allow public client flows' is OFF on the Azure AD app." -ForegroundColor Red
    Write-Host "  Toggle it ON: Azure Portal → App registrations → CAHP Compliance Hub →" -ForegroundColor Red
    Write-Host "  Authentication → Settings → Allow public client flows → Yes → Save" -ForegroundColor Red
    Write-Host ""
    Write-Error "Original error: $_"
    exit 1
}
Write-Host "  ✓ Connected" -ForegroundColor Green
Write-Host ""

$ListTitle = "AuditLog"

$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "→ Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/$ListTitle" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Append-only audit log of every CRUD operation across CAHP Compliance Hub. Each row records who changed what, when, and what the before/after values were." `
        -EnableVersioning $true | Out-Null
    Write-Host "  ✓ List created" -ForegroundColor Green
} else {
    Write-Host "→ List '$ListTitle' already exists; verifying columns" -ForegroundColor Yellow
}

# Column definitions
$columns = @(
    @{ Display = "Action"; Internal = "Action"; Type = "Choice"; Choices = @("CREATE", "UPDATE", "DELETE"); InView = $true }
    @{ Display = "Entity Type"; Internal = "EntityType"; Type = "Text"; InView = $true }
    @{ Display = "Entity ID"; Internal = "EntityId"; Type = "Text"; InView = $true }
    @{ Display = "Entity Title"; Internal = "EntityTitle"; Type = "Text"; InView = $true }
    @{ Display = "Change Summary"; Internal = "ChangeSummary"; Type = "Note"; InView = $true }
    @{ Display = "Before JSON"; Internal = "BeforeJSON"; Type = "Note"; InView = $false }
    @{ Display = "After JSON"; Internal = "AfterJSON"; Type = "Note"; InView = $false }
)

foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListTitle -Identity $col.Internal -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  → $($col.Display) exists, skipping" -ForegroundColor DarkGray
        continue
    }

    try {
        $params = @{
            List         = $ListTitle
            DisplayName  = $col.Display
            InternalName = $col.Internal
            Type         = $col.Type
        }
        if ($col.InView) { $params.AddToDefaultView = $true }
        if ($col.Type -eq "Choice") { $params.Choices = $col.Choices }

        Add-PnPField @params -ErrorAction Stop | Out-Null
        Write-Host "  + $($col.Display) [$($col.Type)]" -ForegroundColor Cyan
    } catch {
        Write-Host "  ! Failed: $($col.Display) — $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "REMEMBER: Toggle 'Allow public client flows' back to OFF on the Azure AD" -ForegroundColor Yellow
Write-Host "app to keep the SPA sign-in flow clean:" -ForegroundColor Yellow
Write-Host "  Azure Portal → App registrations → CAHP Compliance Hub →" -ForegroundColor Yellow
Write-Host "  Authentication → Settings → Allow public client flows → No → Save" -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
