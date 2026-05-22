<#
.SYNOPSIS
    Provisions the Access List SharePoint list used by Settings -> Access List
    to manage who can sign into the app and at what role. Seeds the original
    four team members on first run.

.DESCRIPTION
    Schema (Title is the email/UPN — case-insensitive match):
      - Title             (Text)    — email / UPN (lookup key)
      - AccessRole        (Choice)  — Admin / Contributor / Accounting
      - AccessDisplayName (Text)
      - AccessOrg         (Text)
      - AccessActive      (Yes/No)  — defaults to Yes; uncheck to deny without deleting
      - AccessNotes       (Note)

    On first run, seeds:
      Brandy Turner (Admin), Stan (Admin), Bryan DeBruin (Admin),
      Lori Heckman (Contributor).

    Subsequent runs are no-ops if the list and its rows already exist —
    safe to re-run, but it won't try to merge changes you've made in-app.

    The app has a hardcoded fallback to these same four people in
    src/lib/roleMap.ts, so the app stays usable even if this list is
    deleted or unreachable.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-access-list.ps1 `
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
Write-Host "  Access List provisioning" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
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

$ListTitle = "Access List"
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "-> Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/AccessList" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Who can sign in and at what role. Managed in-app at Settings -> Access List. Title column = email / UPN." `
        -EnableVersioning $true | Out-Null
    Write-Host "  List created" -ForegroundColor Green
} else {
    Write-Host "-> List already exists; verifying columns" -ForegroundColor Yellow
}

$columns = @(
    @{ Display = "Role";         Internal = "AccessRole";        Type = "Choice";  InView = $true;  Choices = @("Admin","Contributor","Accounting") }
    @{ Display = "Display Name"; Internal = "AccessDisplayName"; Type = "Text";    InView = $true }
    @{ Display = "Org";          Internal = "AccessOrg";         Type = "Text";    InView = $true }
    @{ Display = "Active";       Internal = "AccessActive";      Type = "Boolean"; InView = $true }
    @{ Display = "Notes";        Internal = "AccessNotes";       Type = "Note";    InView = $false }
)

foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListTitle -Identity $col.Internal -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  -> $($col.Display) exists, skipping" -ForegroundColor DarkGray
        continue
    }
    try {
        $params = @{
            List         = $ListTitle
            DisplayName  = $col.Display
            InternalName = $col.Internal
            Type         = $col.Type
        }
        if ($col.InView)            { $params.AddToDefaultView = $true }
        if ($col.Type -eq "Choice") { $params.Choices = $col.Choices }
        Add-PnPField @params -ErrorAction Stop | Out-Null
        Write-Host "  + $($col.Display) [$($col.Type)]" -ForegroundColor Cyan
    } catch {
        Write-Host "  ! Failed: $($col.Display) -- $_" -ForegroundColor Red
    }
}

# =============================================================================
# Seed the original four team members if the list is empty
# =============================================================================

$existingItems = Get-PnPListItem -List $ListTitle -PageSize 500 -Fields "ID","Title" -ErrorAction SilentlyContinue
$itemCount = ($existingItems | Measure-Object).Count

if ($itemCount -gt 0) {
    Write-Host ""
    Write-Host "  -> List has $itemCount existing entr(y/ies); skipping seed" -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "-> Seeding the original four team members..." -ForegroundColor White

    $seed = @(
        @{ Title="bturner@newshirepm.com";     Role="Admin";       DisplayName="Brandy Turner"; Org="NewShire";   Active=$true }
        @{ Title="stan@vanrockre.com";         Role="Admin";       DisplayName="Stan";          Org="VanRock";    Active=$true }
        @{ Title="bdebruin@redcedarhomes.com"; Role="Admin";       DisplayName="Bryan DeBruin"; Org="Red Cedar";  Active=$true }
        @{ Title="lheckman@redcedarhomes.com"; Role="Contributor"; DisplayName="Lori Heckman";  Org="Red Cedar";  Active=$true }
    )

    foreach ($s in $seed) {
        Add-PnPListItem -List $ListTitle -Values @{
            Title             = $s.Title
            AccessRole        = $s.Role
            AccessDisplayName = $s.DisplayName
            AccessOrg         = $s.Org
            AccessActive      = $s.Active
        } | Out-Null
    }

    Write-Host "  Seeded 4 entr(y/ies)" -ForegroundColor Green
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. Settings -> Access List now shows the seeded entries with" -ForegroundColor Yellow
Write-Host "     full add / edit / remove controls." -ForegroundColor Yellow
Write-Host "  3. Adding a new email + saving grants access immediately --" -ForegroundColor Yellow
Write-Host "     no redeploy needed. The new user signs in with their M365" -ForegroundColor Yellow
Write-Host "     account and lands inside the app." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
