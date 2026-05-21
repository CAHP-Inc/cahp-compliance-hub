<#
.SYNOPSIS
    Provisions the Email Templates SharePoint list used by the in-app email
    composer's template picker.

.DESCRIPTION
    One row per template. Schema:
      - Title              (Text)    — template name (shown in picker)
      - TemplateSubject    (Text)    — subject line (variables OK)
      - TemplateBody       (Note)    — body text (variables OK)
      - TemplateNotes      (Note)    — internal note about when to use
      - TemplateSortOrder  (Number)  — display order in the picker

    Idempotent — safe to re-run.

    IMPORTANT: For the Compose Email feature to work, an Azure AD tenant
    admin must also grant the 'Mail.Send' delegated permission on the
    CAHP Compliance Hub app registration. This script doesn't touch
    Azure AD; it only sets up the SharePoint list. After consent is
    granted, every signed-in user gets a token with Mail.Send silently.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-email-templates.ps1 `
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
Write-Host "  Email Templates list provisioning" -ForegroundColor Cyan
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

$ListTitle = "Email Templates"
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "-> Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/EmailTemplates" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Subject + body templates for the in-app Compose Email modal. Variables like {{contact}}, {{property}}, {{user}} get substituted at send time." `
        -EnableVersioning $true | Out-Null
    Write-Host "  List created" -ForegroundColor Green
} else {
    Write-Host "-> List already exists; verifying columns" -ForegroundColor Yellow
}

$columns = @(
    @{ Display = "Subject";    Internal = "TemplateSubject";   Type = "Text";   InView = $true }
    @{ Display = "Body";       Internal = "TemplateBody";      Type = "Note";   InView = $false }
    @{ Display = "Notes";      Internal = "TemplateNotes";     Type = "Note";   InView = $false }
    @{ Display = "Sort Order"; Internal = "TemplateSortOrder"; Type = "Number"; InView = $true }
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
        if ($col.InView) { $params.AddToDefaultView = $true }
        Add-PnPField @params -ErrorAction Stop | Out-Null
        Write-Host "  + $($col.Display) [$($col.Type)]" -ForegroundColor Cyan
    } catch {
        Write-Host "  ! Failed: $($col.Display) -- $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. In the Azure portal, grant the Mail.Send delegated permission on" -ForegroundColor Yellow
Write-Host "     the CAHP Compliance Hub app registration, then click 'Grant" -ForegroundColor Yellow
Write-Host "     admin consent'. Without that step, the Compose Email modal" -ForegroundColor Yellow
Write-Host "     will fail with a 403 on send." -ForegroundColor Yellow
Write-Host "  2. Hard-refresh the app and sign out / sign back in so the new" -ForegroundColor Yellow
Write-Host "     token grants Mail.Send silently." -ForegroundColor Yellow
Write-Host "  3. Go to Settings -> Email Templates to add your templates." -ForegroundColor Yellow
Write-Host "  4. From the Contacts page, click 'Compose Email' or the 'Email'" -ForegroundColor Yellow
Write-Host "     button next to any contact with an email address." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
