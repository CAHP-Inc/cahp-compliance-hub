<#
.SYNOPSIS
    Provisions the Property Notes SharePoint list for PR-08d.

.DESCRIPTION
    Creates the "Property Notes" list with two custom columns:
      - NoteBody (multi-line plain text) — the note content
      - Property (lookup → Properties Registry) — the property this note is about

    Idempotent — re-running is safe; existing list and columns are detected and skipped.

    REQUIRES: the CAHP Provisioning Shell Azure AD app (separate from the SPA app).
    No toggle dance needed — the Provisioning Shell app has public client flows on permanently.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app,
    NOT the SPA app. (You set this up in PR-07.)

.EXAMPLE
    .\provision-property-notes.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"

.NOTES
    Authored by Brandy Turner / NewShire Property Management
    Part of: CAHP Compliance Hub — Phase 1, PR-08d
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
Write-Host "  PR-08d: Property Notes list provisioning" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}
Write-Host "  ✓ Connected" -ForegroundColor Green
Write-Host ""

$ListTitle = "Property Notes"

# Create the list if it doesn't exist
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "→ Creating list '$ListTitle'..." -ForegroundColor White
    New-PnPList -Title $ListTitle -Template GenericList -Url "lists/PropertyNotes" | Out-Null
    Set-PnPList -Identity $ListTitle `
        -Description "Per-property notes log. Each row is a note attached to a property, with auto-captured author and timestamp." `
        -EnableVersioning $true | Out-Null
    Write-Host "  ✓ List created" -ForegroundColor Green
} else {
    Write-Host "→ List '$ListTitle' already exists; verifying columns" -ForegroundColor Yellow
}

# NoteBody column (multi-line plain text)
$noteBody = Get-PnPField -List $ListTitle -Identity "NoteBody" -ErrorAction SilentlyContinue
if (-not $noteBody) {
    Add-PnPField -List $ListTitle -DisplayName "Note Body" -InternalName "NoteBody" -Type Note -AddToDefaultView | Out-Null
    Write-Host "  + NoteBody [Note]" -ForegroundColor Cyan
} else {
    Write-Host "  → NoteBody exists, skipping" -ForegroundColor DarkGray
}

# Property lookup column (FieldXml is the reliable way to set the lookup target)
$propertyField = Get-PnPField -List $ListTitle -Identity "Property" -ErrorAction SilentlyContinue
if (-not $propertyField) {
    $propertyList = Get-PnPList -Identity "Properties Registry" -ErrorAction Stop
    if (-not $propertyList) {
        Write-Error "Properties Registry list not found — cannot create lookup field."
        exit 1
    }
    $listId = $propertyList.Id.ToString()
    $fieldXml = @"
<Field
  Type='Lookup'
  DisplayName='Property'
  Name='Property'
  List='{$listId}'
  ShowField='Title'
  Required='FALSE'
  EnforceUniqueValues='FALSE'
  Indexed='TRUE'
/>
"@
    Add-PnPFieldFromXml -List $ListTitle -FieldXml $fieldXml | Out-Null

    # Add to default view manually since FieldXml doesn't support AddToDefaultView
    $view = Get-PnPView -List $ListTitle -Identity "All Items"
    $viewFields = $view.ViewFields
    if ($viewFields -notcontains "Property") {
        Set-PnPView -List $ListTitle -Identity "All Items" -Fields ($viewFields + "Property") | Out-Null
    }

    Write-Host "  + Property [Lookup → Properties Registry]" -ForegroundColor Cyan
} else {
    Write-Host "  → Property exists, skipping" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Disconnect-PnPOnline
