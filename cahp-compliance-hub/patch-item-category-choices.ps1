# =============================================================================
# Patch: ItemCategory choices update — fix PowerShell type coercion
# =============================================================================
# The previous script logged "updated 15 options" but the underlying call
# failed silently on type coercion. This patch uses an explicit [string[]]
# cast so PnP accepts the array.
# =============================================================================

$ErrorActionPreference = 'Stop'
$siteUrl      = 'https://vanrockre.sharepoint.com/sites/CAHPComplianceHub'
$provShellApp = '63567714-59eb-4d4f-b3f0-f827e58d9a59'
$listName     = 'Outstanding Items Checklist'

Write-Host '================================================================' -ForegroundColor Cyan
Write-Host '  Patch: ItemCategory choices (typed array)' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan

Connect-PnPOnline -Url $siteUrl -ClientId $provShellApp -Interactive
Write-Host '  Connected' -ForegroundColor Green

# Explicit [string[]] cast — fixes "Object[] cannot be converted to String[]"
[string[]]$NEW_CATEGORIES = @(
  'Operating Agreement',
  'Articles of Incorporation',
  'EIN Confirmation',
  'Certificate of Existence',
  'Certificate of Authorization',
  '501(c)(3) Determination',
  'Deed',
  'Rent Roll',
  'LURA',
  'AMI Certification',
  'Org Chart',
  'Income Documentation',
  'Signed Submittal',
  'Determination Letter',
  'Other'
)

$catField = Get-PnPField -List $listName -Identity 'ItemCategory' -ErrorAction SilentlyContinue
if (-not $catField) {
    Write-Host '  ! ItemCategory field not found.' -ForegroundColor Red
    exit 1
}

# Method 1: Set-PnPField with explicitly typed array
try {
    Set-PnPField -List $listName -Identity 'ItemCategory' -Values @{ Choices = $NEW_CATEGORIES } -ErrorAction Stop | Out-Null
    Write-Host "  ItemCategory choices updated ($($NEW_CATEGORIES.Count) options)" -ForegroundColor Green
} catch {
    Write-Host "  Set-PnPField failed: $_" -ForegroundColor Yellow
    Write-Host "  Falling back to CSOM direct update..." -ForegroundColor Yellow

    # Method 2: CSOM fallback — set property directly on the field
    $ctx = Get-PnPContext
    $list = Get-PnPList -Identity $listName
    $field = Get-PnPField -List $listName -Identity 'ItemCategory'
    $choiceField = [Microsoft.SharePoint.Client.FieldChoice]$field.TypedObject
    $choiceField.Choices = $NEW_CATEGORIES
    $choiceField.Update()
    $ctx.ExecuteQuery()
    Write-Host "  ItemCategory choices updated via CSOM ($($NEW_CATEGORIES.Count) options)" -ForegroundColor Green
}

# Verify
$verify = Get-PnPField -List $listName -Identity 'ItemCategory'
Write-Host ''
Write-Host '  Current choices on the field:' -ForegroundColor Cyan
$verify.Choices | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }

Write-Host ''
Write-Host '================================================================' -ForegroundColor Green
Write-Host '  Patch complete.' -ForegroundColor Green
Write-Host '================================================================' -ForegroundColor Green
