<#
.SYNOPSIS
    Remove Outstanding Items whose Titles match the OLD (pre-SC-PT-401-O)
    checklist template, leaving adhoc / manually-created items untouched.

.DESCRIPTION
    The old hardcoded DOR_FILING_CHECKLIST had 12 specific titles that the
    Property wizard and Filing Checklist Generator used verbatim when
    creating Outstanding Items. The new SC PT-401-O template has different
    titles (e.g., 'Recorded Property Deed(s)' instead of 'Property Deed(s)'),
    so an exact-title match against the old set never picks up an item
    created from the new template — making this safe to run even after
    you've started using the new templates.

    Defaults to DRY RUN. The script prints every match it would delete so
    you can verify before running for real. Pass -Execute to actually
    delete the matched rows.

    Idempotent — re-running after a successful execute is a no-op.

    Adhoc items (titles outside the old template set) are NEVER touched.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.PARAMETER Execute
    Pass this switch to actually delete the matched rows. Without it the
    script runs in dry-run mode and only prints what it would do.

.EXAMPLE
    # Dry run first — see what would be deleted
    .\cleanup-old-template-items.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"

.EXAMPLE
    # Actually delete after reviewing the dry-run output
    .\cleanup-old-template-items.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59" `
      -Execute

.NOTES
    After running with -Execute, repopulate each property's Outstanding
    Items by opening the property in the app and clicking 'Generate Filing
    Checklist' — that picks up the new SC PT-401-O template list.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [switch]$Execute
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

# Exact titles from the OLD hardcoded DOR_FILING_CHECKLIST (commit history:
# src/lib/filing-checklist.ts before 8bd8846). These are the only items the
# script touches. Any other Title is treated as adhoc and left alone.
$oldTemplateTitles = @(
    "CAHP Operating Agreement (Non Profit OA)"
    "CAHP 501(c)(3) Determination Letter"
    "CAHP EIN Confirmation"
    "CAHP Articles of Incorporation"
    "CAHP Certificate of Existence (COE)"
    "Entity Certification Letter (Cert of Authorization)"
    "Entity EIN Confirmation"
    "Entity Operating Agreement"
    "Entity Articles of Organization"
    "Property Deed(s)"
    "Rent Roll (current year)"
    "IRS Determination Letter (property-specific, if applicable)"
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Cleanup OLD-template Outstanding Items" -ForegroundColor Cyan
Write-Host "  Mode:   $(if ($Execute) { 'EXECUTE (will delete)' } else { 'DRY RUN (no changes)' })" -ForegroundColor $(if ($Execute) { 'Yellow' } else { 'Green' })
Write-Host "  Site:   $SiteUrl" -ForegroundColor Cyan
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

# Verify the Outstanding list exists. The app uses the title
# "Outstanding Items Checklist" (see LIST_NAMES.Outstanding in types.ts).
$listTitle = "Outstanding Items Checklist"
if (-not (Get-PnPList -Identity $listTitle -ErrorAction SilentlyContinue)) {
    Write-Error "'$listTitle' list not found at this site. Aborting."
    exit 1
}

Write-Host "-> Pre-fetching all rows from '$listTitle'..." -ForegroundColor White
$allItems = Get-PnPListItem -List $listTitle -PageSize 500 -Fields "ID","Title","ItemStatus","PropertyLookupId"
Write-Host "   Fetched $($allItems.Count) row(s)" -ForegroundColor Green
Write-Host ""

# Bucket items by whether Title matches an old template title
$titleSet = @{}
foreach ($t in $oldTemplateTitles) { $titleSet[$t] = $true }

$matches = @()
$adhoc = 0
foreach ($it in $allItems) {
    $title = [string]$it["Title"]
    if ($titleSet.ContainsKey($title)) {
        $matches += $it
    } else {
        $adhoc++
    }
}

Write-Host "  Matches old template titles: $($matches.Count)" -ForegroundColor Yellow
Write-Host "  Adhoc / new-template items (untouched): $adhoc" -ForegroundColor Green
Write-Host ""

if ($matches.Count -eq 0) {
    Write-Host "Nothing to clean up. Exiting." -ForegroundColor Green
    Disconnect-PnPOnline
    exit 0
}

# Show a sample of what would be deleted (first 20)
Write-Host "Sample of items that match (showing up to 20):" -ForegroundColor White
$sampleSize = [Math]::Min(20, $matches.Count)
for ($i = 0; $i -lt $sampleSize; $i++) {
    $m = $matches[$i]
    $title  = [string]$m["Title"]
    $status = [string]$m["ItemStatus"]
    $pid    = [string]$m["PropertyLookupId"]
    Write-Host ("  [{0}] {1}  (status: {2}, propertyId: {3})" -f $m.Id, $title, $status, $pid) -ForegroundColor DarkGray
}
if ($matches.Count -gt 20) {
    Write-Host "  ... and $($matches.Count - 20) more" -ForegroundColor DarkGray
}
Write-Host ""

if (-not $Execute) {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  DRY RUN complete. Nothing was changed." -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To delete these $($matches.Count) row(s), re-run with -Execute:" -ForegroundColor Yellow
    Write-Host "  .\cleanup-old-template-items.ps1 -SiteUrl '$SiteUrl' -ClientId '$ClientId' -Execute" -ForegroundColor Yellow
    Write-Host ""
    Disconnect-PnPOnline
    exit 0
}

# Execute path — delete matched rows
Write-Host "-> Deleting $($matches.Count) row(s)..." -ForegroundColor Yellow
$deleted = 0
$failed  = 0
foreach ($m in $matches) {
    try {
        Remove-PnPListItem -List $listTitle -Identity $m.Id -Force | Out-Null
        $deleted++
        if (($deleted % 25) -eq 0) {
            Write-Host "   ... $deleted deleted" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "   ! Failed to delete row $($m.Id): $_" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Deleted: $deleted" -ForegroundColor White
Write-Host "  Failed:  $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'White' })
Write-Host "  Adhoc items (untouched): $adhoc" -ForegroundColor Green
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app." -ForegroundColor Yellow
Write-Host "  2. For each property, open it and click 'Generate Filing" -ForegroundColor Yellow
Write-Host "     Checklist' — that creates new Outstanding Items from the" -ForegroundColor Yellow
Write-Host "     SC PT-401-O template list, with the new per-row checkboxes" -ForegroundColor Yellow
Write-Host "     so you can skip anything that doesn't apply." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
