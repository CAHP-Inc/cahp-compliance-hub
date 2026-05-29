<#
.SYNOPSIS
    Restore document Owner / Property tags wiped by the multi-value column
    re-creation, using the upload audit-log entries as the source of truth.

.DESCRIPTION
    Every time the app uploaded a document, it wrote a CREATE entry to the
    AuditLog list that includes the original metadata in AfterJSON
    (PropertyLookupId, OwnerLookupId, Title, etc.). When provision-document-
    library-columns.ps1 dropped + recreated the Owner / Property columns as
    multi-value, the live tag values were lost — but the audit history wasn't.

    This script reads every CREATE entry for the 9 document libraries, parses
    the original tags out of AfterJSON, finds the current SharePoint listItem
    by ID, and re-applies the tag if the column is currently empty. Existing
    tags are NEVER overwritten — so files you re-tagged manually in the last
    couple of days stay as-is.

    Defaults to DRY RUN. Pass -Execute to actually write the restored tags.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID.

.PARAMETER Execute
    Pass to actually apply the restored tags. Without it, runs as dry-run.

.EXAMPLE
    .\recover-tags-from-auditlog.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59"

.EXAMPLE
    .\recover-tags-from-auditlog.ps1 `
      -SiteUrl "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub" `
      -ClientId "63567714-59eb-4d4f-b3f0-f827e58d9a59" -Execute
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

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Recover document tags from AuditLog" -ForegroundColor Cyan
Write-Host "  Mode: $(if ($Execute) { 'EXECUTE (will write tags)' } else { 'DRY RUN (no changes)' })" -ForegroundColor $(if ($Execute) { 'Yellow' } else { 'Green' })
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Try to reuse an existing PnP connection; otherwise open one.
# Get-PnPConnection throws (not just errors) when there's no session, so wrap it.
$existing = $null
try {
    $existing = Get-PnPConnection -ErrorAction Stop
} catch {
    $existing = $null
}
if (-not $existing) {
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    } catch {
        Write-Error "Connection failed: $_"
        exit 1
    }
    Write-Host "  Connected" -ForegroundColor Green
} else {
    Write-Host "  Reusing existing PnP connection" -ForegroundColor Green
}
Write-Host ""

# Libraries we care about (anything else in AuditLog with these EntityType
# values gets ignored). Mirrors PROPERTY_LINKED_LIBRARIES + CAHP Entity Documents.
$targetLibraries = @(
    "AMI Certification Letters",
    "DOR Correspondence",
    "DOR Submittal Packages",
    "Land Use Restriction Agreements",
    "Operating Agreements",
    "Org Charts",
    "Property Deeds",
    "Supporting Documentation",
    "CAHP Entity Documents"
)
$libSet = @{}
foreach ($l in $targetLibraries) { $libSet[$l] = $true }

# Pull AuditLog entries — CREATE only, the relevant libraries only
Write-Host "-> Loading AuditLog..." -ForegroundColor White
$audit = Get-PnPListItem -List "AuditLog" -PageSize 1000 `
    -Fields "ID","Title","Action","EntityType","EntityId","EntityTitle","AfterJSON","Created"
$creates = $audit | Where-Object {
    [string]$_["Action"] -eq "CREATE" -and $libSet.ContainsKey([string]$_["EntityType"])
}
Write-Host "   $($audit.Count) total entries, $($creates.Count) CREATE rows for document libraries" -ForegroundColor Green
Write-Host ""

# Group by library so we batch-fetch the current listItems
$creatsByLib = $creates | Group-Object { [string]$_["EntityType"] }

$totalScanned = 0
$totalCandidates = 0
$totalAlreadyTagged = 0
$totalNotFound = 0
$totalNoMetadata = 0
$totalRestored = 0
$totalFailed = 0
$totalDuplicates = 0

foreach ($group in $creatsByLib) {
    $lib = $group.Name
    Write-Host "===== $lib =====" -ForegroundColor Cyan

    # Pre-fetch current items by id
    $currentItems = @{}
    try {
        Get-PnPListItem -List $lib -PageSize 500 -Fields "ID","FileLeafRef","Owner","Property" `
            | ForEach-Object { $currentItems[[string]$_.Id] = $_ }
    } catch {
        Write-Host "   ! Could not load library: $_" -ForegroundColor Red
        continue
    }

    # Track which itemIds we've already processed so we don't re-apply a tag
    # from an OLDER audit entry that has been superseded by a newer one.
    $seenItemIds = @{}

    # Sort newest first so the latest CREATE wins on duplicates
    $sorted = $group.Group | Sort-Object { [datetime]$_["Created"] } -Descending

    foreach ($entry in $sorted) {
        $totalScanned++

        $itemId = [string]$entry["EntityId"]
        $entryTitle = [string]$entry["EntityTitle"]
        $afterJson = [string]$entry["AfterJSON"]

        if (-not $itemId) { continue }
        if ($seenItemIds.ContainsKey($itemId)) {
            $totalDuplicates++
            continue
        }
        $seenItemIds[$itemId] = $true

        # Parse the original metadata out of AfterJSON
        if (-not $afterJson) { $totalNoMetadata++; continue }
        try { $after = $afterJson | ConvertFrom-Json } catch { $totalNoMetadata++; continue }

        $origPropId = $null
        $origOwnerId = $null
        if ($after.PSObject.Properties.Name -contains "PropertyLookupId") {
            $origPropId = [string]$after.PropertyLookupId
        }
        if ($after.PSObject.Properties.Name -contains "OwnerLookupId") {
            $origOwnerId = [string]$after.OwnerLookupId
        }

        if (-not $origPropId -and -not $origOwnerId) {
            $totalNoMetadata++
            continue
        }

        $totalCandidates++

        # Does this file still exist?
        $current = $currentItems[$itemId]
        if (-not $current) {
            Write-Host "  - $entryTitle [id $itemId]: listItem no longer exists; skipping" -ForegroundColor DarkGray
            $totalNotFound++
            continue
        }

        $hasProperty = $current["Property"] -and ($current["Property"].Count -gt 0)
        $hasOwner    = $current["Owner"]    -and ($current["Owner"].Count -gt 0)
        $patch = @{}

        if ($origPropId -and -not $hasProperty) {
            $patch["Property"] = @([int]$origPropId)
        }
        if ($origOwnerId -and -not $hasOwner) {
            $patch["Owner"] = @([int]$origOwnerId)
        }

        if ($patch.Count -eq 0) {
            $totalAlreadyTagged++
            continue
        }

        $tagDesc = @()
        if ($patch.ContainsKey("Property")) { $tagDesc += "Property=$origPropId" }
        if ($patch.ContainsKey("Owner"))    { $tagDesc += "Owner=$origOwnerId" }
        $tagDescStr = $tagDesc -join ", "

        $currentFilename = [string]$current["FileLeafRef"]
        Write-Host "  + $currentFilename [id $itemId]: restore $tagDescStr" -ForegroundColor Green

        if ($Execute) {
            try {
                Set-PnPListItem -List $lib -Identity $itemId -Values $patch | Out-Null
                $totalRestored++
            } catch {
                Write-Host "      ! failed: $_" -ForegroundColor Red
                $totalFailed++
            }
        } else {
            $totalRestored++  # count as 'would restore' in dry-run summary
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Audit CREATE rows scanned:    $totalScanned" -ForegroundColor White
Write-Host "  Duplicates skipped:           $totalDuplicates (older entries for same itemId)" -ForegroundColor White
Write-Host "  Rows with restorable tag:     $totalCandidates" -ForegroundColor White
Write-Host "  Already tagged (left alone):  $totalAlreadyTagged" -ForegroundColor White
Write-Host "  Item missing in SharePoint:   $totalNotFound" -ForegroundColor White
Write-Host "  Rows with no usable metadata: $totalNoMetadata" -ForegroundColor White
if ($Execute) {
    Write-Host "  Tags restored:                $totalRestored" -ForegroundColor Green
    Write-Host "  Failed restores:              $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'White' })
} else {
    Write-Host "  Tags that WOULD be restored:  $totalRestored" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "DRY RUN — no changes were made." -ForegroundColor Green
    Write-Host "Re-run with -Execute to apply." -ForegroundColor Yellow
}
Write-Host ""
