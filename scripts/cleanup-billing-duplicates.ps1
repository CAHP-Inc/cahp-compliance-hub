<#
.SYNOPSIS
    Removes duplicate / stale "% of Savings" rows from the Billing Tracker, leaving
    ONE row per parcel (Tax Map ID) + tax year — matching what the Billing page shows.
    Deletes to the Recycle Bin (recoverable).

.DESCRIPTION
    Mirrors the app's dedupe:
      1. Group % rows (non-N/A, with a savings/amount) by Property + resolved TMID + year.
         Resolved TMID = the row's BillTaxMapID, else the linked submittal's TaxMapID.
         Keep the most complete (entered tax bills > explicit TMID > newest); delete the rest.
      2. Drop a "whole property" (no-TMID) row when a real parceled row exists for the
         same property + year.

    DRY RUN by default — shows what it would delete. Add -Apply to actually delete.

.EXAMPLE
    .\cleanup-billing-duplicates.ps1 -SiteUrl "https://newshirepmcom.sharepoint.com/sites/CAHPComplianceHub" -ClientId "7f310acf-12b1-4ba9-a113-c027614268b9"
    .\cleanup-billing-duplicates.ps1 -SiteUrl "..." -ClientId "..." -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [switch]$Apply
)
$ErrorActionPreference = "Stop"

if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
    $cleanPnP = "C:\Users\brand\PnPModules\PnP.PowerShell\3.2.0\PnP.PowerShell.psd1"
    if (Test-Path $cleanPnP) { Import-Module $cleanPnP -Force }
    elseif (Get-Module -ListAvailable -Name PnP.PowerShell) { Import-Module PnP.PowerShell }
    else { Write-Error "PnP.PowerShell not available. Dot-source Start-CleanSession.ps1 first."; exit 1 }
}

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
Write-Host "Connected." -ForegroundColor Green

# submittal id -> its TaxMapID (for backfilling rows that lack an explicit TMID)
$submTmid = @{}
foreach ($s in Get-PnPListItem -List "Submittals Tracker" -PageSize 2000 -Fields "ID", "TaxMapID") {
    $t = $s.FieldValues["TaxMapID"]
    if ($t) { $submTmid[[string]$s.Id] = [string]$t.LookupId }
}

$items = Get-PnPListItem -List "Billing Tracker" -PageSize 2000 -Fields `
    "ID", "Title", "Property", "BillTaxMapID", "BillSubmittal", "cahpTaxYear", "BillingType", "BillApprovedAbatement", "AmountBilled", "LastFullTaxBill", "BillingStatus"

$rows = New-Object System.Collections.Generic.List[object]
foreach ($it in $items) {
    $fv = $it.FieldValues
    $type = [string]$fv["BillingType"]
    if ($type -ne "" -and $type -ne "Percent of Savings") { continue }       # % rows (or untyped legacy) only
    if ([string]$fv["BillingStatus"] -eq "N/A") { continue }
    $sav = [double]($fv["BillApprovedAbatement"] ?? 0); $amt = [double]($fv["AmountBilled"] ?? 0)
    if (-not ($sav -gt 0 -or $amt -gt 0)) { continue }
    $pid = $(if ($fv["Property"]) { [string]$fv["Property"].LookupId } else { "" })
    $explicit = $(if ($fv["BillTaxMapID"]) { [string]$fv["BillTaxMapID"].LookupId } else { "" })
    $submId = $(if ($fv["BillSubmittal"]) { [string]$fv["BillSubmittal"].LookupId } else { "" })
    $tmid = $(if ($explicit) { $explicit } elseif ($submId -and $submTmid.ContainsKey($submId)) { $submTmid[$submId] } else { "" })
    $score = $(if ($fv["LastFullTaxBill"] -ne $null) { 4 } else { 0 }) + $(if ($explicit) { 2 } else { 0 })
    $rows.Add([pscustomobject]@{ Id = [int]$it.Id; Title = [string]$fv["Title"]; pid = $pid; tmid = $tmid; year = [string]$fv["cahpTaxYear"]; score = $score })
}

# 1) keep best per (pid | tmid | year)
$best = @{}
foreach ($r in $rows) {
    $k = "$($r.pid)|$($r.tmid)|$($r.year)"
    if (-not $best.ContainsKey($k)) { $best[$k] = $r }
    else { $c = $best[$k]; if ($r.score -gt $c.score -or ($r.score -eq $c.score -and $r.Id -gt $c.Id)) { $best[$k] = $r } }
}
# 2) among survivors, drop whole-property rows where a parceled row exists for that pid|year
$parceledPY = @{}
foreach ($r in $best.Values) { if ($r.tmid -ne "") { $parceledPY["$($r.pid)|$($r.year)"] = $true } }
$keepIds = @($best.Values | Where-Object { $_.tmid -ne "" -or -not $parceledPY.ContainsKey("$($_.pid)|$($_.year)") } | ForEach-Object { $_.Id })
$toDelete = $rows | Where-Object { $_.Id -notin $keepIds }

Write-Host ("`n% rows: {0}   keep: {1}   delete: {2}`n" -f $rows.Count, $keepIds.Count, @($toDelete).Count) -ForegroundColor Yellow
$toDelete | Sort-Object pid, year | Format-Table Id, Title, @{n='TMID';e={$_.tmid}}, year, score -AutoSize

if (-not $Apply) {
    Write-Host "DRY RUN — nothing deleted. Re-run with -Apply to delete the rows above (to the Recycle Bin)." -ForegroundColor Cyan
} else {
    foreach ($d in $toDelete) { Remove-PnPListItem -List "Billing Tracker" -Identity $d.Id -Recycle -Force | Out-Null }
    Write-Host ("Deleted {0} duplicate/stale rows (recoverable in the Recycle Bin)." -f @($toDelete).Count) -ForegroundColor Green
}
Disconnect-PnPOnline
