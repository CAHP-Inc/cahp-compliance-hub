<#
.SYNOPSIS
    Provisions Communication Property Links + Communication Owner Links
    junction lists so a single Owner Communication can span multiple
    properties and owners. Migrates existing single-value links into
    the new junctions on first run.

.DESCRIPTION
    Creates two new SharePoint lists:
      - Communication Property Links (CommLookup + PropertyLookup)
      - Communication Owner Links   (CommLookup + OwnerLookup)

    Then walks every existing Owner Communications row and writes a junction
    row for its current CommPropertyLookupId / CommOwnerLookupId if one
    doesn't already exist.

    Idempotent — safe to re-run.

.PARAMETER SiteUrl
    Full URL of the CAHP Compliance Hub SharePoint site.

.PARAMETER ClientId
    Azure AD App Registration Client ID — use the CAHP Provisioning Shell app.

.EXAMPLE
    .\provision-comm-links.ps1 `
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
Write-Host "  Communication Property Links + Communication Owner Links" -ForegroundColor Cyan
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

# Pre-flight: make sure the parent lists exist
$commsList = Get-PnPList -Identity "Owner Communications" -ErrorAction SilentlyContinue
if (-not $commsList) {
    Write-Error "'Owner Communications' list not found. Aborting."
    exit 1
}
$propsList = Get-PnPList -Identity "Properties Registry" -ErrorAction SilentlyContinue
if (-not $propsList) {
    Write-Error "'Properties Registry' list not found. Aborting."
    exit 1
}
$ownersList = Get-PnPList -Identity "Owners" -ErrorAction SilentlyContinue
if (-not $ownersList) {
    Write-Error "'Owners' list not found. Aborting (run provision-owners.ps1 first)."
    exit 1
}

$commsListId  = $commsList.Id.ToString()
$propsListId  = $propsList.Id.ToString()
$ownersListId = $ownersList.Id.ToString()

# =============================================================================
# Part 1: Communication Property Links
# =============================================================================

$PropLinksTitle = "Communication Property Links"
$propLinks = Get-PnPList -Identity $PropLinksTitle -ErrorAction SilentlyContinue
if (-not $propLinks) {
    Write-Host "-> Creating list '$PropLinksTitle'..." -ForegroundColor White
    New-PnPList -Title $PropLinksTitle -Template GenericList -Url "lists/CommunicationPropertyLinks" | Out-Null
    Set-PnPList -Identity $PropLinksTitle `
        -Description "Junction list. One row per (Owner Communication, Property) linkage. Lets a single comm log span multiple properties." `
        -EnableVersioning $true | Out-Null
    Write-Host "  Junction list created" -ForegroundColor Green
} else {
    Write-Host "-> List '$PropLinksTitle' already exists; verifying columns" -ForegroundColor Yellow
}

if (-not (Get-PnPField -List $PropLinksTitle -Identity "Comm" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Comm' Name='Comm' List='{$commsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $PropLinksTitle -FieldXml $xml | Out-Null
    Write-Host "  + Comm [Lookup -> Owner Communications]" -ForegroundColor Cyan
}
if (-not (Get-PnPField -List $PropLinksTitle -Identity "Property" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Property' Name='Property' List='{$propsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $PropLinksTitle -FieldXml $xml | Out-Null
    Write-Host "  + Property [Lookup -> Properties Registry]" -ForegroundColor Cyan
}

# Add Comm + Property to default view
$propLinksView = Get-PnPView -List $PropLinksTitle -Identity "All Items" -ErrorAction SilentlyContinue
if ($propLinksView) {
    $viewFields = $propLinksView.ViewFields
    $newFields = $viewFields
    if ($newFields -notcontains "Comm") { $newFields = $newFields + "Comm" }
    if ($newFields -notcontains "Property") { $newFields = $newFields + "Property" }
    if ($newFields.Count -ne $viewFields.Count) {
        Set-PnPView -List $PropLinksTitle -Identity "All Items" -Fields $newFields | Out-Null
    }
}

# =============================================================================
# Part 2: Communication Owner Links
# =============================================================================

$OwnerLinksTitle = "Communication Owner Links"
$ownerLinksList = Get-PnPList -Identity $OwnerLinksTitle -ErrorAction SilentlyContinue
if (-not $ownerLinksList) {
    Write-Host "-> Creating list '$OwnerLinksTitle'..." -ForegroundColor White
    New-PnPList -Title $OwnerLinksTitle -Template GenericList -Url "lists/CommunicationOwnerLinks" | Out-Null
    Set-PnPList -Identity $OwnerLinksTitle `
        -Description "Junction list. One row per (Owner Communication, Owner) linkage. Lets a comm log involve multiple owner entities." `
        -EnableVersioning $true | Out-Null
    Write-Host "  Junction list created" -ForegroundColor Green
} else {
    Write-Host "-> List '$OwnerLinksTitle' already exists; verifying columns" -ForegroundColor Yellow
}

if (-not (Get-PnPField -List $OwnerLinksTitle -Identity "Comm" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Comm' Name='Comm' List='{$commsListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $OwnerLinksTitle -FieldXml $xml | Out-Null
    Write-Host "  + Comm [Lookup -> Owner Communications]" -ForegroundColor Cyan
}
if (-not (Get-PnPField -List $OwnerLinksTitle -Identity "Owner" -ErrorAction SilentlyContinue)) {
    $xml = @"
<Field Type='Lookup' DisplayName='Owner' Name='Owner' List='{$ownersListId}' ShowField='Title' Required='FALSE' Indexed='TRUE' />
"@
    Add-PnPFieldFromXml -List $OwnerLinksTitle -FieldXml $xml | Out-Null
    Write-Host "  + Owner [Lookup -> Owners]" -ForegroundColor Cyan
}

$ownerLinksView = Get-PnPView -List $OwnerLinksTitle -Identity "All Items" -ErrorAction SilentlyContinue
if ($ownerLinksView) {
    $viewFields = $ownerLinksView.ViewFields
    $newFields = $viewFields
    if ($newFields -notcontains "Comm") { $newFields = $newFields + "Comm" }
    if ($newFields -notcontains "Owner") { $newFields = $newFields + "Owner" }
    if ($newFields.Count -ne $viewFields.Count) {
        Set-PnPView -List $OwnerLinksTitle -Identity "All Items" -Fields $newFields | Out-Null
    }
}

# =============================================================================
# Part 3: Migrate existing single-lookup values into the junctions
# =============================================================================

Write-Host ""
Write-Host "-> Migrating existing CommPropertyLookupId / CommOwnerLookupId values..." -ForegroundColor White

$existingComms = Get-PnPListItem -List "Owner Communications" -PageSize 500 -Fields "ID","CommProperty","CommOwner" -ErrorAction SilentlyContinue
$existingPropLinks = Get-PnPListItem -List $PropLinksTitle -PageSize 500 -Fields "ID","Comm","Property" -ErrorAction SilentlyContinue
$existingOwnerLinks = Get-PnPListItem -List $OwnerLinksTitle -PageSize 500 -Fields "ID","Comm","Owner" -ErrorAction SilentlyContinue

$migratedProps = 0
$migratedOwners = 0
$skipped = 0

foreach ($c in ($existingComms ?? @())) {
    $commId = $c.Id
    $propRef = $c["CommProperty"]
    $ownerRef = $c["CommOwner"]

    if ($propRef -and $propRef.LookupId) {
        $pId = $propRef.LookupId
        $alreadyLinked = $false
        foreach ($row in ($existingPropLinks ?? @())) {
            $rowComm = $row["Comm"]
            $rowProp = $row["Property"]
            if ($rowComm -and $rowProp -and $rowComm.LookupId -eq $commId -and $rowProp.LookupId -eq $pId) {
                $alreadyLinked = $true
                break
            }
        }
        if ($alreadyLinked) {
            $skipped++
        } else {
            Add-PnPListItem -List $PropLinksTitle -Values @{
                Title = "Comm $commId <-> Property $pId"
                Comm = $commId
                Property = $pId
            } | Out-Null
            $migratedProps++
        }
    }

    if ($ownerRef -and $ownerRef.LookupId) {
        $oId = $ownerRef.LookupId
        $alreadyLinked = $false
        foreach ($row in ($existingOwnerLinks ?? @())) {
            $rowComm = $row["Comm"]
            $rowOwner = $row["Owner"]
            if ($rowComm -and $rowOwner -and $rowComm.LookupId -eq $commId -and $rowOwner.LookupId -eq $oId) {
                $alreadyLinked = $true
                break
            }
        }
        if ($alreadyLinked) {
            $skipped++
        } else {
            Add-PnPListItem -List $OwnerLinksTitle -Values @{
                Title = "Comm $commId <-> Owner $oId"
                Comm = $commId
                Owner = $oId
            } | Out-Null
            $migratedOwners++
        }
    }
}

Write-Host "  Migration: $migratedProps property link(s), $migratedOwners owner link(s) created; $skipped already linked" -ForegroundColor Green

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT:" -ForegroundColor Yellow
Write-Host "  1. Hard-refresh the app (Ctrl+Shift+R)." -ForegroundColor Yellow
Write-Host "  2. When you log a communication, you can now check off every" -ForegroundColor Yellow
Write-Host "     property and owner entity it touches." -ForegroundColor Yellow
Write-Host "  3. Open an existing comm and click Edit -- the multi-select" -ForegroundColor Yellow
Write-Host "     pickers show its current links and let you add or remove more." -ForegroundColor Yellow
Write-Host ""

Disconnect-PnPOnline
