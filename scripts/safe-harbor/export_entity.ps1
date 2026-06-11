<#
.SYNOPSIS
    Export one certifiable LLC's entity facts from the CAHP Compliance Hub
    (Properties Registry / Owners / Ownership Structure) into a Safe Harbor
    entity config JSON consumed by generate_cert.py.

.DESCRIPTION
    Pulls live data from the hub so the certification letter's boilerplate
    (entity name, EIN, county, DOR account, nonprofit managing member, the
    1% Class C interest, etc.) comes from the source of truth instead of
    being hand-typed.

    It ALSO validates the §12-37-220(B)(11)(e) exemption chain: the property's
    ownership must include a CAHP-flagged 501(c)(3) instrumentality (an Owner
    with IsCAHPEntity = true). If none is found the script writes the JSON but
    prints a loud WARNING — without that nonprofit member the entity does not
    qualify for the exemption and the letter should not be filed.

    Any field the hub does not have is left as "" in the JSON, which prints as
    an underscored blank in the letter for manual completion. Existing values in
    an -OutPath file are preserved unless -Overwrite is passed (so hand-entered
    fields like signatory phone/email or the operating-agreement date survive a
    re-export).

.PARAMETER EntityTitle
    Exact Title of the Property row to certify (e.g., "IV SPB II LLC").

.PARAMETER OutPath
    Where to write the entity JSON. Defaults to
    ..\..\safe-harbor-output\<slug>.entity.json (gitignored).

.EXAMPLE
    .\export_entity.ps1 -EntityTitle "IV SPB II LLC"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EntityTitle,
    [string]$SiteUrl  = "https://vanrockre.sharepoint.com/sites/CAHPComplianceHub",
    [string]$ClientId = "63567714-59eb-4d4f-b3f0-f827e58d9a59",
    [string]$OutPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell not installed. Run: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

function Slug([string]$s) { ($s -replace '[^A-Za-z0-9]+', '_').Trim('_') }

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutPath) {
    $repoRoot = Resolve-Path (Join-Path $scriptDir "..\..")
    $outDir = Join-Path $repoRoot "safe-harbor-output"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $OutPath = Join-Path $outDir ("{0}.entity.json" -f (Slug $EntityTitle))
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Export Safe Harbor entity config: $EntityTitle" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "  Out:  $OutPath" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}
Write-Host "  Connected" -ForegroundColor Green

# ---- Property row -------------------------------------------------------- #
$allProps = Get-PnPListItem -List "Properties Registry" -PageSize 500 `
    -Fields "ID","Title","PropertyAddress","LegalEntity","UnitCount","AMIProgram",
            "cahpCounty","cahpState","DORAccountID","PropertyEIN"
$prop = $allProps | Where-Object { [string]$_["Title"] -eq $EntityTitle } | Select-Object -First 1
if (-not $prop) {
    Write-Error "Property '$EntityTitle' not found in Properties Registry."
    exit 1
}
$propId = $prop.Id
Write-Host "  Found property (ID $propId)" -ForegroundColor Green

$counties = @()
if ($prop["cahpCounty"]) {
    $counties = ($prop["cahpCounty"] -split ',') | ForEach-Object {
        ($_ -replace '\(SC\)|\(NC\)', '').Trim()
    } | Where-Object { $_ }
}

# ---- Ownership rows for this property ------------------------------------ #
$ownership = Get-PnPListItem -List "Ownership Structure" -PageSize 500 `
    -Fields "ID","OwnerLookupId","ParentOwnerLookupId","LinkedPropertyLookupId",
            "RelationshipType","OwnershipPercent","MemberClass"
$mine = $ownership | Where-Object { [int]$_["LinkedPropertyLookupId"].LookupId -eq $propId }

# ---- Owners (indexed) ---------------------------------------------------- #
$owners = Get-PnPListItem -List "Owners" -PageSize 500 `
    -Fields "ID","Title","OwnerType","OwnerState","TaxID","IsCAHPEntity","IsTaxExempt","SponsorName"
$ownerById = @{}
foreach ($o in $owners) { $ownerById[[int]$o.Id] = $o }

# Locate the CAHP nonprofit instrumentality member (IsCAHPEntity = true).
$cahpMember = $null; $cahpRow = $null
foreach ($r in $mine) {
    $oid = if ($r["OwnerLookupId"]) { [int]$r["OwnerLookupId"].LookupId } else { $null }
    if ($oid -and $ownerById.ContainsKey($oid) -and $ownerById[$oid]["IsCAHPEntity"]) {
        $cahpMember = $ownerById[$oid]; $cahpRow = $r; break
    }
}

# CAHP parent corp = a CAHP-flagged Nonprofit owner (holds the 501(c)(3) EIN).
$cahpParent = $owners | Where-Object {
    $_["IsCAHPEntity"] -and ($_["OwnerType"] -eq "Nonprofit" -or $_["IsTaxExempt"])
} | Select-Object -First 1

# ---- Merge with existing file (preserve hand-entered fields) ------------- #
$existing = $null
if ((Test-Path $OutPath) -and (-not $Overwrite)) {
    try { $existing = Get-Content $OutPath -Raw | ConvertFrom-Json } catch { $existing = $null }
}
function Pick($new, $old, $default = "") {
    if ($new) { return $new }
    if ($old) { return $old }
    return $default
}
$e = $existing

$config = [ordered]@{
    company = [ordered]@{
        legalName    = $EntityTitle
        stateType    = Pick $null $e.company.stateType "South Carolina limited liability company"
        ein          = Pick ([string]$prop["PropertyEIN"]) $e.company.ein
        dorAccountId = Pick ([string]$prop["DORAccountID"]) $e.company.dorAccountId
    }
    property = [ordered]@{
        description    = Pick $null $e.property.description "scattered-site residential rental units"
        addressLine    = Pick ([string]$prop["PropertyAddress"]) $e.property.addressLine
        counties       = if ($counties.Count) { $counties } elseif ($e.property.counties) { $e.property.counties } else { @() }
        state          = Pick ([string]$prop["cahpState"]) $e.property.state "SC"
        taxMapParcels  = if ($e.property.taxMapParcels) { $e.property.taxMapParcels } else { @() }
    }
    nonprofit = [ordered]@{
        managingMemberName = Pick ($(if ($cahpMember) { [string]$cahpMember["Title"] })) $e.nonprofit.managingMemberName "CAHP SC, LLC"
        parentName         = Pick ($(if ($cahpParent) { [string]$cahpParent["Title"] })) $e.nonprofit.parentName "Carolina Affordable Housing Project Inc."
        parentEin          = Pick ($(if ($cahpParent) { [string]$cahpParent["TaxID"] })) $e.nonprofit.parentEin $(if ($State -eq 'SC') { '99-4885069' } else { '' })
        ownershipPercent   = if ($cahpRow -and $null -ne $cahpRow["OwnershipPercent"]) { [double]$cahpRow["OwnershipPercent"] } elseif ($null -ne $e.nonprofit.ownershipPercent) { $e.nonprofit.ownershipPercent } else { $null }
        memberClass        = Pick ($(if ($cahpRow) { [string]$cahpRow["MemberClass"] })) $e.nonprofit.memberClass "Class C"
        isTaxExempt        = $true
    }
    manager = [ordered]@{
        name                  = Pick $null $e.manager.name "VanRock Holdings, LLC"
        operatingAgreementDate = Pick $null $e.manager.operatingAgreementDate
    }
    # The certifying party is the property management company (as authorized agent
    # for the owner). The signature block is left blank and filled at signing.
    certification = [ordered]@{
        relationshipToOwner = Pick $null $e.certification.relationshipToOwner "property manager and authorized agent"
    }
    filing = [ordered]@{
        taxYear                     = if ($e.filing.taxYear) { $e.filing.taxYear } else { (Get-Date).Year }
        filingType                  = Pick $null $e.filing.filingType "Annual Renewal Certification"
        annualCertificationDeadline = Pick $null $e.filing.annualCertificationDeadline "October 1"
    }
}

$config | ConvertTo-Json -Depth 6 | Set-Content -Path $OutPath -Encoding UTF8
Write-Host "  Wrote $OutPath" -ForegroundColor Green

# ---- Exemption-chain validation ------------------------------------------ #
Write-Host ""
if ($cahpMember) {
    Write-Host "  ✓ Exemption chain OK: '$([string]$cahpMember["Title"])' (CAHP instrumentality) " `
               "is a member at $($config.nonprofit.ownershipPercent)% $($config.nonprofit.memberClass)." -ForegroundColor Green
} else {
    Write-Host "  ====================  WARNING  ====================" -ForegroundColor Yellow
    Write-Host "  No CAHP-flagged 501(c)(3) instrumentality found in the ownership of" -ForegroundColor Yellow
    Write-Host "  '$EntityTitle'. Without that nonprofit managing member the entity does" -ForegroundColor Yellow
    Write-Host "  NOT qualify for the §12-37-220(B)(11)(e) exemption. The JSON was written" -ForegroundColor Yellow
    Write-Host "  with template defaults — verify the ownership in the hub before filing." -ForegroundColor Yellow
    Write-Host "  ===================================================" -ForegroundColor Yellow
}

# Flag still-blank fields the generator will render as underscores.
$blanks = @()
if (-not $config.company.ein)               { $blanks += "company.ein" }
if (-not $config.company.dorAccountId)      { $blanks += "company.dorAccountId" }
if (-not $config.nonprofit.parentEin)       { $blanks += "nonprofit.parentEin" }
if (-not $config.manager.operatingAgreementDate) { $blanks += "manager.operatingAgreementDate" }
if ($config.property.taxMapParcels.Count -eq 0) { $blanks += "property.taxMapParcels" }
if ($blanks.Count) {
    Write-Host ""
    Write-Host "  Fields still blank (edit the JSON to fill, or they print as blanks):" -ForegroundColor DarkYellow
    $blanks | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkYellow }
}
Write-Host ""
Write-Host "  Next: python scripts/safe-harbor/generate_cert.py --rent-roll <roll>.xlsx ``" -ForegroundColor Cyan
Write-Host "          --entity `"$OutPath`"" -ForegroundColor Cyan
