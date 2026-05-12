<#
.SYNOPSIS
    Read-only inventory of an existing SharePoint site's Lists, Document Libraries, and columns.

.DESCRIPTION
    Prints every non-hidden list/library on the site with its base type, item count, and column
    schema (display name, internal name, type, required flag, and choice options where applicable).
    Makes NO changes to SharePoint. Safe to run on production sites.

.PARAMETER SiteUrl
    Full URL of the SharePoint site to inventory.

.PARAMETER OutFile
    Optional path to write the output to a text file in addition to the console.

.EXAMPLE
    .\inventory-sharepoint.ps1 -SiteUrl "https://newshirepm.sharepoint.com/sites/CAHPHub"

.EXAMPLE
    .\inventory-sharepoint.ps1 -SiteUrl "https://newshirepm.sharepoint.com/sites/CAHPHub" -OutFile inventory.txt
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutFile,

    [Parameter(Mandatory = $false, HelpMessage = "Azure AD Application (Client) ID. Defaults to PNP_MANAGEMENT_SHELL_CLIENTID env var.")]
    [string]$ClientId = $env:PNP_MANAGEMENT_SHELL_CLIENTID
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module is not installed. Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}

# Capture all output for both console and optional file
$output = New-Object System.Text.StringBuilder

function Write-Both {
    param([string]$Text, [string]$Color = "Gray")
    Write-Host $Text -ForegroundColor $Color
    $null = $output.AppendLine($Text)
}

Write-Both ""
Write-Both "================================================================" "Cyan"
Write-Both "  CAHP Compliance Hub — SharePoint Site Inventory" "Cyan"
Write-Both "  Site: $SiteUrl" "Cyan"
Write-Both "  Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" "Cyan"
Write-Both "================================================================" "Cyan"

try {
    if (-not $ClientId) {
        Write-Error "ClientId is required. Either pass -ClientId <guid> or set the PNP_MANAGEMENT_SHELL_CLIENTID environment variable.`n`nUse your CAHP Compliance Hub Azure AD app's Application (client) ID — the same value you have in .env.local as VITE_AZURE_CLIENT_ID."
        exit 1
    }
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}

# System lists to skip (built-in SharePoint plumbing, not user data)
$systemLists = @(
    'Form Templates', 'Site Assets', 'Site Pages', 'Style Library',
    'Customized Reports', 'Master Page Gallery', 'Theme Gallery',
    'Web Part Gallery', 'List Template Gallery', 'Solution Gallery',
    'User Information List', 'TaxonomyHiddenList', 'fpdatasources',
    'Composed Looks', 'Workflows', 'wfpub', 'wfsvc', 'IWConvertedForms',
    'Maintenance Log Library', 'Long Running Operation Status'
)

$lists = Get-PnPList | Where-Object {
    -not $_.Hidden -and $_.Title -notin $systemLists
} | Sort-Object -Property BaseType, Title

$listCount = ($lists | Where-Object { $_.BaseType -eq 'GenericList' }).Count
$libCount = ($lists | Where-Object { $_.BaseType -eq 'DocumentLibrary' }).Count

Write-Both ""
Write-Both "Found $listCount Lists and $libCount Document Libraries (total $($lists.Count) items)" "White"

foreach ($list in $lists) {
    $type = if ($list.BaseType -eq 'DocumentLibrary') { 'Document Library' } else { 'List' }

    Write-Both ""
    Write-Both "----------------------------------------------------------------" "Yellow"
    Write-Both "  $($list.Title)" "Yellow"
    Write-Both "  Type: $type | Items: $($list.ItemCount) | URL: $($list.RootFolder.ServerRelativeUrl)" "Gray"
    Write-Both "----------------------------------------------------------------" "Yellow"

    # System columns to skip — focus on user-defined ones plus useful builtins
    $skipColumns = @(
        '_ColorTag', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagWrittenTime',
        '_ComplianceTagUserId', '_DisplayName', '_UIVersionString', '_dlc_DocId',
        '_dlc_DocIdUrl', '_dlc_DocIdPersistId', '_HasCopyDestinations', '_CopySource',
        'ContentType', 'ContentTypeId', 'Attachments', '_ModerationStatus',
        '_ModerationComments', 'FileSystemObjectType', 'ServerRedirectedEmbedUrl',
        'ServerRedirectedEmbedUri', 'AppAuthor', 'AppEditor', 'FolderChildCount',
        'ItemChildCount', 'ParentVersionString', 'ParentLeafName', 'SortBehavior',
        'CheckoutUser', 'ProgId', 'ScopeId', 'File_x0020_Type', 'HTML_x0020_File_x0020_Type',
        '_EditMenuTableStart', '_EditMenuTableStart2', '_EditMenuTableEnd', 'LinkFilenameNoMenu',
        'LinkFilename', 'LinkFilename2', 'DocIcon', 'ServerUrl', 'EncodedAbsUrl',
        'BaseName', 'FileSizeDisplay', 'MetaInfo', 'Restricted', 'OriginatorId',
        'NoExecute', 'ContentVersion', 'Order', 'WorkflowVersion', 'WorkflowInstanceID',
        'GUID', 'WorkflowInstance', 'FileLeafRef', 'FileDirRef', 'Last_x0020_Modified',
        'Created_x0020_Date', 'FSObjType', '_HasEncryptedContent', 'PermMask',
        'SelectTitle', 'SelectFilename', 'Edit', 'owshiddenversion', 'UniqueId',
        '_Level', '_IsCurrentVersion', 'ItemChildCount', 'FolderChildCount',
        'TaxCatchAll', 'TaxCatchAllLabel', '_SharedFileIndex', '_CopySource',
        '_CheckinComment', 'Combine', 'RepairDocument', 'TemplateUrl', 'xd_ProgID',
        'xd_Signature', 'AccessPolicy', 'LinkCheckedOutTitle', 'LinkTitleNoMenu',
        'LinkTitle', 'LinkTitle2', '_SourceUrl', '_SharedFileIndex'
    )

    $fields = Get-PnPField -List $list.Title | Where-Object {
        -not $_.Hidden -and
        -not $_.ReadOnlyField -and
        $_.InternalName -notin $skipColumns
    } | Sort-Object Title

    if ($fields.Count -eq 0) {
        Write-Both "  (no user-defined columns)" "DarkGray"
        continue
    }

    foreach ($field in $fields) {
        $req = if ($field.Required) { " *required*" } else { "" }
        $defaultsLine = ""

        $choicesLine = ""
        if ($field.TypeAsString -eq "Choice" -or $field.TypeAsString -eq "MultiChoice") {
            $choicesLine = "    Choices: $($field.Choices -join ' | ')"
        }

        $lookupLine = ""
        if ($field.TypeAsString -eq "Lookup") {
            try {
                $lookupListId = $field.LookupList
                if ($lookupListId) {
                    $targetList = Get-PnPList -Identity $lookupListId -ErrorAction SilentlyContinue
                    if ($targetList) {
                        $lookupLine = "    Lookup → $($targetList.Title).$($field.LookupField)"
                    } else {
                        $lookupLine = "    Lookup → (list id: $lookupListId)"
                    }
                }
            } catch {
                # ignore lookup resolution errors
            }
        }

        Write-Both "  • $($field.Title)" "White"
        Write-Both "      Internal: $($field.InternalName) | Type: $($field.TypeAsString)$req" "Gray"
        if ($choicesLine) { Write-Both $choicesLine "DarkGray" }
        if ($lookupLine) { Write-Both $lookupLine "DarkGray" }
    }
}

Write-Both ""
Write-Both "================================================================" "Cyan"
Write-Both "  Inventory complete." "Green"
Write-Both "================================================================" "Cyan"
Write-Both ""

if ($OutFile) {
    $output.ToString() | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Host "Written to: $OutFile" -ForegroundColor Green
}

Disconnect-PnPOnline
