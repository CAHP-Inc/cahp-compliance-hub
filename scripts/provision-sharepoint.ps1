<#
.SYNOPSIS
    Provisions SharePoint Lists and Document Libraries for CAHP Compliance Hub.

.DESCRIPTION
    Creates 14 SharePoint Lists (the data layer) and 13 Document Libraries (file storage)
    per the Platform Specification v1.0 / database schema sheet.

    The script is IDEMPOTENT: running multiple times does not create duplicates. New columns
    or new lists added to this script will be created on re-run; existing items are skipped.

    Authentication is interactive — a browser window will open for sign-in. You must be a
    Site Owner or Site Collection Admin on the target site.

.PARAMETER SiteUrl
    Full URL of the SharePoint site to provision into.
    Example: https://newshirepm.sharepoint.com/sites/CAHPHub

.PARAMETER SkipDocumentLibraries
    Skip creating document libraries (provision only the data lists).

.EXAMPLE
    .\provision-sharepoint.ps1 -SiteUrl "https://newshirepm.sharepoint.com/sites/CAHPHub"

.NOTES
    Authored by Brandy Turner / NewShire Property Management
    Companion to: CAHP_Compliance_Hub_Platform_Specification_v1.0.docx
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Full URL of the SharePoint site, e.g. https://newshirepm.sharepoint.com/sites/CAHPHub")]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDocumentLibraries
)

$ErrorActionPreference = "Stop"

# ============================================================
# PREREQ CHECK
# ============================================================
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module is not installed. Install with:`n`n  Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber`n"
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  CAHP Compliance Hub — SharePoint Provisioning" -ForegroundColor Cyan
Write-Host "  Authored by Brandy Turner / NewShire Property Management" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Target site: $SiteUrl" -ForegroundColor White
Write-Host ""

# ============================================================
# CONNECT
# ============================================================
Write-Host "Connecting to SharePoint (browser may open for sign-in)..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}
Write-Host "  ✓ Connected" -ForegroundColor Green
Write-Host ""

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Ensure-CAHPList {
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [string]$Description,
        [hashtable[]]$Columns = @()
    )

    Write-Host ""
    Write-Host "→ $Title" -ForegroundColor White

    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        New-PnPList -Title $Title -Template GenericList -Url "lists/$Title" | Out-Null
        $list = Get-PnPList -Identity $Title
        Set-PnPList -Identity $Title -Description $Description -EnableVersioning $true | Out-Null
        Write-Host "  ✓ List created (with versioning enabled)" -ForegroundColor Green
    } else {
        Write-Host "  → List exists; verifying columns" -ForegroundColor Yellow
    }

    foreach ($col in $Columns) {
        $existing = Get-PnPField -List $Title -Identity $col.InternalName -ErrorAction SilentlyContinue
        if ($existing) {
            continue
        }

        try {
            $params = @{
                List         = $Title
                DisplayName  = $col.DisplayName
                InternalName = $col.InternalName
                Type         = $col.Type
            }
            if ($col.Required) { $params.Required = $true }
            if ($col.AddToDefaultView) { $params.AddToDefaultView = $true }

            if ($col.Type -eq 'Choice' -and $col.Choices) {
                $params.Choices = $col.Choices
            }

            Add-PnPField @params -ErrorAction Stop | Out-Null
            Write-Host "    + $($col.DisplayName) [$($col.Type)]" -ForegroundColor Cyan
        } catch {
            Write-Host "    ! Failed: $($col.DisplayName) — $_" -ForegroundColor Red
        }
    }
}

function Ensure-CAHPDocumentLibrary {
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [string]$Description
    )

    Write-Host ""
    Write-Host "→ $Title (Document Library)" -ForegroundColor White

    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        New-PnPList -Title $Title -Template DocumentLibrary -Url "$($Title -replace '\s','')" | Out-Null
        Set-PnPList -Identity $Title -Description $Description -EnableVersioning $true | Out-Null
        Write-Host "  ✓ Library created" -ForegroundColor Green
    } else {
        Write-Host "  → Library exists" -ForegroundColor Yellow
    }

    # Add PropertyID metadata column for filtering documents by property
    $existing = Get-PnPField -List $Title -Identity "PropertyID" -ErrorAction SilentlyContinue
    if (-not $existing) {
        try {
            Add-PnPField -List $Title -DisplayName "Property ID" -InternalName "PropertyID" -Type Text -AddToDefaultView | Out-Null
            Write-Host "    + Property ID metadata column" -ForegroundColor Cyan
        } catch {
            Write-Host "    ! Failed to add Property ID column: $_" -ForegroundColor Red
        }
    }

    # Add ExpirationDate column for documents that expire (LURAs, Insurance, AMI Certs, etc.)
    $existing = Get-PnPField -List $Title -Identity "ExpirationDate" -ErrorAction SilentlyContinue
    if (-not $existing) {
        try {
            Add-PnPField -List $Title -DisplayName "Expiration Date" -InternalName "ExpirationDate" -Type DateTime | Out-Null
            Write-Host "    + Expiration Date column" -ForegroundColor Cyan
        } catch {
            Write-Host "    ! Failed to add Expiration Date: $_" -ForegroundColor Red
        }
    }
}

# ============================================================
# DATA LISTS (14)
# ============================================================
# Foreign keys are stored as Text(36) UUIDs. Application enforces referential integrity.
# Built-in columns (Title, Created, Modified, Author, Editor) are not redefined.
# ============================================================

Write-Host ""
Write-Host "----------------------------------------------------------------" -ForegroundColor White
Write-Host "  PHASE 1: DATA LISTS (14)" -ForegroundColor White
Write-Host "----------------------------------------------------------------" -ForegroundColor White

# 1. Users — application users with role
Ensure-CAHPList -Title "Users" -Description "Application users — Admin, Contributor, Accounting roles" -Columns @(
    @{ DisplayName = "Email"; InternalName = "Email"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Role"; InternalName = "UserRole"; Type = "Choice"; Required = $true; Choices = @("Admin", "Contributor", "Accounting"); AddToDefaultView = $true }
    @{ DisplayName = "Org"; InternalName = "Org"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "Active"; InternalName = "Active"; Type = "Boolean"; AddToDefaultView = $true }
)

# 2. Owners — people, LLCs, and nonprofits
Ensure-CAHPList -Title "Owners" -Description "All persons, LLCs, and nonprofits with ownership interest. Joint owners stored as separate Individual rows." -Columns @(
    @{ DisplayName = "Owner Type"; InternalName = "OwnerType"; Type = "Choice"; Required = $true; Choices = @("Individual", "LLC", "Nonprofit"); AddToDefaultView = $true }
    @{ DisplayName = "State"; InternalName = "OwnerState"; Type = "Text" }
    @{ DisplayName = "Tax ID"; InternalName = "TaxID"; Type = "Text" }
    @{ DisplayName = "Contact Email"; InternalName = "ContactEmail"; Type = "Text" }
    @{ DisplayName = "Owner Notes"; InternalName = "OwnerNotes"; Type = "Note" }
    @{ DisplayName = "Is CAHP Entity"; InternalName = "IsCAHP"; Type = "Boolean"; AddToDefaultView = $true }
)

# 3. Properties — core records
Ensure-CAHPList -Title "Properties" -Description "Properties managed by CAHP. Title = display name (e.g., '144 W Henry')." -Columns @(
    @{ DisplayName = "Entity Name"; InternalName = "EntityName"; Type = "Text"; Required = $true }
    @{ DisplayName = "Address"; InternalName = "FullAddress"; Type = "Text"; Required = $true }
    @{ DisplayName = "State"; InternalName = "PropState"; Type = "Choice"; Required = $true; Choices = @("SC", "NC"); AddToDefaultView = $true }
    @{ DisplayName = "County"; InternalName = "County"; Type = "Text" }
    @{ DisplayName = "Units"; InternalName = "Units"; Type = "Number"; AddToDefaultView = $true }
    @{ DisplayName = "AMI Program"; InternalName = "AMIProgram"; Type = "Choice"; Choices = @("20/50", "40/60", "Mixed", "None"); AddToDefaultView = $true }
    @{ DisplayName = "Owner Display"; InternalName = "OwnerDisplay"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "CAHP Fee %"; InternalName = "CAHPFeePercent"; Type = "Number"; Required = $true }
    @{ DisplayName = "Annual Filing Required"; InternalName = "AnnualFilingRequired"; Type = "Boolean"; AddToDefaultView = $true }
    @{ DisplayName = "Lifecycle"; InternalName = "Lifecycle"; Type = "Choice"; Required = $true; Choices = @("Pre-Operational", "Active", "Non-Compliant", "Pending Disposition", "Disposed"); AddToDefaultView = $true }
    @{ DisplayName = "Acquisition Date"; InternalName = "AcquisitionDate"; Type = "DateTime" }
    @{ DisplayName = "Disposition Date"; InternalName = "DispositionDate"; Type = "DateTime" }
    @{ DisplayName = "Disposition Type"; InternalName = "DispositionType"; Type = "Choice"; Choices = @("Sale", "Withdrawal", "Lost AMI", "Other") }
    @{ DisplayName = "Disposition Notes"; InternalName = "DispositionNotes"; Type = "Note" }
)

# 4. Ownership — recursive junction (recursive: subject can be a Property OR an Owner)
Ensure-CAHPList -Title "Ownership" -Description "Junction: who owns what. SubjectType=Property: row says 'owner X has Y% of property Z'. SubjectType=Owner: 'owner X has Y% of LLC Z' (recursive)." -Columns @(
    @{ DisplayName = "Subject Type"; InternalName = "SubjectType"; Type = "Choice"; Required = $true; Choices = @("Property", "Owner"); AddToDefaultView = $true }
    @{ DisplayName = "Subject ID"; InternalName = "SubjectId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Owner ID"; InternalName = "OwnerId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Percentage"; InternalName = "Pct"; Type = "Number"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Role Label"; InternalName = "OwnerRole"; Type = "Text" }
    @{ DisplayName = "Effective Date"; InternalName = "EffectiveDate"; Type = "DateTime" }
    @{ DisplayName = "End Date"; InternalName = "EndDate"; Type = "DateTime" }
    @{ DisplayName = "Change Reason"; InternalName = "ChangeReason"; Type = "Choice"; Choices = @("Initial Filing", "Buy-In", "Buy-Out", "Estate", "Other") }
)

# 5. Submittals — DOR filings
Ensure-CAHPList -Title "Submittals" -Description "DOR property tax abatement filings. One per property per tax year for annual; plus initials and amendments." -Columns @(
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Tax Year"; InternalName = "TaxYear"; Type = "Number"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Filing Type"; InternalName = "FilingType"; Type = "Choice"; Required = $true; Choices = @("Initial", "Annual", "Amendment"); AddToDefaultView = $true }
    @{ DisplayName = "Status"; InternalName = "SubmittalStatus"; Type = "Choice"; Required = $true; Choices = @("Draft", "Filed", "Letter Received", "Responded - Awaiting DOR", "Approved", "Denied", "Withdrawn"); AddToDefaultView = $true }
    @{ DisplayName = "Filed Date"; InternalName = "FiledDate"; Type = "DateTime" }
    @{ DisplayName = "Confirmation Number"; InternalName = "ConfirmationNumber"; Type = "Text" }
    @{ DisplayName = "Assigned User ID"; InternalName = "AssignedUserId"; Type = "Text" }
    @{ DisplayName = "Due Date"; InternalName = "DueDate"; Type = "DateTime" }
    @{ DisplayName = "Next Action Note"; InternalName = "NextActionNote"; Type = "Note" }
    @{ DisplayName = "Org Chart Snapshot"; InternalName = "OrgChartSnapshot"; Type = "Note" }
    @{ DisplayName = "Tax Savings Amount"; InternalName = "TaxSavingsAmount"; Type = "Currency" }
)

# 6. Correspondence — DOR letters
Ensure-CAHPList -Title "Correspondence" -Description "DOR letters in and out — request for info, approvals, denials, audit notices." -Columns @(
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Letter Date"; InternalName = "LetterDate"; Type = "DateTime"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Letter Type"; InternalName = "LetterType"; Type = "Choice"; Required = $true; Choices = @("Letter from DOR", "Response Sent", "Approval", "Denial", "Audit Notice"); AddToDefaultView = $true }
    @{ DisplayName = "Summary"; InternalName = "LetterSummary"; Type = "Note" }
    @{ DisplayName = "Response Due"; InternalName = "ResponseDue"; Type = "DateTime" }
    @{ DisplayName = "Status"; InternalName = "CorrespondenceStatus"; Type = "Choice"; Choices = @("Awaiting My Response", "Awaiting DOR", "Closed"); AddToDefaultView = $true }
    @{ DisplayName = "Related Submittal ID"; InternalName = "RelatedSubmittalId"; Type = "Text" }
)

# 7. Communications — owner emails, calls, meetings (non-DOR)
Ensure-CAHPList -Title "Communications" -Description "Non-DOR communications: owner emails, phone calls, meetings, vendor calls. Separate timeline from DOR Correspondence." -Columns @(
    @{ DisplayName = "Date"; InternalName = "CommDate"; Type = "DateTime"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Type"; InternalName = "CommType"; Type = "Choice"; Required = $true; Choices = @("Email In", "Email Out", "Phone Call", "Meeting"); AddToDefaultView = $true }
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text" }
    @{ DisplayName = "Owner ID"; InternalName = "OwnerId"; Type = "Text" }
    @{ DisplayName = "Subject"; InternalName = "CommSubject"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "Summary"; InternalName = "CommSummary"; Type = "Note" }
    @{ DisplayName = "Participants"; InternalName = "Participants"; Type = "Text" }
    @{ DisplayName = "Follow-Up Due"; InternalName = "FollowUpDue"; Type = "DateTime" }
    @{ DisplayName = "Status"; InternalName = "CommStatus"; Type = "Choice"; Choices = @("Open · Needs Follow-up", "Closed"); AddToDefaultView = $true }
)

# 8. Outstanding — task list
Ensure-CAHPList -Title "Outstanding" -Description "Master task list: manual to-dos plus auto-created cascades from other modules (DOR letters, document expirations, annual cycles)." -Columns @(
    @{ DisplayName = "Description"; InternalName = "TaskDescription"; Type = "Note" }
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text" }
    @{ DisplayName = "Submittal ID"; InternalName = "SubmittalId"; Type = "Text" }
    @{ DisplayName = "Assigned User ID"; InternalName = "AssignedUserId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Due Date"; InternalName = "DueDate"; Type = "DateTime"; AddToDefaultView = $true }
    @{ DisplayName = "Priority"; InternalName = "Priority"; Type = "Choice"; Choices = @("Low", "Normal", "High", "Urgent"); AddToDefaultView = $true }
    @{ DisplayName = "Status"; InternalName = "TaskStatus"; Type = "Choice"; Required = $true; Choices = @("Not Started", "In Progress", "Blocked", "Done"); AddToDefaultView = $true }
    @{ DisplayName = "Completed At"; InternalName = "CompletedAt"; Type = "DateTime" }
    @{ DisplayName = "Completed By User ID"; InternalName = "CompletedByUserId"; Type = "Text" }
    @{ DisplayName = "Created Via"; InternalName = "CreatedVia"; Type = "Choice"; Choices = @("manual", "wf-01", "wf-03", "wf-04", "wf-05", "wf-08") }
)

# 9. Documents — metadata only (files live in Document Libraries)
Ensure-CAHPList -Title "Documents" -Description "Document metadata. Actual file binaries live in Document Libraries with matching PropertyID metadata tag." -Columns @(
    @{ DisplayName = "SharePoint URL"; InternalName = "SharePointURL"; Type = "URL"; Required = $true }
    @{ DisplayName = "Filename"; InternalName = "Filename"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Category"; InternalName = "DocCategory"; Type = "Choice"; Required = $true; Choices = @("Operating Agreement", "LURA", "AMI Cert", "Rent Roll", "Insurance", "DOR Correspondence", "Owner Communication", "Entity Formation", "Governance", "Submittal Package", "Tax Certificate", "Compliance", "Other"); AddToDefaultView = $true }
    @{ DisplayName = "Size (bytes)"; InternalName = "SizeBytes"; Type = "Number" }
    @{ DisplayName = "Version"; InternalName = "DocVersion"; Type = "Text" }
    @{ DisplayName = "Uploaded By User ID"; InternalName = "UploadedByUserId"; Type = "Text" }
    @{ DisplayName = "Upload Date"; InternalName = "UploadDate"; Type = "DateTime"; AddToDefaultView = $true }
    @{ DisplayName = "Expiration Date"; InternalName = "ExpirationDate"; Type = "DateTime"; AddToDefaultView = $true }
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "Submittal ID"; InternalName = "SubmittalId"; Type = "Text" }
    @{ DisplayName = "Description"; InternalName = "DocDescription"; Type = "Note" }
    @{ DisplayName = "Archived"; InternalName = "Archived"; Type = "Boolean" }
)

# 10. Billing — CAHP fee invoices
Ensure-CAHPList -Title "Billing" -Description "CAHP fee invoices. Auto-created on Submittal Approval." -Columns @(
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Submittal ID"; InternalName = "SubmittalId"; Type = "Text"; Required = $true }
    @{ DisplayName = "Tax Year"; InternalName = "TaxYear"; Type = "Number"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Tax Savings Amount"; InternalName = "TaxSavingsAmount"; Type = "Currency" }
    @{ DisplayName = "CAHP Fee %"; InternalName = "CAHPFeePercent"; Type = "Number" }
    @{ DisplayName = "CAHP Fee Amount"; InternalName = "CAHPFeeAmount"; Type = "Currency"; AddToDefaultView = $true }
    @{ DisplayName = "QB Invoice Number"; InternalName = "QBInvoiceNumber"; Type = "Text" }
    @{ DisplayName = "Status"; InternalName = "BillingStatus"; Type = "Choice"; Required = $true; Choices = @("Pending", "Invoiced", "Paid", "Closed", "Disputed"); AddToDefaultView = $true }
    @{ DisplayName = "Invoiced Date"; InternalName = "InvoicedDate"; Type = "DateTime" }
)

# 11. Disbursements — refund passthrough to owners
Ensure-CAHPList -Title "Disbursements" -Description "DOR refunds → CAHP fee deducted → balance disbursed to owner." -Columns @(
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Submittal ID"; InternalName = "SubmittalId"; Type = "Text"; Required = $true }
    @{ DisplayName = "Tax Year"; InternalName = "TaxYear"; Type = "Number"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Refund Received Amount"; InternalName = "RefundReceivedAmount"; Type = "Currency"; AddToDefaultView = $true }
    @{ DisplayName = "Date Received"; InternalName = "DateReceived"; Type = "DateTime" }
    @{ DisplayName = "CAHP Fee Amount"; InternalName = "CAHPFeeAmount"; Type = "Currency" }
    @{ DisplayName = "Owner Owed Amount"; InternalName = "OwnerOwedAmount"; Type = "Currency"; AddToDefaultView = $true }
    @{ DisplayName = "Date Paid"; InternalName = "DatePaid"; Type = "DateTime" }
    @{ DisplayName = "Owner Paid Amount"; InternalName = "OwnerPaidAmount"; Type = "Currency" }
    @{ DisplayName = "Payment Method"; InternalName = "PaymentMethod"; Type = "Choice"; Choices = @("ACH", "Check", "Wire") }
    @{ DisplayName = "Payment Reference"; InternalName = "PaymentReference"; Type = "Text" }
    @{ DisplayName = "Status"; InternalName = "DisbursementStatus"; Type = "Choice"; Required = $true; Choices = @("Refund Received", "Owed to Owner", "Paid", "Hold"); AddToDefaultView = $true }
)

# 12. AuditLog — immutable change log
Ensure-CAHPList -Title "AuditLog" -Description "Immutable audit trail. Every CREATE/UPDATE/DELETE/APPROVE/UPLOAD writes a row. Write-once enforcement comes in PR-07 / Phase 4." -Columns @(
    @{ DisplayName = "Timestamp"; InternalName = "EventTimestamp"; Type = "DateTime"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Actor User ID"; InternalName = "ActorUserId"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "Action"; InternalName = "EventAction"; Type = "Choice"; Required = $true; Choices = @("CREATE", "UPDATE", "DELETE", "APPROVE", "UPLOAD", "SIGN_IN", "SIGN_OUT"); AddToDefaultView = $true }
    @{ DisplayName = "Entity Type"; InternalName = "EntityType"; Type = "Text"; AddToDefaultView = $true }
    @{ DisplayName = "Entity ID"; InternalName = "EntityId"; Type = "Text" }
    @{ DisplayName = "Description"; InternalName = "EventDescription"; Type = "Note" }
    @{ DisplayName = "Before Value"; InternalName = "BeforeValue"; Type = "Note" }
    @{ DisplayName = "After Value"; InternalName = "AfterValue"; Type = "Note" }
    @{ DisplayName = "Reason"; InternalName = "EventReason"; Type = "Text" }
    @{ DisplayName = "IP Address"; InternalName = "IPAddress"; Type = "Text" }
)

# 13. Notifications — My Day feed
Ensure-CAHPList -Title "Notifications" -Description "In-app notifications shown in My Day inbox. Phase 2 may add email delivery." -Columns @(
    @{ DisplayName = "Recipient User ID"; InternalName = "RecipientUserId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Type"; InternalName = "NotificationType"; Type = "Choice"; Required = $true; Choices = @("DOR Letter", "Outstanding Overdue", "Document Expiring", "Submittal Approved", "Mention", "System"); AddToDefaultView = $true }
    @{ DisplayName = "Related Entity Type"; InternalName = "RelatedEntityType"; Type = "Text" }
    @{ DisplayName = "Related Entity ID"; InternalName = "RelatedEntityId"; Type = "Text" }
    @{ DisplayName = "Body"; InternalName = "NotificationBody"; Type = "Note" }
    @{ DisplayName = "Read At"; InternalName = "ReadAt"; Type = "DateTime" }
    @{ DisplayName = "Dismissed At"; InternalName = "DismissedAt"; Type = "DateTime" }
)

# 14. Notes — free-form per property
Ensure-CAHPList -Title "Notes" -Description "Free-form notes per property. Supports pinning." -Columns @(
    @{ DisplayName = "Property ID"; InternalName = "PropertyId"; Type = "Text"; Required = $true; AddToDefaultView = $true }
    @{ DisplayName = "Note Body"; InternalName = "NoteBody"; Type = "Note" }
    @{ DisplayName = "Pinned"; InternalName = "Pinned"; Type = "Boolean"; AddToDefaultView = $true }
)

# ============================================================
# DOCUMENT LIBRARIES (13)
# ============================================================
if (-not $SkipDocumentLibraries) {
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor White
    Write-Host "  PHASE 2: DOCUMENT LIBRARIES (13)" -ForegroundColor White
    Write-Host "----------------------------------------------------------------" -ForegroundColor White

    Ensure-CAHPDocumentLibrary -Title "Operating Agreements" -Description "Operating Agreement documents and OA Amendments adding CAHP as Managing Member."
    Ensure-CAHPDocumentLibrary -Title "LURAs" -Description "Land Use Restriction Agreements recorded against each affordable housing property."
    Ensure-CAHPDocumentLibrary -Title "AMI Certifications" -Description "Area Median Income certifications proving compliance with set-aside requirements."
    Ensure-CAHPDocumentLibrary -Title "Rent Rolls" -Description "Property rent rolls submitted with annual filings."
    Ensure-CAHPDocumentLibrary -Title "Insurance Certificates" -Description "Property insurance certificates with expiration tracking."
    Ensure-CAHPDocumentLibrary -Title "DOR Correspondence Files" -Description "PDF attachments of DOR letters in and out."
    Ensure-CAHPDocumentLibrary -Title "Owner Communication Files" -Description "Attachments from owner emails, meeting notes, signed documents."
    Ensure-CAHPDocumentLibrary -Title "Entity Formation" -Description "Articles of Organization, EIN letters, S.O.S. filings for property entities and CAHP entities."
    Ensure-CAHPDocumentLibrary -Title "Governance" -Description "CAHP board resolutions, meeting minutes, 990 tax filings, bylaws."
    Ensure-CAHPDocumentLibrary -Title "Submittal Packages" -Description "Complete filing packages submitted to DOR."
    Ensure-CAHPDocumentLibrary -Title "Tax Certificates" -Description "Property tax certificates and county-issued documentation."
    Ensure-CAHPDocumentLibrary -Title "Compliance Documents" -Description "Fair housing audits, compliance reports, regulatory submissions."
    Ensure-CAHPDocumentLibrary -Title "Backups" -Description "Nightly JSON exports from Power Automate and manual quarterly snapshots."
}

# ============================================================
# DONE
# ============================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Provisioning complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor White
Write-Host ""
Write-Host "  1. Visit your site and click 'Site contents' to verify:" -ForegroundColor Gray
Write-Host "     $SiteUrl/_layouts/15/viewlsts.aspx" -ForegroundColor Yellow
Write-Host ""
Write-Host "  2. Set VITE_SHAREPOINT_SITE as a GitHub Actions repository variable." -ForegroundColor Gray
Write-Host "     Format for Microsoft Graph (hostname:/path):" -ForegroundColor Gray

# Compute the Graph-compatible site identifier
$uri = [System.Uri]$SiteUrl
$graphSiteId = "$($uri.Host):$($uri.AbsolutePath)"
Write-Host "     $graphSiteId" -ForegroundColor Yellow
Write-Host ""
Write-Host "  3. Come back to Claude for PR-04 (Graph SDK data layer)." -ForegroundColor Gray
Write-Host ""

Disconnect-PnPOnline
