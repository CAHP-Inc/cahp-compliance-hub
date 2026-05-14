/**
 * TypeScript types matching the actual CAHP Compliance Hub SharePoint schema.
 *
 * Generated from inventory snapshot taken 2026-05-12 against:
 *   https://vanrockre.sharepoint.com/sites/CAHPComplianceHub
 *
 * Field names use SharePoint internal names (PropertyAddress, cahpState, etc.) —
 * those are what Graph API returns in the `fields` object.
 *
 * Lookup columns appear as `{FieldName}LookupId` (a numeric ID as string).
 * User columns appear as `{FieldName}LookupId` too — same pattern.
 *
 * Update this file whenever columns are added/removed/renamed in SharePoint.
 */

// =============================================================================
// SHARED TYPES
// =============================================================================

/** Common shape of any SharePoint list item from Graph API */
export interface SharePointListItem<TFields = Record<string, unknown>> {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl?: string;
  createdBy?: {
    user?: {
      displayName?: string;
      email?: string;
      id?: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName?: string;
      email?: string;
      id?: string;
    };
  };
  fields: TFields & {
    id?: string;
    Title?: string;
    Created?: string;
    Modified?: string;
    AuthorLookupId?: string;
    EditorLookupId?: string;
  };
}

/** State (used across most lists) */
export type CahpState = 'SC' | 'NC';

/** Tax year (string-typed choice in SharePoint) */
export type CahpTaxYear = '2023' | '2024' | '2025' | '2026' | '2027' | '2028';

// =============================================================================
// LIST: Properties Registry
// =============================================================================

export type AMIProgram = '20/50' | '40/60' | 'Mixed' | 'None';

export type CAHPLanguageStatus = 'Yes' | 'No' | 'In Progress' | 'Needs Revision';

export type OwnerGroup =
  | 'VanRock Holdings'
  | 'Red Cedar'
  | 'AmRock'
  | 'Troy Hampton'
  | 'Deepak'
  | 'Damon Lilly'
  | 'Other';

export type VerificationStatus =
  | 'Inherited - Unverified'
  | 'Verified'
  | 'Needs Follow-Up'
  | 'N/A';

export type LURAExecutedStatus = 'Yes' | 'No' | 'In Progress' | 'N/A';

export type PropertyStatus =
  | 'Active'
  | 'Pending'
  | 'Withdrawn'
  | 'Removed from Program'
  | 'Sold';

export interface PropertyFields {
  Title: string;                    // Property Name
  PropertyAddress?: string;
  LegalEntity?: string;
  UnitCount?: number;
  AMIProgram?: AMIProgram;
  CAHPLanguageAdded?: CAHPLanguageStatus;
  cahpCounty?: string;
  cahpOwnerGroup?: OwnerGroup;
  cahpState?: CahpState;
  cahpVerificationStatus?: VerificationStatus;
  DORAccountID?: string;
  DateAddedToCAHP?: string;
  LURAExecuted?: LURAExecutedStatus;
  OpAgreementVersion?: string;
  PropertyNotes?: string;
  PropertyStatus?: PropertyStatus;
  RemovedReason?: string;
}

export type Property = SharePointListItem<PropertyFields>;

// =============================================================================
// LIST: Submittals Tracker
// =============================================================================

export type SubmittalStatusValue =
  | 'Draft'
  | 'Package Mailed (NC)'
  | 'Filed'
  | 'Letter Received - Action Needed'
  | 'Responded - Awaiting DOR'
  | 'Approved'
  | 'Denied'
  | 'Withdrawn';

export type FilingMethod = 'Online Portal (SC)' | 'Paper Mail (NC)';

// PR-10a — spec §3.6.3
export type SubmittalFilingType = 'Initial' | 'Annual' | 'Amendment';

export interface SubmittalFields {
  Title: string;                              // Submittal Label
  PropertyLookupId?: string;                  // → Properties Registry
  AssignedToLookupId?: string;                // User picker
  cahpTaxYear?: CahpTaxYear;
  cahpState?: CahpState;
  SubmittalStatus?: SubmittalStatusValue;
  FilingType?: SubmittalFilingType;           // PR-10a — Initial / Annual / Amendment
  FilingMethod?: FilingMethod;
  DateFiled?: string;
  ConfirmationNumber?: string;
  MailTrackingNumber?: string;
  NextAction?: string;
  NextActionDue?: string;
  ApprovedAbatement?: number;
  SubmittalNotes?: string;
  // PR-10c — org chart version freeze on Draft→Filed transition (JSON snapshot)
  OrgChartSnapshotJSON?: string;
  OrgChartSnapshotDate?: string;
}

export type Submittal = SharePointListItem<SubmittalFields>;

// =============================================================================
// LIST: Compliance Deadlines
// =============================================================================

export type DeadlineType =
  | 'IRS 990 Filing'
  | 'Annual Recertification'
  | 'Rent Roll Review'
  | 'AMI Cert Renewal'
  | 'State Compliance Report'
  | 'Property Tax Filing'
  | 'Operating Agreement Review'
  | 'Other';

export type DeadlineStatus =
  | 'Upcoming'
  | 'In Progress'
  | 'Completed'
  | 'Overdue'
  | 'Missed';

export type DeadlineRecurrence = 'One-Time' | 'Annual' | 'Quarterly' | 'Monthly';

export type DeadlineAppliesTo =
  | 'CAHP Entity'
  | 'All Properties'
  | 'Specific Property'
  | 'SC Portfolio'
  | 'NC Portfolio';

export type ResponsibleParty =
  | 'Brandy'
  | 'Chris'
  | 'Brian'      // Note: SharePoint has 'Brian' — kept as-is to match existing data
  | 'John'
  | 'Aljon'
  | 'Other';

export interface ComplianceDeadlineFields {
  Title: string;                              // Deadline Description
  DeadlineType?: DeadlineType;
  DeadlineStatus?: DeadlineStatus;
  DueDate?: string;
  CompletionDate?: string;
  Recurrence?: DeadlineRecurrence;
  AppliesTo?: DeadlineAppliesTo;
  ResponsibleParty?: ResponsibleParty;
  PropertyLookupId?: string;                  // → Properties Registry
  cahpState?: CahpState;
  DeadlineNotes?: string;
}

export type ComplianceDeadline = SharePointListItem<ComplianceDeadlineFields>;

// =============================================================================
// LIST: DOR Correspondence Log
// =============================================================================

export type CorrespondenceDirection = 'Inbound (from DOR)' | 'Outbound (to DOR)';

export type LetterType =
  | 'Initial Acknowledgment'
  | 'Additional Info Request'
  | 'Org Chart Request'
  | 'Approval'
  | 'Denial'
  | 'Withdrawal Notice'
  | 'Refund Notice'
  | 'Other';

export interface CorrespondenceFields {
  Title: string;                              // Correspondence Subject
  Direction?: CorrespondenceDirection;
  LetterType?: LetterType;
  PropertyLookupId?: string;                  // → Properties Registry
  CorrSubmittalLookupId?: string;             // → Submittals Tracker (PR-11a)
  DateReceived?: string;
  DateResponded?: string;
  ResponseDue?: string;
  RequestSummary?: string;
  ResponseNotes?: string;
  cahpTaxYear?: CahpTaxYear;
  cahpState?: CahpState;
}

export type Correspondence = SharePointListItem<CorrespondenceFields>;

// =============================================================================
// LIST: Billing Tracker
// =============================================================================

export type BillingStatusValue =
  | 'Pending Approval'
  | 'Ready to Invoice'
  | 'Invoiced'
  | 'Paid'
  | 'Disputed';

export type QBSyncStatus = 'Not Synced' | 'Synced' | 'Discrepancy';

export interface BillingFields {
  Title: string;                              // Billing Reference
  PropertyLookupId?: string;                  // → Properties Registry
  cahpTaxYear?: CahpTaxYear;
  AmountBilled?: number;
  BillApprovedAbatement?: number;
  CAHPFeePercent?: number;
  InvoiceDate?: string;
  InvoiceNumber?: string;
  BillingStatus?: BillingStatusValue;
  QBSyncStatus?: QBSyncStatus;
  BillingNotes?: string;
}

export type Billing = SharePointListItem<BillingFields>;

// =============================================================================
// LIST: Disbursements (provisioned PR-09d — DOR refund passthroughs to owners)
// =============================================================================

export type DisbursementStatus = 'Pending' | 'Issued' | 'Cleared' | 'Voided';

export interface DisbursementFields {
  Title: string;                              // Disbursement Reference
  DisbProperty?: string;                       // Lookup → Properties Registry (raw name for now)
  DisbPropertyLookupId?: string;               // Standard SP lookup-id field shape
  DisbSubmittalLookupId?: string;              // → Submittals Tracker
  DisbOwnerLookupId?: string;                  // → Owners
  DisbAmount?: number;
  DisbStatus?: DisbursementStatus;
  DisbIssueDate?: string;
  DisbClearDate?: string;
  DisbCheckNum?: string;
  DisbNotes?: string;
}

export type Disbursement = SharePointListItem<DisbursementFields>;

// =============================================================================
// LIST: Outstanding Items Checklist
// =============================================================================

export type ItemCategory =
  | 'Operating Agreement'
  | 'LURA'
  | 'AMI Certification'
  | 'Org Chart'
  | 'Deed'
  | 'Income Documentation'
  | 'Signed Submittal'
  | 'Determination Letter'
  | 'Other';

export type ItemStatus = 'Requested' | 'Overdue' | 'Received' | 'Not Applicable';

export interface OutstandingItemFields {
  Title: string;                              // Item Needed
  PropertyLookupId?: string;                  // → Properties Registry
  ItemCategory?: ItemCategory;
  ItemStatus?: ItemStatus;
  DateRequested?: string;
  DateReceivedItem?: string;
  FollowUpCount?: number;
  ItemOwner?: string;
  ItemNotes?: string;
}

export type OutstandingItem = SharePointListItem<OutstandingItemFields>;

// =============================================================================
// LIST: Known Issues Log
// =============================================================================

export type IssueCategory =
  | 'Operating Agreement'
  | 'LURA'
  | 'DOR Filing'
  | 'Entity Mismatch'
  | 'Missing Documentation'
  | 'Compliance'
  | 'Billing'
  | 'Org Chart'
  | 'Other';

export type ResolutionStatus =
  | 'Open'
  | 'In Progress'
  | 'Blocked'
  | 'Resolved'
  | 'Accepted Risk';

export type IssueSeverity = 'Critical' | 'High' | 'Medium' | 'Low';

export interface KnownIssueFields {
  Title: string;                              // Issue Summary
  IssueCategory?: IssueCategory;
  ResolutionStatus?: ResolutionStatus;
  Severity?: IssueSeverity;
  PropertyLookupId?: string;                  // → Properties Registry
  IssueOwnerLookupId?: string;                // User picker
  IdentifiedDate?: string;
  DateResolved?: string;
  ResolutionNotes?: string;
}

export type KnownIssue = SharePointListItem<KnownIssueFields>;

// =============================================================================
// LIST: Owners — entity master (PR-09a)
// =============================================================================

export type OwnerType = 'Individual' | 'LLC' | 'Nonprofit';

export interface OwnerFields {
  Title?: string;                  // Legal entity name (e.g., "VanRock Holdings LLC")
  OwnerType?: OwnerType;
  OwnerState?: string;             // State of formation (LLC/Nonprofit) or residence (Individual)
  TaxID?: string;                  // EIN or SSN — masked in UI
  ContactEmail?: string;
  OwnerNotes?: string;
}

export type Owner = SharePointListItem<OwnerFields>;

// =============================================================================
// LIST: Ownership Structure (junction table — refactored in PR-09a)
// =============================================================================

export type RelationshipType =
  | 'Managing Member'
  | 'Member'
  | 'Owner'           // legacy — still rendered for old rows
  | 'Subsidiary'      // legacy
  | 'Beneficial Owner';  // legacy — beneficial ownership is now computed by the recursive engine

/**
 * Primary relationship types shown in the new ownership UI.
 * Legacy types (Owner, Subsidiary, Beneficial Owner) remain in the full union so existing rows
 * with those values display correctly, but new entries only get these two clean options.
 */
export const PRIMARY_RELATIONSHIP_TYPES: RelationshipType[] = ['Managing Member', 'Member'];

export interface OwnershipFields {
  Title?: string;                              // Auto-derived from Owner lookup; legacy rows have free-text
  OwnerLookupId?: string;                      // NEW PR-09a: Lookup → Owners (this row's entity)
  ParentOwnerLookupId?: string;                // NEW PR-09a: Lookup → Owners (parent entity if this is a member-of-member chain)
  ParentEntity?: string;                       // Legacy free-text parent — preserved for back-compat
  LinkedPropertyLookupId?: string;             // → Properties Registry
  RelationshipType?: RelationshipType;
  OwnershipPercent?: number;
  EffectiveDate?: string;
  SourceDocument?: string;
  EntityNotes?: string;
}

export type Ownership = SharePointListItem<OwnershipFields>;

// =============================================================================
// LIST: AuditLog (PR-07)
// =============================================================================

export type AuditAction = 'CREATE' | 'UPDATE' | 'DELETE';

export interface AuditLogFields {
  Title: string;                // Event summary (e.g., "Updated Property: 135 Oakwood")
  Action: AuditAction;
  EntityType?: string;          // Logical entity name (e.g., "Property", "Compliance Deadline")
  EntityId?: string;            // SharePoint item ID of affected record
  EntityTitle?: string;         // Human-readable name of affected record at time of change
  ChangeSummary?: string;       // Field-by-field readable diff
  BeforeJSON?: string;          // Full record JSON before change (forensic detail)
  AfterJSON?: string;           // Full record JSON after change
}

export type AuditLog = SharePointListItem<AuditLogFields>;

// =============================================================================
// LIST: Property Notes (PR-08d)
// =============================================================================

export interface PropertyNoteFields {
  Title?: string;                 // SP-required, auto-derived from NoteBody first 80 chars
  NoteBody?: string;              // The note content
  PropertyLookupId?: string;      // Lookup → Properties Registry
}

export type PropertyNote = SharePointListItem<PropertyNoteFields>;

// =============================================================================
// LIST NAMES — for reference and to catch typos at the call site
// =============================================================================

export const LIST_NAMES = {
  Properties: 'Properties Registry',
  Submittals: 'Submittals Tracker',
  ComplianceDeadlines: 'Compliance Deadlines',
  Correspondence: 'DOR Correspondence Log',
  Billing: 'Billing Tracker',
  Outstanding: 'Outstanding Items Checklist',
  KnownIssues: 'Known Issues Log',
  Ownership: 'Ownership Structure',
  AuditLog: 'AuditLog',
  PropertyNotes: 'Property Notes',
  Owners: 'Owners',
  Disbursements: 'Disbursements',
} as const;

export type ListName = (typeof LIST_NAMES)[keyof typeof LIST_NAMES];
