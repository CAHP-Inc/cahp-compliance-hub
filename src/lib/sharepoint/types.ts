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

export type AMIProgram = '20/50' | '40/60' | '50/80' | '60/80' | 'Mixed' | 'None';

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
  PropertyEIN?: string;             // Federal EIN — used for DOR submittal lookups
  DateAddedToCAHP?: string;
  LURAExecuted?: LURAExecutedStatus;
  OpAgreementVersion?: string;
  PropertyNotes?: string;
  PropertyStatus?: PropertyStatus;
  RemovedReason?: string;
  // Owner contact — Lookup → Contacts list. Set on the Property Overview tab via picker.
  PropertyOwnerContactLookupId?: string;
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
  TaxMapIDLookupId?: string;                  // → Tax Map IDs list (per-parcel filing)
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
  ResponsibleParty?: ResponsibleParty;        // Legacy hardcoded enum — superseded by AssignedTo
  AssignedTo?: string;                        // Free text — supports team members or outside parties (vendors, counsel, owners)
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

export type DisbPaymentMethod = 'ACH' | 'Check' | 'Wire' | 'Other';

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
  DisbPaymentMethod?: DisbPaymentMethod;       // PR-12b
  DisbNotes?: string;
}

export type Disbursement = SharePointListItem<DisbursementFields>;

// =============================================================================
// LIST: Owner Communications (PR-13a)
// =============================================================================

export type CommType = 'Email' | 'Phone' | 'Meeting' | 'SMS' | 'Other';
export type CommDirection = 'Inbound' | 'Outbound';
export type CommStatus = 'Open' | 'Closed';

export interface OwnerCommunicationFields {
  Title: string;                              // Subject
  CommPropertyLookupId?: string;              // → Properties Registry
  CommOwnerLookupId?: string;                 // → Owners
  CommType?: CommType;
  CommDirection?: CommDirection;
  CommDate?: string;
  CommResponseDue?: string;
  CommParticipants?: string;
  CommStatus?: CommStatus;
  CommNotes?: string;
}

export type OwnerCommunication = SharePointListItem<OwnerCommunicationFields>;

// =============================================================================
// LIST: Notifications (PR-15a)
// =============================================================================

export type NotificationType =
  | 'TaskAssigned'
  | 'DeadlineApproaching'
  | 'SubmittalUpdate'
  | 'OwnerInvited'
  | 'SystemAlert'
  | 'Other';

export interface NotificationFields {
  Title: string;                              // Notification headline
  NotifAssignedTo?: string;                   // UPN of recipient
  NotifType?: NotificationType;
  NotifTargetType?: string;                   // Entity type (Submittal, OutstandingItem, etc.)
  NotifTargetId?: string;                     // Entity ID
  NotifIsRead?: boolean;
  NotifUrl?: string;                          // Click target — relative path within app
}

export type Notification = SharePointListItem<NotificationFields>;

// =============================================================================
// LIST: Outstanding Items Checklist
// =============================================================================

export type ItemCategory =
  | 'Operating Agreement'
  | 'Articles of Incorporation'      // NEW — nonprofit + LLC formation
  | 'EIN Confirmation'               // NEW
  | 'Certificate of Existence'       // NEW — state-issued
  | 'Certificate of Authorization'   // NEW — state-issued
  | '501(c)(3) Determination'        // NEW — nonprofit-specific
  | 'Deed'
  | 'Rent Roll'                      // NEW
  | 'LURA'
  | 'AMI Certification'
  | 'Org Chart'
  | 'Income Documentation'
  | 'Signed Submittal'
  | 'Determination Letter'           // generic — property determinations
  | 'Other';

export type ItemStatus =
  | 'Not Started'        // PR-11b — primary new status
  | 'In Progress'        // PR-11b
  | 'Blocked'            // PR-11b
  | 'Done'               // PR-11b
  | 'Requested'          // legacy (maps to Not Started in kanban)
  | 'Overdue'            // legacy — overdue is a visual treatment in PR-11b
  | 'Received'           // legacy (maps to Done in kanban)
  | 'Not Applicable';    // legacy — closed bucket

export type ItemPriority = 'Critical' | 'High' | 'Medium' | 'Low';

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
  // PR-11b additions
  DueDate?: string;
  Priority?: ItemPriority;
  AssignedTo?: string;
  // PR-14b additions — link to fulfilling document
  RelatedDocUrl?: string;
  RelatedDocFilename?: string;
  RelatedDocLibrary?: string;
  // PR-14b — submittal linkage when item is part of a filing checklist
  RelatedSubmittalLookupId?: string;
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

export type OwnerType = 'Individual' | 'LLC' | 'Nonprofit' | 'Trust' | 'Corporation' | 'Limited Partnership' | 'General Partnership';

export interface OwnerFields {
  Title?: string;                  // Legal entity name (e.g., "VanRock Holdings LLC")
  OwnerType?: OwnerType;
  OwnerState?: string;             // State of formation (LLC/Nonprofit) or residence (Individual)
  TaxID?: string;                  // EIN or SSN — masked in UI
  ContactEmail?: string;
  OwnerNotes?: string;
  // PR-17 — org chart display fields
  SponsorName?: string;            // The human owner/principal behind the entity (e.g., "Deepak Maheshwari" for Marwar Ventures)
  OwnerAddress?: string;           // Address of the entity (shown on detailed org charts)
  IsTaxExempt?: boolean;           // For nonprofits — adds "IRC § 501(c)(3) Tax-Exempt" line to org chart
  EntityDescription?: string;      // Optional override for auto-derived "South Carolina LLC" label on org charts
}

export type Owner = SharePointListItem<OwnerFields>;

// =============================================================================
// LIST: Ownership Structure (junction table — refactored in PR-09a)
// =============================================================================

export type RelationshipType =
  | 'Managing Member'
  | 'Sole Member'      // wholly-owned single-member LLC (e.g., CAHP SC LLC owned 100% by CAHP Inc)
  | 'Member'
  | 'Owner'           // legacy — still rendered for old rows
  | 'Subsidiary'      // legacy
  | 'Beneficial Owner';  // legacy — beneficial ownership is now computed by the recursive engine

/**
 * Primary relationship types shown in the new ownership UI.
 * Legacy types (Owner, Subsidiary, Beneficial Owner) remain in the full union so existing rows
 * with those values display correctly, but new entries only get these two clean options.
 */
export const PRIMARY_RELATIONSHIP_TYPES: RelationshipType[] = ['Managing Member', 'Sole Member', 'Member'];

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
  // PR-17 — per-relationship metadata for org chart display
  MemberClass?: string;                        // Class A / Class B / Class C / N/A — appears on org chart card
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
  TaxMapIDs: 'Tax Map IDs',
  Deeds: 'Property Deeds',
  DeedParcelLinks: 'Deed Parcel Links',
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
  Communications: 'Owner Communications',
  Notifications: 'Notifications',
  DocumentMetadata: 'Document Metadata',
  Contacts: 'Contacts',
  ContactOwnerLinks: 'Contact Owner Links',
} as const;

export type ListName = (typeof LIST_NAMES)[keyof typeof LIST_NAMES];

// =============================================================================
// LIST: Tax Map IDs (per-parcel tracking for multi-parcel properties)
// =============================================================================

export type ParcelStatus = 'Active' | 'Inactive' | 'Split' | 'Merged';

export interface TaxMapIDFields {
  Title: string;                          // The tax map / parcel ID
  LinkedPropertyLookupId?: string;        // → Properties Registry
  ParcelAddress?: string;                 // Physical street address for this parcel
  County?: string;
  Acreage?: number;
  LegalDescription?: string;
  ParcelStatus?: ParcelStatus;
  ParcelNotes?: string;
}

export type TaxMapID = SharePointListItem<TaxMapIDFields>;

// =============================================================================
// LIST: Deeds (metadata on Property Deeds library; many-to-many with Tax Map IDs)
// =============================================================================

export interface DeedFields {
  Title?: string;                         // Library auto-sets to filename
  FileLeafRef?: string;                   // Filename
  GranteeOwnerLookupId?: string;          // → Owners
  BookPage?: string;                      // The thing SCDOR cares about
  DateRecorded?: string;
}

export type Deed = SharePointListItem<DeedFields>;

// Junction list — one row per (deed, parcel) link
// Deed lookup points at Property Deeds library.
export interface DeedParcelLinkFields {
  Title?: string;
  DeedLookupId?: string;                  // → Property Deeds library item
  TaxMapIDLookupId?: string;              // → Tax Map IDs
}

export type DeedParcelLink = SharePointListItem<DeedParcelLinkFields>;

// =============================================================================
// LIST: Contacts — people we ping (property owners, attorneys, vendors)
//
// Separate from Owners (entity records). A Contact CAN be linked to an Owner
// (e.g., Deepak Maheshwari → Marwar Ventures LLC), but doesn't have to be.
// Contacts feed the AssigneePicker dropdown alongside team members so items
// can be assigned to external folks without bespoke text matching.
// =============================================================================

export type ContactRole =
  | 'Property Owner'
  | 'Sponsor'
  | 'Attorney'
  | 'Accountant'
  | 'Property Manager'
  | 'Vendor'
  | 'Lender'
  | 'Other';

export interface ContactFields {
  Title?: string;                         // Display name (e.g., "Deepak Maheshwari")
  ContactEmail?: string;                  // Primary contact email — used for assignee matching
  ContactPhone?: string;
  ContactRole?: ContactRole;
  /**
   * Legacy single-owner linkage. Kept populated as the "primary" owner
   * for back-compat with existing rows and SharePoint default views, but
   * the source of truth for all (contact, owner) links is the Contact
   * Owner Links junction list (one row per relationship). A contact can
   * represent multiple Owner entities (e.g., Deepak → Marwar Ventures,
   * Maheshwari Holdings, Deepak Family Trust).
   */
  ContactOwnerLookupId?: string;
  ContactNotes?: string;
}

export type Contact = SharePointListItem<ContactFields>;

// =============================================================================
// LIST: Contact Owner Links — junction between Contacts and Owners (many-to-many)
//
// One row per (contact, owner) association. Lets a single contact represent
// any number of Owner entities, and lets us walk in both directions.
// =============================================================================

export interface ContactOwnerLinkFields {
  Title?: string;                         // Auto-derived ("Contact 12 ↔ Owner 7")
  ContactLookupId?: string;               // → Contacts
  OwnerLookupId?: string;                 // → Owners
}

export type ContactOwnerLink = SharePointListItem<ContactOwnerLinkFields>;
