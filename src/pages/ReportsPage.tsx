import { useMemo, useState } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type Submittal,
  type OutstandingItem,
  type ComplianceDeadline,
  type Owner,
  type Ownership,
  type Correspondence,
  type OwnerCommunication,
  type Contact,
  type ContactOwnerLink,
} from '../lib/sharepoint';
import { formatDateOnly, formatDateET, parseDateOnly } from '../lib/dates';
import {
  REPORTS,
  blobToBase64,
  buildCSV,
  buildJSON,
  buildPDF,
  buildXLSX,
  timestampedFilename,
  triggerDownload,
  type ReportCategory,
  type ReportDescriptor,
  type RunOptions,
  type RunResult,
} from '../lib/reports';
import { createBrandedPDF } from '../lib/reports-pdf';
import { PROPERTY_LINKED_LIBRARIES } from './../components/UploadDocumentModal';
import { Icon } from '../components/ui/Icon';
import { RunReportModal } from '../components/RunReportModal';
import { SafeHarborCertModal } from '../components/SafeHarborCertModal';
import { ComposeEmailModal } from '../components/ComposeEmailModal';
import type { EmailAttachment } from '../lib/email';

// =============================================================================
// Category presentation
// =============================================================================

const CATEGORY_ICONS: Record<ReportCategory, 'star' | 'file' | 'calendar' | 'check' | 'folder'> = {
  'Owner Reports': 'star',
  'DOR Audit Pack': 'file',
  'Year-End / Annual': 'calendar',
  'Operational': 'check',
  'Backup and Export': 'folder',
};

const AUDIENCE_STYLES = {
  internal: 'bg-teal-50 text-teal-800 border-teal-200',
  owner: 'bg-gold-100 text-gold-900 border-gold-200',
} as const;

// =============================================================================
// Doc library raw shape (one of the eight property-linked libraries)
// =============================================================================

interface DocItemRaw {
  id: string;
  webUrl?: string;
  lastModifiedDateTime: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    PropertyLookupId?: string;
    OwnerLookupId?: string;
    Modified?: string;
    Created?: string;
    File_x0020_Size?: string | number;
    ExpirationDate?: string;
  };
}

// =============================================================================
// Reports page
// =============================================================================

export function ReportsPage() {
  // Core data
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const compliance = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });
  const comms = useSharePointList<OwnerCommunication>(LIST_NAMES.Communications, { top: 500 });
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const contactOwnerLinks = useSharePointList<ContactOwnerLink>(LIST_NAMES.ContactOwnerLinks, { top: 2000 });

  // Documents — 8 libraries
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });
  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];

  // UI state
  const [openDescriptor, setOpenDescriptor] = useState<ReportDescriptor | null>(null);
  const [safeHarborOpen, setSafeHarborOpen] = useState(false);
  const [runResult, setRunResult] = useState<
    | { kind: 'success'; descriptor: ReportDescriptor; result: RunResult; options: RunOptions }
    | { kind: 'error'; descriptor: ReportDescriptor; message: string }
    | null
  >(null);
  const [composeOpen, setComposeOpen] = useState(false);
  const [composeAttachment, setComposeAttachment] = useState<EmailAttachment | null>(null);

  // Derived
  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const ownersById = useMemo(() => {
    if (!owners.data) return new Map<string, Owner>();
    return new Map(owners.data.map((o) => [String(o.id), o]));
  }, [owners.data]);

  const contactsById = useMemo(() => {
    if (!contacts.data) return new Map<string, Contact>();
    return new Map(contacts.data.map((c) => [String(c.id), c]));
  }, [contacts.data]);

  /** owner id → array of contact ids linked to that owner (legacy field + junction list) */
  const contactsByOwnerId = useMemo(() => {
    const map = new Map<string, Set<string>>();
    (contacts.data ?? []).forEach((c) => {
      const oid = c.fields.ContactOwnerLookupId ? String(c.fields.ContactOwnerLookupId) : null;
      if (oid) {
        if (!map.has(oid)) map.set(oid, new Set());
        map.get(oid)!.add(String(c.id));
      }
    });
    (contactOwnerLinks.data ?? []).forEach((link) => {
      const oid = link.fields.OwnerLookupId ? String(link.fields.OwnerLookupId) : null;
      const cid = link.fields.ContactLookupId ? String(link.fields.ContactLookupId) : null;
      if (oid && cid) {
        if (!map.has(oid)) map.set(oid, new Set());
        map.get(oid)!.add(cid);
      }
    });
    return map;
  }, [contacts.data, contactOwnerLinks.data]);

  const grouped = useMemo(() => {
    const map = new Map<ReportCategory, ReportDescriptor[]>();
    REPORTS.forEach((r) => {
      const list = map.get(r.category) ?? [];
      list.push(r);
      map.set(r.category, list);
    });
    return map;
  }, []);

  const allDocs = useMemo(() => {
    const docs: { library: string; doc: DocItemRaw }[] = [];
    libraries.forEach((lib, idx) => {
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      (lib.data ?? []).forEach((d) => docs.push({ library: libraryName, doc: d }));
    });
    return docs;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data]);

  // Truncation warning — every fetched list is capped at 500. If any list returns
  // exactly that, larger reports may be incomplete. Surface a banner so the user knows.
  const truncatedLists = useMemo(() => {
    const hits: string[] = [];
    const check = (name: string, data: unknown[] | null | undefined, cap: number) => {
      if (data && data.length >= cap) hits.push(name);
    };
    check('Properties', properties.data, 500);
    check('Submittals', submittals.data, 500);
    check('Outstanding', outstanding.data, 500);
    check('Compliance Deadlines', compliance.data, 500);
    check('Owners', owners.data, 500);
    check('Ownership Structure', ownership.data, 500);
    check('Correspondence', correspondence.data, 500);
    check('Communications', comms.data, 500);
    check('Contacts', contacts.data, 500);
    check('Contact Owner Links', contactOwnerLinks.data, 2000);
    libraries.forEach((lib, idx) => check(PROPERTY_LINKED_LIBRARIES[idx], lib.data, 500));
    return hits;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    properties.data, submittals.data, outstanding.data, compliance.data,
    owners.data, ownership.data, correspondence.data, comms.data,
    contacts.data, contactOwnerLinks.data,
    lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data,
  ]);

  const coreLoading =
    properties.loading || submittals.loading || outstanding.loading ||
    compliance.loading || owners.loading || ownership.loading ||
    correspondence.loading || comms.loading || contacts.loading;
  const docsLoading = libraries.some((l) => l.loading);

  // ---------------------------------------------------------------------------
  // Execute a configured report. Receives the RunOptions from the modal.
  // ---------------------------------------------------------------------------

  const handleRun = async (descriptor: ReportDescriptor, options: RunOptions) => {
    const ctx: ReportContext = {
      properties: properties.data ?? [],
      submittals: submittals.data ?? [],
      outstanding: outstanding.data ?? [],
      compliance: compliance.data ?? [],
      owners: owners.data ?? [],
      ownership: ownership.data ?? [],
      correspondence: correspondence.data ?? [],
      comms: comms.data ?? [],
      contacts: contacts.data ?? [],
      allDocs,
      propertiesById,
      ownersById,
      contactsById,
      contactsByOwnerId,
    };

    try {
      const result = await dispatchReport(descriptor, ctx, options);
      triggerDownload(result.blob, result.filename);

      // Audit-log audit-pack exports — DOR-defensive trail.
      if (
        descriptor.id === 'property-audit-pack' ||
        descriptor.id === 'portfolio-audit-pack' ||
        descriptor.id === 'org-chart-history'
      ) {
        void auditLogExport(descriptor, result, options);
      }

      setRunResult({ kind: 'success', descriptor, result, options });
    } catch (err) {
      setRunResult({
        kind: 'error',
        descriptor,
        message: err instanceof Error ? err.message : String(err),
      });
      // Re-throw so the modal stays open on the error path
      throw err;
    }
  };

  // ---------------------------------------------------------------------------
  // Send via Email handoff — convert the just-generated blob to a Graph
  // attachment, resolve any contacts linked to the scope, open the compose
  // modal pre-populated.
  // ---------------------------------------------------------------------------

  const handleSendViaEmail = async () => {
    if (runResult?.kind !== 'success') return;
    const { result } = runResult;
    const contentBase64 = await blobToBase64(result.blob);
    setComposeAttachment({
      filename: result.filename,
      contentType: result.contentType,
      contentBase64,
    });
    setComposeOpen(true);
  };

  // Default contact / property pre-selection for the compose modal
  const composeDefaults = useMemo(() => {
    if (runResult?.kind !== 'success') return null;
    const { descriptor, options, result } = runResult;
    const propertyIds: string[] = [];
    const contactIds = new Set<string>();

    if (options.scope.kind === 'property') {
      propertyIds.push(options.scope.propertyId);
      const p = propertiesById.get(options.scope.propertyId);
      const cid = p?.fields.PropertyOwnerContactLookupId;
      if (cid) contactIds.add(String(cid));
    } else if (options.scope.kind === 'owner') {
      const linked = contactsByOwnerId.get(options.scope.ownerId);
      linked?.forEach((cid) => contactIds.add(cid));
    }

    const scopeName = result.scopeLabel ?? 'portfolio';
    return {
      defaultPropertyIds: propertyIds,
      defaultContactIds: Array.from(contactIds),
      defaultSubject: `${descriptor.name} — ${scopeName}`,
      defaultBody:
        `Hi,\n\nAttaching the ${descriptor.name.toLowerCase()} as of ${new Date().toLocaleDateString()}.\n\nLet me know if you have any questions.\n\nThanks,\n{{user}}`,
    };
  }, [runResult, propertiesById, contactsByOwnerId]);

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Reports</h1>
        <p className="text-sm text-gray-500 mt-1">
          Pre-built reports. Click <strong>Configure & Run</strong> to choose scope, date range, and format —
          then download or send the file straight to an owner.
        </p>
      </div>

      {coreLoading && (
        <div className="mb-4 bg-blue-50 border border-blue-200 rounded-md p-3 text-sm text-blue-900">
          Loading data… reports won't run until this completes.
        </div>
      )}

      {!coreLoading && docsLoading && (
        <div className="mb-4 bg-blue-50 border border-blue-200 rounded-md p-3 text-sm text-blue-900">
          Loading document libraries… document-based reports (Expiration Calendar, Untagged Documents,
          Library Snapshot) may run with incomplete data until this finishes.
        </div>
      )}

      {truncatedLists.length > 0 && (
        <div className="mb-4 bg-amber-50 border border-amber-200 rounded-md p-3 text-xs text-amber-900">
          <strong>Heads up:</strong> these lists hit the 500-row fetch cap and may be incomplete in
          backup / portfolio reports: {truncatedLists.join(', ')}. Run targeted scopes (single property
          or single owner) for full coverage.
        </div>
      )}

      {runResult && (
        <RunResultBanner
          runResult={runResult}
          onSend={handleSendViaEmail}
          onDismiss={() => setRunResult(null)}
        />
      )}

      <div className="space-y-6">
        {Array.from(grouped.entries()).map(([category, reports]) => (
          <CategorySection
            key={category}
            category={category}
            reports={reports}
            onRun={(d) => {
              setRunResult(null);
              if (d.id === 'safe-harbor-certification') {
                setSafeHarborOpen(true);
                return;
              }
              setOpenDescriptor(d);
            }}
          />
        ))}
      </div>

      {safeHarborOpen && <SafeHarborCertModal onClose={() => setSafeHarborOpen(false)} />}

      {openDescriptor && (
        <RunReportModal
          descriptor={openDescriptor}
          properties={properties.data ?? []}
          owners={owners.data ?? []}
          onClose={() => setOpenDescriptor(null)}
          onRun={(opts) => handleRun(openDescriptor, opts)}
        />
      )}

      {composeOpen && composeAttachment && composeDefaults && (
        <ComposeEmailModal
          onClose={() => {
            setComposeOpen(false);
            setComposeAttachment(null);
          }}
          onSuccess={() => {
            setComposeOpen(false);
            setComposeAttachment(null);
          }}
          defaultPropertyIds={composeDefaults.defaultPropertyIds}
          defaultContactIds={composeDefaults.defaultContactIds}
          defaultSubject={composeDefaults.defaultSubject}
          defaultBody={composeDefaults.defaultBody}
          defaultAttachments={[composeAttachment]}
        />
      )}
    </div>
  );
}

// =============================================================================
// Run result banner
// =============================================================================

function RunResultBanner({
  runResult,
  onSend,
  onDismiss,
}: {
  runResult:
    | { kind: 'success'; descriptor: ReportDescriptor; result: RunResult; options: RunOptions }
    | { kind: 'error'; descriptor: ReportDescriptor; message: string };
  onSend: () => void;
  onDismiss: () => void;
}) {
  if (runResult.kind === 'error') {
    return (
      <div className="mb-4 bg-red-50 border border-red-200 rounded-md p-3 flex items-start gap-2">
        <Icon name="alert" size={16} className="text-error flex-shrink-0 mt-0.5" />
        <div className="flex-1">
          <p className="text-sm font-semibold text-red-900">
            {runResult.descriptor.name} failed
          </p>
          <p className="text-xs text-red-800 mt-0.5 font-mono-data">{runResult.message}</p>
        </div>
        <button
          onClick={onDismiss}
          className="text-error hover:bg-red-100 rounded p-0.5"
          aria-label="Dismiss"
        >
          <Icon name="x" size={14} />
        </button>
      </div>
    );
  }

  const { result } = runResult;
  return (
    <div className="mb-4 bg-green-50 border border-green-200 rounded-md p-3 flex items-start gap-2">
      <Icon name="check" size={16} className="text-success flex-shrink-0 mt-0.5" />
      <div className="flex-1 min-w-0">
        <p className="text-sm font-semibold text-green-900">
          Downloaded {result.rowCount} row{result.rowCount === 1 ? '' : 's'}
          {result.scopeLabel && <> — scope: {result.scopeLabel}</>}
        </p>
        <p className="text-xs text-green-800 mt-0.5 font-mono-data truncate">{result.filename}</p>
      </div>
      <button
        onClick={onSend}
        className="px-2.5 py-1 bg-teal-700 hover:bg-teal-900 text-white rounded text-xs font-medium flex items-center gap-1.5 flex-shrink-0"
      >
        <Icon name="mail" size={11} />
        Send via Email
      </button>
      <button
        onClick={onDismiss}
        className="text-success hover:bg-green-100 rounded p-0.5 flex-shrink-0"
        aria-label="Dismiss"
      >
        <Icon name="x" size={14} />
      </button>
    </div>
  );
}

// =============================================================================
// Category + card render
// =============================================================================

function CategorySection({
  category,
  reports,
  onRun,
}: {
  category: ReportCategory;
  reports: ReportDescriptor[];
  onRun: (d: ReportDescriptor) => void;
}) {
  return (
    <section>
      <div className="flex items-center gap-2 mb-3">
        <Icon name={CATEGORY_ICONS[category]} size={16} className="text-teal-700" />
        <h2 className="text-base font-semibold text-teal-700">{category}</h2>
        <span className="text-[10px] font-semibold text-gray-500 bg-gray-100 px-1.5 py-0.5 rounded-full font-mono-data">
          {reports.length}
        </span>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
        {reports.map((r) => (
          <ReportCard key={r.id} report={r} onRun={() => onRun(r)} />
        ))}
      </div>
    </section>
  );
}

function ReportCard({
  report,
  onRun,
}: {
  report: ReportDescriptor;
  onRun: () => void;
}) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card flex flex-col">
      <div className="flex items-start justify-between gap-2 mb-2">
        <h3 className="text-sm font-semibold text-gray-900">{report.name}</h3>
        <span
          className={`text-[9px] font-semibold px-1.5 py-0.5 rounded uppercase tracking-wider whitespace-nowrap border ${AUDIENCE_STYLES[report.audience]}`}
        >
          {report.audience === 'owner' ? 'Owner-facing' : 'Internal'}
        </span>
      </div>
      <p className="text-xs text-gray-600 mb-3 flex-1">{report.description}</p>
      <button
        onClick={onRun}
        className="w-full px-3 py-1.5 rounded-md text-xs font-medium flex items-center justify-center gap-1.5 bg-teal-700 hover:bg-teal-900 text-white transition-colors"
      >
        <Icon name="settings" size={11} />
        Configure & Run
      </button>
    </div>
  );
}

// =============================================================================
// Report Context — passed to every runner
// =============================================================================

interface ReportContext {
  properties: Property[];
  submittals: Submittal[];
  outstanding: OutstandingItem[];
  compliance: ComplianceDeadline[];
  owners: Owner[];
  ownership: Ownership[];
  correspondence: Correspondence[];
  comms: OwnerCommunication[];
  contacts: Contact[];
  allDocs: { library: string; doc: DocItemRaw }[];
  propertiesById: Map<string, Property>;
  ownersById: Map<string, Owner>;
  contactsById: Map<string, Contact>;
  contactsByOwnerId: Map<string, Set<string>>;
}

// =============================================================================
// Dispatcher — routes to the right runner by descriptor.id
// =============================================================================

async function dispatchReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  switch (descriptor.id) {
    case 'annual-filing-report':       return runAnnualFiling(descriptor, ctx, options);
    case 'compliance-status':          return runComplianceStatus(descriptor, ctx, options);
    case 'outstanding-items-by-assignee': return runOutstandingByAssignee(descriptor, ctx, options);
    case 'outstanding-items-by-owner-contact': return runOutstandingByOwnerContact(descriptor, ctx, options);
    case 'property-holdings':          return runPropertyHoldings(descriptor, ctx, options);
    case 'owner-quarterly-statement':  return runOwnerQuarterlyStatement(descriptor, ctx, options);
    case 'property-audit-pack':        return runPropertyAuditPack(descriptor, ctx, options);
    case 'portfolio-audit-pack':       return runPortfolioAuditPack(descriptor, ctx, options);
    case 'org-chart-history':          return runOrgChartHistory(descriptor, ctx, options);
    case 'dor-correspondence-log':     return runDORCorrespondenceLog(descriptor, ctx, options);
    case 'communications-history':     return runCommunicationsHistory(descriptor, ctx, options);
    case 'document-expiration-calendar': return runDocumentExpirationCalendar(descriptor, ctx, options);
    case 'untagged-documents-report':  return runUntaggedDocuments(descriptor, ctx, options);
    case 'full-database-export':       return runFullDatabaseExport(descriptor, ctx, options);
    case 'sharepoint-library-snapshot': return runSharePointLibrarySnapshot(descriptor, ctx, options);
    default:
      throw new Error(`No runner registered for report '${descriptor.id}'`);
  }
}

// =============================================================================
// Scope helpers
// =============================================================================

function scopeLabelFor(options: RunOptions, ctx: ReportContext): string {
  switch (options.scope.kind) {
    case 'portfolio': return 'Portfolio';
    case 'property': {
      const p = ctx.propertiesById.get(options.scope.propertyId);
      return p?.fields.Title ?? options.scope.propertyTitle ?? 'Property';
    }
    case 'owner': {
      const o = ctx.ownersById.get(options.scope.ownerId);
      return o?.fields.Title ?? options.scope.ownerTitle ?? 'Owner';
    }
    case 'state': return options.scope.state;
  }
}

/** Returns the property IDs in scope. Owner scope walks the ownership tree. */
function scopedPropertyIds(options: RunOptions, ctx: ReportContext): Set<string> {
  const ids = new Set<string>();
  if (options.scope.kind === 'portfolio') {
    ctx.properties.forEach((p) => ids.add(String(p.id)));
  } else if (options.scope.kind === 'property') {
    ids.add(options.scope.propertyId);
  } else if (options.scope.kind === 'state') {
    const target = options.scope.state;
    ctx.properties.forEach((p) => {
      if (p.fields.cahpState === target) ids.add(String(p.id));
    });
  } else if (options.scope.kind === 'owner') {
    const ownerId = options.scope.ownerId;
    // Direct ownership
    ctx.ownership.forEach((rel) => {
      if (String(rel.fields.OwnerLookupId) === ownerId && rel.fields.LinkedPropertyLookupId) {
        ids.add(String(rel.fields.LinkedPropertyLookupId));
      }
    });
    // Indirect via member chains — walk down from this owner through parent links
    const downstream = collectDownstreamOwnerIds(ownerId, ctx.ownership);
    downstream.forEach((descendantOwnerId) => {
      ctx.ownership.forEach((rel) => {
        if (
          String(rel.fields.OwnerLookupId) === descendantOwnerId &&
          rel.fields.LinkedPropertyLookupId
        ) {
          ids.add(String(rel.fields.LinkedPropertyLookupId));
        }
      });
    });
  }
  return ids;
}

/** Collect every owner that descends from a parent owner via ParentOwnerLookupId. */
function collectDownstreamOwnerIds(
  rootOwnerId: string,
  ownership: Ownership[],
): Set<string> {
  const out = new Set<string>();
  const queue = [rootOwnerId];
  while (queue.length > 0) {
    const cur = queue.shift()!;
    ownership.forEach((rel) => {
      if (
        String(rel.fields.ParentOwnerLookupId) === cur &&
        rel.fields.OwnerLookupId &&
        !out.has(String(rel.fields.OwnerLookupId))
      ) {
        out.add(String(rel.fields.OwnerLookupId));
        queue.push(String(rel.fields.OwnerLookupId));
      }
    });
  }
  return out;
}

// =============================================================================
// Audit logging
// =============================================================================

async function auditLogExport(
  descriptor: ReportDescriptor,
  result: RunResult,
  options: RunOptions,
): Promise<void> {
  try {
    await createListItem(LIST_NAMES.AuditLog, {
      Title: `Exported ${descriptor.name}: ${result.scopeLabel ?? 'Portfolio'}`,
      Action: 'CREATE',
      EntityType: 'Report Export',
      EntityId: descriptor.id,
      EntityTitle: result.filename,
      ChangeSummary:
        `Report: ${descriptor.name}\n` +
        `Scope: ${result.scopeLabel ?? 'Portfolio'} (${options.scope.kind})\n` +
        `Rows: ${result.rowCount}\n` +
        `Format: ${options.format}`,
    });
  } catch (err) {
    // Best-effort — don't bubble; user already has the file.
    // eslint-disable-next-line no-console
    console.warn('[Reports] Audit log write failed:', err);
  }
}

// =============================================================================
// Runners — Year-End / Annual
// =============================================================================

async function runAnnualFiling(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const year = options.taxYear ?? String(new Date().getFullYear());
  const scopeIds = scopedPropertyIds(options, ctx);
  const inScope = (id: string | undefined) => !id || scopeIds.has(String(id));

  const rows = ctx.submittals
    .filter((s) => {
      if (s.fields.cahpTaxYear) return s.fields.cahpTaxYear === year;
      const d = s.fields.DateFiled ?? s.createdDateTime;
      if (!d) return false;
      return new Date(d).getFullYear() === Number(year);
    })
    .filter((s) => inScope(s.fields.PropertyLookupId))
    .map((s) => {
      const propertyTitle = s.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? ''
        : '';
      return {
        Property: propertyTitle,
        'Submittal Title': s.fields.Title ?? '',
        'Tax Year': s.fields.cahpTaxYear ?? '',
        State: s.fields.cahpState ?? '',
        'Filing Type': s.fields.FilingType ?? '',
        'Filing Method': s.fields.FilingMethod ?? '',
        Status: s.fields.SubmittalStatus ?? '',
        'Date Filed': formatDateOnly(s.fields.DateFiled, ''),
        'Confirmation #': s.fields.ConfirmationNumber ?? '',
        'Approved Abatement': s.fields.ApprovedAbatement ?? '',
        'Next Action': s.fields.NextAction ?? '',
      };
    });

  return finalizeTabular(descriptor, options, rows, `Tax Year ${year}`, ctx);
}

async function runComplianceStatus(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const inScope = (id: string) => scopeIds.has(id);
  const now = Date.now();
  const atRiskWindow = now + 30 * 24 * 60 * 60 * 1000;

  const rows = ctx.properties
    .filter((p) => inScope(String(p.id)))
    .map((p) => {
      const propertyId = String(p.id);
      const myDeadlines = ctx.compliance.filter((c) => String(c.fields.PropertyLookupId) === propertyId);
      const open = myDeadlines.filter((c) => c.fields.DeadlineStatus !== 'Completed');
      const overdue = open.filter((c) => {
        if (!c.fields.DueDate) return false;
        return new Date(c.fields.DueDate).getTime() < now;
      });
      const atRisk = open.filter((c) => {
        if (!c.fields.DueDate) return false;
        const t = new Date(c.fields.DueDate).getTime();
        return t >= now && t <= atRiskWindow;
      });
      const nextDeadline = open
        .filter((c) => c.fields.DueDate)
        .sort((a, b) => new Date(a.fields.DueDate!).getTime() - new Date(b.fields.DueDate!).getTime())[0];

      let status: string;
      if (overdue.length > 0) status = 'OVERDUE';
      else if (atRisk.length > 0) status = 'AT RISK';
      else if (open.length > 0) status = 'On Track';
      else status = 'No deadlines';

      return {
        Property: p.fields.Title ?? '',
        State: p.fields.cahpState ?? '',
        'Property Status': p.fields.PropertyStatus ?? '',
        Units: p.fields.UnitCount ?? '',
        'AMI Program': p.fields.AMIProgram ?? '',
        'Open Deadlines': open.length,
        'Overdue Deadlines': overdue.length,
        'At Risk (≤30 days)': atRisk.length,
        'Next Deadline': nextDeadline?.fields.Title ?? '',
        'Next Deadline Due': nextDeadline?.fields.DueDate
          ? formatDateOnly(nextDeadline.fields.DueDate)
          : '',
        'Compliance Status': status,
      };
    });

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

// =============================================================================
// Runners — Operational
// =============================================================================

async function runOutstandingByAssignee(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const isClosed = (s: string | undefined) => s === 'Done' || s === 'Received' || s === 'Not Applicable';

  const rows = ctx.outstanding
    .filter((o) => !isClosed(o.fields.ItemStatus))
    .filter((o) => {
      if (options.scope.kind === 'portfolio') return true;
      return o.fields.PropertyLookupId && scopeIds.has(String(o.fields.PropertyLookupId));
    })
    .map((o) => {
      const property = o.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(o.fields.PropertyLookupId))
        : null;
      const overdue = o.fields.DueDate && new Date(o.fields.DueDate).getTime() < Date.now();
      const base: Record<string, unknown> = {
        'Assigned To': o.fields.AssignedTo ?? '(unassigned)',
        Item: o.fields.Title ?? '',
        Property: property?.fields.Title ?? '',
        Status: o.fields.ItemStatus ?? '',
        Priority: o.fields.Priority ?? '',
        Category: o.fields.ItemCategory ?? '',
        'Due Date': formatDateOnly(o.fields.DueDate, ''),
        Overdue: overdue ? 'YES' : '',
        'Date Requested': o.fields.DateRequested ? formatDateET(o.fields.DateRequested) : '',
        'Has Document': o.fields.RelatedDocUrl ? 'YES' : '',
      };
      if (options.includeInternalColumns ?? true) {
        base.Notes = o.fields.ItemNotes ?? '';
      }
      return base;
    })
    .sort((a, b) => {
      if (a['Assigned To'] !== b['Assigned To']) {
        return String(a['Assigned To']).localeCompare(String(b['Assigned To']));
      }
      if (a.Overdue !== b.Overdue) return String(b.Overdue).localeCompare(String(a.Overdue));
      return String(a['Due Date']).localeCompare(String(b['Due Date']));
    });

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

async function runOutstandingByOwnerContact(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const isClosed = (s: string | undefined) => s === 'Done' || s === 'Received' || s === 'Not Applicable';
  const includeInternal = options.includeInternalColumns ?? false;

  // Build rows keyed by (ownerContactName, property, item). Skip items whose
  // property has no PropertyOwnerContactLookupId — those are internal-only items.
  const rows = ctx.outstanding
    .filter((o) => !isClosed(o.fields.ItemStatus))
    .filter((o) => {
      if (options.scope.kind === 'portfolio') return true;
      return o.fields.PropertyLookupId && scopeIds.has(String(o.fields.PropertyLookupId));
    })
    .map((o) => {
      const property = o.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(o.fields.PropertyLookupId))
        : null;
      const contactId = property?.fields.PropertyOwnerContactLookupId
        ? String(property.fields.PropertyOwnerContactLookupId)
        : null;
      const contact = contactId ? ctx.contactsById.get(contactId) : null;
      if (!contact) return null;
      const overdue = o.fields.DueDate && new Date(o.fields.DueDate).getTime() < Date.now();
      const base: Record<string, unknown> = {
        'Owner Contact': contact.fields.Title ?? '',
        Email: contact.fields.ContactEmail ?? '',
        Property: property?.fields.Title ?? '',
        Item: o.fields.Title ?? '',
        Category: o.fields.ItemCategory ?? '',
        Priority: o.fields.Priority ?? '',
        'Due Date': formatDateOnly(o.fields.DueDate, ''),
        Overdue: overdue ? 'YES' : '',
      };
      if (includeInternal) {
        base.Status = o.fields.ItemStatus ?? '';
        base['Assigned To'] = o.fields.AssignedTo ?? '';
        base.Notes = o.fields.ItemNotes ?? '';
      }
      return base;
    })
    .filter((r): r is Record<string, unknown> => r !== null)
    .sort((a, b) => {
      const oc = String(a['Owner Contact']).localeCompare(String(b['Owner Contact']));
      if (oc !== 0) return oc;
      const pr = String(a.Property).localeCompare(String(b.Property));
      if (pr !== 0) return pr;
      return String(a['Due Date']).localeCompare(String(b['Due Date']));
    });

  if (options.format === 'pdf') {
    const scopeLabel = scopeLabelFor(options, ctx);
    const pdf = createBrandedPDF({
      title: 'Outstanding Items by Owner Contact',
      subtitle: scopeLabel,
      rightLabel: `${rows.length} open items`,
    });
    if (rows.length === 0) {
      pdf.paragraph('No outstanding items match this scope. Nothing currently owed.');
    } else {
      // Group by contact for the body
      const byContact = new Map<string, Record<string, unknown>[]>();
      rows.forEach((r) => {
        const key = String(r['Owner Contact'] || '(unknown)');
        if (!byContact.has(key)) byContact.set(key, []);
        byContact.get(key)!.push(r);
      });
      Array.from(byContact.entries()).sort((a, b) => a[0].localeCompare(b[0])).forEach(
        ([contactName, items]) => {
          pdf.heading(contactName);
          const headers = includeInternal
            ? ['Property', 'Item', 'Category', 'Due', 'Overdue', 'Status']
            : ['Property', 'Item', 'Category', 'Due', 'Overdue'];
          const tableRows = items.map((r) => {
            const cells = [
              String(r.Property ?? ''),
              String(r.Item ?? ''),
              String(r.Category ?? ''),
              String(r['Due Date'] ?? ''),
              String(r.Overdue ?? ''),
            ];
            if (includeInternal) cells.push(String(r.Status ?? ''));
            return cells;
          });
          pdf.table(headers, tableRows);
        },
      );
    }
    const built = buildPDF(pdf.raw());
    pdf.build(); // ensure footers
    return {
      rowCount: rows.length,
      filename: timestampedFilename(descriptor.filenameBase, 'pdf', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

async function runDocumentExpirationCalendar(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const windowDays = options.expirationWindow ?? 180;
  const scopeIds = scopedPropertyIds(options, ctx);
  const inScope = (id: string | undefined) =>
    options.scope.kind === 'portfolio' || (!!id && scopeIds.has(String(id)));

  const now = Date.now();
  const cutoff = now + windowDays * 24 * 60 * 60 * 1000;

  const bucketOf = (daysUntil: number): string => {
    if (daysUntil < 0) return 'Overdue';
    if (daysUntil <= 30) return '0-30 days';
    if (daysUntil <= 90) return '31-90 days';
    if (daysUntil <= 180) return '91-180 days';
    return '181-365 days';
  };

  const rows: Record<string, unknown>[] = [];

  ctx.compliance.forEach((c) => {
    if (!c.fields.DueDate) return;
    if (c.fields.DeadlineStatus === 'Completed') return;
    if (!inScope(c.fields.PropertyLookupId)) return;
    const due = new Date(c.fields.DueDate).getTime();
    if (due > cutoff) return;
    const property = c.fields.PropertyLookupId
      ? ctx.propertiesById.get(String(c.fields.PropertyLookupId))
      : null;
    const daysUntil = Math.floor((due - now) / (24 * 60 * 60 * 1000));
    rows.push({
      Bucket: bucketOf(daysUntil),
      Source: 'Compliance Deadline',
      Document: c.fields.Title ?? '',
      'Document Type': c.fields.DeadlineType ?? '',
      Property: property?.fields.Title ?? '(portfolio-wide)',
      'Expires / Due': formatDateOnly(c.fields.DueDate),
      'Days Until': daysUntil,
      Status: c.fields.DeadlineStatus ?? '',
      'Responsible Party': c.fields.ResponsibleParty ?? '',
    });
  });

  ctx.allDocs.forEach(({ library, doc }) => {
    if (!doc.fields.ExpirationDate) return;
    if (!inScope(doc.fields.PropertyLookupId)) return;
    const exp = new Date(doc.fields.ExpirationDate).getTime();
    if (isNaN(exp) || exp > cutoff) return;
    const property = doc.fields.PropertyLookupId
      ? ctx.propertiesById.get(String(doc.fields.PropertyLookupId))
      : null;
    const daysUntil = Math.floor((exp - now) / (24 * 60 * 60 * 1000));
    rows.push({
      Bucket: bucketOf(daysUntil),
      Source: 'Document',
      Document: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      'Document Type': library,
      Property: property?.fields.Title ?? '(unscoped)',
      'Expires / Due': formatDateET(doc.fields.ExpirationDate),
      'Days Until': daysUntil,
      Status: '',
      'Responsible Party': '',
    });
  });

  const bucketOrder = ['Overdue', '0-30 days', '31-90 days', '91-180 days', '181-365 days'];
  rows.sort((a, b) => {
    const ba = bucketOrder.indexOf(String(a.Bucket));
    const bb = bucketOrder.indexOf(String(b.Bucket));
    if (ba !== bb) return ba - bb;
    return Number(a['Days Until']) - Number(b['Days Until']);
  });

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

async function runUntaggedDocuments(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const rows = ctx.allDocs
    .filter(({ doc }) => !doc.fields.PropertyLookupId && !doc.fields.OwnerLookupId)
    .map(({ library, doc }) => ({
      Library: library,
      Filename: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      Uploaded: doc.fields.Created ? formatDateET(doc.fields.Created) : '',
      Modified: doc.fields.Modified ? formatDateET(doc.fields.Modified) : '',
      'Size (bytes)': doc.fields.File_x0020_Size ?? '',
      URL: doc.webUrl ?? '',
    }))
    .sort((a, b) => a.Library.localeCompare(b.Library) || a.Filename.localeCompare(b.Filename));
  return finalizeTabular(descriptor, options, rows, 'Untagged', ctx);
}

async function runCommunicationsHistory(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const ownerId = options.scope.kind === 'owner' ? options.scope.ownerId : null;
  const includeInternal = options.includeInternalColumns ?? true;
  const fromTs = options.dateFrom ? parseDateOnly(options.dateFrom)?.getTime() ?? -Infinity : -Infinity;
  const toTs = options.dateTo ? parseDateOnly(options.dateTo)?.getTime() ?? Infinity : Infinity;

  const rows = ctx.comms
    .filter((c) => {
      const t = c.fields.CommDate ? new Date(c.fields.CommDate).getTime() : null;
      if (t == null) return false;
      if (t < fromTs || t > toTs) return false;
      if (options.scope.kind === 'property' || options.scope.kind === 'state') {
        return c.fields.CommPropertyLookupId
          ? scopeIds.has(String(c.fields.CommPropertyLookupId))
          : false;
      }
      if (ownerId) {
        return String(c.fields.CommOwnerLookupId) === ownerId;
      }
      return true;
    })
    .map((c) => {
      const property = c.fields.CommPropertyLookupId
        ? ctx.propertiesById.get(String(c.fields.CommPropertyLookupId))
        : null;
      const owner = c.fields.CommOwnerLookupId
        ? ctx.ownersById.get(String(c.fields.CommOwnerLookupId))
        : null;
      const base: Record<string, unknown> = {
        Date: c.fields.CommDate ? formatDateET(c.fields.CommDate) : '',
        Type: c.fields.CommType ?? '',
        Direction: c.fields.CommDirection ?? '',
        Subject: c.fields.Title ?? '',
        Property: property?.fields.Title ?? '',
        Owner: owner?.fields.Title ?? '',
        Participants: c.fields.CommParticipants ?? '',
        Status: c.fields.CommStatus ?? '',
        'Response Due': c.fields.CommResponseDue ? formatDateOnly(c.fields.CommResponseDue) : '',
      };
      if (includeInternal) base.Notes = c.fields.CommNotes ?? '';
      return base;
    })
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date)));

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

// =============================================================================
// Runners — Owner Reports
// =============================================================================

async function runPropertyHoldings(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const includeInternal = options.includeInternalColumns ?? false;
  const onlyOwnerId = options.scope.kind === 'owner' ? options.scope.ownerId : null;

  const rows: Record<string, unknown>[] = [];
  ctx.properties
    .filter((p) => scopeIds.has(String(p.id)))
    .forEach((p) => {
      const propertyId = String(p.id);
      const ownerships = ctx.ownership.filter(
        (o) => String(o.fields.LinkedPropertyLookupId) === propertyId,
      );
      // Optionally further filter to one specific owner when owner-scoped
      const filtered = onlyOwnerId
        ? ownerships.filter((o) => String(o.fields.OwnerLookupId) === onlyOwnerId)
        : ownerships;

      if (filtered.length === 0) {
        const row: Record<string, unknown> = {
          Property: p.fields.Title ?? '',
          State: p.fields.cahpState ?? '',
          'Property Status': p.fields.PropertyStatus ?? '',
          Owner: '(no matching ownership records)',
          'Owner Type': '',
          Role: '',
          'Ownership %': '',
          'Effective Date': '',
        };
        if (includeInternal) row['Entity Notes'] = '';
        rows.push(row);
      } else {
        filtered
          .sort((a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0))
          .forEach((rel) => {
            const owner = rel.fields.OwnerLookupId
              ? ctx.ownersById.get(String(rel.fields.OwnerLookupId))
              : null;
            const row: Record<string, unknown> = {
              Property: p.fields.Title ?? '',
              State: p.fields.cahpState ?? '',
              'Property Status': p.fields.PropertyStatus ?? '',
              Owner: owner?.fields.Title ?? rel.fields.Title ?? '',
              'Owner Type': owner?.fields.OwnerType ?? '',
              Role: rel.fields.RelationshipType ?? '',
              'Ownership %': rel.fields.OwnershipPercent ?? '',
              'Effective Date': rel.fields.EffectiveDate
                ? formatDateOnly(rel.fields.EffectiveDate)
                : '',
            };
            if (includeInternal) row['Entity Notes'] = rel.fields.EntityNotes ?? '';
            rows.push(row);
          });
      }
    });

  const scopeLabel = scopeLabelFor(options, ctx);

  if (options.format === 'pdf') {
    const pdf = createBrandedPDF({
      title: 'Property Holdings Statement',
      subtitle: scopeLabel,
      rightLabel: `${rows.length} rows`,
    });
    if (rows.length === 0) {
      pdf.paragraph('No holdings found for this scope.');
    } else {
      pdf.table(
        ['Property', 'State', 'Owner', 'Type', 'Role', '%', 'Effective'],
        rows.map((r) => [
          String(r.Property ?? ''),
          String(r.State ?? ''),
          String(r.Owner ?? ''),
          String(r['Owner Type'] ?? ''),
          String(r.Role ?? ''),
          String(r['Ownership %'] ?? ''),
          String(r['Effective Date'] ?? ''),
        ]),
      );
    }
    const built = buildPDF(pdf.raw());
    pdf.build();
    return {
      rowCount: rows.length,
      filename: timestampedFilename(descriptor.filenameBase, 'pdf', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }

  return finalizeTabular(descriptor, options, rows, scopeLabel, ctx);
}

async function runOwnerQuarterlyStatement(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  if (options.scope.kind !== 'owner') {
    throw new Error('Owner Quarterly Statement requires a single owner scope.');
  }
  const owner = ctx.ownersById.get(options.scope.ownerId);
  if (!owner) throw new Error('Owner not found.');

  const { year, q } = options.quarter ?? { year: new Date().getFullYear(), q: 1 };
  const qStartMonth = (q - 1) * 3;
  const qStart = new Date(year, qStartMonth, 1).getTime();
  const qEnd = new Date(year, qStartMonth + 3, 0, 23, 59, 59).getTime();
  const inQuarter = (iso: string | undefined) => {
    if (!iso) return false;
    const t = new Date(iso).getTime();
    return t >= qStart && t <= qEnd;
  };

  // Scope by all properties this owner has direct or downstream interest in
  const scopeIds = scopedPropertyIds(options, ctx);
  const scopedProperties = ctx.properties.filter((p) => scopeIds.has(String(p.id)));
  const ownerId = options.scope.ownerId;

  // Sections to render
  const holdings = ctx.ownership
    .filter((rel) => {
      const owns = String(rel.fields.OwnerLookupId) === ownerId;
      const inScopeProperty = rel.fields.LinkedPropertyLookupId
        ? scopeIds.has(String(rel.fields.LinkedPropertyLookupId))
        : false;
      return owns && inScopeProperty;
    })
    .map((rel) => {
      const p = rel.fields.LinkedPropertyLookupId
        ? ctx.propertiesById.get(String(rel.fields.LinkedPropertyLookupId))
        : null;
      return {
        property: p?.fields.Title ?? '',
        state: p?.fields.cahpState ?? '',
        units: p?.fields.UnitCount ?? '',
        role: rel.fields.RelationshipType ?? '',
        percent: rel.fields.OwnershipPercent ?? '',
      };
    });

  const filings = ctx.submittals
    .filter((s) => s.fields.PropertyLookupId && scopeIds.has(String(s.fields.PropertyLookupId)))
    .filter((s) => inQuarter(s.fields.DateFiled) || inQuarter(s.createdDateTime));

  const correspondenceItems = ctx.correspondence
    .filter((c) => c.fields.PropertyLookupId && scopeIds.has(String(c.fields.PropertyLookupId)))
    .filter((c) => inQuarter(c.fields.DateReceived) || inQuarter(c.fields.DateResponded));

  const ownerComms = ctx.comms
    .filter((c) => String(c.fields.CommOwnerLookupId) === ownerId)
    .filter((c) => inQuarter(c.fields.CommDate));

  const openItems = ctx.outstanding
    .filter((o) => {
      const closed = o.fields.ItemStatus === 'Done' || o.fields.ItemStatus === 'Received' || o.fields.ItemStatus === 'Not Applicable';
      return !closed;
    })
    .filter((o) => o.fields.PropertyLookupId && scopeIds.has(String(o.fields.PropertyLookupId)));

  const upcomingDeadlines = ctx.compliance
    .filter((c) => c.fields.PropertyLookupId && scopeIds.has(String(c.fields.PropertyLookupId)))
    .filter((c) => c.fields.DeadlineStatus !== 'Completed' && c.fields.DueDate)
    .sort((a, b) => new Date(a.fields.DueDate!).getTime() - new Date(b.fields.DueDate!).getTime())
    .slice(0, 12);

  const scopeLabel = owner.fields.Title ?? 'Owner';
  const periodLabel = `Q${q} ${year}`;

  const pdf = createBrandedPDF({
    title: 'Owner Quarterly Statement',
    subtitle: scopeLabel,
    rightLabel: periodLabel,
  });

  pdf.kv('Owner', owner.fields.Title ?? '');
  pdf.kv('Owner Type', owner.fields.OwnerType ?? '');
  pdf.kv('State', owner.fields.OwnerState ?? '');
  if (owner.fields.SponsorName) pdf.kv('Sponsor', owner.fields.SponsorName);
  pdf.kv('Properties in Scope', String(scopedProperties.length));
  pdf.kv('Period', periodLabel);

  pdf.heading('Holdings');
  if (holdings.length === 0) {
    pdf.paragraph('No active holdings on file for this owner.');
  } else {
    pdf.table(
      ['Property', 'State', 'Units', 'Role', '%'],
      holdings.map((h) => [String(h.property), String(h.state), String(h.units), String(h.role), String(h.percent)]),
    );
  }

  pdf.heading(`Filing Activity — ${periodLabel}`);
  if (filings.length === 0) {
    pdf.paragraph('No submittals filed in this period.');
  } else {
    pdf.table(
      ['Property', 'Submittal', 'Tax Year', 'Status', 'Date Filed'],
      filings.map((s) => [
        ctx.propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? '',
        s.fields.Title ?? '',
        s.fields.cahpTaxYear ?? '',
        s.fields.SubmittalStatus ?? '',
        formatDateOnly(s.fields.DateFiled, ''),
      ]),
    );
  }

  pdf.heading(`DOR Correspondence — ${periodLabel}`);
  if (correspondenceItems.length === 0) {
    pdf.paragraph('No DOR correspondence logged in this period.');
  } else {
    pdf.table(
      ['Property', 'Subject', 'Direction', 'Letter Type', 'Received', 'Responded'],
      correspondenceItems.map((c) => [
        ctx.propertiesById.get(String(c.fields.PropertyLookupId))?.fields.Title ?? '',
        c.fields.Title ?? '',
        c.fields.Direction ?? '',
        c.fields.LetterType ?? '',
        c.fields.DateReceived ? formatDateET(c.fields.DateReceived) : '',
        c.fields.DateResponded ? formatDateET(c.fields.DateResponded) : '',
      ]),
    );
  }

  pdf.heading(`Communications — ${periodLabel}`);
  if (ownerComms.length === 0) {
    pdf.paragraph('No owner communications logged in this period.');
  } else {
    pdf.table(
      ['Date', 'Type', 'Direction', 'Subject', 'Status'],
      ownerComms.map((c) => [
        c.fields.CommDate ? formatDateET(c.fields.CommDate) : '',
        c.fields.CommType ?? '',
        c.fields.CommDirection ?? '',
        c.fields.Title ?? '',
        c.fields.CommStatus ?? '',
      ]),
    );
  }

  pdf.heading('Open Items You Owe');
  if (openItems.length === 0) {
    pdf.paragraph('Nothing currently outstanding from your side. Thank you.');
  } else {
    pdf.table(
      ['Property', 'Item', 'Category', 'Priority', 'Due Date'],
      openItems.map((o) => [
        ctx.propertiesById.get(String(o.fields.PropertyLookupId ?? ''))?.fields.Title ?? '',
        o.fields.Title ?? '',
        o.fields.ItemCategory ?? '',
        o.fields.Priority ?? '',
        formatDateOnly(o.fields.DueDate, ''),
      ]),
    );
  }

  pdf.heading('Upcoming Deadlines (next 12)');
  if (upcomingDeadlines.length === 0) {
    pdf.paragraph('No upcoming compliance deadlines in scope.');
  } else {
    pdf.table(
      ['Property', 'Deadline', 'Type', 'Due Date'],
      upcomingDeadlines.map((c) => [
        ctx.propertiesById.get(String(c.fields.PropertyLookupId ?? ''))?.fields.Title ?? '(portfolio)',
        c.fields.Title ?? '',
        c.fields.DeadlineType ?? '',
        formatDateOnly(c.fields.DueDate, ''),
      ]),
    );
  }

  const built = buildPDF(pdf.raw());
  pdf.build();

  const totalRows =
    holdings.length + filings.length + correspondenceItems.length +
    ownerComms.length + openItems.length + upcomingDeadlines.length;

  return {
    rowCount: totalRows,
    filename: timestampedFilename(
      descriptor.filenameBase,
      'pdf',
      `${scopeLabel}-${periodLabel}`,
    ),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel: `${scopeLabel} · ${periodLabel}`,
  };
}

// =============================================================================
// Runners — DOR Audit Pack
// =============================================================================

async function runPropertyAuditPack(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  if (options.scope.kind !== 'property') {
    throw new Error('Per-Property Audit Pack requires a single property scope.');
  }
  const property = ctx.propertiesById.get(options.scope.propertyId);
  if (!property) throw new Error('Property not found.');
  return bundleAuditPack(descriptor, [property], ctx, options);
}

async function runPortfolioAuditPack(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const scoped = ctx.properties.filter((p) => scopeIds.has(String(p.id)));
  return bundleAuditPack(descriptor, scoped, ctx, options);
}

async function bundleAuditPack(
  descriptor: ReportDescriptor,
  scopedProperties: Property[],
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopedIds = new Set(scopedProperties.map((p) => String(p.id)));
  const inScope = (id: string | undefined) => (id ? scopedIds.has(String(id)) : false);

  const propertiesSheet = scopedProperties.map((p) => ({
    'Property ID': p.id,
    Title: p.fields.Title ?? '',
    'Legal Entity': p.fields.LegalEntity ?? '',
    State: p.fields.cahpState ?? '',
    County: p.fields.cahpCounty ?? '',
    Address: p.fields.PropertyAddress ?? '',
    Units: p.fields.UnitCount ?? '',
    'AMI Program': p.fields.AMIProgram ?? '',
    'Property Status': p.fields.PropertyStatus ?? '',
    'CAHP Language Added': p.fields.CAHPLanguageAdded ?? '',
    'LURA Executed': p.fields.LURAExecuted ?? '',
    'Verification Status': p.fields.cahpVerificationStatus ?? '',
  }));

  const submittalsSheet = ctx.submittals
    .filter((s) => inScope(s.fields.PropertyLookupId))
    .map((s) => ({
      Property: ctx.propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? '',
      Title: s.fields.Title ?? '',
      'Tax Year': s.fields.cahpTaxYear ?? '',
      State: s.fields.cahpState ?? '',
      'Filing Type': s.fields.FilingType ?? '',
      Status: s.fields.SubmittalStatus ?? '',
      'Date Filed': formatDateOnly(s.fields.DateFiled, ''),
      'Confirmation #': s.fields.ConfirmationNumber ?? '',
      'Approved Abatement': s.fields.ApprovedAbatement ?? '',
    }));

  const ownershipSheet = ctx.ownership
    .filter((o) => inScope(o.fields.LinkedPropertyLookupId))
    .map((rel) => {
      const owner = rel.fields.OwnerLookupId
        ? ctx.ownersById.get(String(rel.fields.OwnerLookupId))
        : null;
      return {
        Property: ctx.propertiesById.get(String(rel.fields.LinkedPropertyLookupId))?.fields.Title ?? '',
        Owner: owner?.fields.Title ?? rel.fields.Title ?? '',
        'Owner Type': owner?.fields.OwnerType ?? '',
        Role: rel.fields.RelationshipType ?? '',
        'Ownership %': rel.fields.OwnershipPercent ?? '',
        'Effective Date': rel.fields.EffectiveDate ? formatDateOnly(rel.fields.EffectiveDate) : '',
      };
    });

  const complianceSheet = ctx.compliance
    .filter((c) => !c.fields.PropertyLookupId || inScope(c.fields.PropertyLookupId))
    .map((c) => ({
      Property: c.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(c.fields.PropertyLookupId))?.fields.Title ?? ''
        : '(portfolio-wide)',
      Deadline: c.fields.Title ?? '',
      Type: c.fields.DeadlineType ?? '',
      Status: c.fields.DeadlineStatus ?? '',
      'Due Date': c.fields.DueDate ? formatDateOnly(c.fields.DueDate) : '',
      'Responsible Party': c.fields.ResponsibleParty ?? '',
    }));

  const correspondenceSheet = ctx.correspondence
    .filter((c) => inScope(c.fields.PropertyLookupId))
    .map((c) => ({
      Property: ctx.propertiesById.get(String(c.fields.PropertyLookupId))?.fields.Title ?? '',
      Subject: c.fields.Title ?? '',
      Direction: c.fields.Direction ?? '',
      'Letter Type': c.fields.LetterType ?? '',
      'Date Received': c.fields.DateReceived ? formatDateET(c.fields.DateReceived) : '',
      'Date Responded': c.fields.DateResponded ? formatDateET(c.fields.DateResponded) : '',
      'Response Due': c.fields.ResponseDue ? formatDateET(c.fields.ResponseDue) : '',
    }));

  const outstandingSheet = ctx.outstanding
    .filter((o) => inScope(o.fields.PropertyLookupId))
    .map((o) => ({
      Property: ctx.propertiesById.get(String(o.fields.PropertyLookupId))?.fields.Title ?? '',
      Item: o.fields.Title ?? '',
      Category: o.fields.ItemCategory ?? '',
      Status: o.fields.ItemStatus ?? '',
      Priority: o.fields.Priority ?? '',
      'Due Date': formatDateOnly(o.fields.DueDate, ''),
      'Has Document': o.fields.RelatedDocUrl ? 'YES' : '',
      'Document Filename': o.fields.RelatedDocFilename ?? '',
    }));

  const documentsSheet = ctx.allDocs
    .filter(({ doc }) => inScope(doc.fields.PropertyLookupId))
    .map(({ library, doc }) => ({
      Property: ctx.propertiesById.get(String(doc.fields.PropertyLookupId))?.fields.Title ?? '',
      Library: library,
      Filename: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      Uploaded: doc.fields.Created ? formatDateET(doc.fields.Created) : '',
      URL: doc.webUrl ?? '',
    }));

  const orgChartsSheet = ctx.submittals
    .filter((s) => inScope(s.fields.PropertyLookupId))
    .filter((s) => s.fields.OrgChartSnapshotJSON || s.fields.OrgChartSnapshotDate)
    .map((s) => ({
      Property: ctx.propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? '',
      Submittal: s.fields.Title ?? '',
      'Snapshot Date': s.fields.OrgChartSnapshotDate ? formatDateET(s.fields.OrgChartSnapshotDate) : '',
      'Snapshot Size (chars)': s.fields.OrgChartSnapshotJSON?.length ?? 0,
    }));

  const sheets: Record<string, Record<string, unknown>[]> = {
    Properties: propertiesSheet,
    Submittals: submittalsSheet,
    Ownership: ownershipSheet,
    Compliance: complianceSheet,
    Correspondence: correspondenceSheet,
    Outstanding: outstandingSheet,
    Documents: documentsSheet,
    'Org Chart Snapshots': orgChartsSheet,
  };

  const built = await buildXLSX(sheets);
  const scopeLabel = scopeLabelFor(options, ctx);
  const total = Object.values(sheets).reduce((sum, rows) => sum + rows.length, 0);

  return {
    rowCount: total,
    filename: timestampedFilename(descriptor.filenameBase, 'xlsx', scopeLabel),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel,
  };
}

async function runOrgChartHistory(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  if (options.scope.kind !== 'property') {
    throw new Error('Org Chart History requires a single property scope.');
  }
  const propertyId = options.scope.propertyId;
  const property = ctx.propertiesById.get(propertyId);
  if (!property) throw new Error('Property not found.');

  const snapshots = ctx.submittals
    .filter((s) => String(s.fields.PropertyLookupId) === propertyId)
    .filter((s) => s.fields.OrgChartSnapshotJSON || s.fields.OrgChartSnapshotDate)
    .sort((a, b) => {
      const aT = a.fields.OrgChartSnapshotDate ?? a.fields.DateFiled ?? a.createdDateTime;
      const bT = b.fields.OrgChartSnapshotDate ?? b.fields.DateFiled ?? b.createdDateTime;
      return new Date(aT).getTime() - new Date(bT).getTime();
    });

  const scopeLabel = property.fields.Title ?? 'Property';

  if (options.format === 'json') {
    const bundle = {
      property: {
        id: property.id,
        title: property.fields.Title,
        legalEntity: property.fields.LegalEntity,
        state: property.fields.cahpState,
      },
      generatedAt: new Date().toISOString(),
      snapshotCount: snapshots.length,
      snapshots: snapshots.map((s) => {
        let parsed: unknown = null;
        try {
          if (s.fields.OrgChartSnapshotJSON) parsed = JSON.parse(s.fields.OrgChartSnapshotJSON);
        } catch {
          parsed = s.fields.OrgChartSnapshotJSON;
        }
        return {
          submittalId: s.id,
          submittalTitle: s.fields.Title,
          submittalStatus: s.fields.SubmittalStatus,
          taxYear: s.fields.cahpTaxYear,
          dateFiled: s.fields.DateFiled,
          snapshotDate: s.fields.OrgChartSnapshotDate,
          snapshot: parsed,
        };
      }),
    };
    const built = buildJSON(bundle);
    return {
      rowCount: snapshots.length,
      filename: timestampedFilename(descriptor.filenameBase, 'json', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }

  // PDF — one page per snapshot, each a textual ownership tree summary.
  // Vector chart rendering would require the full OwnershipNode tree the
  // engine builds on PropertyDetail; the per-submittal snapshot is the
  // frozen JSON, which we surface as a labeled tree here.
  const pdf = createBrandedPDF({
    title: 'Org Chart History',
    subtitle: scopeLabel,
    rightLabel: `${snapshots.length} snapshot${snapshots.length === 1 ? '' : 's'}`,
  });

  if (snapshots.length === 0) {
    pdf.paragraph('No org chart snapshots have been frozen for this property.');
  } else {
    snapshots.forEach((s, idx) => {
      if (idx > 0) pdf.newPage();
      pdf.heading(s.fields.Title ?? `Snapshot ${idx + 1}`);
      pdf.kv('Tax Year', s.fields.cahpTaxYear ?? '');
      pdf.kv('Status', s.fields.SubmittalStatus ?? '');
      pdf.kv('Date Filed', formatDateOnly(s.fields.DateFiled, ''));
      pdf.kv('Snapshot Date', formatDateOnly(s.fields.OrgChartSnapshotDate, ''));
      pdf.kv('Confirmation #', s.fields.ConfirmationNumber ?? '');

      pdf.heading('Frozen Ownership Tree');
      const raw = s.fields.OrgChartSnapshotJSON;
      if (!raw) {
        pdf.paragraph('(No snapshot JSON captured for this submittal.)');
      } else {
        try {
          const parsed = JSON.parse(raw) as unknown;
          const lines = formatTreeAsText(parsed);
          lines.forEach((line) => pdf.paragraph(line));
        } catch {
          pdf.paragraph(raw.slice(0, 2000));
        }
      }
    });
  }

  const built = buildPDF(pdf.raw());
  pdf.build();
  return {
    rowCount: snapshots.length,
    filename: timestampedFilename(descriptor.filenameBase, 'pdf', scopeLabel),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel,
  };
}

/** Render a frozen org-chart JSON tree as indented text lines. */
function formatTreeAsText(node: unknown, depth = 0): string[] {
  if (!node || typeof node !== 'object') return [];
  const lines: string[] = [];
  const indent = '  '.repeat(depth);
  const obj = node as Record<string, unknown>;
  const title = (obj.title ?? obj.name ?? obj.Title ?? '(unnamed)') as string;
  const meta: string[] = [];
  if (obj.ownerType ?? obj.OwnerType) meta.push(String(obj.ownerType ?? obj.OwnerType));
  if (obj.role ?? obj.relationshipType ?? obj.RelationshipType) {
    meta.push(String(obj.role ?? obj.relationshipType ?? obj.RelationshipType));
  }
  if (obj.percent != null || obj.ownershipPercent != null || obj.OwnershipPercent != null) {
    meta.push(`${obj.percent ?? obj.ownershipPercent ?? obj.OwnershipPercent}%`);
  }
  lines.push(`${indent}• ${title}${meta.length ? ` — ${meta.join(' / ')}` : ''}`);

  const children = (obj.children ?? obj.members ?? obj.parents) as unknown;
  if (Array.isArray(children)) {
    children.forEach((c) => lines.push(...formatTreeAsText(c, depth + 1)));
  }
  return lines;
}

async function runDORCorrespondenceLog(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  options: RunOptions,
): Promise<RunResult> {
  const scopeIds = scopedPropertyIds(options, ctx);
  const includeInternal = options.includeInternalColumns ?? true;
  const fromTs = options.dateFrom ? parseDateOnly(options.dateFrom)?.getTime() ?? -Infinity : -Infinity;
  const toTs = options.dateTo ? parseDateOnly(options.dateTo)?.getTime() ?? Infinity : Infinity;

  const rows = ctx.correspondence
    .filter((c) => {
      if (options.scope.kind !== 'portfolio') {
        if (!c.fields.PropertyLookupId) return false;
        if (!scopeIds.has(String(c.fields.PropertyLookupId))) return false;
      }
      const dt = c.fields.DateReceived ?? c.fields.DateResponded ?? c.createdDateTime;
      const t = dt ? new Date(dt).getTime() : null;
      if (t == null) return true;
      return t >= fromTs && t <= toTs;
    })
    .map((c) => {
      const property = c.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(c.fields.PropertyLookupId))
        : null;
      const base: Record<string, unknown> = {
        Property: property?.fields.Title ?? '',
        State: c.fields.cahpState ?? '',
        Subject: c.fields.Title ?? '',
        Direction: c.fields.Direction ?? '',
        Channel: c.fields.CorrChannel ?? 'Letter',
        'Letter Type': c.fields.LetterType ?? '',
        'Date Received': c.fields.DateReceived ? formatDateET(c.fields.DateReceived) : '',
        'Date Responded': c.fields.DateResponded ? formatDateET(c.fields.DateResponded) : '',
        'Response Due': c.fields.ResponseDue ? formatDateET(c.fields.ResponseDue) : '',
        'Tax Year': c.fields.cahpTaxYear ?? '',
      };
      if (includeInternal) {
        base['Request Summary'] = c.fields.RequestSummary ?? '';
        base['Response Notes'] = c.fields.ResponseNotes ?? '';
      }
      return base;
    })
    .sort((a, b) => String(b['Date Received']).localeCompare(String(a['Date Received'])));

  return finalizeTabular(descriptor, options, rows, scopeLabelFor(options, ctx), ctx);
}

// =============================================================================
// Runners — Backup and Export
// =============================================================================

async function runFullDatabaseExport(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  _options: RunOptions,
): Promise<RunResult> {
  const sheets: Record<string, Record<string, unknown>[]> = {
    Properties: ctx.properties.map((p) => ({ id: p.id, ...p.fields })),
    Submittals: ctx.submittals.map((s) => ({ id: s.id, ...s.fields })),
    Outstanding: ctx.outstanding.map((o) => ({ id: o.id, ...o.fields })),
    Compliance: ctx.compliance.map((c) => ({ id: c.id, ...c.fields })),
    Owners: ctx.owners.map((o) => ({ id: o.id, ...o.fields })),
    Ownership: ctx.ownership.map((o) => ({ id: o.id, ...o.fields })),
    Correspondence: ctx.correspondence.map((c) => ({ id: c.id, ...c.fields })),
    Communications: ctx.comms.map((c) => ({ id: c.id, ...c.fields })),
  };
  const built = await buildXLSX(sheets);
  const total = Object.values(sheets).reduce((sum, rows) => sum + rows.length, 0);
  return {
    rowCount: total,
    filename: timestampedFilename(descriptor.filenameBase, 'xlsx'),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel: 'Full database',
  };
}

async function runSharePointLibrarySnapshot(
  descriptor: ReportDescriptor,
  ctx: ReportContext,
  _options: RunOptions,
): Promise<RunResult> {
  const snapshot = {
    generatedAt: new Date().toISOString(),
    libraryCount: 8,
    totalDocuments: ctx.allDocs.length,
    libraries: ctx.allDocs.reduce((acc, { library, doc }) => {
      if (!acc[library]) acc[library] = [];
      acc[library].push({
        id: doc.id,
        filename: doc.fields.FileLeafRef ?? doc.fields.Title,
        webUrl: doc.webUrl,
        propertyId: doc.fields.PropertyLookupId,
        ownerId: doc.fields.OwnerLookupId,
        modified: doc.fields.Modified,
        created: doc.fields.Created,
        size: doc.fields.File_x0020_Size,
      });
      return acc;
    }, {} as Record<string, unknown[]>),
  };
  const built = buildJSON(snapshot);
  return {
    rowCount: ctx.allDocs.length,
    filename: timestampedFilename(descriptor.filenameBase, 'json'),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel: 'All libraries',
  };
}

// =============================================================================
// Tabular finalizer — converts rows to CSV / XLSX / JSON / PDF (single table)
// =============================================================================

async function finalizeTabular(
  descriptor: ReportDescriptor,
  options: RunOptions,
  rows: Record<string, unknown>[],
  scopeLabel: string,
  ctx: ReportContext,
): Promise<RunResult> {
  // Suppress unused-arg lint when a runner doesn't need ctx in finalize
  void ctx;

  if (options.format === 'xlsx') {
    const built = await buildXLSX({ [descriptor.name.slice(0, 31)]: rows });
    return {
      rowCount: rows.length,
      filename: timestampedFilename(descriptor.filenameBase, 'xlsx', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }
  if (options.format === 'json') {
    const built = buildJSON(rows);
    return {
      rowCount: rows.length,
      filename: timestampedFilename(descriptor.filenameBase, 'json', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }
  if (options.format === 'pdf') {
    const pdf = createBrandedPDF({
      title: descriptor.name,
      subtitle: scopeLabel,
      rightLabel: `${rows.length} rows`,
    });
    if (rows.length === 0) {
      pdf.paragraph('No data matches the chosen scope and filters.');
    } else {
      const headers = Object.keys(rows[0]);
      const tableRows = rows.map((r) => headers.map((h) => String(r[h] ?? '')));
      pdf.table(headers, tableRows);
    }
    const built = buildPDF(pdf.raw());
    pdf.build();
    return {
      rowCount: rows.length,
      filename: timestampedFilename(descriptor.filenameBase, 'pdf', scopeLabel),
      blob: built.blob,
      contentType: built.contentType,
      scopeLabel,
    };
  }
  // default — CSV
  const built = buildCSV(rows);
  return {
    rowCount: rows.length,
    filename: timestampedFilename(descriptor.filenameBase, 'csv', scopeLabel),
    blob: built.blob,
    contentType: built.contentType,
    scopeLabel,
  };
}
