import { useState, useMemo } from 'react';
import {
  useSharePointList,
  LIST_NAMES,
  type Property,
  type Submittal,
  type OutstandingItem,
  type ComplianceDeadline,
  type Owner,
  type Ownership,
  type Correspondence,
  type OwnerCommunication,
} from '../lib/sharepoint';
import { formatDateOnly, formatDateET } from '../lib/dates';
import {
  REPORTS,
  downloadCSV,
  downloadJSON,
  downloadXLSX,
  timestampedFilename,
  type ReportDescriptor,
  type ReportCategory,
  type ReportStatus,
} from '../lib/reports';
import { PROPERTY_LINKED_LIBRARIES } from '../components/UploadDocumentModal';
import { Icon } from '../components/ui/Icon';

const STATUS_LABEL: Record<ReportStatus, string> = {
  'available': 'Available',
  'pending-pr14b': 'Coming soon',
};

const STATUS_STYLES: Record<ReportStatus, string> = {
  'available': 'bg-success/10 text-success border-success/30',
  'pending-pr14b': 'bg-gold-100 text-gold-900 border-gold-200',
};

const CATEGORY_ICONS: Record<ReportCategory, 'star' | 'file' | 'calendar' | 'check' | 'folder'> = {
  'Owner Reports': 'star',
  'DOR Audit Pack': 'file',
  'Year-End / Annual': 'calendar',
  'Operational': 'check',
  'Backup and Export': 'folder',
};

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

  // Documents — 8 libraries, lazy-fetched. For large reports we'll fetch on click.
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });
  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];

  const [runningId, setRunningId] = useState<string | null>(null);
  const [runResult, setRunResult] = useState<{ id: string; success: boolean; message: string } | null>(null);

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const ownersById = useMemo(() => {
    if (!owners.data) return new Map<string, Owner>();
    return new Map(owners.data.map((o) => [String(o.id), o]));
  }, [owners.data]);

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

  const handleRun = async (descriptor: ReportDescriptor) => {
    if (descriptor.status !== 'available') return;
    setRunningId(descriptor.id);
    setRunResult(null);

    try {
      let rowCount = 0;
      const ctx = {
        properties: properties.data ?? [],
        submittals: submittals.data ?? [],
        outstanding: outstanding.data ?? [],
        compliance: compliance.data ?? [],
        owners: owners.data ?? [],
        ownership: ownership.data ?? [],
        correspondence: correspondence.data ?? [],
        comms: comms.data ?? [],
        allDocs,
        propertiesById,
        ownersById,
      };

      switch (descriptor.id) {
        case 'annual-filing-report':
          rowCount = await runAnnualFilingReport(descriptor, ctx); break;
        case 'compliance-status':
          rowCount = await runComplianceStatusReport(descriptor, ctx); break;
        case 'outstanding-items-by-owner':
          rowCount = await runOutstandingItemsByOwnerReport(descriptor, ctx); break;
        case 'property-holdings':
          rowCount = await runPropertyHoldingsReport(descriptor, ctx); break;
        case 'property-audit-pack':
          rowCount = await runPropertyAuditPack(descriptor, ctx); break;
        case 'portfolio-audit-pack':
          rowCount = await runPortfolioAuditPack(descriptor, ctx); break;
        case 'org-chart-history':
          rowCount = await runOrgChartHistoryReport(descriptor, ctx); break;
        case 'document-expiration-calendar':
          rowCount = await runDocumentExpirationCalendar(descriptor, ctx); break;
        case 'untagged-documents-report':
          rowCount = await runUntaggedDocumentsReport(descriptor, ctx); break;
        case 'full-database-export':
          rowCount = await runFullDatabaseExport(descriptor, ctx); break;
        case 'sharepoint-library-snapshot':
          rowCount = await runSharePointLibrarySnapshot(descriptor, ctx); break;
        default:
          throw new Error(`No runner registered for report '${descriptor.id}'`);
      }

      setRunResult({
        id: descriptor.id,
        success: true,
        message: `Downloaded ${rowCount} row${rowCount === 1 ? '' : 's'}.`,
      });
    } catch (err) {
      setRunResult({
        id: descriptor.id,
        success: false,
        message: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setRunningId(null);
    }
  };

  const loading = properties.loading || submittals.loading || outstanding.loading || compliance.loading;

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Reports</h1>
        <p className="text-sm text-gray-500 mt-1">
          Pre-built reports. Click Run Now to download as CSV.
        </p>
      </div>

      {loading && (
        <div className="mb-4 bg-blue-50 border border-blue-200 rounded-md p-3 text-sm text-blue-900">
          Loading data… reports that need data won't run until this completes.
        </div>
      )}

      {runResult && (
        <div
          className={`mb-4 rounded-md p-3 flex items-start gap-2 ${
            runResult.success ? 'bg-green-50 border border-green-200' : 'bg-red-50 border border-red-200'
          }`}
        >
          <Icon
            name={runResult.success ? 'check' : 'alert'}
            size={16}
            className={runResult.success ? 'text-success flex-shrink-0 mt-0.5' : 'text-error flex-shrink-0 mt-0.5'}
          />
          <p className={`text-sm ${runResult.success ? 'text-green-900' : 'text-red-900'}`}>
            {runResult.message}
          </p>
        </div>
      )}

      {/* Category sections */}
      <div className="space-y-6">
        {Array.from(grouped.entries()).map(([category, reports]) => (
          <CategorySection
            key={category}
            category={category}
            reports={reports}
            runningId={runningId}
            onRun={handleRun}
          />
        ))}
      </div>
    </div>
  );
}

function CategorySection({
  category,
  reports,
  runningId,
  onRun,
}: {
  category: ReportCategory;
  reports: ReportDescriptor[];
  runningId: string | null;
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
          <ReportCard key={r.id} report={r} running={runningId === r.id} onRun={() => onRun(r)} />
        ))}
      </div>
    </section>
  );
}

function ReportCard({
  report,
  running,
  onRun,
}: {
  report: ReportDescriptor;
  running: boolean;
  onRun: () => void;
}) {
  const isAvailable = report.status === 'available';

  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card flex flex-col">
      <div className="flex items-start justify-between gap-2 mb-2">
        <h3 className="text-sm font-semibold text-gray-900">{report.name}</h3>
        <span
          className={`text-[9px] font-semibold px-1.5 py-0.5 rounded uppercase tracking-wider whitespace-nowrap border ${STATUS_STYLES[report.status]}`}
        >
          {STATUS_LABEL[report.status]}
        </span>
      </div>
      <p className="text-xs text-gray-600 mb-3 flex-1">{report.description}</p>
      <div className="flex items-center gap-2 mt-auto">
        <button
          onClick={onRun}
          disabled={!isAvailable || running}
          className={`flex-1 px-3 py-1.5 rounded-md text-xs font-medium flex items-center justify-center gap-1.5 transition-colors ${
            isAvailable
              ? 'bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50'
              : 'bg-gray-100 text-gray-400 cursor-not-allowed'
          }`}
        >
          {running && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
          {running ? 'Running…' : 'Run Now'}
        </button>
        <button
          disabled
          title="Recurring scheduled delivery — deferred"
          className="px-3 py-1.5 rounded-md text-xs font-medium border border-gray-200 text-gray-400 cursor-not-allowed flex items-center gap-1"
        >
          <Icon name="calendar" size={11} />
          Schedule
        </button>
      </div>
    </div>
  );
}

// =============================================================================
// Report Runners — pure functions that take data + return row count after download
// =============================================================================

// =============================================================================
// Report Context — passed to every runner
// =============================================================================

interface DocItemRawForReport {
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

interface ReportContext {
  properties: Property[];
  submittals: Submittal[];
  outstanding: OutstandingItem[];
  compliance: ComplianceDeadline[];
  owners: Owner[];
  ownership: Ownership[];
  correspondence: Correspondence[];
  comms: OwnerCommunication[];
  allDocs: { library: string; doc: DocItemRawForReport }[];
  propertiesById: Map<string, Property>;
  ownersById: Map<string, Owner>;
}

// =============================================================================
// Runners — Year-End / Annual
// =============================================================================

async function runAnnualFilingReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  const thisYear = new Date().getFullYear();
  const rows = ctx.submittals
    .filter((s) => {
      const date = s.fields.DateFiled ?? s.createdDateTime;
      if (!date) return false;
      return new Date(date).getFullYear() === thisYear;
    })
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
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

async function runComplianceStatusReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  const rows = ctx.properties.map((p) => {
    const propertyId = String(p.id);
    const myDeadlines = ctx.compliance.filter((c) => String(c.fields.PropertyLookupId) === propertyId);
    const overdue = myDeadlines.filter((c) => {
      if (!c.fields.DueDate) return false;
      if (c.fields.DeadlineStatus === 'Completed') return false;
      return new Date(c.fields.DueDate).getTime() < Date.now();
    });
    const nextDeadline = myDeadlines
      .filter((c) => c.fields.DeadlineStatus !== 'Completed' && c.fields.DueDate)
      .sort((a, b) => new Date(a.fields.DueDate!).getTime() - new Date(b.fields.DueDate!).getTime())[0];

    return {
      Property: p.fields.Title ?? '',
      State: p.fields.cahpState ?? '',
      'Property Status': p.fields.PropertyStatus ?? '',
      Units: p.fields.UnitCount ?? '',
      'AMI Program': p.fields.AMIProgram ?? '',
      'Open Deadlines': myDeadlines.filter((c) => c.fields.DeadlineStatus !== 'Completed').length,
      'Overdue Deadlines': overdue.length,
      'Next Deadline': nextDeadline?.fields.Title ?? '',
      'Next Deadline Due': nextDeadline?.fields.DueDate
        ? formatDateOnly(nextDeadline.fields.DueDate)
        : '',
      'Compliance Status': overdue.length > 0 ? 'OVERDUE' : myDeadlines.length > 0 ? 'On Track' : 'No deadlines',
    };
  });
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

// =============================================================================
// Runners — Operational
// =============================================================================

async function runOutstandingItemsByOwnerReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  const isClosed = (s: string | undefined) =>
    s === 'Done' || s === 'Received' || s === 'Not Applicable';

  const rows = ctx.outstanding
    .filter((o) => !isClosed(o.fields.ItemStatus))
    .map((o) => {
      const property = o.fields.PropertyLookupId
        ? ctx.propertiesById.get(String(o.fields.PropertyLookupId))
        : null;
      const overdue =
        o.fields.DueDate && new Date(o.fields.DueDate).getTime() < Date.now();
      return {
        'Assigned To': o.fields.AssignedTo ?? '(unassigned)',
        Item: o.fields.Title ?? '',
        Property: property?.fields.Title ?? '',
        Status: o.fields.ItemStatus ?? '',
        Priority: o.fields.Priority ?? '',
        Category: o.fields.ItemCategory ?? '',
        'Due Date': formatDateOnly(o.fields.DueDate, ''),
        Overdue: overdue ? 'YES' : '',
        'Date Requested': o.fields.DateRequested
          ? formatDateET(o.fields.DateRequested)
          : '',
        'Has Document': o.fields.RelatedDocUrl ? 'YES' : '',
        Notes: o.fields.ItemNotes ?? '',
      };
    })
    .sort((a, b) => {
      if (a['Assigned To'] !== b['Assigned To']) {
        return String(a['Assigned To']).localeCompare(String(b['Assigned To']));
      }
      if (a.Overdue !== b.Overdue) return b.Overdue.localeCompare(a.Overdue);
      return String(a['Due Date']).localeCompare(String(b['Due Date']));
    });
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

async function runDocumentExpirationCalendar(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  const now = Date.now();
  const oneYear = now + 365 * 24 * 60 * 60 * 1000;

  // Surrogate "expiration" sources, since not every doc has an explicit ExpirationDate field:
  // 1. AMI Cert Renewal compliance deadlines
  // 2. Outstanding Items with category AMI Certification + DueDate
  // 3. Any document with an ExpirationDate field
  const rows: Record<string, unknown>[] = [];

  ctx.compliance.forEach((c) => {
    if (!c.fields.DueDate) return;
    if (c.fields.DeadlineStatus === 'Completed') return;
    const due = new Date(c.fields.DueDate).getTime();
    if (due > oneYear) return;
    const property = c.fields.PropertyLookupId
      ? ctx.propertiesById.get(String(c.fields.PropertyLookupId))
      : null;
    rows.push({
      Source: 'Compliance Deadline',
      Document: c.fields.Title ?? '',
      'Document Type': c.fields.DeadlineType ?? '',
      Property: property?.fields.Title ?? '(portfolio-wide)',
      'Expires / Due': formatDateOnly(c.fields.DueDate),
      'Days Until': Math.floor((due - now) / (24 * 60 * 60 * 1000)),
      Status: c.fields.DeadlineStatus ?? '',
      'Responsible Party': c.fields.ResponsibleParty ?? '',
    });
  });

  ctx.allDocs.forEach(({ library, doc }) => {
    if (!doc.fields.ExpirationDate) return;
    const exp = new Date(doc.fields.ExpirationDate).getTime();
    if (isNaN(exp) || exp > oneYear) return;
    const property = doc.fields.PropertyLookupId
      ? ctx.propertiesById.get(String(doc.fields.PropertyLookupId))
      : null;
    rows.push({
      Source: 'Document',
      Document: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      'Document Type': library,
      Property: property?.fields.Title ?? '(unscoped)',
      'Expires / Due': formatDateET(doc.fields.ExpirationDate),
      'Days Until': Math.floor((exp - now) / (24 * 60 * 60 * 1000)),
      Status: '',
      'Responsible Party': '',
    });
  });

  rows.sort((a, b) => Number(a['Days Until']) - Number(b['Days Until']));
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

async function runUntaggedDocumentsReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  const rows = ctx.allDocs
    .filter(({ doc }) => !doc.fields.PropertyLookupId && !doc.fields.OwnerLookupId)
    .map(({ library, doc }) => ({
      Library: library,
      Filename: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      'Uploaded': doc.fields.Created ? formatDateET(doc.fields.Created) : '',
      'Modified': doc.fields.Modified ? formatDateET(doc.fields.Modified) : '',
      'Size (bytes)': doc.fields.File_x0020_Size ?? '',
      URL: doc.webUrl ?? '',
    }))
    .sort((a, b) => a.Library.localeCompare(b.Library) || a.Filename.localeCompare(b.Filename));
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

// =============================================================================
// Runners — Owner Reports
// =============================================================================

async function runPropertyHoldingsReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  // For each property, list each owner with their relationship + ownership %
  const rows: Record<string, unknown>[] = [];
  ctx.properties.forEach((p) => {
    const propertyId = String(p.id);
    const ownerships = ctx.ownership.filter(
      (o) => String(o.fields.LinkedPropertyLookupId) === propertyId
    );
    if (ownerships.length === 0) {
      rows.push({
        Property: p.fields.Title ?? '',
        State: p.fields.cahpState ?? '',
        'Property Status': p.fields.PropertyStatus ?? '',
        Owner: '(no ownership records)',
        'Owner Type': '',
        Role: '',
        'Ownership %': '',
        'Effective Date': '',
      });
    } else {
      ownerships
        .sort((a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0))
        .forEach((rel) => {
          const owner = rel.fields.OwnerLookupId
            ? ctx.ownersById.get(String(rel.fields.OwnerLookupId))
            : null;
          rows.push({
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
          });
        });
    }
  });
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

// =============================================================================
// Runners — DOR Audit Pack
// =============================================================================

/**
 * Per-Property Audit Pack — multi-sheet Excel bundle containing everything
 * needed to defend a single property in a DOR audit. Prompts for property selection.
 */
async function runPropertyAuditPack(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  // Prompt for a property
  const choices = ctx.properties
    .slice()
    .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  if (choices.length === 0) throw new Error('No properties in the system.');

  const label = choices
    .map((p, i) => `${i + 1}. ${p.fields.Title}`)
    .join('\n');
  const input = window.prompt(
    `Per-Property Audit Pack\n\nEnter the number of the property to bundle:\n\n${label}`,
    '1'
  );
  if (!input) throw new Error('Cancelled.');
  const idx = parseInt(input.trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= choices.length) throw new Error('Invalid property selection.');
  const property = choices[idx];
  const propertyId = String(property.id);

  return await bundleAuditPack(
    [property],
    propertyId,
    ctx,
    `${descriptor.filenameBase}-${(property.fields.Title ?? 'property').replace(/[^\w-]+/g, '-')}`
  );
}

async function runPortfolioAuditPack(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  // Bundle ALL properties — every sheet aggregates across the portfolio
  return await bundleAuditPack(ctx.properties, null, ctx, descriptor.filenameBase);
}

/**
 * Shared audit pack builder. If propertyId is provided, scopes to that property.
 * Otherwise spans the full portfolio.
 */
async function bundleAuditPack(
  scopedProperties: Property[],
  propertyId: string | null,
  ctx: ReportContext,
  filenameBase: string
): Promise<number> {
  const scopedIds = new Set(scopedProperties.map((p) => String(p.id)));
  const inScope = (id: string | undefined) => (id ? scopedIds.has(String(id)) : false);

  // Sheet: Properties
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

  // Sheet: Submittals
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

  // Sheet: Ownership
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
        'Effective Date': rel.fields.EffectiveDate
          ? formatDateOnly(rel.fields.EffectiveDate)
          : '',
      };
    });

  // Sheet: Compliance
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

  // Sheet: Correspondence
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

  // Sheet: Outstanding Items
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

  // Sheet: Documents Inventory
  const documentsSheet = ctx.allDocs
    .filter(({ doc }) => inScope(doc.fields.PropertyLookupId))
    .map(({ library, doc }) => ({
      Property: ctx.propertiesById.get(String(doc.fields.PropertyLookupId))?.fields.Title ?? '',
      Library: library,
      Filename: doc.fields.FileLeafRef ?? doc.fields.Title ?? '',
      Uploaded: doc.fields.Created ? formatDateET(doc.fields.Created) : '',
      URL: doc.webUrl ?? '',
    }));

  // Sheet: Org Chart Snapshots — derived from submittal-level snapshots
  const orgChartsSheet = ctx.submittals
    .filter((s) => inScope(s.fields.PropertyLookupId))
    .filter((s) => s.fields.OrgChartSnapshotJSON || s.fields.OrgChartSnapshotDate)
    .map((s) => ({
      Property: ctx.propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? '',
      Submittal: s.fields.Title ?? '',
      'Snapshot Date': s.fields.OrgChartSnapshotDate
        ? formatDateET(s.fields.OrgChartSnapshotDate)
        : '',
      'Snapshot Size (chars)': s.fields.OrgChartSnapshotJSON?.length ?? 0,
    }));

  const sheets: Record<string, Record<string, unknown>[]> = {
    'Properties': propertiesSheet,
    'Submittals': submittalsSheet,
    'Ownership': ownershipSheet,
    'Compliance': complianceSheet,
    'Correspondence': correspondenceSheet,
    'Outstanding': outstandingSheet,
    'Documents': documentsSheet,
    'Org Chart Snapshots': orgChartsSheet,
  };

  await downloadXLSX(sheets, timestampedFilename(filenameBase, 'xlsx'));

  // Suppress unused-var warning for portfolio audit pack
  void propertyId;

  // Return total row count across sheets
  return Object.values(sheets).reduce((sum, rows) => sum + rows.length, 0);
}

async function runOrgChartHistoryReport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  // Prompt for property
  const choices = ctx.properties
    .slice()
    .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  if (choices.length === 0) throw new Error('No properties in the system.');

  const label = choices.map((p, i) => `${i + 1}. ${p.fields.Title}`).join('\n');
  const input = window.prompt(
    `Org Chart History\n\nEnter the number of the property:\n\n${label}`,
    '1'
  );
  if (!input) throw new Error('Cancelled.');
  const idx = parseInt(input.trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= choices.length) throw new Error('Invalid selection.');
  const property = choices[idx];
  const propertyId = String(property.id);

  // Collect all submittal snapshots for this property
  const snapshots = ctx.submittals
    .filter((s) => String(s.fields.PropertyLookupId) === propertyId)
    .filter((s) => s.fields.OrgChartSnapshotJSON || s.fields.OrgChartSnapshotDate)
    .map((s) => {
      let parsed: unknown = null;
      try {
        if (s.fields.OrgChartSnapshotJSON) {
          parsed = JSON.parse(s.fields.OrgChartSnapshotJSON);
        }
      } catch {
        // Leave as raw string
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
    });

  const bundle = {
    property: {
      id: property.id,
      title: property.fields.Title,
      legalEntity: property.fields.LegalEntity,
      state: property.fields.cahpState,
    },
    generatedAt: new Date().toISOString(),
    snapshotCount: snapshots.length,
    snapshots,
  };

  const safeTitle = (property.fields.Title ?? 'property').replace(/[^\w-]+/g, '-');
  downloadJSON(bundle, timestampedFilename(`${descriptor.filenameBase}-${safeTitle}`, 'json'));
  return snapshots.length;
}

// =============================================================================
// Runners — Backup and Export
// =============================================================================

async function runFullDatabaseExport(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
  // Multi-sheet Excel — every SharePoint list as its own sheet
  const sheets: Record<string, Record<string, unknown>[]> = {
    'Properties': ctx.properties.map((p) => ({ id: p.id, ...p.fields })),
    'Submittals': ctx.submittals.map((s) => ({ id: s.id, ...s.fields })),
    'Outstanding': ctx.outstanding.map((o) => ({ id: o.id, ...o.fields })),
    'Compliance': ctx.compliance.map((c) => ({ id: c.id, ...c.fields })),
    'Owners': ctx.owners.map((o) => ({ id: o.id, ...o.fields })),
    'Ownership': ctx.ownership.map((o) => ({ id: o.id, ...o.fields })),
    'Correspondence': ctx.correspondence.map((c) => ({ id: c.id, ...c.fields })),
    'Communications': ctx.comms.map((c) => ({ id: c.id, ...c.fields })),
  };

  await downloadXLSX(sheets, timestampedFilename(descriptor.filenameBase, 'xlsx'));
  return Object.values(sheets).reduce((sum, rows) => sum + rows.length, 0);
}

async function runSharePointLibrarySnapshot(
  descriptor: ReportDescriptor,
  ctx: ReportContext
): Promise<number> {
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

  downloadJSON(snapshot, timestampedFilename(descriptor.filenameBase, 'json'));
  return ctx.allDocs.length;
}
