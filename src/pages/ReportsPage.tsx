import { useState, useMemo } from 'react';
import {
  useSharePointList,
  LIST_NAMES,
  type Property,
  type Submittal,
  type OutstandingItem,
  type ComplianceDeadline,
} from '../lib/sharepoint';
import {
  REPORTS,
  downloadCSV,
  timestampedFilename,
  type ReportDescriptor,
  type ReportCategory,
  type ReportStatus,
} from '../lib/reports';
import { Icon } from '../components/ui/Icon';

const STATUS_LABEL: Record<ReportStatus, string> = {
  'available': 'Available',
  'pending-billing': 'Pending Billing Module',
  'pending-pr14b': 'Coming in PR-14b',
};

const STATUS_STYLES: Record<ReportStatus, string> = {
  'available': 'bg-success/10 text-success border-success/30',
  'pending-billing': 'bg-gray-100 text-gray-500 border-gray-200',
  'pending-pr14b': 'bg-gold-100 text-gold-900 border-gold-200',
};

const CATEGORY_ICONS: Record<ReportCategory, 'star' | 'file' | 'calendar' | 'dollar' | 'check' | 'folder'> = {
  'Owner Reports': 'star',
  'DOR Audit Pack': 'file',
  'Year-End / Annual': 'calendar',
  'Financial': 'dollar',
  'Operational': 'check',
  'Backup and Export': 'folder',
};

export function ReportsPage() {
  // Data for the reports that actually run
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const compliance = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });

  const [runningId, setRunningId] = useState<string | null>(null);
  const [runResult, setRunResult] = useState<{ id: string; success: boolean; message: string } | null>(null);

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const grouped = useMemo(() => {
    const map = new Map<ReportCategory, ReportDescriptor[]>();
    REPORTS.forEach((r) => {
      const list = map.get(r.category) ?? [];
      list.push(r);
      map.set(r.category, list);
    });
    return map;
  }, []);

  const handleRun = async (descriptor: ReportDescriptor) => {
    if (descriptor.status !== 'available') return;
    setRunningId(descriptor.id);
    setRunResult(null);

    try {
      let rowCount = 0;

      switch (descriptor.id) {
        case 'annual-filing-report': {
          rowCount = await runAnnualFilingReport(descriptor, submittals.data ?? [], propertiesById);
          break;
        }
        case 'compliance-status': {
          rowCount = await runComplianceStatusReport(descriptor, properties.data ?? [], compliance.data ?? []);
          break;
        }
        case 'outstanding-items-by-owner': {
          rowCount = await runOutstandingItemsByOwnerReport(descriptor, outstanding.data ?? [], propertiesById);
          break;
        }
        default: {
          throw new Error(`No runner registered for report '${descriptor.id}'`);
        }
      }

      setRunResult({
        id: descriptor.id,
        success: true,
        message: `Downloaded ${rowCount} row${rowCount === 1 ? '' : 's'} as CSV.`,
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
          Pre-built reports across six categories. Click Run Now to download as CSV.
          Schedule (recurring delivery) and Custom Report Builder are deferred to a future phase.
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

async function runAnnualFilingReport(
  descriptor: ReportDescriptor,
  submittals: Submittal[],
  propertiesById: Map<string, Property>
): Promise<number> {
  const thisYear = new Date().getFullYear();
  const rows = submittals
    .filter((s) => {
      const date = s.fields.DateFiled ?? s.createdDateTime;
      if (!date) return false;
      return new Date(date).getFullYear() === thisYear;
    })
    .map((s) => {
      const propertyTitle = s.fields.PropertyLookupId
        ? propertiesById.get(String(s.fields.PropertyLookupId))?.fields.Title ?? ''
        : '';
      return {
        Property: propertyTitle,
        'Submittal Title': s.fields.Title ?? '',
        'Tax Year': s.fields.cahpTaxYear ?? '',
        State: s.fields.cahpState ?? '',
        'Filing Type': s.fields.FilingType ?? '',
        'Filing Method': s.fields.FilingMethod ?? '',
        Status: s.fields.SubmittalStatus ?? '',
        'Date Filed': s.fields.DateFiled ? new Date(s.fields.DateFiled).toLocaleDateString() : '',
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
  properties: Property[],
  compliance: ComplianceDeadline[]
): Promise<number> {
  const rows = properties.map((p) => {
    const propertyId = String(p.id);
    const myDeadlines = compliance.filter((c) => String(c.fields.PropertyLookupId) === propertyId);
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
      'Open Deadlines': myDeadlines.filter((c) => c.fields.DeadlineStatus !== 'Completed').length,
      'Overdue Deadlines': overdue.length,
      'Next Deadline': nextDeadline?.fields.Title ?? '',
      'Next Deadline Due': nextDeadline?.fields.DueDate
        ? new Date(nextDeadline.fields.DueDate).toLocaleDateString()
        : '',
      'Compliance Status': overdue.length > 0 ? 'OVERDUE' : myDeadlines.length > 0 ? 'On Track' : 'No deadlines',
    };
  });
  downloadCSV(rows, timestampedFilename(descriptor.filenameBase, 'csv'));
  return rows.length;
}

async function runOutstandingItemsByOwnerReport(
  descriptor: ReportDescriptor,
  outstanding: OutstandingItem[],
  propertiesById: Map<string, Property>
): Promise<number> {
  const isClosed = (s: string | undefined) =>
    s === 'Done' || s === 'Received' || s === 'Not Applicable';

  const rows = outstanding
    .filter((o) => !isClosed(o.fields.ItemStatus))
    .map((o) => {
      const property = o.fields.PropertyLookupId
        ? propertiesById.get(String(o.fields.PropertyLookupId))
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
        'Due Date': o.fields.DueDate ? new Date(o.fields.DueDate).toLocaleDateString() : '',
        Overdue: overdue ? 'YES' : '',
        'Date Requested': o.fields.DateRequested
          ? new Date(o.fields.DateRequested).toLocaleDateString()
          : '',
        Notes: o.fields.ItemNotes ?? '',
      };
    })
    // Sort by assignee, then by overdue first, then by due date
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
