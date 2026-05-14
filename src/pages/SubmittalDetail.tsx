import { useMemo } from 'react';
import { useParams, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  LIST_NAMES,
  type Submittal,
  type Property,
  type Correspondence,
  type OutstandingItem,
  type AuditLog,
  type SubmittalStatusValue,
  type SubmittalFilingType,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar } from '../components/detail';

const STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  'Draft': 'bg-gray-100 text-gray-800',
  'Package Mailed (NC)': 'bg-indigo-100 text-indigo-800',
  'Filed': 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-amber-100 text-amber-800',
  'Responded - Awaiting DOR': 'bg-purple-100 text-purple-800',
  'Approved': 'bg-green-100 text-green-800',
  'Denied': 'bg-red-100 text-red-800',
  'Withdrawn': 'bg-gray-100 text-gray-500',
};

const FILING_TYPE_STYLES: Record<SubmittalFilingType, string> = {
  'Initial': 'bg-teal-100 text-teal-800',
  'Annual': 'bg-blue-100 text-blue-800',
  'Amendment': 'bg-amber-100 text-amber-800',
};

// Pipeline visualization — spec §3.6.4
const PIPELINE_STAGES: { status: SubmittalStatusValue; label: string }[] = [
  { status: 'Draft', label: 'Draft' },
  { status: 'Filed', label: 'Filed' },
  { status: 'Letter Received - Action Needed', label: 'Letter Received' },
  { status: 'Responded - Awaiting DOR', label: 'Responded' },
  { status: 'Approved', label: 'Approved' },
];

const TERMINAL_STATUSES: SubmittalStatusValue[] = ['Approved', 'Denied', 'Withdrawn'];

export function SubmittalDetail() {
  const { id } = useParams<{ id: string }>();

  const { data: submittal, loading, error } = useSharePointItem<Submittal>(LIST_NAMES.Submittals, id);
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const auditLog = useSharePointList<AuditLog>(LIST_NAMES.AuditLog, { top: 200 });

  const property = useMemo(() => {
    if (!submittal || !properties.data || !submittal.fields.PropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(submittal.fields.PropertyLookupId)) ?? null;
  }, [submittal, properties.data]);

  // Related correspondence — filter by property
  const relatedCorrespondence = useMemo(() => {
    if (!submittal || !correspondence.data) return [];
    const propertyId = submittal.fields.PropertyLookupId;
    if (!propertyId) return [];
    return correspondence.data
      .filter((c) => String(c.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => {
        const da = a.fields.DateReceived ? new Date(a.fields.DateReceived).getTime() : 0;
        const db = b.fields.DateReceived ? new Date(b.fields.DateReceived).getTime() : 0;
        return db - da;
      });
  }, [submittal, correspondence.data]);

  // Action plan — outstanding items for this property still open
  const actionPlan = useMemo(() => {
    if (!submittal || !outstanding.data) return [];
    const propertyId = submittal.fields.PropertyLookupId;
    if (!propertyId) return [];
    return outstanding.data
      .filter(
        (o) =>
          String(o.fields.PropertyLookupId) === String(propertyId) &&
          o.fields.ItemStatus !== 'Received' &&
          o.fields.ItemStatus !== 'Not Applicable'
      );
  }, [submittal, outstanding.data]);

  // Activity for this submittal
  const submittalActivity = useMemo(() => {
    if (!submittal || !auditLog.data) return [];
    return auditLog.data
      .filter(
        (a) =>
          a.fields.EntityType === 'Submittal' &&
          String(a.fields.EntityId) === String(submittal.id)
      )
      .sort(
        (a, b) =>
          new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
      );
  }, [submittal, auditLog.data]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading submittal…</span>
        </div>
      </div>
    );
  }

  if (error || !submittal) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Submittals" parentTo="/submittals" currentLabel="Submittal Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load submittal</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const f = submittal.fields;
  const statusIdx = f.SubmittalStatus
    ? PIPELINE_STAGES.findIndex((s) => s.status === f.SubmittalStatus)
    : -1;
  const isTerminal = f.SubmittalStatus ? TERMINAL_STATUSES.includes(f.SubmittalStatus) : false;
  const isDenied = f.SubmittalStatus === 'Denied';
  const isWithdrawn = f.SubmittalStatus === 'Withdrawn';

  return (
    <div>
      <BreadcrumbBar parentLabel="Submittals" parentTo="/submittals" currentLabel={f.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{f.Title}</h1>
            {f.FilingType && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${FILING_TYPE_STYLES[f.FilingType]}`}>
                {f.FilingType}
              </span>
            )}
            {f.SubmittalStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[f.SubmittalStatus]}`}>
                {f.SubmittalStatus}
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500">
            {property ? (
              <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                {property.fields.Title}
              </Link>
            ) : (
              <span className="italic text-gray-400">unlinked</span>
            )}
            {f.cahpTaxYear && ` · Tax Year ${f.cahpTaxYear}`}
            {f.cahpState && ` · ${f.cahpState}`}
          </p>
        </div>
        <div className="flex items-center gap-2">
          <span className="px-3 py-1.5 border border-dashed border-gray-300 text-gray-500 rounded-md text-xs font-medium italic">
            Editing & status transitions ship in PR-10b
          </span>
        </div>
      </div>

      {/* Status pipeline visualization */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-5">
        <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">Status Pipeline</div>
        {isTerminal && (isDenied || isWithdrawn) ? (
          <div className={`p-3 rounded-md ${isDenied ? 'bg-red-50 border border-red-200' : 'bg-gray-50 border border-gray-200'}`}>
            <p className="text-sm font-semibold">{isDenied ? 'Denied' : 'Withdrawn'}</p>
            <p className="text-xs text-gray-600 mt-1">
              {isDenied
                ? 'DOR denied this submittal. No billing record will be created. Appeal possible.'
                : 'Submittal withdrawn before DOR action. No billing record will be created.'}
            </p>
          </div>
        ) : (
          <div className="flex items-center gap-1 overflow-x-auto pb-1">
            {PIPELINE_STAGES.map((stage, idx) => {
              const reached = statusIdx >= idx;
              const isCurrent = statusIdx === idx;
              return (
                <div key={stage.status} className="flex items-center gap-1 flex-1 min-w-fit">
                  <div
                    className={`flex-1 px-2 py-1.5 rounded-md text-center text-xs font-medium whitespace-nowrap transition-colors ${
                      isCurrent
                        ? 'bg-teal-700 text-white'
                        : reached
                          ? 'bg-success text-white'
                          : 'bg-gray-100 text-gray-400'
                    }`}
                  >
                    {reached && !isCurrent && <Icon name="check" size={10} className="inline mr-1" />}
                    {stage.label}
                  </div>
                  {idx < PIPELINE_STAGES.length - 1 && (
                    <div className={`w-3 h-0.5 ${reached && idx < statusIdx ? 'bg-success' : 'bg-gray-200'}`} />
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Metadata + Action Plan two-column */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
        {/* Left two-thirds: metadata */}
        <div className="lg:col-span-2 bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <h3 className="text-sm font-semibold text-gray-700 mb-3">Submittal Metadata</h3>
          <dl className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-2 text-sm">
            <Row label="Filing Type" value={f.FilingType} />
            <Row label="Tax Year" value={f.cahpTaxYear} mono />
            <Row label="State" value={f.cahpState} mono />
            <Row label="Filing Method" value={f.FilingMethod} />
            <Row label="Date Filed" value={formatDate(f.DateFiled)} mono />
            <Row label="Confirmation #" value={f.ConfirmationNumber} mono />
            <Row label="Mail Tracking #" value={f.MailTrackingNumber} mono />
            <Row label="Approved Abatement" value={f.ApprovedAbatement != null ? `$${f.ApprovedAbatement.toLocaleString()}` : undefined} mono />
            <Row label="Next Action" value={f.NextAction} />
            <Row label="Next Action Due" value={formatDate(f.NextActionDue)} mono />
          </dl>
          {f.SubmittalNotes && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1">Notes</div>
              <p className="text-sm text-gray-700 whitespace-pre-wrap">{f.SubmittalNotes}</p>
            </div>
          )}
          {f.OrgChartSnapshotJSON && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <div className="flex items-center gap-2 text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1">
                <Icon name="check" size={12} className="text-success" />
                Org Chart Snapshot Frozen
              </div>
              <p className="text-xs text-gray-600">
                {f.OrgChartSnapshotDate
                  ? `Captured ${new Date(f.OrgChartSnapshotDate).toLocaleString()}`
                  : 'Snapshot captured (date unknown)'}
                {' · '}
                <span className="font-mono-data">{f.OrgChartSnapshotJSON.length.toLocaleString()} bytes</span>
              </p>
            </div>
          )}
        </div>

        {/* Right one-third: action plan */}
        <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-semibold text-gray-700">Action Plan</h3>
            <span className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">
              {actionPlan.length} open
            </span>
          </div>
          {actionPlan.length === 0 ? (
            <p className="text-xs text-gray-500 italic">No open outstanding items for this property.</p>
          ) : (
            <ul className="space-y-2 text-sm">
              {actionPlan.slice(0, 8).map((o) => (
                <li key={o.id} className="flex items-start gap-2">
                  <Icon
                    name={o.fields.ItemStatus === 'Overdue' ? 'alert' : 'inbox'}
                    size={12}
                    className={`mt-1 flex-shrink-0 ${o.fields.ItemStatus === 'Overdue' ? 'text-error' : 'text-gray-400'}`}
                  />
                  <div className="flex-1 min-w-0">
                    <div className="text-gray-900 truncate">{o.fields.Title}</div>
                    <div className="text-[11px] text-gray-500 font-mono-data">
                      {o.fields.ItemStatus ?? 'Requested'}
                      {o.fields.DateRequested && ` · requested ${new Date(o.fields.DateRequested).toLocaleDateString()}`}
                    </div>
                  </div>
                </li>
              ))}
              {actionPlan.length > 8 && (
                <li className="text-[11px] text-gray-400 italic">…and {actionPlan.length - 8} more</li>
              )}
            </ul>
          )}
        </div>
      </div>

      {/* Correspondence Thread */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">DOR Correspondence</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {relatedCorrespondence.length === 0
              ? 'No DOR correspondence logged for this property.'
              : `${relatedCorrespondence.length} letter${relatedCorrespondence.length === 1 ? '' : 's'} on file. Full CRUD ships in PR-11a.`}
          </p>
        </div>
        {relatedCorrespondence.length > 0 && (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Direction</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Received</th>
                <th className="px-4 py-3 text-left">Response Due</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {relatedCorrespondence.map((c) => (
                <tr key={c.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.Direction || '—'}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.LetterType || '—'}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.DateReceived)}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.ResponseDue)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Activity */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Activity</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {submittalActivity.length === 0
              ? 'No audit events for this submittal yet.'
              : `${submittalActivity.length} event${submittalActivity.length === 1 ? '' : 's'}, newest first`}
          </p>
        </div>
        {submittalActivity.length > 0 && (
          <ul className="divide-y divide-gray-100">
            {submittalActivity.map((row) => (
              <li key={row.id} className="px-4 py-3 flex items-start gap-3 text-sm">
                <span className="inline-block px-1.5 py-0.5 rounded text-[10px] font-bold bg-gray-100 text-gray-800 flex-shrink-0">
                  {row.fields.Action}
                </span>
                <div className="flex-1 min-w-0">
                  <div className="text-gray-900">{row.fields.Title}</div>
                  {row.fields.ChangeSummary && (
                    <pre className="mt-1 text-xs text-gray-600 whitespace-pre-wrap font-mono-data">
                      {row.fields.ChangeSummary}
                    </pre>
                  )}
                </div>
                <div className="text-xs text-gray-500 font-mono-data flex-shrink-0 text-right">
                  <div>{new Date(row.createdDateTime).toLocaleDateString()}</div>
                  <div className="text-gray-400">{row.createdBy?.user?.displayName ?? ''}</div>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

function Row({ label, value, mono }: { label: string; value?: string | number; mono?: boolean }) {
  const hasValue = value !== null && value !== undefined && value !== '';
  return (
    <div className="flex items-start gap-3 py-0.5">
      <dt className="text-sm text-gray-500 w-40 flex-shrink-0">{label}</dt>
      <dd className={`text-sm flex-1 ${mono ? 'font-mono-data' : ''} ${hasValue ? 'text-gray-900' : 'text-gray-300'}`}>
        {hasValue ? String(value) : '—'}
      </dd>
    </div>
  );
}

function formatDate(iso: string | undefined): string | undefined {
  if (!iso) return undefined;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString();
}
