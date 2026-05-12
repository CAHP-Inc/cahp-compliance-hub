import { useState, useMemo } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  LIST_NAMES,
  type Property,
  type Submittal,
  type PropertyStatus,
  type SubmittalStatusValue,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const STATUS_STYLES: Record<PropertyStatus, string> = {
  Active: 'bg-green-100 text-green-800 border-green-200',
  Pending: 'bg-yellow-100 text-yellow-800 border-yellow-200',
  Withdrawn: 'bg-gray-100 text-gray-700 border-gray-200',
  'Removed from Program': 'bg-red-100 text-red-800 border-red-200',
  Sold: 'bg-blue-100 text-blue-800 border-blue-200',
};

const SUBMITTAL_STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  Draft: 'bg-gray-100 text-gray-700',
  'Package Mailed (NC)': 'bg-blue-100 text-blue-800',
  Filed: 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-yellow-100 text-yellow-800',
  'Responded - Awaiting DOR': 'bg-amber-100 text-amber-800',
  Approved: 'bg-green-100 text-green-800',
  Denied: 'bg-red-100 text-red-800',
  Withdrawn: 'bg-gray-100 text-gray-700',
};

type TabId = 'overview' | 'submittals' | 'documents' | 'compliance' | 'notes';

const TABS: { id: TabId; label: string; shipped: boolean }[] = [
  { id: 'overview', label: 'Overview', shipped: true },
  { id: 'submittals', label: 'Submittals', shipped: true },
  { id: 'documents', label: 'Documents', shipped: false },
  { id: 'compliance', label: 'Compliance', shipped: false },
  { id: 'notes', label: 'Notes', shipped: false },
];

export function PropertyDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState<TabId>('overview');

  const { data: property, loading, error } = useSharePointItem<Property>(
    LIST_NAMES.Properties,
    id
  );

  // Fetch all submittals; filter client-side by PropertyLookupId. For 18 records this
  // is faster than dealing with SharePoint indexed-column requirements on lookups.
  const { data: allSubmittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, {
    top: 500,
  });

  const relatedSubmittals = useMemo(() => {
    if (!allSubmittals || !id) return [];
    return allSubmittals.filter((s) => String(s.fields.PropertyLookupId) === String(id));
  }, [allSubmittals, id]);

  if (loading) {
    return (
      <div className="text-center py-12 text-gray-500">
        <div className="inline-flex items-center gap-3">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading property…</span>
        </div>
      </div>
    );
  }

  if (error || !property) {
    return (
      <div>
        <BreadcrumbBar />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Property not found
          </div>
          <p className="text-sm text-red-700 mb-3">
            {error ? error.message : `No property with ID ${id} exists in Properties Registry.`}
          </p>
          <button
            onClick={() => navigate('/properties')}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            ← Back to Properties
          </button>
        </div>
      </div>
    );
  }

  const f = property.fields;

  return (
    <div>
      <BreadcrumbBar propertyName={f.Title} />

      {/* Property header */}
      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-3xl font-bold text-teal-700">{f.Title}</h1>
            {f.PropertyStatus && (
              <span
                className={`inline-block px-2.5 py-0.5 rounded text-xs font-semibold border ${
                  STATUS_STYLES[f.PropertyStatus]
                }`}
              >
                {f.PropertyStatus}
              </span>
            )}
            {f.cahpVerificationStatus === 'Inherited - Unverified' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-yellow-50 text-yellow-800 border border-yellow-200">
                UNVERIFIED
              </span>
            )}
            {f.cahpVerificationStatus === 'Needs Follow-Up' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-error/10 text-error border border-error/20">
                NEEDS FOLLOW-UP
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500 mt-1">
            {[
              f.LegalEntity,
              f.cahpState,
              f.cahpCounty?.replace(/\s*\([^)]*\)\s*/g, ''),
              f.UnitCount ? `${f.UnitCount} units` : null,
              f.AMIProgram && f.AMIProgram !== 'None' ? `${f.AMIProgram} AMI` : null,
              f.cahpOwnerGroup,
            ]
              .filter(Boolean)
              .join(' · ')}
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
            disabled
            title="Watch button ships in PR-05c"
          >
            <Icon name="star" size={14} />
            Watch
          </button>
          <button
            className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            disabled
            title="Inline editing ships in PR-05b"
          >
            <Icon name="settings" size={14} />
            Edit
          </button>
        </div>
      </div>

      {/* Tab strip */}
      <div className="border-b border-gray-200 mb-6">
        <div className="flex gap-1">
          {TABS.map((t) => {
            const active = activeTab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => t.shipped && setActiveTab(t.id)}
                disabled={!t.shipped}
                className={`px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
                  active
                    ? 'border-teal-700 text-teal-700'
                    : t.shipped
                      ? 'border-transparent text-gray-600 hover:text-teal-700 hover:border-gray-300'
                      : 'border-transparent text-gray-400 cursor-not-allowed'
                }`}
                title={!t.shipped ? 'Coming in PR-05b' : undefined}
              >
                {t.label}
                {!t.shipped && <span className="ml-1.5 text-[10px] text-gray-400">soon</span>}
              </button>
            );
          })}
        </div>
      </div>

      {activeTab === 'overview' && <OverviewTab property={property} />}
      {activeTab === 'submittals' && <SubmittalsTab submittals={relatedSubmittals} />}
    </div>
  );
}

// =============================================================================
// Tab: Overview
// =============================================================================

function OverviewTab({ property }: { property: Property }) {
  const f = property.fields;

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Section title="Property Identity">
        <Field label="Property Name" value={f.Title} />
        <Field label="Legal Entity" value={f.LegalEntity} />
        <Field label="Property Address" value={f.PropertyAddress} />
        <Field label="DOR Account ID" value={f.DORAccountID} mono />
        <Field label="Date Added to CAHP" value={formatDate(f.DateAddedToCAHP)} />
      </Section>

      <Section title="Affordability">
        <Field label="AMI Program" value={f.AMIProgram} />
        <Field label="CAHP Language Added" value={f.CAHPLanguageAdded} />
        <Field label="LURA Executed" value={f.LURAExecuted} />
        <Field label="Operating Agreement Version" value={f.OpAgreementVersion} mono />
      </Section>

      <Section title="Status & Verification">
        <Field label="Property Status" value={f.PropertyStatus} />
        <Field label="Verification Status" value={f.cahpVerificationStatus} />
        <Field label="Removed Reason" value={f.RemovedReason} />
      </Section>

      <Section title="Location & Ownership">
        <Field label="State" value={f.cahpState} mono />
        <Field label="County" value={f.cahpCounty} />
        <Field label="Units" value={f.UnitCount?.toString()} mono />
        <Field label="Owner Group" value={f.cahpOwnerGroup} />
      </Section>

      {f.PropertyNotes && (
        <Section title="Notes" fullWidth>
          <p className="text-sm text-gray-700 whitespace-pre-wrap leading-relaxed">
            {f.PropertyNotes}
          </p>
        </Section>
      )}
    </div>
  );
}

// =============================================================================
// Tab: Submittals
// =============================================================================

function SubmittalsTab({ submittals }: { submittals: Submittal[] }) {
  if (submittals.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500">No submittals on file for this property yet.</p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
          <tr>
            <th className="px-4 py-3 text-left">Submittal</th>
            <th className="px-4 py-3 text-left">Tax Year</th>
            <th className="px-4 py-3 text-left">Filing Method</th>
            <th className="px-4 py-3 text-left">Status</th>
            <th className="px-4 py-3 text-left">Filed Date</th>
            <th className="px-4 py-3 text-right">Approved Abatement</th>
            <th className="px-4 py-3 text-left">Next Action</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {submittals.map((s) => (
            <tr key={s.id} className="hover:bg-gray-50 transition-colors">
              <td className="px-4 py-3 font-medium text-gray-900">{s.fields.Title}</td>
              <td className="px-4 py-3 text-gray-700 font-mono-data">{s.fields.cahpTaxYear || '—'}</td>
              <td className="px-4 py-3 text-gray-700 text-xs">{s.fields.FilingMethod || '—'}</td>
              <td className="px-4 py-3">
                {s.fields.SubmittalStatus ? (
                  <span
                    className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                      SUBMITTAL_STATUS_STYLES[s.fields.SubmittalStatus] || 'bg-gray-100'
                    }`}
                  >
                    {s.fields.SubmittalStatus}
                  </span>
                ) : (
                  '—'
                )}
              </td>
              <td className="px-4 py-3 text-gray-700 text-xs">{formatDate(s.fields.DateFiled)}</td>
              <td className="px-4 py-3 text-right font-mono-data">
                {s.fields.ApprovedAbatement
                  ? `$${s.fields.ApprovedAbatement.toLocaleString()}`
                  : '—'}
              </td>
              <td className="px-4 py-3 text-gray-700 text-xs">{s.fields.NextAction || '—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// =============================================================================
// Building blocks
// =============================================================================

function BreadcrumbBar({ propertyName }: { propertyName?: string }) {
  return (
    <nav className="mb-4 text-sm">
      <Link to="/properties" className="text-teal-700 hover:text-teal-900 font-medium">
        ← Properties
      </Link>
      {propertyName && (
        <>
          <span className="mx-2 text-gray-400">/</span>
          <span className="text-gray-700">{propertyName}</span>
        </>
      )}
    </nav>
  );
}

function Section({
  title,
  children,
  fullWidth,
}: {
  title: string;
  children: React.ReactNode;
  fullWidth?: boolean;
}) {
  return (
    <div className={`bg-white border border-gray-200 rounded-lg p-5 shadow-card ${fullWidth ? 'lg:col-span-2' : ''}`}>
      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">{title}</h3>
      <div className="space-y-2.5">{children}</div>
    </div>
  );
}

function Field({ label, value, mono }: { label: string; value?: string | null; mono?: boolean }) {
  return (
    <div className="flex items-start gap-3">
      <dt className="text-sm text-gray-500 w-44 flex-shrink-0">{label}</dt>
      <dd className={`text-sm text-gray-900 flex-1 ${mono ? 'font-mono-data' : ''}`}>
        {value || <span className="text-gray-300">—</span>}
      </dd>
    </div>
  );
}

function formatDate(iso: string | undefined | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
