import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Submittal,
  type Property,
  type SubmittalStatusValue,
  type SubmittalFilingType,
  type CahpTaxYear,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

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

const ACTIONABLE_STATUSES: SubmittalStatusValue[] = [
  'Letter Received - Action Needed',
  'Responded - Awaiting DOR',
];

export function Submittals() {
  const navigate = useNavigate();
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<SubmittalStatusValue | 'All'>('All');
  const [yearFilter, setYearFilter] = useState<CahpTaxYear | 'All'>('All');
  const [filingTypeFilter, setFilingTypeFilter] = useState<SubmittalFilingType | 'All'>('All');

  const loading = submittals.loading || properties.loading;
  const error = submittals.error || properties.error;

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const filtered = useMemo(() => {
    if (!submittals.data) return [];
    return submittals.data
      .filter((s) => {
        const f = s.fields;
        if (search) {
          const propName = f.PropertyLookupId
            ? propertiesById.get(String(f.PropertyLookupId))?.fields.Title ?? ''
            : '';
          const hay = `${f.Title ?? ''} ${propName} ${f.ConfirmationNumber ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (statusFilter !== 'All' && f.SubmittalStatus !== statusFilter) return false;
        if (yearFilter !== 'All' && f.cahpTaxYear !== yearFilter) return false;
        if (filingTypeFilter !== 'All' && f.FilingType !== filingTypeFilter) return false;
        return true;
      })
      .sort((a, b) => {
        // Sort by year DESC, then drafts/unfiled bubble up (Infinity DateFiled = on top within year)
        const ya = Number(a.fields.cahpTaxYear ?? 0);
        const yb = Number(b.fields.cahpTaxYear ?? 0);
        if (yb !== ya) return yb - ya;
        const da = a.fields.DateFiled ? new Date(a.fields.DateFiled).getTime() : Infinity;
        const db = b.fields.DateFiled ? new Date(b.fields.DateFiled).getTime() : Infinity;
        return db - da;
      });
  }, [submittals.data, search, statusFilter, yearFilter, filingTypeFilter, propertiesById]);

  const stats = useMemo(() => {
    if (!submittals.data) return null;
    return {
      total: submittals.data.length,
      drafts: submittals.data.filter((s) => s.fields.SubmittalStatus === 'Draft').length,
      inFlight: submittals.data.filter(
        (s) =>
          s.fields.SubmittalStatus &&
          ['Filed', 'Letter Received - Action Needed', 'Responded - Awaiting DOR'].includes(s.fields.SubmittalStatus)
      ).length,
      needsAction: submittals.data.filter(
        (s) => s.fields.SubmittalStatus && ACTIONABLE_STATUSES.includes(s.fields.SubmittalStatus)
      ).length,
    };
  }, [submittals.data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Submittals</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading submittals…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Submittals</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load submittals</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (!submittals.data || !stats) return null;

  const years = Array.from(
    new Set(submittals.data.map((s) => s.fields.cahpTaxYear).filter(Boolean))
  ).sort().reverse() as CahpTaxYear[];

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Submittals</h1>
          <p className="text-sm text-gray-500 mt-1">
            DOR property tax abatement filings. One submittal per property per tax year for annual filings; additional for amendments.
          </p>
        </div>
      </div>

      {/* KPI row */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total Submittals" value={stats.total} />
        <KPI label="Drafts" value={stats.drafts} accent="warning" />
        <KPI label="In Flight" value={stats.inFlight} />
        <KPI label="Needs Action" value={stats.needsAction} accent={stats.needsAction > 0 ? 'danger' : 'default'} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search title, property, confirmation #…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as SubmittalStatusValue | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All statuses</option>
          {(Object.keys(STATUS_STYLES) as SubmittalStatusValue[]).map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </select>
        <select
          value={yearFilter}
          onChange={(e) => setYearFilter(e.target.value as CahpTaxYear | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white font-mono-data"
        >
          <option value="All">All years</option>
          {years.map((y) => (<option key={y} value={y}>{y}</option>))}
        </select>
        <select
          value={filingTypeFilter}
          onChange={(e) => setFilingTypeFilter(e.target.value as SubmittalFilingType | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All filing types</option>
          <option value="Initial">Initial</option>
          <option value="Annual">Annual</option>
          <option value="Amendment">Amendment</option>
        </select>
        {filtered.length !== submittals.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {submittals.data.length}</span>
        )}
      </div>

      {/* Table */}
      {submittals.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No submittals yet</p>
          <p className="text-sm text-blue-800">
            Submittals are created automatically by the Property Creation Wizard, or you can add them from a property's Submittals tab.
          </p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Year</th>
                <th className="px-4 py-3 text-left">Filing Type</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Filed</th>
                <th className="px-4 py-3 text-left">Confirmation #</th>
                <th className="px-4 py-3 text-left">Next Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((s) => {
                const property = s.fields.PropertyLookupId
                  ? propertiesById.get(String(s.fields.PropertyLookupId))
                  : null;
                return (
                  <tr
                    key={s.id}
                    onClick={() => navigate(`/submittals/${s.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">
                      {property?.fields.Title ?? <span className="text-gray-400 italic">(unlinked)</span>}
                      {s.fields.Title && property?.fields.Title !== s.fields.Title && (
                        <div className="text-[11px] text-gray-500 font-mono-data mt-0.5 truncate max-w-xs">{s.fields.Title}</div>
                      )}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{s.fields.cahpTaxYear ?? '—'}</td>
                    <td className="px-4 py-3">
                      {s.fields.FilingType ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${FILING_TYPE_STYLES[s.fields.FilingType]}`}>
                          {s.fields.FilingType}
                        </span>
                      ) : <span className="text-gray-300 text-xs">—</span>}
                    </td>
                    <td className="px-4 py-3">
                      {s.fields.SubmittalStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[s.fields.SubmittalStatus]}`}>
                          {s.fields.SubmittalStatus}
                        </span>
                      ) : <span className="text-gray-300 text-xs">—</span>}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {s.fields.DateFiled ? new Date(s.fields.DateFiled).toLocaleDateString() : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{s.fields.ConfirmationNumber || '—'}</td>
                    <td className="px-4 py-3 text-xs text-gray-600 max-w-xs truncate">{s.fields.NextAction || '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function KPI({
  label,
  value,
  accent = 'default',
}: {
  label: string;
  value: number;
  accent?: 'default' | 'warning' | 'danger';
}) {
  const accentClass =
    accent === 'danger' ? 'text-error' : accent === 'warning' ? 'text-warning' : 'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}
