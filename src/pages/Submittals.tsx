import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Submittal,
  type TaxMapID,
  type Property,
  type SubmittalStatusValue,
  type SubmittalFilingType,
  type CahpTaxYear,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateOnly } from '../lib/dates';
import { NewSubmittalModal } from '../components/NewSubmittalModal';
import { DORDeadlinesCard } from '../components/DORDeadlinesCard';

const STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  'Draft': 'bg-gray-100 text-gray-800',
  'Package Mailed (NC)': 'bg-indigo-100 text-indigo-800',
  'Filed': 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-amber-100 text-amber-800',
  'Responded - Awaiting DOR': 'bg-purple-100 text-purple-800',
  'Approved': 'bg-green-100 text-green-800',
  'Invoiced': 'bg-teal-100 text-teal-800',
  'Paid': 'bg-emerald-100 text-emerald-900',
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

/** Gold "SAHA" chip — parcel was previously approved for abatement under SAHA. */
function SahaBadge() {
  return (
    <span
      title="Parcel previously approved for abatement under SAHA"
      className="ml-1.5 inline-block px-1 py-0.5 rounded text-[9px] font-semibold bg-gold-100 text-gold-900 align-middle font-sans"
    >
      SAHA
    </span>
  );
}

export function Submittals() {
  const navigate = useNavigate();
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });

  const taxMapIdsById = useMemo(() => {
    const m = new Map<string, TaxMapID>();
    (taxMapIDs.data ?? []).forEach((t) => m.set(String(t.id), t));
    return m;
  }, [taxMapIDs.data]);

  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<SubmittalStatusValue | 'All'>('All');
  const [yearFilter, setYearFilter] = useState<CahpTaxYear | 'All'>('All');
  const [filingTypeFilter, setFilingTypeFilter] = useState<SubmittalFilingType | 'All'>('All');
  const [sahaFilter, setSahaFilter] = useState<'All' | 'SAHA' | 'NonSAHA'>('All');
  const [newSubmittalOpen, setNewSubmittalOpen] = useState(false);
  const [groupByProperty, setGroupByProperty] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set());

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
        if (sahaFilter !== 'All') {
          const parcel = f.TaxMapIDLookupId ? taxMapIdsById.get(String(f.TaxMapIDLookupId)) : undefined;
          const isSaha = !!parcel?.fields.PriorSAHAAbatement;
          if (sahaFilter === 'SAHA' && !isSaha) return false;
          if (sahaFilter === 'NonSAHA' && isSaha) return false;
        }
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
  }, [submittals.data, search, statusFilter, yearFilter, filingTypeFilter, sahaFilter, propertiesById, taxMapIdsById]);

  /**
   * Group filtered submittals by property + year + filing type.
   * Each group represents one "filing campaign" — what shows up as one MyDORWAY filing
   * if it covers a single parcel, or one parent filing covering N parcels (Townes at Converse).
   */
  type SubmittalGroup = {
    key: string;
    propertyId: string;
    propertyTitle: string;
    year: string;
    filingType: string;
    submittals: Submittal[];
    counts: { total: number; draft: number; filed: number; approved: number; denied: number };
    headlineStatus: string;
  };
  const grouped: SubmittalGroup[] = useMemo(() => {
    const groups = new Map<string, SubmittalGroup>();
    filtered.forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      const year = s.fields.cahpTaxYear ?? '';
      const ft = s.fields.FilingType ?? '';
      const key = `${pid}::${year}::${ft}`;
      let g = groups.get(key);
      if (!g) {
        g = {
          key,
          propertyId: pid,
          propertyTitle: propertiesById.get(pid)?.fields.Title ?? '(unknown property)',
          year,
          filingType: ft,
          submittals: [],
          counts: { total: 0, draft: 0, filed: 0, approved: 0, denied: 0 },
          headlineStatus: '',
        };
        groups.set(key, g);
      }
      g.submittals.push(s);
      g.counts.total++;
      const status = s.fields.SubmittalStatus;
      if (status === 'Draft') g.counts.draft++;
      // Invoiced / Paid are past Approved — still count as approved for the headline.
      else if (status === 'Approved' || status === 'Invoiced' || status === 'Paid') g.counts.approved++;
      else if (status === 'Denied') g.counts.denied++;
      else if (status) g.counts.filed++;
    });
    // Compute headline status per group: uniform → that status; otherwise "Mixed"
    groups.forEach((g) => {
      const { total, draft, filed, approved, denied } = g.counts;
      if (approved === total) g.headlineStatus = 'Approved';
      else if (denied === total) g.headlineStatus = 'Denied';
      else if (draft === total) g.headlineStatus = 'Draft';
      else if (filed === total) g.headlineStatus = 'Filed';
      else g.headlineStatus = 'Mixed';
    });
    return Array.from(groups.values()).sort((a, b) => {
      // Year DESC, then property name
      if (a.year !== b.year) return (b.year || '').localeCompare(a.year || '');
      return a.propertyTitle.localeCompare(b.propertyTitle);
    });
  }, [filtered, propertiesById]);

  const toggleGroup = (key: string) => {
    setExpandedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

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
            DOR property tax abatement filings. One submittal per tax map ID per tax year.
          </p>
        </div>
        <button
          onClick={() => setNewSubmittalOpen(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
        >
          <Icon name="plus" size={14} />
          New Submittal
        </button>
      </div>

      {/* DOR Deadlines — RFI responses owed + DOR responses awaited, overdue first */}
      <DORDeadlinesCard submittals={submittals.data} propertiesById={propertiesById} />

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
        <select
          value={sahaFilter}
          onChange={(e) => setSahaFilter(e.target.value as 'All' | 'SAHA' | 'NonSAHA')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
          title="Filter by prior SAHA abatement"
        >
          <option value="All">All parcels (SAHA)</option>
          <option value="SAHA">SAHA only</option>
          <option value="NonSAHA">Exclude SAHA</option>
        </select>
        <label className="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer ml-2 select-none">
          <input
            type="checkbox"
            checked={groupByProperty}
            onChange={(e) => setGroupByProperty(e.target.checked)}
            className="rounded border-gray-300 text-teal-700 focus:ring-teal-500"
          />
          Group by Property
        </label>
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
                <th className="px-4 py-3 text-left">Tax Map ID</th>
                <th className="px-4 py-3 text-left">Year</th>
                <th className="px-4 py-3 text-left">Filing Type</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Filed</th>
                <th className="px-4 py-3 text-left">Confirmation #</th>
                <th className="px-4 py-3 text-left">Next Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {groupByProperty ? (
                grouped.flatMap((g) => {
                  const isExpanded = expandedGroups.has(g.key);
                  const headlineColor =
                    g.headlineStatus === 'Approved' ? 'bg-green-100 text-green-800'
                    : g.headlineStatus === 'Denied' ? 'bg-red-100 text-red-800'
                    : g.headlineStatus === 'Draft' ? 'bg-gray-100 text-gray-700'
                    : g.headlineStatus === 'Mixed' ? 'bg-purple-100 text-purple-800'
                    : 'bg-blue-100 text-blue-800';
                  const rows: JSX.Element[] = [];
                  // Group header row
                  rows.push(
                    <tr
                      key={g.key}
                      onClick={() => toggleGroup(g.key)}
                      className="bg-teal-50 hover:bg-teal-100 cursor-pointer border-l-4 border-teal-700"
                    >
                      <td className="px-4 py-3 font-semibold text-teal-900">
                        <span className="inline-block w-4 text-center text-teal-700 mr-1">{isExpanded ? '▾' : '▸'}</span>
                        {g.propertyTitle}
                      </td>
                      <td className="px-4 py-3 font-mono-data text-xs text-gray-700">
                        {g.counts.total} {g.counts.total === 1 ? 'parcel' : 'parcels'}
                      </td>
                      <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{g.year || '—'}</td>
                      <td className="px-4 py-3">
                        {g.filingType ? (
                          <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${FILING_TYPE_STYLES[g.filingType as SubmittalFilingType] ?? 'bg-gray-100 text-gray-700'}`}>
                            {g.filingType}
                          </span>
                        ) : <span className="text-gray-300 text-xs">—</span>}
                      </td>
                      <td className="px-4 py-3">
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${headlineColor}`}>
                          {g.headlineStatus}
                        </span>
                        {g.headlineStatus === 'Mixed' && (
                          <div className="text-[10px] text-gray-600 font-mono-data mt-0.5">
                            {g.counts.approved > 0 && <span className="text-green-700">{g.counts.approved}A</span>}
                            {g.counts.approved > 0 && (g.counts.filed + g.counts.denied + g.counts.draft > 0) && ' / '}
                            {g.counts.filed > 0 && <span className="text-blue-700">{g.counts.filed}F</span>}
                            {g.counts.filed > 0 && (g.counts.denied + g.counts.draft > 0) && ' / '}
                            {g.counts.denied > 0 && <span className="text-red-700">{g.counts.denied}D</span>}
                            {g.counts.denied > 0 && g.counts.draft > 0 && ' / '}
                            {g.counts.draft > 0 && <span className="text-gray-700">{g.counts.draft}Dr</span>}
                          </div>
                        )}
                      </td>
                      <td className="px-4 py-3" colSpan={3}>
                        <span className="text-xs text-gray-500 italic">
                          {isExpanded ? 'Click to collapse' : `Click to expand ${g.counts.total} submittals`}
                        </span>
                      </td>
                    </tr>
                  );
                  // Detail rows when expanded
                  if (isExpanded) {
                    g.submittals.forEach((s) => {
                      rows.push(
                        <tr
                          key={`${g.key}-${s.id}`}
                          onClick={() => navigate(`/submittals/${s.id}`)}
                          className="hover:bg-gray-50 cursor-pointer bg-white"
                        >
                          <td className="px-4 py-2 pl-12 text-gray-700 text-xs">
                            {s.fields.Title || <span className="text-gray-400 italic">(no title)</span>}
                          </td>
                          <td className="px-4 py-2 font-mono-data text-xs text-gray-700">
                            {(() => {
                              if (!s.fields.TaxMapIDLookupId) return <span className="text-gray-400 italic font-sans">unassigned</span>;
                              const t = taxMapIdsById.get(String(s.fields.TaxMapIDLookupId));
                              if (!t) return <span className="text-gray-400 italic font-sans">missing</span>;
                              return <>{t.fields.Title}{t.fields.PriorSAHAAbatement && <SahaBadge />}</>;
                            })()}
                          </td>
                          <td className="px-4 py-2 text-gray-700 font-mono-data text-xs">{s.fields.cahpTaxYear ?? '—'}</td>
                          <td className="px-4 py-2">
                            {s.fields.FilingType ? (
                              <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${FILING_TYPE_STYLES[s.fields.FilingType]}`}>
                                {s.fields.FilingType}
                              </span>
                            ) : <span className="text-gray-300 text-xs">—</span>}
                          </td>
                          <td className="px-4 py-2">
                            {s.fields.SubmittalStatus ? (
                              <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${STATUS_STYLES[s.fields.SubmittalStatus]}`}>
                                {s.fields.SubmittalStatus}
                              </span>
                            ) : <span className="text-gray-300 text-xs">—</span>}
                          </td>
                          <td className="px-4 py-2 text-gray-700 font-mono-data text-xs">{formatDateOnly(s.fields.DateFiled)}</td>
                          <td className="px-4 py-2 text-gray-700 font-mono-data text-xs">{s.fields.ConfirmationNumber || '—'}</td>
                          <td className="px-4 py-2 text-xs text-gray-600 max-w-xs truncate">{s.fields.NextAction || '—'}</td>
                        </tr>
                      );
                    });
                  }
                  return rows;
                })
              ) : (
                filtered.map((s) => {
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
                      <td className="px-4 py-3 font-mono-data text-xs text-gray-700">
                        {(() => {
                          if (!s.fields.TaxMapIDLookupId) {
                            return <span className="text-gray-400 italic font-sans">unassigned</span>;
                          }
                          const t = taxMapIdsById.get(String(s.fields.TaxMapIDLookupId));
                          if (!t) return <span className="text-gray-400 italic font-sans">missing</span>;
                          return <>{t.fields.Title}{t.fields.PriorSAHAAbatement && <SahaBadge />}</>;
                        })()}
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
                        {formatDateOnly(s.fields.DateFiled)}
                      </td>
                      <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{s.fields.ConfirmationNumber || '—'}</td>
                      <td className="px-4 py-3 text-xs text-gray-600 max-w-xs truncate">{s.fields.NextAction || '—'}</td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      )}

      {newSubmittalOpen && (
        <NewSubmittalModal
          onClose={() => setNewSubmittalOpen(false)}
          onCreated={(id) => {
            submittals.refetch?.();
            navigate(`/submittals/${id}`);
          }}
        />
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
