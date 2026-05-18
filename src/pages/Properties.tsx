import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import { useSharePointList, LIST_NAMES, type Property, type PropertyStatus, type CahpState, type Submittal, type SubmittalStatusValue } from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const STATUS_STYLES: Record<PropertyStatus, string> = {
  Active: 'bg-green-100 text-green-800',
  Pending: 'bg-yellow-100 text-yellow-800',
  Withdrawn: 'bg-gray-100 text-gray-700',
  'Removed from Program': 'bg-red-100 text-red-800',
  Sold: 'bg-blue-100 text-blue-800',
};

const FILING_STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  'Draft': 'bg-gray-100 text-gray-800',
  'Package Mailed (NC)': 'bg-indigo-100 text-indigo-800',
  'Filed': 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-amber-100 text-amber-800',
  'Responded - Awaiting DOR': 'bg-purple-100 text-purple-800',
  'Approved': 'bg-green-100 text-green-800',
  'Denied': 'bg-red-100 text-red-800',
  'Withdrawn': 'bg-gray-100 text-gray-500',
};

export function Properties() {
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useSharePointList<Property>(LIST_NAMES.Properties, {
    top: 200,
  });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const [search, setSearch] = useState('');
  const [stateFilter, setStateFilter] = useState<CahpState | 'All'>('All');
  const [statusFilter, setStatusFilter] = useState<PropertyStatus | 'All'>('All');

  /**
   * Build a map of propertyId → most recent submittal for the property.
   * "Most recent" = highest tax year, then latest DateFiled as a tiebreaker.
   * This lets us surface each property's current filing posture on the table.
   */
  const latestSubmittalByProperty = useMemo(() => {
    if (!submittals.data) return new Map<string, Submittal>();
    const map = new Map<string, Submittal>();
    submittals.data.forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      if (!pid) return;
      const existing = map.get(pid);
      if (!existing) {
        map.set(pid, s);
        return;
      }
      const existingYear = Number(existing.fields.cahpTaxYear ?? 0);
      const newYear = Number(s.fields.cahpTaxYear ?? 0);
      if (newYear > existingYear) {
        map.set(pid, s);
      } else if (newYear === existingYear) {
        const existingDate = existing.fields.DateFiled ?? '';
        const newDate = s.fields.DateFiled ?? '';
        if (newDate > existingDate) {
          map.set(pid, s);
        }
      }
    });
    return map;
  }, [submittals.data]);

  const filtered = useMemo(() => {
    if (!data) return [];
    const result = data.filter((p) => {
      const f = p.fields;
      if (search) {
        const hay = `${f.Title ?? ''} ${f.LegalEntity ?? ''} ${f.PropertyAddress ?? ''}`.toLowerCase();
        if (!hay.includes(search.toLowerCase())) return false;
      }
      if (stateFilter !== 'All' && f.cahpState !== stateFilter) return false;
      if (statusFilter !== 'All' && f.PropertyStatus !== statusFilter) return false;
      return true;
    });
    // Client-side sort — SharePoint won't sort server-side because Title isn't indexed
    return result.sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [data, search, stateFilter, statusFilter]);

  const stats = useMemo(() => {
    if (!data) return null;
    return {
      total: data.length,
      active: data.filter((p) => p.fields.PropertyStatus === 'Active').length,
      sc: data.filter((p) => p.fields.cahpState === 'SC').length,
      nc: data.filter((p) => p.fields.cahpState === 'NC').length,
      units: data.reduce((sum, p) => sum + (p.fields.UnitCount ?? 0), 0),
    };
  }, [data]);

  if (loading) return <LoadingState />;
  if (error) return <ErrorState error={error} onRetry={refetch} />;
  if (!data || !stats) return null;

  return (
    <div>
      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Properties</h1>
          <p className="text-sm text-gray-500 mt-1">
            {stats.total} properties under CAHP management · {stats.units.toLocaleString()} units total
          </p>
        </div>
        <button
          onClick={() => navigate('/properties/new')}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors"
        >
          <Icon name="plus" size={16} />
          New Property
        </button>
      </div>

      {/* KPI cards */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
        <KPICard label="Total Properties" value={stats.total} />
        <KPICard label="Active" value={stats.active} accent="success" />
        <KPICard label="Total Units" value={stats.units.toLocaleString()} />
        <KPICard label="SC" value={stats.sc} />
        <KPICard label="NC" value={stats.nc} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search property, entity, address…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <Select value={stateFilter} onChange={(v) => setStateFilter(v as CahpState | 'All')} options={['All', 'SC', 'NC']} />
        <Select
          value={statusFilter}
          onChange={(v) => setStatusFilter(v as PropertyStatus | 'All')}
          options={['All', 'Active', 'Pending', 'Withdrawn', 'Removed from Program', 'Sold']}
        />
        {filtered.length !== data.length && (
          <span className="text-xs text-gray-500 px-1">
            {filtered.length} of {data.length}
          </span>
        )}
      </div>

      {/* Table */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Legal Entity</th>
                <th className="px-4 py-3 text-left">State</th>
                <th className="px-4 py-3 text-left">County</th>
                <th className="px-4 py-3 text-right">Units</th>
                <th className="px-4 py-3 text-left">AMI</th>
                <th className="px-4 py-3 text-left">Owner Group</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Filing Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((p) => (
                <tr
                  key={p.id}
                  onClick={() => navigate(`/properties/${p.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">
                    {p.fields.Title}
                    {p.fields.cahpVerificationStatus === 'Inherited - Unverified' && (
                      <span
                        className="ml-2 inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold bg-yellow-50 text-yellow-800 align-middle"
                        title="Inherited data — needs verification"
                      >
                        UNVERIFIED
                      </span>
                    )}
                    {p.fields.cahpVerificationStatus === 'Needs Follow-Up' && (
                      <span className="ml-2 inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold bg-error/10 text-error align-middle">
                        FOLLOW-UP
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-gray-700">{p.fields.LegalEntity || '—'}</td>
                  <td className="px-4 py-3">
                    <span className="font-mono-data text-xs font-semibold text-teal-700">
                      {p.fields.cahpState || '—'}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-700 text-xs">
                    {p.fields.cahpCounty?.replace(/\s*\([^)]*\)\s*/g, '') || '—'}
                  </td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {p.fields.UnitCount ?? '—'}
                  </td>
                  <td className="px-4 py-3 text-gray-700 text-xs">{p.fields.AMIProgram || '—'}</td>
                  <td className="px-4 py-3 text-gray-700 text-xs">{p.fields.cahpOwnerGroup || '—'}</td>
                  <td className="px-4 py-3">
                    {p.fields.PropertyStatus ? (
                      <span
                        className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                          STATUS_STYLES[p.fields.PropertyStatus] || 'bg-gray-100 text-gray-700'
                        }`}
                      >
                        {p.fields.PropertyStatus}
                      </span>
                    ) : (
                      '—'
                    )}
                  </td>
                  <td className="px-4 py-3">
                    {(() => {
                      const sub = latestSubmittalByProperty.get(p.id);
                      if (!sub) {
                        return <span className="text-gray-400 text-xs italic">Not Filed</span>;
                      }
                      const status = sub.fields.SubmittalStatus;
                      const taxYear = sub.fields.cahpTaxYear;
                      const filingType = sub.fields.FilingType;
                      return (
                        <div className="flex flex-col gap-0.5 items-start">
                          {status && (
                            <span
                              className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                                FILING_STATUS_STYLES[status] || 'bg-gray-100 text-gray-700'
                              }`}
                            >
                              {status}
                            </span>
                          )}
                          {(taxYear || filingType) && (
                            <span className="text-[10px] text-gray-500 font-mono-data">
                              {taxYear ?? ''}{taxYear && filingType ? ' · ' : ''}{filingType ?? ''}
                            </span>
                          )}
                        </div>
                      );
                    })()}
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={9} className="px-4 py-8 text-center text-gray-500 text-sm">
                    No properties match the current filters.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <p className="text-xs text-gray-400 mt-4 text-center">
        Click any property to drill into its detail page · Click <strong>New Property</strong> to add one.
      </p>
    </div>
  );
}

function KPICard({ label, value, accent }: { label: string; value: string | number; accent?: 'success' }) {
  const accentClass = accent === 'success' ? 'text-success' : 'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}

function Select({ value, onChange, options }: { value: string; onChange: (v: string) => void; options: readonly string[] }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
    >
      {options.map((o) => (
        <option key={o} value={o}>
          {o}
        </option>
      ))}
    </select>
  );
}

function LoadingState() {
  return (
    <div>
      <h1 className="text-2xl font-bold text-teal-700 mb-6">Properties</h1>
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin"></div>
          <span className="text-sm">Loading properties from SharePoint…</span>
        </div>
      </div>
    </div>
  );
}

function ErrorState({ error, onRetry }: { error: Error; onRetry: () => void }) {
  return (
    <div>
      <h1 className="text-2xl font-bold text-teal-700 mb-6">Properties</h1>
      <div className="bg-red-50 border border-red-200 rounded-lg p-6">
        <div className="font-semibold text-error mb-2 flex items-center gap-2">
          <Icon name="alert" size={18} />
          Failed to load properties
        </div>
        <p className="text-sm text-red-700 mb-3 font-mono-data">{error.message}</p>
        <div className="text-xs text-red-600 space-y-1">
          <p>Common causes:</p>
          <ul className="list-disc list-inside ml-2 space-y-0.5">
            <li>
              <code className="font-mono-data">VITE_SHAREPOINT_SITE</code> env var is missing or
              wrong (should be <code>vanrockre.sharepoint.com:/sites/CAHPComplianceHub</code>)
            </li>
            <li>
              The list name doesn't match exactly — expecting <code>Properties Registry</code>
            </li>
            <li>Insufficient SharePoint permissions on the site</li>
            <li>MSAL token couldn't be acquired silently for SharePoint scopes</li>
          </ul>
        </div>
        <button
          onClick={onRetry}
          className="mt-4 text-sm text-teal-700 hover:text-teal-900 font-medium underline"
        >
          Retry
        </button>
      </div>
    </div>
  );
}
