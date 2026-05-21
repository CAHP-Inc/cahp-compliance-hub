import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import { useSharePointList, LIST_NAMES, type Property, type PropertyStatus, type CahpState, type Submittal, type SubmittalStatusValue, type TaxMapID, type Contact } from '../lib/sharepoint';
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
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });

  // contactId → contact, for quick lookup when rendering the Owner Contact column
  const contactsById = useMemo(() => {
    const m = new Map<string, Contact>();
    (contacts.data ?? []).forEach((c) => m.set(String(c.id), c));
    return m;
  }, [contacts.data]);

  /**
   * Parcels per property — used to indicate multi-parcel filings.
   */
  const parcelCountByProperty = useMemo(() => {
    const map = new Map<string, number>();
    (taxMapIDs.data ?? []).forEach((t) => {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
      if (!pid) return;
      map.set(pid, (map.get(pid) ?? 0) + 1);
    });
    return map;
  }, [taxMapIDs.data]);

  /**
   * For each property's latest year+filing-type, count submittals by status.
   * Used to show "X of N Filed" style multi-parcel breakdown.
   */
  const filingAggregateByProperty = useMemo(() => {
    if (!submittals.data) return new Map<string, { total: number; filed: number; approved: number; denied: number; draft: number; year?: string; filingType?: string }>();
    // Find max-year + tiebreaking latest filing per property first
    const latestKey = new Map<string, { year: string; filingType: string }>();
    submittals.data.forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      if (!pid) return;
      const year = s.fields.cahpTaxYear ?? '';
      const ft = s.fields.FilingType ?? '';
      const cur = latestKey.get(pid);
      if (!cur || (year > cur.year) || (year === cur.year && ft > cur.filingType)) {
        latestKey.set(pid, { year, filingType: ft });
      }
    });
    // Aggregate counts for that latest year+filing-type
    const result = new Map<string, { total: number; filed: number; approved: number; denied: number; draft: number; year?: string; filingType?: string }>();
    latestKey.forEach((key, pid) => {
      const matching = submittals.data!.filter(
        (s) =>
          String(s.fields.PropertyLookupId ?? '') === pid &&
          (s.fields.cahpTaxYear ?? '') === key.year &&
          (s.fields.FilingType ?? '') === key.filingType
      );
      const agg = {
        total: matching.length,
        filed: 0,
        approved: 0,
        denied: 0,
        draft: 0,
        year: key.year || undefined,
        filingType: key.filingType || undefined,
      };
      matching.forEach((s) => {
        const status = s.fields.SubmittalStatus;
        if (status === 'Draft') agg.draft++;
        else if (status === 'Approved') agg.approved++;
        else if (status === 'Denied') agg.denied++;
        else if (status) agg.filed++; // Filed, Letter Received, Responded etc.
      });
      result.set(pid, agg);
    });
    return result;
  }, [submittals.data]);

  const [search, setSearch] = useState('');
  const [stateFilter, setStateFilter] = useState<CahpState | 'All'>('All');
  const [statusFilter, setStatusFilter] = useState<PropertyStatus | 'All'>('All');
  // Owner contact filter — value is the Contact's listItem ID (string) or 'All' / 'None'.
  const [contactFilter, setContactFilter] = useState<string>('All');
  const [sortBy, setSortBy] = useState<'name' | 'filingStatus'>('name');

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
      if (contactFilter !== 'All') {
        const propContactId = f.PropertyOwnerContactLookupId ? String(f.PropertyOwnerContactLookupId) : '';
        if (contactFilter === 'None') {
          if (propContactId) return false;
        } else if (propContactId !== contactFilter) {
          return false;
        }
      }
      return true;
    });
    // Client-side sort — SharePoint won't sort server-side because Title isn't indexed
    if (sortBy === 'filingStatus') {
      // Status order per workflow: Draft → Filed → Letter Received → Responded → Denied → Approved → Withdrawn
      // Properties without any submittal go to the bottom.
      const ORDER: Record<string, number> = {
        'Draft': 1,
        'Filed': 2,
        'Letter Received - Action Needed': 3,
        'Letter Received': 3,
        'Responded - Awaiting DOR': 4,
        'Denied': 5,
        'Approved': 6,
        'Withdrawn': 7,
      };
      return result.sort((a, b) => {
        const aSub = latestSubmittalByProperty.get(a.id);
        const bSub = latestSubmittalByProperty.get(b.id);
        const aStatus = aSub?.fields.SubmittalStatus ?? '';
        const bStatus = bSub?.fields.SubmittalStatus ?? '';
        const aRank = ORDER[aStatus] ?? 99;
        const bRank = ORDER[bStatus] ?? 99;
        if (aRank !== bRank) return aRank - bRank;
        // Tiebreaker: property name
        return (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '');
      });
    }
    return result.sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [data, search, stateFilter, statusFilter, contactFilter, sortBy, latestSubmittalByProperty]);

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
        <select
          value={contactFilter}
          onChange={(e) => setContactFilter(e.target.value)}
          className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          title="Filter by owner contact"
        >
          <option value="All">All contacts</option>
          <option value="None">— No contact set —</option>
          {[...(contacts.data ?? [])]
            .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''))
            .map((c) => (
              <option key={c.id} value={String(c.id)}>{c.fields.Title}</option>
            ))}
        </select>
        <div className="flex items-center gap-1">
          <span className="text-[11px] text-gray-500 uppercase tracking-wider font-semibold">Sort:</span>
          <select
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value as 'name' | 'filingStatus')}
            className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          >
            <option value="name">Property name</option>
            <option value="filingStatus">Filing status</option>
          </select>
        </div>
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
                <th className="px-4 py-3 text-left">Owner Contact</th>
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
                    {(() => {
                      const list = (p.fields.cahpCounty ?? '').split(',').map((s) => s.trim()).filter(Boolean);
                      if (list.length === 0) return '—';
                      // Strip the "(SC)"/"(NC)" suffix per county for the column display since
                      // the State column already shows that info; full label stays on the detail page.
                      const stripped = list.map((c) => c.replace(/\s*\([^)]*\)\s*/g, ''));
                      if (stripped.length === 1) return stripped[0];
                      return (
                        <div className="flex flex-wrap gap-1">
                          {stripped.map((c) => (
                            <span key={c} className="px-1 py-0.5 rounded bg-teal-50 text-teal-800 text-[10px] font-medium">
                              {c}
                            </span>
                          ))}
                        </div>
                      );
                    })()}
                  </td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {p.fields.UnitCount ?? '—'}
                  </td>
                  <td className="px-4 py-3 text-gray-700 text-xs">{p.fields.AMIProgram || '—'}</td>
                  <td className="px-4 py-3 text-xs">
                    {(() => {
                      const cId = p.fields.PropertyOwnerContactLookupId
                        ? String(p.fields.PropertyOwnerContactLookupId)
                        : '';
                      const contact = cId ? contactsById.get(cId) : undefined;
                      if (!contact) return <span className="text-gray-400">—</span>;
                      return (
                        <div className="min-w-0">
                          <div className="text-gray-900 truncate">{contact.fields.Title}</div>
                          {contact.fields.ContactEmail && (
                            <div className="text-[11px] text-gray-500 font-mono-data truncate">
                              {contact.fields.ContactEmail}
                            </div>
                          )}
                        </div>
                      );
                    })()}
                  </td>
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
                      const agg = filingAggregateByProperty.get(p.id);
                      const sub = latestSubmittalByProperty.get(p.id);
                      const parcelCount = parcelCountByProperty.get(p.id) ?? 0;

                      if (!sub || !agg) {
                        return (
                          <div className="flex flex-col gap-0.5 items-start">
                            <span className="text-gray-400 text-xs italic">Not Filed</span>
                            {parcelCount > 1 && (
                              <span className="text-[10px] text-gray-400 font-mono-data">
                                {parcelCount} parcels
                              </span>
                            )}
                          </div>
                        );
                      }

                      const status = sub.fields.SubmittalStatus;
                      const year = agg.year ?? sub.fields.cahpTaxYear;
                      const filingType = agg.filingType ?? sub.fields.FilingType;
                      const multiParcel = agg.total > 1;

                      // Decide what the "headline" status is
                      // Mixed → show "Mixed"; uniform → show that single status
                      const allApproved = agg.approved === agg.total;
                      const allDenied = agg.denied === agg.total;
                      const allDraft = agg.draft === agg.total;
                      const allFiled = agg.filed === agg.total;
                      const isMixed = multiParcel && !allApproved && !allDenied && !allDraft && !allFiled;
                      const headlineStatus = isMixed ? 'Mixed' : status;

                      return (
                        <div className="flex flex-col gap-0.5 items-start">
                          {headlineStatus && (
                            <span
                              className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                                FILING_STATUS_STYLES[headlineStatus as SubmittalStatusValue] || 'bg-purple-100 text-purple-800'
                              }`}
                            >
                              {headlineStatus}
                            </span>
                          )}
                          {(year || filingType) && (
                            <span className="text-[10px] text-gray-500 font-mono-data">
                              {year ?? ''}{year && filingType ? ' · ' : ''}{filingType ?? ''}
                            </span>
                          )}
                          {multiParcel && (
                            <span
                              className="text-[10px] text-gray-600 font-mono-data"
                              title={`${agg.draft} Draft / ${agg.filed} Filed / ${agg.approved} Approved / ${agg.denied} Denied`}
                            >
                              {isMixed ? (
                                <>
                                  {agg.approved > 0 && <span className="text-green-700">{agg.approved}A</span>}
                                  {agg.approved > 0 && (agg.filed + agg.denied + agg.draft > 0) && ' / '}
                                  {agg.filed > 0 && <span className="text-blue-700">{agg.filed}F</span>}
                                  {agg.filed > 0 && (agg.denied + agg.draft > 0) && ' / '}
                                  {agg.denied > 0 && <span className="text-red-700">{agg.denied}D</span>}
                                  {agg.denied > 0 && agg.draft > 0 && ' / '}
                                  {agg.draft > 0 && <span className="text-gray-600">{agg.draft}Dr</span>}
                                  {' of '}{agg.total}
                                </>
                              ) : (
                                <>{agg.total} of {parcelCount} parcels</>
                              )}
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
