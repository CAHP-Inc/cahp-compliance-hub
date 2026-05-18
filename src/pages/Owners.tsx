import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  countPropertiesForOwner,
  countLLCsOwnedBy,
  type Owner,
  type Ownership,
  type OwnerType,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const TYPE_STYLES: Record<OwnerType, string> = {
  Individual: 'bg-blue-100 text-blue-800',
  LLC: 'bg-purple-100 text-purple-800',
  Nonprofit: 'bg-teal-100 text-teal-800',
  Trust: 'bg-amber-100 text-amber-800',
  Corporation: 'bg-indigo-100 text-indigo-800',
  'Limited Partnership': 'bg-rose-100 text-rose-800',
  'General Partnership': 'bg-fuchsia-100 text-fuchsia-800',
};

export function Owners() {
  const navigate = useNavigate();
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });

  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState<OwnerType | 'All'>('All');
  const [stateFilter, setStateFilter] = useState<string>('All');

  const loading = owners.loading || ownership.loading;
  const error = owners.error || ownership.error;

  const states = useMemo(() => {
    if (!owners.data) return [];
    return Array.from(new Set(owners.data.map((o) => o.fields.OwnerState).filter(Boolean))) as string[];
  }, [owners.data]);

  const filtered = useMemo(() => {
    if (!owners.data) return [];
    return owners.data.filter((o) => {
      const f = o.fields;
      if (search && !f.Title?.toLowerCase().includes(search.toLowerCase())) return false;
      if (typeFilter !== 'All' && f.OwnerType !== typeFilter) return false;
      if (stateFilter !== 'All' && f.OwnerState !== stateFilter) return false;
      return true;
    });
  }, [owners.data, search, typeFilter, stateFilter]);

  const kpis = useMemo(() => {
    if (!owners.data || !ownership.data) return null;
    return {
      total: owners.data.length,
      individuals: owners.data.filter((o) => o.fields.OwnerType === 'Individual').length,
      llcs: owners.data.filter((o) => o.fields.OwnerType === 'LLC').length,
      nonprofits: owners.data.filter((o) => o.fields.OwnerType === 'Nonprofit').length,
    };
  }, [owners.data, ownership.data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Owners</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading owners…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Owners</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load owners</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (!owners.data || !kpis) return null;

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Owners</h1>
          <p className="text-sm text-gray-500 mt-1">
            Entity master — every person, LLC, and nonprofit with ownership interest in CAHP properties.
            {' '}<strong>This is the single source of truth.</strong> Edit an owner here; changes cascade to every property where they hold interest.
          </p>
        </div>
        <button
          onClick={() => navigate('/owners/new')}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors flex-shrink-0"
        >
          <Icon name="plus" size={16} />
          New Owner
        </button>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total Owners" value={kpis.total} />
        <KPI label="Individuals" value={kpis.individuals} />
        <KPI label="LLCs" value={kpis.llcs} />
        <KPI label="Nonprofits" value={kpis.nonprofits} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search by name…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={typeFilter}
          onChange={(e) => setTypeFilter(e.target.value as OwnerType | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All types</option>
          <option value="Individual">Individual</option>
          <option value="LLC">LLC</option>
          <option value="Nonprofit">Nonprofit</option>
        </select>
        {states.length > 0 && (
          <select
            value={stateFilter}
            onChange={(e) => setStateFilter(e.target.value)}
            className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
          >
            <option value="All">All states</option>
            {states.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        )}
        {filtered.length !== owners.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {owners.data.length}</span>
        )}
      </div>

      {/* Empty state */}
      {owners.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No owners yet</p>
          <p className="text-sm text-blue-800 mb-4">
            Add your first owner — that's an individual, LLC, or nonprofit that holds interest in a property.
          </p>
          <button
            onClick={() => navigate('/owners/new')}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium inline-flex items-center gap-2 transition-colors"
          >
            <Icon name="plus" size={16} />
            Add First Owner
          </button>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Name</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">State</th>
                <th className="px-4 py-3 text-right">Properties</th>
                <th className="px-4 py-3 text-right">LLCs Owned</th>
                <th className="px-4 py-3 text-left">Contact</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((owner) => {
                const propCounts = ownership.data
                  ? countPropertiesForOwner(owner.id, ownership.data, owners.data!)
                  : { direct: 0, indirect: 0 };
                const llcsOwned = ownership.data ? countLLCsOwnedBy(owner.id, ownership.data) : 0;
                return (
                  <tr
                    key={owner.id}
                    onClick={() => navigate(`/owners/${owner.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">{owner.fields.Title}</td>
                    <td className="px-4 py-3">
                      {owner.fields.OwnerType ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${TYPE_STYLES[owner.fields.OwnerType]}`}>
                          {owner.fields.OwnerType}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{owner.fields.OwnerState || '—'}</td>
                    <td className="px-4 py-3 text-right font-mono-data text-xs">
                      {propCounts.direct > 0 ? (
                        <span className="text-gray-900 font-semibold">{propCounts.direct}</span>
                      ) : (
                        <span className="text-gray-300">0</span>
                      )}
                      {propCounts.indirect > 0 && (
                        <span className="text-gray-400"> · {propCounts.indirect} via</span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-right font-mono-data text-xs">
                      {llcsOwned > 0 ? llcsOwned : <span className="text-gray-300">0</span>}
                    </td>
                    <td className="px-4 py-3 text-gray-600 text-xs">{owner.fields.ContactEmail || '—'}</td>
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

function KPI({ label, value }: { label: string; value: number }) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className="text-3xl font-bold mt-1 text-teal-700">{value}</div>
    </div>
  );
}
