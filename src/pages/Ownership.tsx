import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Ownership,
  type RelationshipType,
  type Property,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const RELATIONSHIP_STYLES: Record<RelationshipType, string> = {
  'Managing Member': 'bg-teal-100 text-teal-800',
  'Sole Member': 'bg-emerald-100 text-emerald-800',
  Member: 'bg-blue-100 text-blue-800',
  Owner: 'bg-purple-100 text-purple-800',
  Subsidiary: 'bg-amber-100 text-amber-800',
  'Beneficial Owner': 'bg-pink-100 text-pink-800',
};

export function Ownership() {
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useSharePointList<Ownership>(
    LIST_NAMES.Ownership,
    { top: 500 }
  );
  const { data: allProperties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 200 });

  const propertiesById = useMemo(() => {
    const map = new Map<string, Property>();
    allProperties?.forEach((p) => map.set(String(p.id), p));
    return map;
  }, [allProperties]);

  const [search, setSearch] = useState('');
  const [relationshipFilter, setRelationshipFilter] = useState<RelationshipType | 'All'>('All');
  const [propertyFilter, setPropertyFilter] = useState<string>('All');

  const filtered = useMemo(() => {
    if (!data) return [];
    return data.filter((o) => {
      const f = o.fields;
      if (search) {
        const hay = `${f.Title ?? ''} ${f.ParentEntity ?? ''}`.toLowerCase();
        if (!hay.includes(search.toLowerCase())) return false;
      }
      if (relationshipFilter !== 'All' && f.RelationshipType !== relationshipFilter) return false;
      if (propertyFilter !== 'All' && String(f.LinkedPropertyLookupId) !== propertyFilter) return false;
      return true;
    }).sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [data, search, relationshipFilter, propertyFilter]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Ownership Structure</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading ownership records…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Ownership Structure</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Failed to load
          </div>
          <p className="text-sm text-red-700 mb-3 font-mono-data">{error.message}</p>
          <button onClick={refetch} className="text-sm text-teal-700 hover:text-teal-900 font-medium underline">
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  return (
    <div>
      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Ownership Structure</h1>
          <p className="text-sm text-gray-500 mt-1">
            {data.length} ownership records · maps entities to properties and parent entities
          </p>
        </div>
        <button
          onClick={() => navigate('/ownership/new')}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors"
        >
          <Icon name="plus" size={16} />
          New Ownership Entry
        </button>
      </div>

      {data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-blue-100 text-blue-700 mb-3">
            <Icon name="star" size={24} />
          </div>
          <h2 className="text-base font-semibold text-blue-900 mb-1">No ownership records yet</h2>
          <p className="text-sm text-blue-800 max-w-md mx-auto mb-4">
            Use this module to map who owns what across your portfolio — managing members of each property LLC,
            beneficial owners of those entities, and the chain of ownership above them.
          </p>
          <button
            onClick={() => navigate('/ownership/new')}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium text-sm transition-colors"
          >
            Add the first entry
          </button>
        </div>
      ) : (
        <>
          {/* Filters */}
          <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
            <div className="relative flex-1 min-w-[240px]">
              <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search entity name or parent…"
                className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
              />
            </div>
            <select
              value={relationshipFilter}
              onChange={(e) => setRelationshipFilter(e.target.value as RelationshipType | 'All')}
              className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
            >
              <option value="All">All relationship types</option>
              {(['Managing Member', 'Member', 'Owner', 'Subsidiary', 'Beneficial Owner'] as RelationshipType[]).map((r) => (
                <option key={r} value={r}>{r}</option>
              ))}
            </select>
            <select
              value={propertyFilter}
              onChange={(e) => setPropertyFilter(e.target.value)}
              className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
            >
              <option value="All">All properties</option>
              {allProperties?.map((p) => (
                <option key={p.id} value={p.id}>{p.fields.Title}</option>
              ))}
            </select>
            {filtered.length !== data.length && (
              <span className="text-xs text-gray-500 px-1">{filtered.length} of {data.length}</span>
            )}
          </div>

          <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
                <tr>
                  <th className="px-4 py-3 text-left">Entity</th>
                  <th className="px-4 py-3 text-left">Relationship</th>
                  <th className="px-4 py-3 text-right">%</th>
                  <th className="px-4 py-3 text-left">Parent Entity</th>
                  <th className="px-4 py-3 text-left">Linked Property</th>
                  <th className="px-4 py-3 text-left">Effective</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {filtered.map((o) => {
                  const linkedProperty = o.fields.LinkedPropertyLookupId
                    ? propertiesById.get(String(o.fields.LinkedPropertyLookupId))
                    : null;
                  return (
                    <tr
                      key={o.id}
                      onClick={() => navigate(`/ownership/${o.id}`)}
                      className="hover:bg-gray-50 transition-colors cursor-pointer"
                    >
                      <td className="px-4 py-3 font-medium text-gray-900">{o.fields.Title}</td>
                      <td className="px-4 py-3">
                        {o.fields.RelationshipType ? (
                          <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${RELATIONSHIP_STYLES[o.fields.RelationshipType] || 'bg-gray-100'}`}>
                            {o.fields.RelationshipType}
                          </span>
                        ) : '—'}
                      </td>
                      <td className="px-4 py-3 text-right font-mono-data">
                        {o.fields.OwnershipPercent != null ? `${o.fields.OwnershipPercent}%` : '—'}
                      </td>
                      <td className="px-4 py-3 text-gray-700 text-xs">{o.fields.ParentEntity || '—'}</td>
                      <td className="px-4 py-3 text-gray-700 text-xs">{linkedProperty?.fields.Title || '—'}</td>
                      <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                        {o.fields.EffectiveDate ? new Date(o.fields.EffectiveDate).toLocaleDateString() : '—'}
                      </td>
                    </tr>
                  );
                })}
                {filtered.length === 0 && (
                  <tr>
                    <td colSpan={6} className="px-4 py-8 text-center text-gray-500 text-sm">
                      No ownership records match the current filters.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
