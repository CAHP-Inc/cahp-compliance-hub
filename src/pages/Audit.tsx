import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type AuditLog,
  type AuditAction,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const ACTION_STYLES: Record<AuditAction, string> = {
  CREATE: 'bg-green-100 text-green-800',
  UPDATE: 'bg-blue-100 text-blue-800',
  DELETE: 'bg-red-100 text-red-800',
};

const ENTITY_ROUTES: Record<string, string> = {
  Property: 'properties',
  'Compliance Deadline': 'compliance',
  'Ownership Record': 'ownership',
};

export function Audit() {
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useSharePointList<AuditLog>(LIST_NAMES.AuditLog, {
    top: 500,
  });

  const [search, setSearch] = useState('');
  const [actionFilter, setActionFilter] = useState<AuditAction | 'All'>('All');
  const [entityFilter, setEntityFilter] = useState<string>('All');
  const [expandedId, setExpandedId] = useState<string | null>(null);

  const entityTypes = useMemo(() => {
    if (!data) return [];
    const set = new Set<string>();
    data.forEach((row) => {
      if (row.fields.EntityType) set.add(row.fields.EntityType);
    });
    return Array.from(set).sort();
  }, [data]);

  const filtered = useMemo(() => {
    if (!data) return [];
    return data
      .filter((row) => {
        const f = row.fields;
        if (search) {
          const hay = `${f.Title ?? ''} ${f.EntityTitle ?? ''} ${f.ChangeSummary ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (actionFilter !== 'All' && f.Action !== actionFilter) return false;
        if (entityFilter !== 'All' && f.EntityType !== entityFilter) return false;
        return true;
      })
      // Sort newest first by Created column (from SharePoint built-in)
      .sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
  }, [data, search, actionFilter, entityFilter]);

  const stats = useMemo(() => {
    if (!data) return null;
    const since = Date.now() - 24 * 60 * 60 * 1000;
    return {
      total: data.length,
      last24h: data.filter((r) => new Date(r.createdDateTime).getTime() >= since).length,
      creates: data.filter((r) => r.fields.Action === 'CREATE').length,
      updates: data.filter((r) => r.fields.Action === 'UPDATE').length,
      deletes: data.filter((r) => r.fields.Action === 'DELETE').length,
    };
  }, [data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Audit Log</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading audit log…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Audit Log</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Failed to load audit log
          </div>
          <p className="text-sm text-red-700 mb-3 font-mono-data">{error.message}</p>
          <p className="text-xs text-red-600 mb-3">
            Likely cause: AuditLog list doesn't exist yet. Run the provisioning script.
          </p>
          <button onClick={refetch} className="text-sm text-teal-700 hover:text-teal-900 font-medium underline">
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data || !stats) return null;

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Audit Log</h1>
        <p className="text-sm text-gray-500 mt-1">
          Append-only record of every CRUD operation across all modules ·{' '}
          {stats.total} total entries · {stats.last24h} in the last 24 hours
        </p>
      </div>

      {data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-blue-100 text-blue-700 mb-3">
            <Icon name="history" size={24} />
          </div>
          <h2 className="text-base font-semibold text-blue-900 mb-1">No audit entries yet</h2>
          <p className="text-sm text-blue-800 max-w-md mx-auto">
            Every create, update, or delete you make from this point forward will be recorded here.
            Try editing a property or marking a deadline complete to see the audit log in action.
          </p>
        </div>
      ) : (
        <>
          {/* KPI cards */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
            <KPICard label="Total Entries" value={stats.total} />
            <KPICard label="Last 24h" value={stats.last24h} />
            <KPICard label="Creates" value={stats.creates} accent="success" />
            <KPICard label="Updates" value={stats.updates} />
            <KPICard label="Deletes" value={stats.deletes} accent="danger" />
          </div>

          {/* Filters */}
          <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
            <div className="relative flex-1 min-w-[240px]">
              <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search event, entity, or change summary…"
                className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
              />
            </div>
            <select
              value={actionFilter}
              onChange={(e) => setActionFilter(e.target.value as AuditAction | 'All')}
              className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
            >
              <option value="All">All actions</option>
              <option value="CREATE">CREATE</option>
              <option value="UPDATE">UPDATE</option>
              <option value="DELETE">DELETE</option>
            </select>
            <select
              value={entityFilter}
              onChange={(e) => setEntityFilter(e.target.value)}
              className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
            >
              <option value="All">All entity types</option>
              {entityTypes.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
            {filtered.length !== data.length && (
              <span className="text-xs text-gray-500 px-1">{filtered.length} of {data.length}</span>
            )}
          </div>

          {/* Table */}
          <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
                <tr>
                  <th className="px-4 py-3 text-left w-44">When</th>
                  <th className="px-4 py-3 text-left w-24">Action</th>
                  <th className="px-4 py-3 text-left w-40">Entity</th>
                  <th className="px-4 py-3 text-left">Affected Record</th>
                  <th className="px-4 py-3 text-left">Change Summary</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {filtered.map((row) => {
                  const f = row.fields;
                  const isExpanded = expandedId === row.id;
                  const route = f.EntityType ? ENTITY_ROUTES[f.EntityType] : null;
                  const canDrillIn = route && f.EntityId && f.Action !== 'DELETE';
                  return (
                    <>
                      <tr
                        key={row.id}
                        onClick={() => setExpandedId(isExpanded ? null : row.id)}
                        className="hover:bg-gray-50 transition-colors cursor-pointer"
                      >
                        <td className="px-4 py-3 text-gray-700 font-mono-data text-xs whitespace-nowrap">
                          {formatTimestamp(row.createdDateTime)}
                        </td>
                        <td className="px-4 py-3">
                          {f.Action && (
                            <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${ACTION_STYLES[f.Action] || 'bg-gray-100'}`}>
                              {f.Action}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-gray-700 text-xs">{f.EntityType || '—'}</td>
                        <td className="px-4 py-3 text-gray-900">
                          {canDrillIn ? (
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                navigate(`/${route}/${f.EntityId}`);
                              }}
                              className="text-teal-700 hover:text-teal-900 font-medium underline text-left"
                            >
                              {f.EntityTitle}
                            </button>
                          ) : (
                            <span className={f.Action === 'DELETE' ? 'text-gray-500 line-through' : ''}>
                              {f.EntityTitle || '—'}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-gray-700 text-xs whitespace-pre-wrap line-clamp-2">
                          {f.ChangeSummary || '—'}
                        </td>
                      </tr>
                      {isExpanded && (
                        <tr key={`${row.id}-detail`} className="bg-gray-50">
                          <td colSpan={5} className="px-4 py-4">
                            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                              <div>
                                <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-1">
                                  Full Change Summary
                                </div>
                                <pre className="text-xs whitespace-pre-wrap font-mono-data text-gray-700 bg-white p-3 rounded border border-gray-200">
                                  {f.ChangeSummary || '(none)'}
                                </pre>
                              </div>
                              <div>
                                <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-1">
                                  Event
                                </div>
                                <div className="text-sm text-gray-700 bg-white p-3 rounded border border-gray-200 space-y-1">
                                  <div><span className="text-gray-500">Title:</span> {f.Title}</div>
                                  <div><span className="text-gray-500">Entity ID:</span> <code className="font-mono-data text-xs">{f.EntityId || '—'}</code></div>
                                  <div><span className="text-gray-500">Audit ID:</span> <code className="font-mono-data text-xs">{row.id}</code></div>
                                  <div><span className="text-gray-500">Timestamp:</span> <code className="font-mono-data text-xs">{row.createdDateTime}</code></div>
                                </div>
                              </div>
                              {f.BeforeJSON && (
                                <div>
                                  <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-1">
                                    Before
                                  </div>
                                  <pre className="text-[11px] whitespace-pre-wrap font-mono-data text-gray-600 bg-white p-3 rounded border border-gray-200 max-h-64 overflow-y-auto">
                                    {f.BeforeJSON}
                                  </pre>
                                </div>
                              )}
                              {f.AfterJSON && (
                                <div>
                                  <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-1">
                                    After
                                  </div>
                                  <pre className="text-[11px] whitespace-pre-wrap font-mono-data text-gray-600 bg-white p-3 rounded border border-gray-200 max-h-64 overflow-y-auto">
                                    {f.AfterJSON}
                                  </pre>
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </>
                  );
                })}
                {filtered.length === 0 && (
                  <tr>
                    <td colSpan={5} className="px-4 py-8 text-center text-gray-500 text-sm">
                      No audit entries match the current filters.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <p className="text-xs text-gray-400 mt-4 text-center">
            Click any row to expand · click the affected record name to drill into it
          </p>
        </>
      )}
    </div>
  );
}

function KPICard({
  label,
  value,
  accent,
}: {
  label: string;
  value: number;
  accent?: 'success' | 'danger';
}) {
  const accentClass =
    accent === 'success' ? 'text-success' : accent === 'danger' ? 'text-error' : 'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}

function formatTimestamp(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  });
}
