import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Correspondence,
  type Property,
  type LetterType,
  type CorrespondenceDirection,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { LogLetterModal } from '../components/LogLetterModal';

const LETTER_TYPE_STYLES: Record<LetterType, string> = {
  'Initial Acknowledgment': 'bg-blue-100 text-blue-800',
  'Additional Info Request': 'bg-amber-100 text-amber-800',
  'Org Chart Request': 'bg-amber-100 text-amber-800',
  'Approval': 'bg-green-100 text-green-800',
  'Denial': 'bg-red-100 text-red-800',
  'Withdrawal Notice': 'bg-gray-100 text-gray-700',
  'Refund Notice': 'bg-teal-100 text-teal-800',
  'Other': 'bg-gray-100 text-gray-700',
};

const DIRECTION_STYLES: Record<CorrespondenceDirection, string> = {
  'Inbound (from DOR)': 'bg-purple-100 text-purple-800',
  'Outbound (to DOR)': 'bg-blue-100 text-blue-800',
};

export function CorrespondencePage() {
  const navigate = useNavigate();
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState<LetterType | 'All'>('All');
  const [directionFilter, setDirectionFilter] = useState<CorrespondenceDirection | 'All'>('All');
  const [logModalOpen, setLogModalOpen] = useState(false);

  const loading = correspondence.loading || properties.loading;
  const error = correspondence.error || properties.error;

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const filtered = useMemo(() => {
    if (!correspondence.data) return [];
    return correspondence.data
      .filter((c) => {
        const f = c.fields;
        if (search) {
          const propName = f.PropertyLookupId
            ? propertiesById.get(String(f.PropertyLookupId))?.fields.Title ?? ''
            : '';
          const hay = `${f.Title ?? ''} ${propName} ${f.RequestSummary ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (typeFilter !== 'All' && f.LetterType !== typeFilter) return false;
        if (directionFilter !== 'All' && f.Direction !== directionFilter) return false;
        return true;
      })
      .sort((a, b) => {
        const da = a.fields.DateReceived ? new Date(a.fields.DateReceived).getTime() : 0;
        const db = b.fields.DateReceived ? new Date(b.fields.DateReceived).getTime() : 0;
        return db - da;
      });
  }, [correspondence.data, search, typeFilter, directionFilter, propertiesById]);

  const stats = useMemo(() => {
    if (!correspondence.data) return null;
    const today = new Date();
    const responseDueSoon = correspondence.data.filter((c) => {
      if (!c.fields.ResponseDue) return false;
      const due = new Date(c.fields.ResponseDue);
      const days = (due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24);
      return days >= 0 && days <= 14 && !c.fields.DateResponded;
    });
    const overdue = correspondence.data.filter((c) => {
      if (!c.fields.ResponseDue) return false;
      const due = new Date(c.fields.ResponseDue);
      return due.getTime() < today.getTime() && !c.fields.DateResponded;
    });
    return {
      total: correspondence.data.length,
      ytd: correspondence.data.filter((c) => {
        if (!c.fields.DateReceived) return false;
        return new Date(c.fields.DateReceived).getFullYear() === today.getFullYear();
      }).length,
      responseDueSoon: responseDueSoon.length,
      overdue: overdue.length,
    };
  }, [correspondence.data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">DOR Correspondence</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading correspondence…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">DOR Correspondence</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load correspondence</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (!correspondence.data || !stats) return null;

  const handleLogSuccess = () => {
    setLogModalOpen(false);
    correspondence.refetch?.();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">DOR Correspondence</h1>
          <p className="text-sm text-gray-500 mt-1">
            Every formal letter to and from SC/NC Department of Revenue. Logging a letter cascades to the related submittal, an Outstanding Item with deadline, and a filed document.
          </p>
        </div>
        <button
          onClick={() => setLogModalOpen(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors flex-shrink-0"
        >
          <Icon name="plus" size={16} />
          Log Letter
        </button>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total Letters" value={stats.total} />
        <KPI label="YTD" value={stats.ytd} />
        <KPI label="Response Due ≤14d" value={stats.responseDueSoon} accent={stats.responseDueSoon > 0 ? 'warning' : 'default'} />
        <KPI label="Overdue Response" value={stats.overdue} accent={stats.overdue > 0 ? 'danger' : 'default'} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search subject, property, summary…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={typeFilter}
          onChange={(e) => setTypeFilter(e.target.value as LetterType | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All types</option>
          {(Object.keys(LETTER_TYPE_STYLES) as LetterType[]).map((t) => (
            <option key={t} value={t}>{t}</option>
          ))}
        </select>
        <select
          value={directionFilter}
          onChange={(e) => setDirectionFilter(e.target.value as CorrespondenceDirection | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All directions</option>
          <option value="Inbound (from DOR)">Inbound (from DOR)</option>
          <option value="Outbound (to DOR)">Outbound (to DOR)</option>
        </select>
        {filtered.length !== correspondence.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {correspondence.data.length}</span>
        )}
      </div>

      {/* Table */}
      {correspondence.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No correspondence logged yet</p>
          <p className="text-sm text-blue-800 mb-4">
            Click <strong>Log Letter</strong> to record your first DOR communication. The system will create the
            correspondence record, an Outstanding Item if you set a response deadline, and update the related submittal.
          </p>
          <button
            onClick={() => setLogModalOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium inline-flex items-center gap-2 transition-colors"
          >
            <Icon name="plus" size={16} />
            Log First Letter
          </button>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Date</th>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Direction</th>
                <th className="px-4 py-3 text-left">Response Due</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((c) => {
                const property = c.fields.PropertyLookupId
                  ? propertiesById.get(String(c.fields.PropertyLookupId))
                  : null;
                const responseDue = c.fields.ResponseDue ? new Date(c.fields.ResponseDue) : null;
                const isOverdue = responseDue && responseDue.getTime() < Date.now() && !c.fields.DateResponded;
                return (
                  <tr
                    key={c.id}
                    onClick={() => navigate(`/correspondence/${c.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {c.fields.DateReceived ? new Date(c.fields.DateReceived).toLocaleDateString() : '—'}
                    </td>
                    <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                    <td className="px-4 py-3 text-gray-700 text-xs">
                      {property?.fields.Title ?? <span className="text-gray-400 italic">unlinked</span>}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.LetterType ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${LETTER_TYPE_STYLES[c.fields.LetterType]}`}>
                          {c.fields.LetterType}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.Direction ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${DIRECTION_STYLES[c.fields.Direction]}`}>
                          {c.fields.Direction.replace(' (from DOR)', ' ').replace(' (to DOR)', ' ').trim()}
                        </span>
                      ) : '—'}
                    </td>
                    <td className={`px-4 py-3 font-mono-data text-xs ${isOverdue ? 'text-error font-semibold' : 'text-gray-700'}`}>
                      {responseDue ? responseDue.toLocaleDateString() : '—'}
                      {isOverdue && ' ⚠'}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Log Letter modal */}
      {logModalOpen && (
        <LogLetterModal
          onClose={() => setLogModalOpen(false)}
          onSuccess={handleLogSuccess}
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
