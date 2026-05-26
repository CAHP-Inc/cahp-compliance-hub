import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Correspondence,
  type Property,
  type LetterType,
  type CorrespondenceDirection,
  type CorrespondencePropertyLink,
  type CorrChannel,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateET } from '../lib/dates';
import { LogLetterModal } from '../components/LogLetterModal';
import { LogDORCommModal } from '../components/LogDORCommModal';

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

const CHANNEL_STYLES: Record<CorrChannel, string> = {
  Letter:  'bg-indigo-100 text-indigo-800',
  Email:   'bg-blue-100 text-blue-800',
  Phone:   'bg-purple-100 text-purple-800',
  Meeting: 'bg-teal-100 text-teal-800',
  Other:   'bg-gray-100 text-gray-700',
};

export function CorrespondencePage() {
  const navigate = useNavigate();
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const propertyLinks = useSharePointList<CorrespondencePropertyLink>(LIST_NAMES.CorrespondencePropertyLinks, { top: 2000 });

  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState<LetterType | 'All'>('All');
  const [directionFilter, setDirectionFilter] = useState<CorrespondenceDirection | 'All'>('All');
  const [channelFilter, setChannelFilter] = useState<CorrChannel | 'All'>('All');
  const [logLetterOpen, setLogLetterOpen] = useState(false);
  const [logCommOpen, setLogCommOpen] = useState(false);

  const loading = correspondence.loading || properties.loading;
  const error = correspondence.error || properties.error;

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  // corrId → list of {id, title} pairs (junction + legacy single PropertyLookupId)
  const propertyTitlesByCorr = useMemo(() => {
    const map = new Map<string, { id: string; title: string }[]>();
    const add = (cId: string, pId: string) => {
      const title = propertiesById.get(pId)?.fields.Title;
      if (!title) return;
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.some((e) => e.id === pId)) list.push({ id: pId, title });
    };
    (propertyLinks.data ?? []).forEach((l) => {
      if (l.fields.CorrLookupId && l.fields.PropertyLookupId) {
        add(String(l.fields.CorrLookupId), String(l.fields.PropertyLookupId));
      }
    });
    (correspondence.data ?? []).forEach((c) => {
      if (c.fields.PropertyLookupId) {
        add(String(c.id), String(c.fields.PropertyLookupId));
      }
    });
    return map;
  }, [propertyLinks.data, correspondence.data, propertiesById]);

  const filtered = useMemo(() => {
    if (!correspondence.data) return [];
    return correspondence.data
      .filter((c) => {
        const f = c.fields;
        if (search) {
          const propTitles = (propertyTitlesByCorr.get(String(c.id)) ?? []).map((p) => p.title).join(' ');
          const hay = `${f.Title ?? ''} ${propTitles} ${f.RequestSummary ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (typeFilter !== 'All' && f.LetterType !== typeFilter) return false;
        if (directionFilter !== 'All' && f.Direction !== directionFilter) return false;
        if (channelFilter !== 'All') {
          // Default missing channel to 'Letter' so legacy rows still match when filtering
          const channel = (f.CorrChannel ?? 'Letter') as CorrChannel;
          if (channel !== channelFilter) return false;
        }
        return true;
      })
      .sort((a, b) => {
        const da = a.fields.DateReceived ? new Date(a.fields.DateReceived).getTime() : 0;
        const db = b.fields.DateReceived ? new Date(b.fields.DateReceived).getTime() : 0;
        return db - da;
      });
  }, [correspondence.data, search, typeFilter, directionFilter, channelFilter, propertyTitlesByCorr]);

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

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">DOR Correspondence</h1>
          <p className="text-sm text-gray-500 mt-1">
            Formal letters AND general communications (calls, emails, meetings) with SC/NC Department of Revenue.
            Formal letters cascade to Outstanding Items + submittal status; general comms are notes-only.
          </p>
        </div>
        <div className="flex items-center gap-2 flex-shrink-0">
          <button
            onClick={() => setLogCommOpen(true)}
            className="bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 px-3 py-2 rounded-md font-medium flex items-center gap-2 transition-colors text-sm"
            title="Log a phone call, email, or meeting with DOR"
          >
            <Icon name="plus" size={14} />
            Log Comm
          </button>
          <button
            onClick={() => setLogLetterOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors"
            title="Log a formal DOR letter (with cascading Outstanding Items)"
          >
            <Icon name="plus" size={16} />
            Log Letter
          </button>
        </div>
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
        <select
          value={channelFilter}
          onChange={(e) => setChannelFilter(e.target.value as CorrChannel | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All channels</option>
          <option value="Letter">Letter</option>
          <option value="Email">Email</option>
          <option value="Phone">Phone</option>
          <option value="Meeting">Meeting</option>
          <option value="Other">Other</option>
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
          <div className="flex items-center justify-center gap-2 flex-wrap">
            <button
              onClick={() => setLogCommOpen(true)}
              className="bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-md font-medium inline-flex items-center gap-2"
            >
              <Icon name="plus" size={16} />
              Log Communication
            </button>
            <button
              onClick={() => setLogLetterOpen(true)}
              className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium inline-flex items-center gap-2 transition-colors"
            >
              <Icon name="plus" size={16} />
              Log First Letter
            </button>
          </div>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Date</th>
                <th className="px-4 py-3 text-left">Channel</th>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Properties</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Direction</th>
                <th className="px-4 py-3 text-left">Response Due</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((c) => {
                const linkedProps = propertyTitlesByCorr.get(String(c.id)) ?? [];
                const responseDue = c.fields.ResponseDue ? new Date(c.fields.ResponseDue) : null;
                const isOverdue = responseDue && responseDue.getTime() < Date.now() && !c.fields.DateResponded;
                const channel = (c.fields.CorrChannel ?? 'Letter') as CorrChannel;
                return (
                  <tr
                    key={c.id}
                    onClick={() => navigate(`/correspondence/${c.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {c.fields.DateReceived ? formatDateET(c.fields.DateReceived) : '—'}
                    </td>
                    <td className="px-4 py-3">
                      <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${CHANNEL_STYLES[channel]}`}>
                        {channel}
                      </span>
                    </td>
                    <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                    <td className="px-4 py-3 text-xs">
                      {linkedProps.length === 0 ? (
                        <span className="text-gray-400 italic">unlinked</span>
                      ) : (
                        <div className="flex flex-wrap gap-1">
                          {linkedProps.slice(0, 4).map((p) => (
                            <span key={p.id} className="px-1.5 py-0.5 rounded bg-teal-50 text-teal-800 text-[11px] font-medium">
                              {p.title}
                            </span>
                          ))}
                          {linkedProps.length > 4 && (
                            <span className="text-[11px] text-gray-500">+{linkedProps.length - 4} more</span>
                          )}
                        </div>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.LetterType ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${LETTER_TYPE_STYLES[c.fields.LetterType]}`}>
                          {c.fields.LetterType}
                        </span>
                      ) : <span className="text-gray-400">—</span>}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.Direction ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${DIRECTION_STYLES[c.fields.Direction]}`}>
                          {c.fields.Direction.replace(' (from DOR)', ' ').replace(' (to DOR)', ' ').trim()}
                        </span>
                      ) : '—'}
                    </td>
                    <td className={`px-4 py-3 font-mono-data text-xs ${isOverdue ? 'text-error font-semibold' : 'text-gray-700'}`}>
                      {responseDue ? formatDateET(responseDue) : '—'}
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
      {logLetterOpen && (
        <LogLetterModal
          onClose={() => setLogLetterOpen(false)}
          onSuccess={() => {
            setLogLetterOpen(false);
            correspondence.refetch?.();
            propertyLinks.refetch?.();
          }}
        />
      )}

      {/* Log General DOR Communication modal */}
      {logCommOpen && (
        <LogDORCommModal
          onClose={() => setLogCommOpen(false)}
          onSuccess={() => {
            setLogCommOpen(false);
            correspondence.refetch?.();
            propertyLinks.refetch?.();
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
