import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type OwnerCommunication,
  type Property,
  type Owner,
  type CommType,
  type CommStatus,
  type CommunicationPropertyLink,
  type CommunicationOwnerLink,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateET } from '../lib/dates';
import { LogCommunicationModal } from '../components/LogCommunicationModal';

const TYPE_STYLES: Record<CommType, string> = {
  Email: 'bg-blue-100 text-blue-800',
  Phone: 'bg-purple-100 text-purple-800',
  Meeting: 'bg-teal-100 text-teal-800',
  SMS: 'bg-amber-100 text-amber-800',
  Other: 'bg-gray-100 text-gray-700',
};

const STATUS_STYLES: Record<CommStatus, string> = {
  Open: 'bg-amber-100 text-amber-800',
  Closed: 'bg-gray-100 text-gray-500',
};

export function OwnerCommunicationsPage() {
  const navigate = useNavigate();
  const comms = useSharePointList<OwnerCommunication>(LIST_NAMES.Communications, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const propertyLinks = useSharePointList<CommunicationPropertyLink>(LIST_NAMES.CommunicationPropertyLinks, { top: 2000 });
  const ownerLinks = useSharePointList<CommunicationOwnerLink>(LIST_NAMES.CommunicationOwnerLinks, { top: 2000 });

  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState<CommType | 'All'>('All');
  const [statusFilter, setStatusFilter] = useState<CommStatus | 'All'>('All');
  const [logModalOpen, setLogModalOpen] = useState(false);

  const loading = comms.loading || properties.loading || owners.loading;
  const error = comms.error || properties.error || owners.error;

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const ownersById = useMemo(() => {
    if (!owners.data) return new Map<string, Owner>();
    return new Map(owners.data.map((o) => [String(o.id), o]));
  }, [owners.data]);

  // commId → linked entities. Includes the legacy single CommPropertyLookupId
  // / CommOwnerLookupId so existing rows still surface their primary link.
  const propertyTitlesByComm = useMemo(() => {
    const map = new Map<string, { id: string; title: string }[]>();
    const add = (cId: string, pId: string) => {
      const title = propertiesById.get(pId)?.fields.Title;
      if (!title) return;
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.some((e) => e.id === pId)) list.push({ id: pId, title });
    };
    (propertyLinks.data ?? []).forEach((l) => {
      if (l.fields.CommLookupId && l.fields.PropertyLookupId) {
        add(String(l.fields.CommLookupId), String(l.fields.PropertyLookupId));
      }
    });
    (comms.data ?? []).forEach((c) => {
      if (c.fields.CommPropertyLookupId) {
        add(String(c.id), String(c.fields.CommPropertyLookupId));
      }
    });
    return map;
  }, [propertyLinks.data, comms.data, propertiesById]);

  const ownerTitlesByComm = useMemo(() => {
    const map = new Map<string, { id: string; title: string }[]>();
    const add = (cId: string, oId: string) => {
      const title = ownersById.get(oId)?.fields.Title;
      if (!title) return;
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.some((e) => e.id === oId)) list.push({ id: oId, title });
    };
    (ownerLinks.data ?? []).forEach((l) => {
      if (l.fields.CommLookupId && l.fields.OwnerLookupId) {
        add(String(l.fields.CommLookupId), String(l.fields.OwnerLookupId));
      }
    });
    (comms.data ?? []).forEach((c) => {
      if (c.fields.CommOwnerLookupId) {
        add(String(c.id), String(c.fields.CommOwnerLookupId));
      }
    });
    return map;
  }, [ownerLinks.data, comms.data, ownersById]);

  const filtered = useMemo(() => {
    if (!comms.data) return [];
    return comms.data
      .filter((c) => {
        const f = c.fields;
        if (search) {
          const propTitles = (propertyTitlesByComm.get(String(c.id)) ?? []).map((p) => p.title).join(' ');
          const ownerTitles = (ownerTitlesByComm.get(String(c.id)) ?? []).map((o) => o.title).join(' ');
          const hay = `${f.Title ?? ''} ${propTitles} ${ownerTitles} ${f.CommParticipants ?? ''} ${f.CommNotes ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (typeFilter !== 'All' && f.CommType !== typeFilter) return false;
        if (statusFilter !== 'All' && f.CommStatus !== statusFilter) return false;
        return true;
      })
      .sort((a, b) => {
        const da = a.fields.CommDate ? new Date(a.fields.CommDate).getTime() : 0;
        const db = b.fields.CommDate ? new Date(b.fields.CommDate).getTime() : 0;
        return db - da;
      });
  }, [comms.data, search, typeFilter, statusFilter, propertyTitlesByComm, ownerTitlesByComm]);

  const stats = useMemo(() => {
    if (!comms.data) return null;
    const now = new Date();
    const thisYear = now.getFullYear();
    const thisMonth = now.getMonth();
    const ytd = comms.data.filter((c) => {
      if (!c.fields.CommDate) return false;
      return new Date(c.fields.CommDate).getFullYear() === thisYear;
    });
    const needsFollowup = comms.data.filter((c) => {
      if (c.fields.CommStatus === 'Closed') return false;
      if (!c.fields.CommResponseDue) return false;
      return new Date(c.fields.CommResponseDue).getTime() >= Date.now() - 1000 * 60 * 60 * 24 * 30;
    });
    const emailsThisMonth = comms.data.filter((c) => {
      if (c.fields.CommType !== 'Email') return false;
      if (!c.fields.CommDate) return false;
      const d = new Date(c.fields.CommDate);
      return d.getFullYear() === thisYear && d.getMonth() === thisMonth;
    });
    const callsThisMonth = comms.data.filter((c) => {
      if (c.fields.CommType !== 'Phone') return false;
      if (!c.fields.CommDate) return false;
      const d = new Date(c.fields.CommDate);
      return d.getFullYear() === thisYear && d.getMonth() === thisMonth;
    });
    return {
      ytd: ytd.length,
      needsFollowup: needsFollowup.length,
      emailsThisMonth: emailsThisMonth.length,
      callsThisMonth: callsThisMonth.length,
    };
  }, [comms.data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Owner Communications</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading communications…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Owner Communications</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (!comms.data || !stats) return null;

  const handleLogSuccess = () => {
    setLogModalOpen(false);
    comms.refetch?.();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Owner Communications</h1>
          <p className="text-sm text-gray-500 mt-1">
            Non-DOR communications — owner emails, calls, meetings, vendor calls, team meetings.
            Single timeline of every conversation. Setting a follow-up due date auto-creates an Outstanding Item.
          </p>
        </div>
        <button
          onClick={() => setLogModalOpen(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors flex-shrink-0"
        >
          <Icon name="plus" size={16} />
          Log Communication
        </button>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total YTD" value={stats.ytd} />
        <KPI label="Needs Follow-up" value={stats.needsFollowup} accent={stats.needsFollowup > 0 ? 'warning' : 'default'} />
        <KPI label="Emails This Month" value={stats.emailsThisMonth} />
        <KPI label="Calls This Month" value={stats.callsThisMonth} />
      </div>

      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search subject, property, owner, participants…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={typeFilter}
          onChange={(e) => setTypeFilter(e.target.value as CommType | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All types</option>
          {(Object.keys(TYPE_STYLES) as CommType[]).map((t) => (<option key={t} value={t}>{t}</option>))}
        </select>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as CommStatus | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All statuses</option>
          <option value="Open">Open</option>
          <option value="Closed">Closed</option>
        </select>
        {filtered.length !== comms.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {comms.data.length}</span>
        )}
      </div>

      {comms.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No communications logged yet</p>
          <p className="text-sm text-blue-800 mb-4">
            Click Log Communication to record your first interaction.
          </p>
          <button
            onClick={() => setLogModalOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium inline-flex items-center gap-2"
          >
            <Icon name="plus" size={16} />
            Log First Communication
          </button>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Date</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Property / Owner</th>
                <th className="px-4 py-3 text-left">Participants</th>
                <th className="px-4 py-3 text-left">Follow-up</th>
                <th className="px-4 py-3 text-left">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((c) => {
                const linkedProps = propertyTitlesByComm.get(String(c.id)) ?? [];
                const linkedOwners = ownerTitlesByComm.get(String(c.id)) ?? [];
                const responseDue = c.fields.CommResponseDue ? new Date(c.fields.CommResponseDue) : null;
                const isOverdue =
                  responseDue && responseDue.getTime() < Date.now() && c.fields.CommStatus !== 'Closed';
                return (
                  <tr
                    key={c.id}
                    onClick={() => navigate(`/comms/${c.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {c.fields.CommDate ? formatDateET(c.fields.CommDate) : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.CommType ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${TYPE_STYLES[c.fields.CommType]}`}>
                          {c.fields.CommDirection === 'Inbound' ? '← ' : c.fields.CommDirection === 'Outbound' ? '→ ' : ''}
                          {c.fields.CommType}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                    <td className="px-4 py-3 text-xs">
                      {linkedProps.length === 0 && linkedOwners.length === 0 ? (
                        <span className="text-gray-400 italic">unlinked</span>
                      ) : (
                        <div className="space-y-1">
                          {linkedProps.length > 0 && (
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
                          {linkedOwners.length > 0 && (
                            <div className="flex flex-wrap gap-1">
                              {linkedOwners.slice(0, 3).map((o) => (
                                <span key={o.id} className="px-1.5 py-0.5 rounded bg-gold-50 text-gold-900 text-[11px] font-medium">
                                  {o.title}
                                </span>
                              ))}
                              {linkedOwners.length > 3 && (
                                <span className="text-[11px] text-gray-500">+{linkedOwners.length - 3} more</span>
                              )}
                            </div>
                          )}
                        </div>
                      )}
                    </td>
                    <td className="px-4 py-3 text-xs text-gray-600 max-w-xs truncate">
                      {c.fields.CommParticipants || '—'}
                    </td>
                    <td className={`px-4 py-3 font-mono-data text-xs ${isOverdue ? 'text-error font-semibold' : 'text-gray-700'}`}>
                      {responseDue ? formatDateET(responseDue) : '—'}
                      {isOverdue && ' ⚠'}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.CommStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[c.fields.CommStatus]}`}>
                          {c.fields.CommStatus}
                        </span>
                      ) : '—'}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {logModalOpen && (
        <LogCommunicationModal
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
