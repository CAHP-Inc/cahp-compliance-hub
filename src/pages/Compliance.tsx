import { useState, useMemo } from 'react';
import {
  useSharePointList,
  LIST_NAMES,
  type ComplianceDeadline,
  type DeadlineStatus,
  type DeadlineType,
  type ResponsibleParty,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const STATUS_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800',
  'In Progress': 'bg-amber-100 text-amber-800',
  Completed: 'bg-green-100 text-green-800',
  Overdue: 'bg-red-100 text-red-800',
  Missed: 'bg-red-200 text-red-900',
};

export function Compliance() {
  const { data, loading, error, refetch } = useSharePointList<ComplianceDeadline>(
    LIST_NAMES.ComplianceDeadlines,
    { top: 500 }
  );

  const [statusFilter, setStatusFilter] = useState<DeadlineStatus | 'All'>('All');
  const [typeFilter, setTypeFilter] = useState<DeadlineType | 'All'>('All');
  const [ownerFilter, setOwnerFilter] = useState<ResponsibleParty | 'All'>('All');

  const filtered = useMemo(() => {
    if (!data) return [];
    const result = data.filter((d) => {
      const f = d.fields;
      if (statusFilter !== 'All' && f.DeadlineStatus !== statusFilter) return false;
      if (typeFilter !== 'All' && f.DeadlineType !== typeFilter) return false;
      if (ownerFilter !== 'All' && f.ResponsibleParty !== ownerFilter) return false;
      return true;
    });
    // Sort: overdue/missed first, then by due date ascending
    return result.sort((a, b) => {
      const urgencyA = a.fields.DeadlineStatus === 'Overdue' || a.fields.DeadlineStatus === 'Missed' ? 0 : 1;
      const urgencyB = b.fields.DeadlineStatus === 'Overdue' || b.fields.DeadlineStatus === 'Missed' ? 0 : 1;
      if (urgencyA !== urgencyB) return urgencyA - urgencyB;
      const dateA = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Number.MAX_VALUE;
      const dateB = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Number.MAX_VALUE;
      return dateA - dateB;
    });
  }, [data, statusFilter, typeFilter, ownerFilter]);

  const stats = useMemo(() => {
    if (!data) return null;
    return {
      total: data.length,
      overdue: data.filter((d) => d.fields.DeadlineStatus === 'Overdue').length,
      upcoming: data.filter((d) => d.fields.DeadlineStatus === 'Upcoming').length,
      inProgress: data.filter((d) => d.fields.DeadlineStatus === 'In Progress').length,
      completed: data.filter((d) => d.fields.DeadlineStatus === 'Completed').length,
    };
  }, [data]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Compliance</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading deadlines from SharePoint…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Compliance</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Failed to load deadlines
          </div>
          <p className="text-sm text-red-700 mb-3 font-mono-data">{error.message}</p>
          <button
            onClick={refetch}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  if (!data || !stats) return null;

  return (
    <div>
      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Compliance</h1>
          <p className="text-sm text-gray-500 mt-1">
            {stats.total} deadlines tracked · {stats.overdue} overdue · {stats.upcoming} upcoming
          </p>
        </div>
        <button
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          disabled
          title="Create deadline ships in Phase 2"
        >
          <Icon name="plus" size={16} />
          New Deadline
        </button>
      </div>

      {/* KPI cards */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
        <KPICard label="Total" value={stats.total} />
        <KPICard label="Overdue" value={stats.overdue} accent="danger" />
        <KPICard label="Upcoming" value={stats.upcoming} />
        <KPICard label="In Progress" value={stats.inProgress} />
        <KPICard label="Completed" value={stats.completed} accent="success" />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <Select
          value={statusFilter}
          onChange={(v) => setStatusFilter(v as DeadlineStatus | 'All')}
          options={['All', 'Upcoming', 'In Progress', 'Completed', 'Overdue', 'Missed']}
        />
        <Select
          value={typeFilter}
          onChange={(v) => setTypeFilter(v as DeadlineType | 'All')}
          options={[
            'All',
            'IRS 990 Filing',
            'Annual Recertification',
            'Rent Roll Review',
            'AMI Cert Renewal',
            'State Compliance Report',
            'Property Tax Filing',
            'Operating Agreement Review',
            'Other',
          ]}
        />
        <Select
          value={ownerFilter}
          onChange={(v) => setOwnerFilter(v as ResponsibleParty | 'All')}
          options={['All', 'Brandy', 'Chris', 'Brian', 'John', 'Aljon', 'Other']}
        />
        {filtered.length !== data.length && (
          <span className="text-xs text-gray-500 px-1">
            {filtered.length} of {data.length}
          </span>
        )}
      </div>

      {/* Table */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Deadline</th>
              <th className="px-4 py-3 text-left">Type</th>
              <th className="px-4 py-3 text-left">Due Date</th>
              <th className="px-4 py-3 text-left">Status</th>
              <th className="px-4 py-3 text-left">Recurrence</th>
              <th className="px-4 py-3 text-left">Applies To</th>
              <th className="px-4 py-3 text-left">Owner</th>
              <th className="px-4 py-3 text-left">State</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map((d) => (
              <tr key={d.id} className="hover:bg-gray-50 transition-colors">
                <td className="px-4 py-3 font-medium text-gray-900">{d.fields.Title}</td>
                <td className="px-4 py-3 text-gray-700 text-xs">{d.fields.DeadlineType || '—'}</td>
                <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                  {formatDate(d.fields.DueDate) || '—'}
                </td>
                <td className="px-4 py-3">
                  {d.fields.DeadlineStatus ? (
                    <span
                      className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                        STATUS_STYLES[d.fields.DeadlineStatus] || 'bg-gray-100'
                      }`}
                    >
                      {d.fields.DeadlineStatus}
                    </span>
                  ) : (
                    '—'
                  )}
                </td>
                <td className="px-4 py-3 text-gray-700 text-xs">{d.fields.Recurrence || '—'}</td>
                <td className="px-4 py-3 text-gray-700 text-xs">{d.fields.AppliesTo || '—'}</td>
                <td className="px-4 py-3 text-gray-700">{d.fields.ResponsibleParty || '—'}</td>
                <td className="px-4 py-3 font-mono-data text-xs text-teal-700">
                  {d.fields.cahpState || '—'}
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={8} className="px-4 py-8 text-center text-gray-500 text-sm">
                  No deadlines match the current filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
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

function Select({
  value,
  onChange,
  options,
}: {
  value: string;
  onChange: (v: string) => void;
  options: readonly string[];
}) {
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

function formatDate(iso: string | undefined | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
