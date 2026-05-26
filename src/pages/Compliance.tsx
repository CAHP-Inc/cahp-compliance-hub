import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type ComplianceDeadline,
  type DeadlineStatus,
  type DeadlineType,
  type Property,
} from '../lib/sharepoint';
import { EASTERN_TZ } from '../lib/dates';
import { Icon } from '../components/ui/Icon';
import { NewDeadlineModal } from '../components/NewDeadlineModal';

const STATUS_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800',
  'In Progress': 'bg-amber-100 text-amber-800',
  Completed: 'bg-green-100 text-green-800',
  Overdue: 'bg-red-100 text-red-800',
  Missed: 'bg-red-200 text-red-900',
};

type QuickFilter = 'all' | 'overdue' | 'thisMonth' | 'thisYear' | 'ami' | 'sc' | 'nc';

export function Compliance() {
  const navigate = useNavigate();
  const { data, loading, error, refetch } = useSharePointList<ComplianceDeadline>(
    LIST_NAMES.ComplianceDeadlines,
    { top: 500 }
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [statusFilter, setStatusFilter] = useState<DeadlineStatus | 'All'>('All');
  const [typeFilter, setTypeFilter] = useState<DeadlineType | 'All'>('All');
  const [assigneeFilter, setAssigneeFilter] = useState<string>('All');
  const [quickFilter, setQuickFilter] = useState<QuickFilter>('all');
  const [newDeadlineOpen, setNewDeadlineOpen] = useState(false);

  // Build the assignee filter options dynamically from existing data
  // (covers free-text values like "Cogency Global" or "Owner – Deepak")
  const assigneeOptions = useMemo(() => {
    if (!data) return ['All'];
    const set = new Set<string>();
    data.forEach((d) => {
      const v = d.fields.AssignedTo || d.fields.ResponsibleParty;
      if (v) set.add(v);
    });
    return ['All', ...Array.from(set).sort()];
  }, [data]);

  const filtered = useMemo(() => {
    if (!data) return [];
    const now = Date.now();
    const monthEnd = new Date();
    monthEnd.setDate(monthEnd.getDate() + 30);
    const yearEnd = new Date();
    yearEnd.setFullYear(yearEnd.getFullYear() + 1);

    const result = data.filter((d) => {
      const f = d.fields;
      if (statusFilter !== 'All' && f.DeadlineStatus !== statusFilter) return false;
      if (typeFilter !== 'All' && f.DeadlineType !== typeFilter) return false;
      if (assigneeFilter !== 'All') {
        const v = f.AssignedTo || f.ResponsibleParty;
        if (v !== assigneeFilter) return false;
      }
      // Quick filters
      if (quickFilter === 'overdue' && f.DeadlineStatus !== 'Overdue' && f.DeadlineStatus !== 'Missed') return false;
      if (quickFilter === 'thisMonth') {
        if (!f.DueDate) return false;
        if (f.DeadlineStatus === 'Completed') return false;
        const due = new Date(f.DueDate).getTime();
        if (due > monthEnd.getTime()) return false;
      }
      if (quickFilter === 'thisYear') {
        if (!f.DueDate) return false;
        if (f.DeadlineStatus === 'Completed') return false;
        const due = new Date(f.DueDate).getTime();
        if (due > yearEnd.getTime()) return false;
      }
      if (quickFilter === 'ami' && f.DeadlineType !== 'AMI Cert Renewal') return false;
      if (quickFilter === 'sc' && f.cahpState !== 'SC') return false;
      if (quickFilter === 'nc' && f.cahpState !== 'NC') return false;
      void now;
      return true;
    });
    return result.sort((a, b) => {
      const urgencyA = a.fields.DeadlineStatus === 'Overdue' || a.fields.DeadlineStatus === 'Missed' ? 0 : 1;
      const urgencyB = b.fields.DeadlineStatus === 'Overdue' || b.fields.DeadlineStatus === 'Missed' ? 0 : 1;
      if (urgencyA !== urgencyB) return urgencyA - urgencyB;
      const dateA = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Number.MAX_VALUE;
      const dateB = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Number.MAX_VALUE;
      return dateA - dateB;
    });
  }, [data, statusFilter, typeFilter, assigneeFilter, quickFilter]);

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

  // PR-15b — AMI compliance stats — always computed, panel always visible so you can spot data gaps
  const amiStats = useMemo(() => {
    if (!data || !properties.data) return null;
    const amiProperties = properties.data.filter(
      (p) => p.fields.AMIProgram && p.fields.AMIProgram !== 'None'
    );
    const activeAmiProperties = amiProperties.filter((p) => p.fields.PropertyStatus === 'Active');
    const propertyIdsWithAmiDeadline = new Set(
      data
        .filter((d) => d.fields.DeadlineType === 'AMI Cert Renewal' && d.fields.DeadlineStatus !== 'Completed')
        .map((d) => d.fields.PropertyLookupId ? String(d.fields.PropertyLookupId) : null)
        .filter((id): id is string => id !== null)
    );
    const missingAmiDeadline = activeAmiProperties.filter(
      (p) => !propertyIdsWithAmiDeadline.has(String(p.id))
    );

    const now = Date.now();
    const ninetyDays = now + 90 * 24 * 60 * 60 * 1000;
    const amiDueSoon = data.filter((d) => {
      if (d.fields.DeadlineType !== 'AMI Cert Renewal') return false;
      if (d.fields.DeadlineStatus === 'Completed') return false;
      if (!d.fields.DueDate) return false;
      const due = new Date(d.fields.DueDate).getTime();
      return due <= ninetyDays;
    }).length;

    // Also count AMI deadlines without matching properties — could indicate data inconsistency
    const totalAmiDeadlines = data.filter((d) => d.fields.DeadlineType === 'AMI Cert Renewal').length;

    return {
      amiProperties: activeAmiProperties.length,
      totalAmiPropertiesIncludingInactive: amiProperties.length,
      amiDueSoon,
      missingAmiDeadline: missingAmiDeadline.length,
      missingProperties: missingAmiDeadline,
      totalAmiDeadlines,
      propertiesScanned: properties.data.length,
    };
  }, [data, properties.data]);

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
          onClick={() => setNewDeadlineOpen(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors"
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

      {/* AMI compliance focus panel — always shown so data gaps surface */}
      {amiStats && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-4 overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-100 bg-gold-50">
            <h3 className="text-sm font-semibold text-teal-900 flex items-center gap-2">
              <Icon name="star" size={14} className="text-gold-700" />
              AMI Compliance Focus
            </h3>
            <p className="text-xs text-gray-600 mt-0.5">
              Income-restricted housing oversight — tracks AMI Cert Renewal deadlines and gaps in coverage.
            </p>
          </div>
          {amiStats.amiProperties === 0 && amiStats.totalAmiPropertiesIncludingInactive === 0 && amiStats.totalAmiDeadlines === 0 ? (
            <div className="px-4 py-4 bg-amber-50 border-t border-amber-200">
              <p className="text-xs text-amber-900">
                <strong>No AMI-restricted properties detected.</strong> Scanned {amiStats.propertiesScanned} properties — none have <code className="bg-white px-1 rounded">AMIProgram</code> set to a value other than "None".
                If you have income-restricted properties, open them and set the AMI Program field on the Overview tab.
              </p>
            </div>
          ) : (
            <>
              <div className="grid grid-cols-3 divide-x divide-gray-100">
                <div className="px-4 py-3">
                  <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">Active AMI Properties</div>
                  <div className="text-2xl font-bold text-teal-700 mt-1">{amiStats.amiProperties}</div>
                  <div className="text-[11px] text-gray-500 mt-0.5">
                    Income-restricted
                    {amiStats.totalAmiPropertiesIncludingInactive > amiStats.amiProperties &&
                      ` · +${amiStats.totalAmiPropertiesIncludingInactive - amiStats.amiProperties} inactive`}
                  </div>
                </div>
                <div className="px-4 py-3">
                  <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">AMI Certs Due Soon</div>
                  <div className={`text-2xl font-bold mt-1 ${amiStats.amiDueSoon > 0 ? 'text-warning' : 'text-teal-700'}`}>
                    {amiStats.amiDueSoon}
                  </div>
                  <div className="text-[11px] text-gray-500 mt-0.5">Within 90 days</div>
                </div>
                <div className="px-4 py-3">
                  <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">Coverage Gap</div>
                  <div className={`text-2xl font-bold mt-1 ${amiStats.missingAmiDeadline > 0 ? 'text-error' : 'text-success'}`}>
                    {amiStats.missingAmiDeadline}
                  </div>
                  <div className="text-[11px] text-gray-500 mt-0.5">No AMI deadline set</div>
                </div>
              </div>
              {amiStats.missingProperties.length > 0 && (
                <div className="px-4 py-3 border-t border-amber-200 bg-amber-50">
                  <p className="text-xs text-amber-900 font-semibold mb-1">
                    ⚠ AMI properties without an AMI Cert Renewal deadline scheduled:
                  </p>
                  <div className="flex flex-wrap gap-1.5">
                    {amiStats.missingProperties.slice(0, 8).map((p) => (
                      <button
                        key={p.id}
                        onClick={() => navigate(`/properties/${p.id}`)}
                        className="text-[11px] bg-white border border-amber-300 text-amber-900 hover:bg-amber-100 px-2 py-0.5 rounded-full"
                      >
                        {p.fields.Title}
                      </button>
                    ))}
                    {amiStats.missingProperties.length > 8 && (
                      <span className="text-[11px] text-amber-700 self-center">
                        +{amiStats.missingProperties.length - 8} more
                      </span>
                    )}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* Quick filter chips */}
      <div className="mb-4 flex flex-wrap gap-1.5">
        {([
          { id: 'all', label: 'All' },
          { id: 'overdue', label: 'Overdue' },
          { id: 'thisMonth', label: 'Due This Month' },
          { id: 'thisYear', label: 'Due This Year' },
          { id: 'ami', label: 'AMI Only' },
          { id: 'sc', label: 'SC' },
          { id: 'nc', label: 'NC' },
        ] as { id: QuickFilter; label: string }[]).map((chip) => (
          <button
            key={chip.id}
            onClick={() => setQuickFilter(chip.id)}
            className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${
              quickFilter === chip.id
                ? 'bg-teal-700 text-white'
                : 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-50'
            }`}
          >
            {chip.label}
          </button>
        ))}
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
          value={assigneeFilter}
          onChange={(v) => setAssigneeFilter(v)}
          options={assigneeOptions}
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
              <tr
                key={d.id}
                onClick={() => navigate(`/compliance/${d.id}`)}
                className="hover:bg-gray-50 transition-colors cursor-pointer"
              >
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
                <td className="px-4 py-3 text-gray-700">{d.fields.AssignedTo || d.fields.ResponsibleParty || '—'}</td>
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

      {newDeadlineOpen && (
        <NewDeadlineModal
          onClose={() => setNewDeadlineOpen(false)}
          onSuccess={() => {
            setNewDeadlineOpen(false);
            refetch();
          }}
        />
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
  return d.toLocaleDateString('en-US', { timeZone: EASTERN_TZ, year: 'numeric', month: 'short', day: 'numeric' });
}
