import { useMemo } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Property,
  type ComplianceDeadline,
  type AuditLog,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { EASTERN_TZ } from '../lib/dates';

export function Portfolio() {
  const navigate = useNavigate();
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const deadlines = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });
  const auditLog = useSharePointList<AuditLog>(LIST_NAMES.AuditLog, { top: 10 });

  const loading = properties.loading || deadlines.loading;

  const analytics = useMemo(() => {
    if (!properties.data || !deadlines.data) return null;
    const props = properties.data;
    const dlines = deadlines.data;

    // KPIs
    const total = props.length;
    const active = props.filter((p) => p.fields.PropertyStatus === 'Active').length;
    const pending = props.filter((p) => p.fields.PropertyStatus === 'Pending').length;
    const totalUnits = props.reduce((sum, p) => sum + (p.fields.UnitCount ?? 0), 0);

    // 30-day deadline window
    const now = Date.now();
    const thirtyDaysOut = now + 30 * 24 * 60 * 60 * 1000;
    const deadlinesIn30Days = dlines.filter((d) => {
      if (!d.fields.DueDate) return false;
      const due = new Date(d.fields.DueDate).getTime();
      return due >= now && due <= thirtyDaysOut && d.fields.DeadlineStatus !== 'Completed';
    }).length;
    const overdue = dlines.filter((d) => {
      if (!d.fields.DueDate || d.fields.DeadlineStatus === 'Completed') return false;
      return new Date(d.fields.DueDate).getTime() < now;
    }).length;

    // Distributions
    const statusDist = countBy(props, (p) => p.fields.PropertyStatus ?? 'Unset');
    const amiDist = countBy(props, (p) => p.fields.AMIProgram ?? 'Unset');
    const stateDist = countBy(props, (p) => p.fields.cahpState ?? 'Unset');
    const verifyDist = countBy(props, (p) => p.fields.cahpVerificationStatus ?? 'Unset');

    // Needs attention — properties matching any risk rule
    const needsAttention = props.filter((p) => {
      const f = p.fields;
      if (f.PropertyStatus === 'Withdrawn' || f.PropertyStatus === 'Removed from Program' || f.PropertyStatus === 'Sold') {
        return false; // disposed, doesn't need attention
      }
      const issues: string[] = [];
      if (f.PropertyStatus === 'Pending') issues.push('Pending status');
      if (f.cahpVerificationStatus === 'Inherited - Unverified') issues.push('Unverified');
      if (f.cahpVerificationStatus === 'Needs Follow-Up') issues.push('Verification follow-up needed');
      if (f.LURAExecuted === 'In Progress' || f.LURAExecuted === 'No') issues.push('LURA not executed');
      if (f.CAHPLanguageAdded === 'No' || f.CAHPLanguageAdded === 'In Progress') issues.push('CAHP language not in OA');
      if (f.CAHPLanguageAdded === 'Needs Revision') issues.push('CAHP language needs revision');
      if (!f.UnitCount) issues.push('Missing unit count');
      if (!f.DORAccountID) issues.push('Missing DOR Account ID');
      return issues.length > 0 ? { ...p, _issues: issues } : false;
    }).map((p) => {
      const issues: string[] = [];
      const f = p.fields;
      if (f.PropertyStatus === 'Pending') issues.push('Pending');
      if (f.cahpVerificationStatus === 'Inherited - Unverified') issues.push('Unverified');
      if (f.cahpVerificationStatus === 'Needs Follow-Up') issues.push('Verification follow-up');
      if (f.LURAExecuted === 'In Progress' || f.LURAExecuted === 'No') issues.push('LURA pending');
      if (f.CAHPLanguageAdded === 'No' || f.CAHPLanguageAdded === 'In Progress') issues.push('OA language pending');
      if (f.CAHPLanguageAdded === 'Needs Revision') issues.push('OA language revision');
      if (!f.UnitCount) issues.push('No unit count');
      if (!f.DORAccountID) issues.push('No DOR ID');
      return { property: p, issues };
    });

    return {
      total,
      active,
      pending,
      totalUnits,
      deadlinesIn30Days,
      overdue,
      statusDist,
      amiDist,
      stateDist,
      verifyDist,
      needsAttention,
    };
  }, [properties.data, deadlines.data]);

  if (loading || !analytics) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Portfolio</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading portfolio analytics…</span>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Portfolio</h1>
        <p className="text-sm text-gray-500 mt-1">
          Aggregate analytics across all properties · {analytics.total} properties · {analytics.totalUnits.toLocaleString()} units
        </p>
      </div>

      {/* KPI Row */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPICard
          label="Properties"
          value={analytics.total}
          sub={`${analytics.active} active · ${analytics.pending} pending`}
        />
        <KPICard
          label="Units"
          value={analytics.totalUnits.toLocaleString()}
          sub="across all properties"
        />
        <KPICard
          label="Needs Attention"
          value={analytics.needsAttention.length}
          sub="properties with open issues"
          accent={analytics.needsAttention.length > 0 ? 'warning' : 'success'}
        />
        <KPICard
          label="Deadlines (30d)"
          value={analytics.deadlinesIn30Days}
          sub={analytics.overdue > 0 ? `${analytics.overdue} overdue!` : 'on track'}
          accent={analytics.overdue > 0 ? 'danger' : analytics.deadlinesIn30Days > 0 ? 'warning' : 'success'}
        />
      </div>

      {/* Distributions Row 1 */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
        <BreakdownCard
          title="Status"
          data={analytics.statusDist}
          total={analytics.total}
          onItemClick={(label) => navigate(`/properties?status=${encodeURIComponent(label)}`)}
        />
        <BreakdownCard
          title="AMI Program"
          data={analytics.amiDist}
          total={analytics.total}
        />
      </div>

      {/* Distributions Row 2 */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <BreakdownCard
          title="State"
          data={analytics.stateDist}
          total={analytics.total}
        />
        <BreakdownCard
          title="Verification"
          data={analytics.verifyDist}
          total={analytics.total}
        />
      </div>

      {/* Needs Attention */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
              <Icon name="alert" size={14} className="text-warning" />
              Needs Attention
            </h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {analytics.needsAttention.length} properties with open compliance issues
            </p>
          </div>
        </div>
        {analytics.needsAttention.length === 0 ? (
          <div className="px-4 py-8 text-center">
            <p className="text-sm text-success font-medium">All properties are in good standing.</p>
            <p className="text-xs text-gray-500 mt-1">No verification, LURA, or data issues detected.</p>
          </div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Issues</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {analytics.needsAttention.map(({ property: p, issues }) => (
                <tr
                  key={p.id}
                  onClick={() => navigate(`/properties/${p.id}`)}
                  className="hover:bg-gray-50 cursor-pointer transition-colors"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{p.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{p.fields.PropertyStatus ?? '—'}</td>
                  <td className="px-4 py-3">
                    <div className="flex flex-wrap gap-1">
                      {issues.map((issue) => (
                        <span
                          key={issue}
                          className="inline-block px-2 py-0.5 bg-amber-100 text-amber-900 rounded text-[11px] font-medium"
                        >
                          {issue}
                        </span>
                      ))}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Recent Activity */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
          <h3 className="text-sm font-semibold text-gray-700">Recent Activity</h3>
          <Link to="/audit" className="text-xs text-teal-700 hover:text-teal-900 font-medium">
            View full audit log →
          </Link>
        </div>
        {auditLog.loading ? (
          <div className="px-4 py-6 text-center text-sm text-gray-500">Loading…</div>
        ) : !auditLog.data || auditLog.data.length === 0 ? (
          <div className="px-4 py-6 text-center text-sm text-gray-500">No activity yet.</div>
        ) : (
          <ul className="divide-y divide-gray-100">
            {auditLog.data
              .slice()
              .sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime())
              .slice(0, 10)
              .map((row) => (
                <li key={row.id} className="px-4 py-2.5 text-sm flex items-center gap-3">
                  <span
                    className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold flex-shrink-0 ${
                      row.fields.Action === 'CREATE' ? 'bg-green-100 text-green-800' :
                      row.fields.Action === 'DELETE' ? 'bg-red-100 text-red-800' :
                      'bg-blue-100 text-blue-800'
                    }`}
                  >
                    {row.fields.Action}
                  </span>
                  <span className="text-xs text-gray-500 font-mono-data flex-shrink-0 w-20">
                    {row.fields.EntityType}
                  </span>
                  <span className="flex-1 text-gray-700 truncate">{row.fields.EntityTitle}</span>
                  <span className="text-xs text-gray-400 font-mono-data flex-shrink-0">
                    {formatRelative(row.createdDateTime)}
                  </span>
                </li>
              ))}
          </ul>
        )}
      </div>
    </div>
  );
}

// =============================================================================
// Components
// =============================================================================

function KPICard({
  label,
  value,
  sub,
  accent,
}: {
  label: string;
  value: string | number;
  sub: string;
  accent?: 'success' | 'warning' | 'danger';
}) {
  const valueClass =
    accent === 'success' ? 'text-success' :
    accent === 'warning' ? 'text-warning' :
    accent === 'danger' ? 'text-error' :
    'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${valueClass}`}>{value}</div>
      <div className="text-xs text-gray-500 mt-1">{sub}</div>
    </div>
  );
}

function BreakdownCard({
  title,
  data,
  total,
  onItemClick,
}: {
  title: string;
  data: Record<string, number>;
  total: number;
  onItemClick?: (label: string) => void;
}) {
  const entries = Object.entries(data).sort((a, b) => b[1] - a[1]);
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">{title}</h3>
      {entries.length === 0 ? (
        <p className="text-sm text-gray-400">No data.</p>
      ) : (
        <div className="space-y-2">
          {entries.map(([label, count]) => {
            const pct = total > 0 ? (count / total) * 100 : 0;
            return (
              <div
                key={label}
                className={onItemClick ? 'cursor-pointer hover:bg-gray-50 rounded -mx-1 px-1' : ''}
                onClick={() => onItemClick?.(label)}
              >
                <div className="flex items-center justify-between text-sm mb-0.5">
                  <span className="text-gray-700">{label}</span>
                  <span className="font-mono-data text-gray-500 text-xs">
                    {count} · {pct.toFixed(0)}%
                  </span>
                </div>
                <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-teal-500 rounded-full transition-all"
                    style={{ width: `${pct}%` }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Helpers
// =============================================================================

function countBy<T>(items: T[], keyFn: (item: T) => string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const item of items) {
    const key = keyFn(item);
    counts[key] = (counts[key] ?? 0) + 1;
  }
  return counts;
}

function formatRelative(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const diffMin = Math.floor((Date.now() - d.getTime()) / 60000);
  if (diffMin < 1) return 'just now';
  if (diffMin < 60) return `${diffMin}m`;
  const diffHr = Math.floor(diffMin / 60);
  if (diffHr < 24) return `${diffHr}h`;
  const diffDay = Math.floor(diffHr / 24);
  if (diffDay < 7) return `${diffDay}d`;
  return d.toLocaleDateString('en-US', { timeZone: EASTERN_TZ, month: 'short', day: 'numeric' });
}
