import { useMemo } from 'react';
import { Icon } from '../components/ui/Icon';
import { Link } from 'react-router-dom';
import { useSession } from '../lib/session';
import { ROLE_PERMISSIONS } from '../lib/permissions';
import type { Role } from '../lib/permissions';
import { useSharePointList, LIST_NAMES, type Property, type ComplianceDeadline, type DeadlineStatus } from '../lib/sharepoint';

const TODAY = new Date().toLocaleDateString('en-US', {
  weekday: 'long',
  year: 'numeric',
  month: 'long',
  day: 'numeric',
});

const PHASE_1_PROGRESS = [
  { id: 'PR-01', label: 'Repo scaffolding + AppShell + GH Pages deploy', status: 'done' as const },
  { id: 'PR-02', label: 'MSAL auth + role detection from M365', status: 'done' as const },
  { id: 'PR-03', label: 'SharePoint inventory + schema mapping', status: 'done' as const },
  { id: 'PR-04', label: 'Graph SDK data layer + Properties list view', status: 'done' as const },
  { id: 'PR-05a', label: 'Property Detail page (Overview + Submittals tabs)', status: 'done' as const },
  { id: 'PR-05b', label: 'Compliance Deadlines module + My Day widget', status: 'done' as const },
  { id: 'PR-05c', label: 'Inline editing on Property Detail', status: 'done' as const },
  { id: 'PR-06a', label: 'Compliance Deadline editing + Mark Complete', status: 'done' as const },
  { id: 'PR-06b', label: 'Submittals editing + Outstanding Items module', status: 'next' as const },
  { id: 'PR-07', label: 'Audit log + Phase 1 wrap', status: 'pending' as const },
];

const DEADLINE_URGENCY_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800',
  'In Progress': 'bg-amber-100 text-amber-800',
  Completed: 'bg-green-100 text-green-800',
  Overdue: 'bg-red-100 text-red-800',
  Missed: 'bg-red-200 text-red-900',
};

export function MyDay() {
  const { user, role, realRole, setDevRoleOverride } = useSession();
  const { data: properties } = useSharePointList<Property>(
    LIST_NAMES.Properties,
    { top: 200 }
  );
  const { data: allDeadlines } = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });

  const upcomingDeadlines = useMemo(() => {
    if (!allDeadlines) return [];
    const now = Date.now();
    const horizon = now + 90 * 24 * 60 * 60 * 1000; // 90 days out
    return allDeadlines
      .filter((d) => {
        const status = d.fields.DeadlineStatus;
        if (status === 'Completed') return false;
        if (!d.fields.DueDate) return false;
        const due = new Date(d.fields.DueDate).getTime();
        return due <= horizon; // includes already-overdue
      })
      .sort((a, b) => {
        const da = new Date(a.fields.DueDate!).getTime();
        const db = new Date(b.fields.DueDate!).getTime();
        return da - db;
      })
      .slice(0, 5);
  }, [allDeadlines]);

  if (!user || !role) return null;

  const realRoleConfig = realRole ? ROLE_PERMISSIONS[realRole] : null;
  const isViewingAsOverride = realRole && role !== realRole;

  const totalProperties = properties?.length ?? 0;
  const activeProperties = properties?.filter((p) => p.fields.PropertyStatus === 'Active').length ?? 0;
  const totalUnits = properties?.reduce((sum, p) => sum + (p.fields.UnitCount ?? 0), 0) ?? 0;

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">My Day</h1>
        <p className="text-sm text-gray-500 mt-1">{TODAY}</p>
        <p className="text-base text-gray-700 mt-3">
          Good morning, <span className="font-semibold">{user.name.split(' ')[0]}</span>. Welcome to the CAHP Compliance Hub.
        </p>
      </div>

      {/* PR-06a banner */}
      <div className="mb-6 bg-gold-50 border border-gold-200 rounded-lg p-4">
        <div className="flex items-start gap-3">
          <div className="flex-shrink-0 w-8 h-8 rounded-md bg-gold-500 text-teal-900 font-bold text-xs flex items-center justify-center font-mono-data">
            06a
          </div>
          <div className="flex-1">
            <div className="font-semibold text-teal-900">
              PR-06a deployed. Compliance Deadlines are now editable.
            </div>
            <p className="text-sm text-gray-700 mt-1">
              Click any deadline in the widget below, or from the Compliance module, to drill into
              it. <strong>Mark Complete</strong> sets status + completion date in one click. Full
              edit mode for everything else. All writes go straight to SharePoint.
            </p>
          </div>
        </div>
      </div>

      {/* Portfolio at a glance */}
      {properties && (
        <div className="mb-6 grid grid-cols-3 gap-4">
          <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
            <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Properties</div>
            <div className="text-3xl font-bold text-teal-700 mt-1">{totalProperties}</div>
            <div className="text-xs text-gray-500 mt-0.5">{activeProperties} active</div>
          </div>
          <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
            <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Units</div>
            <div className="text-3xl font-bold text-teal-700 mt-1">{totalUnits.toLocaleString()}</div>
            <div className="text-xs text-gray-500 mt-0.5">across all properties</div>
          </div>
          <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
            <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Submittals</div>
            <div className="text-3xl font-bold text-teal-700 mt-1">18</div>
            <div className="text-xs text-gray-500 mt-0.5">tracked in SharePoint</div>
          </div>
        </div>
      )}

      {/* Upcoming Deadlines widget */}
      {upcomingDeadlines.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
          <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
            <div>
              <h2 className="text-base font-semibold text-teal-700 flex items-center gap-2">
                <Icon name="calendar" size={16} />
                Upcoming Deadlines
              </h2>
              <p className="text-xs text-gray-500 mt-0.5">
                Next 90 days · sorted by due date
              </p>
            </div>
            <Link
              to="/compliance"
              className="text-xs text-teal-700 hover:text-teal-900 font-medium"
            >
              View all →
            </Link>
          </div>
          <div className="divide-y divide-gray-100">
            {upcomingDeadlines.map((d) => {
              const dueDate = d.fields.DueDate ? new Date(d.fields.DueDate) : null;
              const daysOut = dueDate ? Math.round((dueDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000)) : null;
              const isOverdue = daysOut !== null && daysOut < 0;
              const isUrgent = daysOut !== null && daysOut >= 0 && daysOut <= 7;
              return (
                <Link
                  key={d.id}
                  to={`/compliance/${d.id}`}
                  className="px-5 py-3 flex items-center gap-3 hover:bg-gray-50 transition-colors"
                >
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium text-gray-900 truncate">{d.fields.Title}</div>
                    <div className="text-xs text-gray-500 mt-0.5">
                      {d.fields.DeadlineType || 'Other'} · {d.fields.ResponsibleParty || 'Unassigned'}
                      {d.fields.cahpState && (
                        <span className="ml-1.5 font-mono-data text-teal-700">{d.fields.cahpState}</span>
                      )}
                    </div>
                  </div>
                  <div className="text-right flex-shrink-0">
                    <div
                      className={`text-sm font-mono-data ${
                        isOverdue ? 'text-error font-bold' : isUrgent ? 'text-amber-700 font-semibold' : 'text-gray-700'
                      }`}
                    >
                      {dueDate?.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
                    </div>
                    <div className="text-[11px] text-gray-500">
                      {isOverdue
                        ? `${Math.abs(daysOut!)} days overdue`
                        : daysOut === 0
                          ? 'Today'
                          : `in ${daysOut} day${daysOut === 1 ? '' : 's'}`}
                    </div>
                  </div>
                  {d.fields.DeadlineStatus && (
                    <span
                      className={`flex-shrink-0 px-2 py-0.5 rounded text-[10px] font-semibold ${
                        DEADLINE_URGENCY_STYLES[d.fields.DeadlineStatus] || 'bg-gray-100'
                      }`}
                    >
                      {d.fields.DeadlineStatus}
                    </span>
                  )}
                </Link>
              );
            })}
          </div>
        </div>
      )}

      {/* Phase 1 progress */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
        <div className="px-5 py-3 border-b border-gray-100">
          <h2 className="text-base font-semibold text-teal-700">Phase 1 Build Progress</h2>
          <p className="text-xs text-gray-500 mt-0.5">Weeks 1-3 — Foundation</p>
        </div>
        <div className="divide-y divide-gray-100">
          {PHASE_1_PROGRESS.map((pr) => (
            <div key={pr.id} className="px-5 py-3 flex items-center gap-3">
              <div className="font-mono-data text-xs font-semibold text-gray-400 w-12">{pr.id}</div>
              <div className="flex-1 text-sm text-gray-700">{pr.label}</div>
              <StatusBadge status={pr.status} />
            </div>
          ))}
        </div>
      </div>

      {/* Dev-only role override (hidden in production builds) */}
      {import.meta.env.DEV && setDevRoleOverride && (
        <div className="bg-yellow-50 border border-yellow-200 rounded-lg shadow-card">
          <div className="px-5 py-3 border-b border-yellow-200">
            <h2 className="text-base font-semibold text-yellow-800">Dev: View As Role</h2>
            <p className="text-xs text-yellow-700 mt-0.5">
              Only visible in <span className="font-mono-data">npm run dev</span>. Lets you preview how the UI changes per role without signing in as different users.
              {realRoleConfig && (
                <>
                  {' '}Your real M365 role: <strong>{realRoleConfig.label}</strong>.
                </>
              )}
              {isViewingAsOverride && (
                <>
                  {' '}Currently overriding to: <strong>{ROLE_PERMISSIONS[role].label}</strong>.
                </>
              )}
            </p>
          </div>
          <div className="p-4 flex flex-wrap gap-2 items-center">
            {(['Admin', 'Contributor', 'Accounting'] as Role[]).map((r) => {
              const cfg = ROLE_PERMISSIONS[r];
              const active = role === r;
              return (
                <button
                  key={r}
                  onClick={() => setDevRoleOverride(r)}
                  className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
                    active ? `${cfg.color} shadow-sm` : 'bg-white text-gray-600 hover:bg-gray-50 border border-gray-200'
                  }`}
                >
                  {cfg.label}
                </button>
              );
            })}
            <button
              onClick={() => setDevRoleOverride(null)}
              className="px-4 py-2 rounded-md text-sm font-medium text-yellow-700 hover:bg-yellow-100"
            >
              Reset to real role
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function StatusBadge({ status }: { status: 'done' | 'next' | 'pending' }) {
  if (status === 'done') {
    return (
      <span className="inline-flex items-center gap-1 text-xs font-semibold text-success">
        <Icon name="check" size={14} />
        Done
      </span>
    );
  }
  if (status === 'next') {
    return (
      <span className="inline-flex items-center gap-1 text-xs font-semibold text-gold-700 bg-gold-50 px-2 py-0.5 rounded">
        Next
      </span>
    );
  }
  return <span className="text-xs text-gray-400">Pending</span>;
}
