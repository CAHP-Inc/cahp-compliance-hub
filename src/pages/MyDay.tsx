import { useMemo } from 'react';
import { Icon } from '../components/ui/Icon';
import { Link } from 'react-router-dom';
import { useSession } from '../lib/session';
import { ROLE_PERMISSIONS } from '../lib/permissions';
import type { Role } from '../lib/permissions';
import { useSharePointList, LIST_NAMES, type Property, type ComplianceDeadline, type DeadlineStatus, type OutstandingItem, type Submittal } from '../lib/sharepoint';

const TODAY = new Date().toLocaleDateString('en-US', {
  weekday: 'long',
  year: 'numeric',
  month: 'long',
  day: 'numeric',
});

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
  const { data: allItems } = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const { data: allSubmittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const propertiesById = useMemo(() => {
    if (!properties) return new Map<string, Property>();
    return new Map(properties.map((p) => [String(p.id), p]));
  }, [properties]);

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

  // Outstanding Items widget — open items, overdue first, then by due date / priority
  const openItems = useMemo(() => {
    if (!allItems) return [];
    const isClosed = (s: string | undefined) =>
      s === 'Done' || s === 'Received' || s === 'Not Applicable';
    const priorityOrder: Record<string, number> = { Critical: 0, High: 1, Medium: 2, Low: 3 };

    return allItems
      .filter((i) => !isClosed(i.fields.ItemStatus))
      .sort((a, b) => {
        const aOverdue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() < Date.now() : false;
        const bOverdue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() < Date.now() : false;
        if (aOverdue !== bOverdue) return aOverdue ? -1 : 1;
        const aP = priorityOrder[a.fields.Priority ?? 'Medium'] ?? 2;
        const bP = priorityOrder[b.fields.Priority ?? 'Medium'] ?? 2;
        if (aP !== bP) return aP - bP;
        const aDue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Infinity;
        const bDue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Infinity;
        return aDue - bDue;
      })
      .slice(0, 8);
  }, [allItems]);

  const openItemsTotal = useMemo(() => {
    if (!allItems) return 0;
    const isClosed = (s: string | undefined) =>
      s === 'Done' || s === 'Received' || s === 'Not Applicable';
    return allItems.filter((i) => !isClosed(i.fields.ItemStatus)).length;
  }, [allItems]);

  // Submittal Next Actions — submittals with a NextAction set that aren't in terminal state
  const submittalsWithNextAction = useMemo(() => {
    if (!allSubmittals) return [];
    const terminal = ['Approved', 'Denied', 'Withdrawn'];
    return allSubmittals
      .filter((s) => {
        if (!s.fields.NextAction) return false;
        if (s.fields.SubmittalStatus && terminal.includes(s.fields.SubmittalStatus)) return false;
        return true;
      })
      .sort((a, b) => {
        const aDue = a.fields.NextActionDue ? new Date(a.fields.NextActionDue).getTime() : Infinity;
        const bDue = b.fields.NextActionDue ? new Date(b.fields.NextActionDue).getTime() : Infinity;
        return aDue - bDue;
      })
      .slice(0, 6);
  }, [allSubmittals]);

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

      {/* Outstanding Items widget */}
      {openItems.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
          <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
            <div>
              <h2 className="text-base font-semibold text-teal-700 flex items-center gap-2">
                <Icon name="inbox" size={16} />
                Outstanding Items
              </h2>
              <p className="text-xs text-gray-500 mt-0.5">
                Open · sorted by urgency · {openItemsTotal} total open
              </p>
            </div>
            <Link
              to="/outstanding-items"
              className="text-xs text-teal-700 hover:text-teal-900 font-medium"
            >
              View all →
            </Link>
          </div>
          <div className="divide-y divide-gray-100">
            {openItems.map((i) => {
              const property = i.fields.PropertyLookupId
                ? propertiesById.get(String(i.fields.PropertyLookupId))
                : null;
              const dueDate = i.fields.DueDate ? new Date(i.fields.DueDate) : null;
              const daysOut = dueDate ? Math.round((dueDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000)) : null;
              const isOverdue = daysOut !== null && daysOut < 0;
              const priority = i.fields.Priority ?? 'Medium';
              const priorityStyle =
                priority === 'Critical' ? 'bg-red-100 text-red-800' :
                priority === 'High' ? 'bg-amber-100 text-amber-800' :
                priority === 'Medium' ? 'bg-blue-100 text-blue-800' :
                'bg-gray-100 text-gray-600';
              return (
                <Link
                  key={i.id}
                  to={`/outstanding-items/${i.id}`}
                  className="px-5 py-3 flex items-center gap-3 hover:bg-gray-50 transition-colors"
                >
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium text-gray-900 truncate">
                      {isOverdue && <span className="text-error mr-1">⚠</span>}
                      {i.fields.Title}
                    </div>
                    <div className="text-xs text-gray-500 mt-0.5">
                      {property ? property.fields.Title : <span className="italic">unassigned property</span>}
                      {i.fields.AssignedTo && <span> · {i.fields.AssignedTo}</span>}
                    </div>
                  </div>
                  <span className={`flex-shrink-0 px-2 py-0.5 rounded text-[10px] font-semibold ${priorityStyle}`}>
                    {priority}
                  </span>
                  <div className="text-right flex-shrink-0 w-20">
                    {dueDate ? (
                      <>
                        <div className={`text-xs font-mono-data ${isOverdue ? 'text-error font-bold' : 'text-gray-700'}`}>
                          {dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                        </div>
                        <div className="text-[10px] text-gray-500">
                          {isOverdue ? `${Math.abs(daysOut!)}d overdue` : daysOut === 0 ? 'Today' : `in ${daysOut}d`}
                        </div>
                      </>
                    ) : (
                      <span className="text-[10px] text-gray-400">no due date</span>
                    )}
                  </div>
                </Link>
              );
            })}
          </div>
        </div>
      )}

      {/* Submittal Next Actions widget */}
      {submittalsWithNextAction.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
          <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
            <div>
              <h2 className="text-base font-semibold text-teal-700 flex items-center gap-2">
                <Icon name="file" size={16} />
                Submittal Next Actions
              </h2>
              <p className="text-xs text-gray-500 mt-0.5">
                Submittals with a pending action
              </p>
            </div>
            <Link
              to="/submittals"
              className="text-xs text-teal-700 hover:text-teal-900 font-medium"
            >
              View all →
            </Link>
          </div>
          <div className="divide-y divide-gray-100">
            {submittalsWithNextAction.map((s) => {
              const property = s.fields.PropertyLookupId
                ? propertiesById.get(String(s.fields.PropertyLookupId))
                : null;
              const dueDate = s.fields.NextActionDue ? new Date(s.fields.NextActionDue) : null;
              const daysOut = dueDate ? Math.round((dueDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000)) : null;
              const isOverdue = daysOut !== null && daysOut < 0;
              return (
                <Link
                  key={s.id}
                  to={`/submittals/${s.id}`}
                  className="px-5 py-3 flex items-center gap-3 hover:bg-gray-50 transition-colors"
                >
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium text-gray-900 truncate">
                      {isOverdue && <span className="text-error mr-1">⚠</span>}
                      {s.fields.Title}
                    </div>
                    <div className="text-xs text-gray-500 mt-0.5 truncate">
                      <span className="italic">{s.fields.NextAction}</span>
                      {property && <span> · {property.fields.Title}</span>}
                      {s.fields.SubmittalStatus && (
                        <span className="ml-1.5 font-mono-data text-teal-700">{s.fields.SubmittalStatus}</span>
                      )}
                    </div>
                  </div>
                  <div className="text-right flex-shrink-0 w-20">
                    {dueDate ? (
                      <>
                        <div className={`text-xs font-mono-data ${isOverdue ? 'text-error font-bold' : 'text-gray-700'}`}>
                          {dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                        </div>
                        <div className="text-[10px] text-gray-500">
                          {isOverdue ? `${Math.abs(daysOut!)}d overdue` : daysOut === 0 ? 'Today' : `in ${daysOut}d`}
                        </div>
                      </>
                    ) : (
                      <span className="text-[10px] text-gray-400">no due date</span>
                    )}
                  </div>
                </Link>
              );
            })}
          </div>
        </div>
      )}

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

