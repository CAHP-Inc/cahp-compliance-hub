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

const PHASE_1_PROGRESS = [
  { id: 'PR-01', label: 'Repo scaffolding + AppShell + GH Pages deploy', status: 'done' as const },
  { id: 'PR-02', label: 'MSAL auth + role detection from M365', status: 'done' as const },
  { id: 'PR-03', label: 'SharePoint inventory + schema mapping', status: 'done' as const },
  { id: 'PR-04', label: 'Graph SDK data layer + Properties list view', status: 'done' as const },
  { id: 'PR-05a', label: 'Property Detail page (Overview + Submittals tabs)', status: 'done' as const },
  { id: 'PR-05b', label: 'Compliance Deadlines module + My Day widget', status: 'done' as const },
  { id: 'PR-05c', label: 'Inline editing on Property Detail', status: 'done' as const },
  { id: 'PR-06a', label: 'Compliance Deadline editing + Mark Complete', status: 'done' as const },
  { id: 'PR-06b', label: 'Ownership Structure module (list + edit + create + delete)', status: 'done' as const },
  { id: 'PR-07', label: 'Audit log — automatic CRUD logging across all modules', status: 'done' as const },
  { id: 'PR-08a', label: 'Property Detail: Compliance + Ownership + Documents tabs', status: 'done' as const },
  { id: 'PR-08b', label: 'Property creation wizard', status: 'done' as const },
  { id: 'PR-08c', label: 'Disposition workflow', status: 'done' as const },
  { id: 'PR-08d', label: 'Property Notes thread (new SharePoint list)', status: 'done' as const },
  { id: 'PR-08e', label: 'Portfolio dashboard + polish', status: 'done' as const },
  { id: 'PR-09a', label: 'Owners entity master + recursive ownership engine', status: 'done' as const },
  { id: 'PR-09b', label: 'Org Chart rendering (3 layouts)', status: 'done' as const },
  { id: 'PR-09c', label: 'Property Wizard 6-step atomic + Audit CSV + remaining sub-tabs', status: 'done' as const },
  { id: 'PR-09d', label: 'Strict spec gap closure (5 lists, wizards, cascade, reason, CAHP Entity)', status: 'done' as const },
  { id: 'PR-10a', label: 'Submittals module — list + detail (read-only)', status: 'done' as const },
  { id: 'PR-10b', label: 'Submittal editing + status pipeline transitions', status: 'done' as const },
  { id: 'PR-10c', label: 'Org chart version freeze (rendered view of snapshot)', status: 'done' as const },
  { id: 'PR-10d', label: 'Approval Workflow modal — tax savings → Billing + Disbursement', status: 'done' as const },
  { id: 'PR-11a', label: 'DOR Correspondence full CRUD with cascade', status: 'done' as const },
  { id: 'PR-11b', label: 'Outstanding Items module (kanban + list)', status: 'done' as const },
  { id: 'PR-11c', label: 'Document upload with SharePoint Document Library routing', status: 'done' as const },
  { id: 'PR-11d', label: 'Untagged Documents queue with bulk-tag actions', status: 'done' as const },
  { id: 'PR-12-removed', label: 'Billing module removed — deferred to future, status tracking only', status: 'done' as const },
  { id: 'PR-13a', label: 'Owner Communications module', status: 'done' as const },
  { id: 'PR-14a', label: 'Reports foundation — 6 category cards + 3 working reports', status: 'done' as const },
  { id: 'PR-14-hardening', label: 'My Day widgets + reports cleanup + CAHP entity documents', status: 'done' as const },
  { id: 'PR-14b', label: 'Filing Checklist + doc linkage + owner docs + expanded categories', status: 'done' as const },
  { id: 'PR-15a', label: 'Notifications system — bell + feed + triggers', status: 'done' as const },
  { id: 'PR-15b', label: 'Compliance AMI Dashboard polish — KPIs + quick filters + coverage gaps', status: 'done' as const },
];

const DEADLINE_URGENCY_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800',
  'In Progress': 'bg-amber-100 text-amber-800',
  Completed: 'bg-green-100 text-green-800',
  Overdue: 'bg-red-100 text-red-800',
  Missed: 'bg-red-200 text-red-900',
};

/**
 * Items originally scoped in earlier PRs that have not yet shipped.
 * Each item shows what it is and where it originally lived in the plan so nothing
 * silently falls off the roadmap. Move items into PHASE_1_PROGRESS as concrete PRs
 * when you decide to schedule them.
 */
const BACKLOG: { feature: string; origin: string; note?: string }[] = [
  { feature: 'Submittals top-level module', origin: 'Phase 2' },
  { feature: 'Submittals editing on detail page', origin: 'Phase 2' },
  { feature: 'Outstanding Items module', origin: 'Phase 2' },
  { feature: 'DOR Correspondence Log module', origin: 'Phase 2' },
  { feature: 'Known Issues Log module', origin: 'Phase 2' },
  { feature: 'Owner Communications module', origin: 'Phase 3' },
  { feature: 'Document library file uploads + browser', origin: 'Phase 2', note: 'PR-08a added READ-ONLY listing per property' },
  { feature: 'Billing Tracker module', origin: 'Phase 3' },
  { feature: 'Reports module', origin: 'Phase 3' },
  { feature: 'Watch / follow button', origin: 'Phase 3', note: 'tied to notifications' },
];

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

      {/* Phase 3 wrap banner */}
      <div className="mb-6 bg-green-50 border border-green-200 rounded-lg p-4">
        <div className="flex items-start gap-3">
          <div className="flex-shrink-0 w-8 h-8 rounded-md bg-success text-white font-bold text-xs flex items-center justify-center font-mono-data">
            ✓
          </div>
          <div className="flex-1">
            <div className="font-semibold text-green-900">
              Phase 3 complete. PR-15a + 15b deployed.
            </div>
            <p className="text-sm text-gray-700 mt-1">
              <strong>Notifications</strong> live — bell icon in the header shows an unread count badge,
              dropdown with the 5 most recent, full feed at <strong>/notifications</strong> with read/unread filters
              and Mark All Read. Triggered automatically when an Outstanding Item is assigned to someone else, when a
              Communication has a follow-up due, and on Submittal status transitions.
              <strong> Compliance</strong> got an AMI Compliance Focus panel showing AMI property count, certs due in
              the next 90 days, and properties with no AMI Cert deadline scheduled (coverage gap detection).
              Quick-filter chips for Overdue / Due This Month / This Year / AMI Only / SC / NC sit above the existing filters.
              <br/><br/>
              <strong>Phase 3 is functionally complete.</strong> Next steps are operational — backfill filing checklists
              for existing properties, clean up untagged docs, onboard the team into active use.
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

      {/* Backlog — outstanding work from earlier planning */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
        <div className="px-5 py-3 border-b border-gray-100">
          <h2 className="text-base font-semibold text-teal-700">Backlog</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Originally scoped but not yet shipped · {BACKLOG.length} items
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-100 text-[11px] font-semibold text-gray-500 uppercase tracking-wider">
              <tr>
                <th className="px-5 py-2 text-left">Feature</th>
                <th className="px-5 py-2 text-left w-32">Originally In</th>
                <th className="px-5 py-2 text-left">Note</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {BACKLOG.map((item) => (
                <tr key={item.feature}>
                  <td className="px-5 py-2 text-gray-700">{item.feature}</td>
                  <td className="px-5 py-2 font-mono-data text-xs text-gray-500">{item.origin}</td>
                  <td className="px-5 py-2 text-xs text-gray-500 italic">{item.note || ''}</td>
                </tr>
              ))}
            </tbody>
          </table>
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
