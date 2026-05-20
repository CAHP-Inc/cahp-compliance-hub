import { useState, useMemo } from 'react';
import { useSession } from '../lib/session';
import {
  ROLE_PERMISSIONS,
  type Role,
  type ModuleId,
} from '../lib/permissions';
import {
  useSharePointList,
  LIST_NAMES,
  type Property,
  type Submittal,
  type Owner,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { ChecklistTemplatesEditor } from '../components/ChecklistTemplatesEditor';

const ROLE_LABELS: Record<Role, string> = {
  'Admin': 'Admin',
  'Contributor': 'Contributor',
  'Accounting': 'Accounting',
};

const ROLE_DESCRIPTIONS: Record<Role, string> = {
  'Admin': 'Full access. Can view, create, edit, delete, approve, upload, and export across all modules.',
  'Contributor': 'Day-to-day operations. Can view, create, edit, and upload. Cannot delete or approve.',
  'Accounting': 'Read-mostly. Can view all financial data and export reports.',
};

const ROLE_BADGE_STYLES: Record<Role, string> = {
  'Admin': 'bg-teal-100 text-teal-900 border-teal-300',
  'Contributor': 'bg-gold-100 text-gold-900 border-gold-300',
  'Accounting': 'bg-blue-100 text-blue-900 border-blue-300',
};

// Access list mirrors roleMap.ts — kept in sync at deploy time
const ACCESS_LIST: { email: string; role: Role; org: string; addedIn: string }[] = [
  { email: 'bturner@newshirepm.com', role: 'Admin', org: 'NewShire', addedIn: 'PR-02' },
  { email: 'stan@vanrockre.com', role: 'Admin', org: 'VanRock', addedIn: 'PR-15a' },
  { email: 'bdebruin@redcedarhomes.com', role: 'Admin', org: 'Red Cedar', addedIn: 'PR-15a' },
  { email: 'lheckman@redcedarhomes.com', role: 'Contributor', org: 'Red Cedar', addedIn: 'PR-15-hotfix' },
];

export function SettingsPage() {
  const { user, role, realRole, setDevRoleOverride } = useSession();
  const [tab, setTab] = useState<'profile' | 'access' | 'permissions' | 'checklist' | 'system'>('profile');

  // System info data — count of stuff in SharePoint
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 1 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 1 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 1 });

  const isAdmin = role === 'Admin';

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Settings</h1>
        <p className="text-sm text-gray-500 mt-1">
          Account profile, access control, permissions reference, and system info.
        </p>
      </div>

      {/* Tabs */}
      <div className="border-b border-gray-200 mb-6 flex flex-wrap">
        {([
          { id: 'profile', label: 'Profile' },
          { id: 'access', label: 'Access List' },
          { id: 'permissions', label: 'Permissions Matrix' },
          { id: 'checklist', label: 'Checklist Templates' },
          { id: 'system', label: 'System Info' },
        ] as const).map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className={`px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
              tab === t.id
                ? 'border-teal-700 text-teal-700'
                : 'border-transparent text-gray-600 hover:text-gray-900'
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {/* Profile tab */}
      {tab === 'profile' && (
        <div className="space-y-6">
          <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
            <h2 className="text-base font-semibold text-teal-900 mb-4">Account</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="Display Name" value={user?.name ?? '—'} />
              <Field label="Email (UPN)" value={user?.email ?? '—'} mono />
              <Field label="Organization" value={user?.org ?? '—'} />
              <Field label="User ID" value={user?.id ?? '—'} mono small />
            </div>
          </div>

          <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
            <h2 className="text-base font-semibold text-teal-900 mb-4">Role & Access</h2>
            <div className="flex items-center gap-3 mb-4">
              <span className="text-sm text-gray-700">Effective Role:</span>
              {realRole ? (
                <span className={`inline-block px-2.5 py-1 rounded text-xs font-semibold border ${ROLE_BADGE_STYLES[realRole]}`}>
                  {ROLE_LABELS[realRole]}
                </span>
              ) : (
                <span className="inline-block px-2.5 py-1 rounded text-xs font-semibold border bg-red-100 text-red-900 border-red-300">
                  No access
                </span>
              )}
              {role && realRole && role !== realRole && (
                <span className="text-xs text-amber-700 italic">(dev override: viewing as {role})</span>
              )}
            </div>
            {realRole && (
              <p className="text-xs text-gray-600">{ROLE_DESCRIPTIONS[realRole]}</p>
            )}
            <div className="mt-4 pt-4 border-t border-gray-100">
              <p className="text-xs text-gray-500">
                Roles are managed in <code className="bg-gray-100 px-1 rounded">src/lib/roleMap.ts</code>. To change a role,
                edit the file, commit, and redeploy. SharePoint group membership is independent.
              </p>
            </div>
          </div>

          {/* Dev tools — only in development builds */}
          {import.meta.env.DEV && setDevRoleOverride && (
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-5 shadow-card">
              <h2 className="text-base font-semibold text-amber-900 mb-2 flex items-center gap-2">
                <Icon name="settings" size={14} />
                Dev Tools — View As Role
              </h2>
              <p className="text-xs text-amber-800 mb-3">
                Override your effective role to test UI under different permission sets. Affects only your local view;
                no changes to SharePoint or backend permissions. Disabled in production.
              </p>
              <div className="flex flex-wrap gap-2">
                {(['Admin', 'Contributor', 'Accounting'] as Role[]).map((r) => (
                  <button
                    key={r}
                    onClick={() => setDevRoleOverride(r)}
                    className={`px-3 py-1.5 rounded text-xs font-medium border ${
                      role === r
                        ? 'bg-amber-700 text-white border-amber-800'
                        : 'bg-white text-amber-900 border-amber-300 hover:bg-amber-100'
                    }`}
                  >
                    View as {r}
                  </button>
                ))}
                <button
                  onClick={() => setDevRoleOverride(null)}
                  className="px-3 py-1.5 rounded text-xs font-medium border bg-white text-gray-700 border-gray-300 hover:bg-gray-50"
                >
                  Reset to real role
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Access List tab */}
      {tab === 'access' && (
        <div className="space-y-4">
          <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900">
            <strong>Source of truth:</strong> Access is controlled by{' '}
            <code className="bg-white px-1 rounded">src/lib/roleMap.ts</code>. To add or remove access, edit that file,
            commit, and push. GitHub Actions deploys the change in ~2 minutes. Users not in this list see an Access Denied
            screen on sign-in, regardless of M365 tenant membership.
          </div>

          <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
                <tr>
                  <th className="px-4 py-3 text-left">Email (UPN)</th>
                  <th className="px-4 py-3 text-left">Organization</th>
                  <th className="px-4 py-3 text-left">Role</th>
                  <th className="px-4 py-3 text-left">Added In</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {ACCESS_LIST.map((entry) => {
                  const isMe = entry.email.toLowerCase() === user?.email.toLowerCase();
                  return (
                    <tr key={entry.email} className={isMe ? 'bg-teal-50/50' : ''}>
                      <td className="px-4 py-3 font-mono-data text-xs">
                        {entry.email}
                        {isMe && <span className="ml-2 text-[10px] font-semibold text-teal-700 bg-teal-100 px-1.5 py-0.5 rounded">YOU</span>}
                      </td>
                      <td className="px-4 py-3 text-sm text-gray-700">{entry.org}</td>
                      <td className="px-4 py-3">
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold border ${ROLE_BADGE_STYLES[entry.role]}`}>
                          {ROLE_LABELS[entry.role]}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-xs text-gray-500 font-mono-data">{entry.addedIn}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {!isAdmin && (
            <p className="text-xs text-gray-500 italic">
              Only Admins can request access changes. Contact bturner@newshirepm.com.
            </p>
          )}
        </div>
      )}

      {/* Permissions Matrix tab */}
      {tab === 'permissions' && <PermissionsMatrix />}

      {/* Checklist Templates tab — edit the list of items the Filing Checklist Generator creates */}
      {tab === 'checklist' && (
        isAdmin ? (
          <ChecklistTemplatesEditor />
        ) : (
          <div className="bg-yellow-50 border border-yellow-200 rounded-md p-4 text-sm text-yellow-900">
            Editing checklist templates requires the Admin role.
          </div>
        )
      )}

      {/* System Info tab */}
      {tab === 'system' && (
        <div className="space-y-6">
          <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
            <h2 className="text-base font-semibold text-teal-900 mb-4">System</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="App Version" value="Phase 3 complete (PR-16)" />
              <Field label="SharePoint Site" value="vanrockre.sharepoint.com/sites/CAHPComplianceHub" mono small />
              <Field label="Azure Tenant" value="VanRock Holdings" />
              <Field label="SPA Client ID" value="eeb92696-399a-4394-858c-ee73de0e94c6" mono small />
            </div>
          </div>

          <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
            <h2 className="text-base font-semibold text-teal-900 mb-4">Data Snapshot</h2>
            <div className="grid grid-cols-3 gap-4">
              <StatCard label="Properties" value={properties.data?.length ?? '…'} />
              <StatCard label="Submittals" value={submittals.data?.length ?? '…'} />
              <StatCard label="Owners" value={owners.data?.length ?? '…'} />
            </div>
            <p className="text-xs text-gray-500 mt-3">
              Top-of-list counts from SharePoint. For full inventory, run the <strong>Full Database Export</strong> report.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
            <h2 className="text-base font-semibold text-teal-900 mb-4">SharePoint Lists</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-xs">
              {Object.entries(LIST_NAMES).map(([key, name]) => (
                <div key={key} className="flex items-center gap-2 py-1">
                  <Icon name="check" size={10} className="text-success" />
                  <span className="font-mono-data text-gray-700">{name}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Field({ label, value, mono, small }: { label: string; value: string; mono?: boolean; small?: boolean }) {
  return (
    <div>
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-0.5">{label}</div>
      <div className={`text-gray-900 ${mono ? 'font-mono-data' : ''} ${small ? 'text-xs' : 'text-sm'}`}>{value}</div>
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: number | string }) {
  return (
    <div className="bg-gray-50 border border-gray-200 rounded p-3 text-center">
      <div className="text-2xl font-bold text-teal-700 font-mono-data">{value}</div>
      <div className="text-[10px] text-gray-500 uppercase tracking-wider mt-0.5">{label}</div>
    </div>
  );
}

function PermissionsMatrix() {
  // Build a matrix of [role] × [module] → actions
  const modules = useMemo(() => {
    const set = new Set<ModuleId>();
    (Object.values(ROLE_PERMISSIONS) as typeof ROLE_PERMISSIONS[Role][]).forEach((rp) => {
      rp.views.forEach((m) => set.add(m));
      Object.keys(rp.actions).forEach((m) => set.add(m as ModuleId));
    });
    return Array.from(set).sort();
  }, []);

  const roles: Role[] = ['Admin', 'Contributor', 'Accounting'];

  return (
    <div className="space-y-4">
      <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900">
        <strong>Read-only reference.</strong> Permissions are defined in{' '}
        <code className="bg-white px-1 rounded">src/lib/permissions.ts</code> and enforced throughout the app via
        <code className="bg-white px-1 rounded">canView()</code> / <code className="bg-white px-1 rounded">canDo()</code> helpers.
      </div>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden overflow-x-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b border-gray-200 font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-3 py-2 text-left">Module</th>
              {roles.map((r) => (
                <th key={r} className="px-3 py-2 text-center min-w-[120px]">
                  <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${ROLE_BADGE_STYLES[r]}`}>
                    {ROLE_LABELS[r]}
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {modules.map((m) => (
              <tr key={m} className="hover:bg-gray-50">
                <td className="px-3 py-2 font-medium text-gray-900">{m}</td>
                {roles.map((r) => {
                  const perm = ROLE_PERMISSIONS[r];
                  const canView = perm.views.includes(m);
                  const actions = perm.actions[m] ?? [];
                  return (
                    <td key={r} className="px-3 py-2 text-center">
                      {canView ? (
                        <div className="flex flex-col items-center gap-0.5">
                          <Icon name="check" size={12} className="text-success" />
                          {actions.length > 0 && (
                            <div className="text-[9px] text-gray-500 font-mono-data">
                              {actions.join(', ')}
                            </div>
                          )}
                        </div>
                      ) : (
                        <Icon name="alert" size={12} className="text-gray-300" />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
