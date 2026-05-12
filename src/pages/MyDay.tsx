import { Icon } from '../components/ui/Icon';
import { useSession } from '../lib/session';
import { ROLE_PERMISSIONS } from '../lib/permissions';
import type { Role } from '../lib/permissions';

const TODAY = new Date().toLocaleDateString('en-US', {
  weekday: 'long',
  year: 'numeric',
  month: 'long',
  day: 'numeric',
});

const PHASE_1_PROGRESS = [
  { id: 'PR-01', label: 'Repo scaffolding + AppShell + GH Pages deploy', status: 'done' as const },
  { id: 'PR-02', label: 'MSAL auth + role detection from M365', status: 'done' as const },
  { id: 'PR-03', label: 'SharePoint provisioning (14 lists)', status: 'next' as const },
  { id: 'PR-04', label: 'Graph SDK data layer + TypeScript types', status: 'pending' as const },
  { id: 'PR-05', label: 'Properties module (list + detail + wizard)', status: 'pending' as const },
  { id: 'PR-06', label: 'Owners module + Ownership engine', status: 'pending' as const },
  { id: 'PR-07', label: 'Audit log + Phase 1 wrap', status: 'pending' as const },
];

export function MyDay() {
  const { user, role, realRole, setDevRoleOverride } = useSession();
  if (!user || !role) return null;

  const realRoleConfig = realRole ? ROLE_PERMISSIONS[realRole] : null;
  const isViewingAsOverride = realRole && role !== realRole;

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">My Day</h1>
        <p className="text-sm text-gray-500 mt-1">{TODAY}</p>
        <p className="text-base text-gray-700 mt-3">
          Good morning, <span className="font-semibold">{user.name.split(' ')[0]}</span>. Welcome to the CAHP Compliance Hub.
        </p>
      </div>

      {/* PR-02 banner */}
      <div className="mb-6 bg-gold-50 border border-gold-200 rounded-lg p-4">
        <div className="flex items-start gap-3">
          <div className="flex-shrink-0 w-8 h-8 rounded-md bg-gold-500 text-teal-900 font-bold text-xs flex items-center justify-center font-mono-data">
            02
          </div>
          <div className="flex-1">
            <div className="font-semibold text-teal-900">PR-02 deployed. M365 sign-in is live.</div>
            <p className="text-sm text-gray-700 mt-1">
              You're signed in as <strong>{user.name}</strong> ({user.email}) with role{' '}
              <strong>{realRoleConfig?.label || 'unknown'}</strong>. Next up: SharePoint provisioning in PR-03.
            </p>
          </div>
        </div>
      </div>

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
