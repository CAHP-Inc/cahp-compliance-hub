import { useNavigate, useLocation } from 'react-router-dom';
import clsx from 'clsx';
import { Icon, IconName } from '../ui/Icon';
import { canView, ModuleId } from '../../lib/permissions';
import { useSession } from '../../lib/session';

interface NavItem {
  id: ModuleId;
  label: string;
  icon: IconName;
  path: string;
  badge?: number;
  badgeStyle?: 'urgent' | 'normal';
}

interface NavGroup {
  label: string;
  items: NavItem[];
}

const NAV_GROUPS: NavGroup[] = [
  {
    label: 'Daily Work',
    items: [
      { id: 'myday', label: 'My Day', icon: 'inbox', path: '/' },
      { id: 'outstanding', label: 'Outstanding', icon: 'alert', path: '/outstanding' },
    ],
  },
  {
    label: 'Portfolio',
    items: [
      { id: 'portfolio', label: 'Portfolio', icon: 'grid', path: '/portfolio' },
      { id: 'properties', label: 'Properties', icon: 'home', path: '/properties' },
      { id: 'owners', label: 'Owners', icon: 'star', path: '/owners' },
      { id: 'cahp-entity', label: 'CAHP Entity', icon: 'star', path: '/cahp-entity' },
    ],
  },
  {
    label: 'Filings & Communications',
    items: [
      { id: 'submittals', label: 'Submittals', icon: 'file', path: '/submittals' },
      { id: 'correspondence', label: 'DOR Correspondence', icon: 'mail', path: '/correspondence' },
      { id: 'comms', label: 'Owner Communications', icon: 'mail', path: '/comms' },
    ],
  },
  {
    label: 'Operations',
    items: [
      { id: 'compliance', label: 'Compliance', icon: 'calendar', path: '/compliance' },
      { id: 'billing', label: 'Billing & Disbursements', icon: 'dollar', path: '/billing' },
      { id: 'documents', label: 'Documents', icon: 'folder', path: '/documents' },
      { id: 'untagged', label: 'Untagged Docs', icon: 'alert', path: '/untagged-documents' },
    ],
  },
  {
    label: 'Insights',
    items: [
      { id: 'reports', label: 'Reports', icon: 'file', path: '/reports' },
      { id: 'audit', label: 'Audit Log', icon: 'history', path: '/audit' },
    ],
  },
  {
    label: 'System',
    items: [
      { id: 'settings', label: 'Settings', icon: 'settings', path: '/settings' },
    ],
  },
];

interface SidebarNavProps {
  open: boolean;
  onClose: () => void;
}

export function SidebarNav({ open, onClose }: SidebarNavProps) {
  const navigate = useNavigate();
  const location = useLocation();
  const { role } = useSession();

  // SidebarNav only renders inside SignInGate's authorized branch, but defensively guard:
  if (!role) return null;

  const handleNavigate = (path: string) => {
    navigate(path);
    onClose();
  };

  const isActive = (path: string) => {
    if (path === '/') return location.pathname === '/';
    return location.pathname === path || location.pathname.startsWith(path + '/');
  };

  return (
    <>
      {open && (
        <div
          className="fixed inset-0 bg-black/35 z-40 animate-fade-in"
          onClick={onClose}
          aria-hidden="true"
        />
      )}
      <nav
        className={clsx(
          'fixed left-0 top-14 bottom-0 w-[260px] bg-white border-r border-gray-200 shadow-drawer z-50',
          'flex flex-col overflow-y-auto',
          'transform transition-transform duration-200 ease-out',
          open ? 'translate-x-0' : '-translate-x-full'
        )}
        aria-label="Main navigation"
      >
        <div className="px-4 py-3.5 border-b border-gray-100 flex items-center justify-between">
          <span className="text-[11px] font-bold text-gray-400 uppercase tracking-wider">
            Navigation
          </span>
          <button
            onClick={onClose}
            className="text-gray-400 hover:text-teal-700 transition-colors p-1"
            aria-label="Close navigation menu"
          >
            <Icon name="x" size={18} />
          </button>
        </div>

        {NAV_GROUPS.map((group) => {
          const visibleItems = group.items.filter((item) => canView(role, item.id));
          if (visibleItems.length === 0) return null;

          return (
            <div key={group.label} className="py-1 border-b border-gray-100 last:border-b-0">
              <div className="px-4 pt-2.5 pb-1 text-[10px] font-bold text-gray-400 uppercase tracking-wider">
                {group.label}
              </div>
              {visibleItems.map((item) => {
                const active = isActive(item.path);
                return (
                  <button
                    key={item.id}
                    onClick={() => handleNavigate(item.path)}
                    className={clsx(
                      'w-full flex items-center gap-3 px-4 py-2.5 text-sm font-medium text-left transition-colors',
                      'border-l-[3px]',
                      active
                        ? 'bg-teal-50 text-teal-700 border-gold-500 font-bold'
                        : 'text-gray-600 border-transparent hover:bg-gray-50 hover:text-teal-700'
                    )}
                  >
                    <Icon name={item.icon} size={18} />
                    <span className="flex-1">{item.label}</span>
                    {item.badge != null && (
                      <span
                        className={clsx(
                          'text-[10px] font-bold px-1.5 py-0.5 rounded-full',
                          item.badgeStyle === 'urgent'
                            ? 'bg-error text-white'
                            : 'bg-gold-500 text-teal-900'
                        )}
                      >
                        {item.badge}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          );
        })}

        <div className="mt-auto px-4 py-3 border-t border-gray-100">
          <div className="flex items-center justify-between text-[11px] text-gray-400">
            <span>CAHP Compliance Hub</span>
            <span className="font-mono-data">v0.1.0</span>
          </div>
        </div>
      </nav>
    </>
  );
}
