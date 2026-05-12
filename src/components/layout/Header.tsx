import { Icon } from '../ui/Icon';
import { useSession } from '../../lib/session';
import { ROLE_PERMISSIONS } from '../../lib/permissions';
import { UserMenu } from './UserMenu';

interface HeaderProps {
  onOpenNav: () => void;
}

export function Header({ onOpenNav }: HeaderProps) {
  const { role } = useSession();
  if (!role) return null;
  const roleConfig = ROLE_PERMISSIONS[role];

  return (
    <header className="fixed top-0 left-0 right-0 z-40 h-14 bg-teal-900 text-white flex items-center px-3 shadow-md">
      {/* Left: hamburger + logo + title */}
      <div className="flex items-center gap-2 flex-shrink-0">
        <button
          onClick={onOpenNav}
          className="w-9 h-9 rounded-md flex items-center justify-center hover:bg-white/10 transition-colors"
          aria-label="Open navigation menu"
        >
          <Icon name="menu" size={22} />
        </button>

        <div className="flex items-center gap-2.5 pl-1">
          <div className="w-8 h-8 rounded-md bg-gold-500 text-teal-900 font-bold text-base flex items-center justify-center shadow-sm">
            C
          </div>
          <div className="hidden sm:block">
            <div className="text-[15px] font-semibold leading-tight">CAHP Compliance Hub</div>
            <div className="text-[11px] text-teal-100 leading-tight">Carolina Affordable Housing Project</div>
          </div>
        </div>
      </div>

      {/* Center: search (Phase 2) */}
      <div className="flex-1 px-4 hidden md:flex justify-center">
        <button
          className="w-full max-w-md flex items-center gap-2 px-3 py-1.5 rounded-md bg-white/10 hover:bg-white/15 text-teal-100 text-sm text-left transition-colors disabled:cursor-not-allowed"
          disabled
          title="Search coming in Phase 2"
        >
          <Icon name="search" size={16} />
          <span className="flex-1">Search properties, owners, submittals…</span>
          <span className="text-[11px] font-mono-data opacity-60">⌘K</span>
        </button>
      </div>

      {/* Right: notifications, role badge, user menu */}
      <div className="flex items-center gap-1 flex-shrink-0">
        <button
          className="w-9 h-9 rounded-md flex items-center justify-center hover:bg-white/10 transition-colors relative"
          aria-label="Notifications"
        >
          <Icon name="bell" size={18} />
        </button>

        <div
          className={`hidden sm:inline-flex px-2 py-1 rounded text-[11px] font-semibold uppercase tracking-wider ${roleConfig.color}`}
        >
          {roleConfig.label}
        </div>

        <UserMenu />
      </div>
    </header>
  );
}
