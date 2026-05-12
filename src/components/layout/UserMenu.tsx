import { useState, useRef, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { useSession } from '../../lib/session';
import { ROLE_PERMISSIONS } from '../../lib/permissions';
import { Icon } from '../ui/Icon';

export function UserMenu() {
  const { instance } = useMsal();
  const { user, role } = useSession();
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    if (open) document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open]);

  const handleSignOut = () => {
    instance.logoutRedirect().catch((err) => {
      console.error('Sign-out failed:', err);
    });
  };

  if (!user || !role) return null;
  const roleConfig = ROLE_PERMISSIONS[role];

  return (
    <div className="relative" ref={menuRef}>
      <button
        onClick={() => setOpen(!open)}
        className="w-9 h-9 rounded-full bg-gold-500 text-teal-900 font-bold text-sm flex items-center justify-center hover:bg-gold-200 transition-colors ml-1"
        aria-label={`User menu for ${user.name}`}
      >
        {user.initials}
      </button>

      {open && (
        <div className="absolute right-0 top-11 w-72 bg-white rounded-lg shadow-xl border border-gray-200 z-50 overflow-hidden text-gray-800">
          <div className="p-4 border-b border-gray-100">
            <div className="font-semibold text-gray-900">{user.name}</div>
            <div className="text-xs text-gray-500 mt-0.5 truncate" title={user.email}>
              {user.email}
            </div>
            <div className="text-xs text-gray-500 mt-1">{user.org}</div>
            <div
              className={`inline-flex px-2 py-1 rounded text-[11px] font-semibold uppercase tracking-wider mt-3 ${roleConfig.color}`}
            >
              {roleConfig.label}
            </div>
          </div>
          <button
            onClick={handleSignOut}
            className="w-full text-left px-4 py-3 text-sm text-gray-700 hover:bg-gray-50 flex items-center gap-2 transition-colors"
          >
            <Icon name="external" size={16} />
            Sign out
          </button>
        </div>
      )}
    </div>
  );
}
