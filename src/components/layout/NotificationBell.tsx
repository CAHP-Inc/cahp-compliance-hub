import { useState, useMemo, useRef, useEffect } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type Notification,
  type NotificationType,
} from '../../lib/sharepoint';
import { useSession } from '../../lib/session';
import { Icon } from '../ui/Icon';

const TYPE_ICONS: Record<NotificationType, 'check' | 'calendar' | 'file' | 'star' | 'alert' | 'inbox'> = {
  TaskAssigned: 'check',
  DeadlineApproaching: 'calendar',
  SubmittalUpdate: 'file',
  OwnerInvited: 'star',
  SystemAlert: 'alert',
  Other: 'inbox',
};

export function NotificationBell() {
  const navigate = useNavigate();
  const { user } = useSession();
  const { data, refetch } = useSharePointList<Notification>(LIST_NAMES.Notifications, { top: 100 });
  const [open, setOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close on outside click
  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const myNotifications = useMemo(() => {
    if (!data || !user) return [];
    const myUpn = user.email.toLowerCase();
    return data
      .filter((n) => (n.fields.NotifAssignedTo ?? '').toLowerCase() === myUpn)
      .sort((a, b) => {
        const da = a.createdDateTime ? new Date(a.createdDateTime).getTime() : 0;
        const db = b.createdDateTime ? new Date(b.createdDateTime).getTime() : 0;
        return db - da;
      });
  }, [data, user]);

  const unread = myNotifications.filter((n) => !n.fields.NotifIsRead);
  const recent = myNotifications.slice(0, 5);

  const handleClick = async (n: Notification) => {
    setOpen(false);
    if (!n.fields.NotifIsRead) {
      try {
        await updateListItem(LIST_NAMES.Notifications, n.id, { NotifIsRead: true });
        refetch?.();
      } catch {
        // non-blocking
      }
    }
    if (n.fields.NotifUrl) {
      navigate(n.fields.NotifUrl);
    }
  };

  return (
    <div className="relative" ref={dropdownRef}>
      <button
        onClick={() => setOpen((v) => !v)}
        className="relative w-9 h-9 rounded-md flex items-center justify-center hover:bg-white/10 transition-colors"
        aria-label={unread.length > 0 ? `${unread.length} unread notifications` : 'Notifications'}
      >
        <Icon name="alert" size={18} />
        {unread.length > 0 && (
          <span className="absolute -top-0.5 -right-0.5 min-w-[16px] h-4 px-1 rounded-full bg-gold-500 text-teal-900 text-[10px] font-bold flex items-center justify-center">
            {unread.length > 9 ? '9+' : unread.length}
          </span>
        )}
      </button>

      {open && (
        <div className="absolute right-0 mt-2 w-80 bg-white text-gray-900 rounded-md shadow-xl border border-gray-200 overflow-hidden z-50">
          <div className="px-4 py-2.5 border-b border-gray-100 flex items-center justify-between bg-gray-50">
            <h3 className="text-sm font-semibold text-teal-700">Notifications</h3>
            {unread.length > 0 && (
              <span className="text-[10px] font-semibold bg-gold-500 text-teal-900 px-1.5 py-0.5 rounded">
                {unread.length} unread
              </span>
            )}
          </div>
          {recent.length === 0 ? (
            <div className="px-4 py-8 text-center">
              <Icon name="inbox" size={20} className="text-gray-300 mx-auto mb-1" />
              <p className="text-xs text-gray-500">No notifications</p>
            </div>
          ) : (
            <ul className="divide-y divide-gray-100 max-h-96 overflow-y-auto">
              {recent.map((n) => {
                const isUnread = !n.fields.NotifIsRead;
                const type = n.fields.NotifType ?? 'Other';
                return (
                  <li
                    key={n.id}
                    onClick={() => handleClick(n)}
                    className={`px-4 py-2.5 flex items-start gap-2 cursor-pointer transition-colors ${
                      isUnread ? 'bg-teal-50/40 hover:bg-teal-50' : 'hover:bg-gray-50'
                    }`}
                  >
                    {isUnread && <div className="w-1.5 h-1.5 rounded-full bg-teal-700 flex-shrink-0 mt-1.5" />}
                    {!isUnread && <div className="w-1.5 h-1.5 flex-shrink-0 mt-1.5" />}
                    <Icon name={TYPE_ICONS[type]} size={12} className="text-teal-700 flex-shrink-0 mt-1" />
                    <div className="flex-1 min-w-0">
                      <div className={`text-xs leading-snug ${isUnread ? 'font-semibold text-gray-900' : 'text-gray-700'}`}>
                        {n.fields.Title}
                      </div>
                      <div className="text-[10px] text-gray-500 font-mono-data mt-0.5">
                        {n.createdDateTime && new Date(n.createdDateTime).toLocaleString()}
                      </div>
                    </div>
                  </li>
                );
              })}
            </ul>
          )}
          <div className="px-4 py-2 border-t border-gray-100 bg-gray-50 text-center">
            <Link
              to="/notifications"
              onClick={() => setOpen(false)}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium"
            >
              View all →
            </Link>
          </div>
        </div>
      )}
    </div>
  );
}
