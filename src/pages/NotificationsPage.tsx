import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type Notification,
  type NotificationType,
} from '../lib/sharepoint';
import { useSession } from '../lib/session';
import { Icon } from '../components/ui/Icon';
import { formatDateTime } from '../lib/dates';

const TYPE_LABEL: Record<NotificationType, string> = {
  TaskAssigned: 'Task Assigned',
  DeadlineApproaching: 'Deadline',
  SubmittalUpdate: 'Submittal',
  OwnerInvited: 'Owner Invited',
  SystemAlert: 'System Alert',
  Other: 'Other',
};

const TYPE_STYLES: Record<NotificationType, string> = {
  TaskAssigned: 'bg-blue-100 text-blue-800',
  DeadlineApproaching: 'bg-amber-100 text-amber-800',
  SubmittalUpdate: 'bg-teal-100 text-teal-800',
  OwnerInvited: 'bg-purple-100 text-purple-800',
  SystemAlert: 'bg-red-100 text-red-800',
  Other: 'bg-gray-100 text-gray-700',
};

const TYPE_ICONS: Record<NotificationType, 'check' | 'calendar' | 'file' | 'star' | 'alert' | 'inbox'> = {
  TaskAssigned: 'check',
  DeadlineApproaching: 'calendar',
  SubmittalUpdate: 'file',
  OwnerInvited: 'star',
  SystemAlert: 'alert',
  Other: 'inbox',
};

export function NotificationsPage() {
  const navigate = useNavigate();
  const { user } = useSession();
  const { data, loading, error, refetch } = useSharePointList<Notification>(
    LIST_NAMES.Notifications,
    { top: 500 }
  );

  const [showRead, setShowRead] = useState(false);
  const [typeFilter, setTypeFilter] = useState<NotificationType | 'All'>('All');
  const [updatingId, setUpdatingId] = useState<string | null>(null);
  const [bulkUpdating, setBulkUpdating] = useState(false);

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

  const filtered = useMemo(() => {
    return myNotifications.filter((n) => {
      if (!showRead && n.fields.NotifIsRead) return false;
      if (typeFilter !== 'All' && n.fields.NotifType !== typeFilter) return false;
      return true;
    });
  }, [myNotifications, showRead, typeFilter]);

  const unreadCount = myNotifications.filter((n) => !n.fields.NotifIsRead).length;

  const handleClick = async (n: Notification) => {
    // Mark as read first, then navigate
    if (!n.fields.NotifIsRead) {
      setUpdatingId(n.id);
      try {
        await updateListItem(LIST_NAMES.Notifications, n.id, { NotifIsRead: true });
      } catch {
        // non-blocking
      } finally {
        setUpdatingId(null);
      }
    }
    if (n.fields.NotifUrl) {
      navigate(n.fields.NotifUrl);
    }
  };

  const handleMarkRead = async (e: React.MouseEvent, n: Notification) => {
    e.stopPropagation();
    setUpdatingId(n.id);
    try {
      await updateListItem(LIST_NAMES.Notifications, n.id, { NotifIsRead: !n.fields.NotifIsRead });
      await refetch();
    } finally {
      setUpdatingId(null);
    }
  };

  const handleMarkAllRead = async () => {
    const unread = myNotifications.filter((n) => !n.fields.NotifIsRead);
    if (unread.length === 0) return;
    setBulkUpdating(true);
    try {
      for (const n of unread) {
        try {
          await updateListItem(LIST_NAMES.Notifications, n.id, { NotifIsRead: true });
        } catch {
          // continue
        }
      }
      await refetch();
    } finally {
      setBulkUpdating(false);
    }
  };

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Notifications</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Notifications</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Notifications</h1>
          <p className="text-sm text-gray-500 mt-1">
            {unreadCount > 0
              ? `${unreadCount} unread`
              : myNotifications.length === 0
                ? 'No notifications'
                : 'All caught up'}
          </p>
        </div>
        {unreadCount > 0 && (
          <button
            onClick={handleMarkAllRead}
            disabled={bulkUpdating}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
          >
            {bulkUpdating && <div className="w-3 h-3 rounded-full border-2 border-gray-500 border-r-transparent animate-spin" />}
            Mark all as read
          </button>
        )}
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <label className="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer">
          <input
            type="checkbox"
            checked={showRead}
            onChange={(e) => setShowRead(e.target.checked)}
          />
          Show read
        </label>
        <select
          value={typeFilter}
          onChange={(e) => setTypeFilter(e.target.value as NotificationType | 'All')}
          className="px-2 py-1 border border-gray-200 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All types</option>
          {(Object.keys(TYPE_LABEL) as NotificationType[]).map((t) => (
            <option key={t} value={t}>{TYPE_LABEL[t]}</option>
          ))}
        </select>
        {filtered.length !== myNotifications.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {myNotifications.length}</span>
        )}
      </div>

      {myNotifications.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <Icon name="inbox" size={32} className="text-blue-400 mx-auto mb-2" />
          <p className="text-base font-semibold text-blue-900 mb-1">No notifications yet</p>
          <p className="text-sm text-blue-800">
            You'll see notifications here when items are assigned to you, when submittals you're tracking
            update, or when system events affect your work.
          </p>
        </div>
      ) : filtered.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500">No notifications match your filters.</p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <ul className="divide-y divide-gray-100">
            {filtered.map((n) => {
              const isUnread = !n.fields.NotifIsRead;
              const type = n.fields.NotifType ?? 'Other';
              return (
                <li
                  key={n.id}
                  onClick={() => handleClick(n)}
                  className={`px-4 py-3 flex items-start gap-3 transition-colors cursor-pointer ${
                    isUnread ? 'bg-teal-50/40 hover:bg-teal-50' : 'hover:bg-gray-50'
                  } ${updatingId === n.id ? 'opacity-60' : ''}`}
                >
                  {isUnread && <div className="w-2 h-2 rounded-full bg-teal-700 flex-shrink-0 mt-2" />}
                  {!isUnread && <div className="w-2 h-2 flex-shrink-0 mt-2" />}
                  <Icon name={TYPE_ICONS[type]} size={14} className="text-teal-700 flex-shrink-0 mt-0.5" />
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                      <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded uppercase tracking-wider ${TYPE_STYLES[type]}`}>
                        {TYPE_LABEL[type]}
                      </span>
                      {n.fields.NotifTargetType && (
                        <span className="text-[10px] text-gray-500 font-mono-data">
                          {n.fields.NotifTargetType}
                        </span>
                      )}
                    </div>
                    <div className={`text-sm ${isUnread ? 'font-semibold text-gray-900' : 'text-gray-700'}`}>
                      {n.fields.Title}
                    </div>
                    <div className="text-[11px] text-gray-500 font-mono-data mt-0.5">
                      {n.createdDateTime && formatDateTime(n.createdDateTime)}
                    </div>
                  </div>
                  <button
                    onClick={(e) => handleMarkRead(e, n)}
                    disabled={updatingId === n.id}
                    className="text-[11px] text-teal-700 hover:text-teal-900 underline flex-shrink-0 disabled:opacity-50"
                  >
                    {isUnread ? 'Mark read' : 'Mark unread'}
                  </button>
                </li>
              );
            })}
          </ul>
        </div>
      )}
    </div>
  );
}
