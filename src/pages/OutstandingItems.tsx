import { useState, useMemo, useEffect } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import {
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type OutstandingItem,
  type Property,
  type ItemStatus,
  type ItemPriority,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateOnly, parseDateOnly, toDateOnlyISO } from '../lib/dates';
import { NewOutstandingItemModal } from '../components/NewOutstandingItemModal';
import { LinkOrUploadDocumentModal } from '../components/LinkOrUploadDocumentModal';
import { ExportOutstandingItemsModal } from '../components/ExportOutstandingItemsModal';
import { AssigneePicker } from '../components/AssigneePicker';

// =============================================================================
// Status / category styles
// =============================================================================

const STATUS_STYLES: Record<ItemStatus, string> = {
  'Not Started': 'bg-gray-100 text-gray-800',
  'In Progress': 'bg-blue-100 text-blue-800',
  'Blocked': 'bg-red-100 text-red-800',
  'Done': 'bg-green-100 text-green-800',
  'Requested': 'bg-gray-100 text-gray-800',         // legacy alias for Not Started
  'Overdue': 'bg-amber-100 text-amber-800',
  'Received': 'bg-green-100 text-green-800',        // legacy alias for Done
  'Not Applicable': 'bg-gray-100 text-gray-500',
};

const PRIORITY_STYLES: Record<ItemPriority, string> = {
  Critical: 'bg-red-100 text-red-800 border-red-200',
  High: 'bg-amber-100 text-amber-800 border-amber-200',
  Medium: 'bg-blue-100 text-blue-800 border-blue-200',
  Low: 'bg-gray-100 text-gray-600 border-gray-200',
};

/** Normalize legacy statuses to one of the four canonical buckets. */
function normalizeStatus(status: ItemStatus | undefined): 'Not Started' | 'In Progress' | 'Blocked' | 'Done' {
  if (!status || status === 'Requested') return 'Not Started';
  if (status === 'Received' || status === 'Not Applicable') return 'Done';
  if (status === 'Overdue') return 'In Progress'; // legacy — overdue is a visual, not a column
  if (status === 'Not Started' || status === 'In Progress' || status === 'Blocked' || status === 'Done') return status;
  return 'Not Started';
}

/** Calendar / chip color per status — yellow / blue / red / green. */
const STATUS_CHIP_COLOR: Record<'Not Started' | 'In Progress' | 'Blocked' | 'Done', string> = {
  'Not Started': 'bg-yellow-100 text-yellow-800',
  'In Progress': 'bg-blue-100 text-blue-800',
  'Blocked': 'bg-red-100 text-red-800',
  'Done': 'bg-green-100 text-green-800',
};

function isOverdue(item: OutstandingItem): boolean {
  if (!item.fields.DueDate) return false;
  const due = parseDateOnly(item.fields.DueDate);
  if (!due) return false;
  const closed = item.fields.ItemStatus === 'Done' || item.fields.ItemStatus === 'Received' || item.fields.ItemStatus === 'Not Applicable';
  return due.getTime() < Date.now() && !closed;
}

// =============================================================================
// Page
// =============================================================================

type ViewMode = 'list' | 'calendar';

export function OutstandingItems() {
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const items = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [view, setView] = useState<ViewMode>('list');
  const [search, setSearch] = useState('');
  const [priorityFilter, setPriorityFilter] = useState<ItemPriority | 'All'>('All');
  const [assigneeFilter, setAssigneeFilter] = useState<string>('All');
  // Seed the property filter from ?property={id} so deep links from MyDay land
  // pre-filtered to the property the user clicked.
  const [propertyFilter, setPropertyFilter] = useState<string>(
    () => searchParams.get('property') || 'All',
  );
  // Specific-day filter — set by clicking a day on the Calendar view to drill
  // into the list. Stored as YYYY-MM-DD; empty string = no filter.
  const [dueDateFilter, setDueDateFilter] = useState<string>('');
  const [showClosed, setShowClosed] = useState(false);
  const [updatingId, setUpdatingId] = useState<string | null>(null);
  const [newItemOpen, setNewItemOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [linkUploadItem, setLinkUploadItem] = useState<OutstandingItem | null>(null);

  const loading = items.loading || properties.loading;
  const error = items.error || properties.error;

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  // Build list of unique assignees
  const assignees = useMemo(() => {
    if (!items.data) return [];
    return Array.from(
      new Set(items.data.map((i) => i.fields.AssignedTo).filter(Boolean))
    ) as string[];
  }, [items.data]);

  const filtered = useMemo(() => {
    if (!items.data) return [];
    return items.data.filter((i) => {
      const f = i.fields;
      if (!showClosed) {
        // Hide Done / Received / Not Applicable when toggle off
        const isClosed =
          f.ItemStatus === 'Done' ||
          f.ItemStatus === 'Received' ||
          f.ItemStatus === 'Not Applicable';
        if (isClosed) return false;
      }
      if (search) {
        const propName = f.PropertyLookupId
          ? propertiesById.get(String(f.PropertyLookupId))?.fields.Title ?? ''
          : '';
        const hay = `${f.Title ?? ''} ${propName} ${f.ItemNotes ?? ''}`.toLowerCase();
        if (!hay.includes(search.toLowerCase())) return false;
      }
      if (priorityFilter !== 'All' && f.Priority !== priorityFilter) return false;
      if (assigneeFilter !== 'All' && f.AssignedTo !== assigneeFilter) return false;
      if (propertyFilter !== 'All' && String(f.PropertyLookupId) !== propertyFilter) return false;
      if (dueDateFilter) {
        // Match items whose DueDate falls on the selected YYYY-MM-DD (local time)
        const itemDate = f.DueDate ? String(f.DueDate).slice(0, 10) : '';
        if (itemDate !== dueDateFilter) return false;
      }
      return true;
    });
  }, [items.data, search, priorityFilter, assigneeFilter, propertyFilter, dueDateFilter, showClosed, propertiesById]);

  const stats = useMemo(() => {
    if (!items.data) return null;
    const open = items.data.filter(
      (i) =>
        i.fields.ItemStatus !== 'Done' &&
        i.fields.ItemStatus !== 'Received' &&
        i.fields.ItemStatus !== 'Not Applicable'
    );
    return {
      total: items.data.length,
      open: open.length,
      overdue: open.filter(isOverdue).length,
      critical: open.filter((i) => i.fields.Priority === 'Critical').length,
    };
  }, [items.data]);

  // Inline assignee update from the list view (saved on blur or Enter)
  const handleAssigneeChange = async (itemId: string, newAssignee: string) => {
    setUpdatingId(itemId);
    try {
      const trimmed = newAssignee.trim();
      await updateListItem(LIST_NAMES.Outstanding, itemId, {
        AssignedTo: trimmed || (null as unknown as undefined),
      });
      await items.refetch?.();
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error('Assignee update failed:', e);
      alert(`Assignee update failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setUpdatingId(null);
    }
  };

  // Inline due-date update from the list view. Takes a YYYY-MM-DD string
  // from <input type="date"> (empty = clear the date).
  const handleDueDateChange = async (itemId: string, newDueDate: string) => {
    setUpdatingId(itemId);
    try {
      const iso = newDueDate ? toDateOnlyISO(newDueDate) : null;
      await updateListItem(LIST_NAMES.Outstanding, itemId, {
        DueDate: iso as unknown as undefined,
      });
      await items.refetch?.();
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error('Due date update failed:', e);
      alert(`Due date update failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setUpdatingId(null);
    }
  };

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Outstanding Items</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading items…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Outstanding Items</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load outstanding items</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (!items.data || !stats) return null;

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Outstanding Items</h1>
          <p className="text-sm text-gray-500 mt-1">
            Master task list. Items are auto-created by the Property Wizard, DOR Correspondence cascade, and other workflows — plus manual entries.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setExportOpen(true)}
            disabled={filtered.length === 0}
            className="bg-white border border-gray-300 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed text-gray-700 px-3 py-1.5 rounded-md text-sm font-medium flex items-center gap-1.5"
            title="Export an assignee-ready list of pending items"
          >
            <Icon name="file" size={14} />
            Export
          </button>
          <button
            onClick={() => setNewItemOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium flex items-center gap-1.5"
          >
            <Icon name="plus" size={14} />
            New Item
          </button>
          <div className="inline-flex bg-gray-100 rounded-md p-0.5">
            {([
              { id: 'list',     label: 'List',     icon: 'check' as const },
              { id: 'calendar', label: 'Calendar', icon: 'calendar' as const },
            ]).map((opt) => (
              <button
                key={opt.id}
                onClick={() => setView(opt.id as ViewMode)}
                className={`px-2.5 py-1 text-xs font-medium rounded inline-flex items-center gap-1.5 ${
                  view === opt.id ? 'bg-white shadow text-teal-700' : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <Icon name={opt.icon} size={12} />
                {opt.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total Items" value={stats.total} />
        <KPI label="Open" value={stats.open} />
        <KPI label="Overdue" value={stats.overdue} accent={stats.overdue > 0 ? 'danger' : 'default'} />
        <KPI label="Critical" value={stats.critical} accent={stats.critical > 0 ? 'danger' : 'default'} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[200px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search items…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={priorityFilter}
          onChange={(e) => setPriorityFilter(e.target.value as ItemPriority | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All priorities</option>
          <option value="Critical">Critical</option>
          <option value="High">High</option>
          <option value="Medium">Medium</option>
          <option value="Low">Low</option>
        </select>
        {assignees.length > 0 && (
          <select
            value={assigneeFilter}
            onChange={(e) => setAssigneeFilter(e.target.value)}
            className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
          >
            <option value="All">All assignees</option>
            {assignees.map((a) => (<option key={a} value={a}>{a}</option>))}
          </select>
        )}
        {properties.data && properties.data.length > 0 && (
          <select
            value={propertyFilter}
            onChange={(e) => setPropertyFilter(e.target.value)}
            className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white max-w-[180px] truncate"
          >
            <option value="All">All properties</option>
            {properties.data.map((p) => (<option key={p.id} value={p.id}>{p.fields.Title}</option>))}
          </select>
        )}
        <label className="flex items-center gap-1.5 text-sm text-gray-700 cursor-pointer px-2">
          <input
            type="checkbox"
            checked={showClosed}
            onChange={(e) => setShowClosed(e.target.checked)}
          />
          Show closed
        </label>
        {filtered.length !== items.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {items.data.length}</span>
        )}
      </div>

      {/* Content — kanban or list */}
      {items.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No outstanding items yet</p>
          <p className="text-sm text-blue-800">
            Items are auto-created when you create a new property (Property Wizard generates 7 document-collection items)
            or log a DOR letter with a response deadline. They'll appear here.
          </p>
        </div>
      ) : view === 'list' ? (
        <>
          {dueDateFilter && (
            <div className="bg-amber-50 border border-amber-200 rounded-md p-2.5 mb-3 text-xs flex items-center justify-between gap-2">
              <span className="text-amber-900">
                Filtered to items due <strong>{formatDateOnly(dueDateFilter)}</strong>
                {' · '}{filtered.length} match{filtered.length === 1 ? '' : 'es'}
              </span>
              <button
                onClick={() => setDueDateFilter('')}
                className="text-amber-700 hover:text-amber-900 font-medium px-2 py-0.5 rounded hover:bg-amber-100"
              >
                Clear date filter ×
              </button>
            </div>
          )}
          <ListView
            items={filtered}
            propertiesById={propertiesById}
            onRowClick={(itemId) => navigate(`/outstanding-items/${itemId}`)}
            onLinkUpload={(item) => setLinkUploadItem(item)}
            onAssigneeChange={handleAssigneeChange}
            onDueDateChange={handleDueDateChange}
            updatingId={updatingId}
          />
        </>
      ) : (
        <CalendarView
          items={filtered}
          propertiesById={propertiesById}
          onItemClick={(itemId) => navigate(`/outstanding-items/${itemId}`)}
          onDayClick={(dateStr) => {
            setDueDateFilter(dateStr);
            setView('list');
          }}
        />
      )}

      {newItemOpen && (
        <NewOutstandingItemModal
          onClose={() => setNewItemOpen(false)}
          onSuccess={() => {
            setNewItemOpen(false);
            items.refetch?.();
          }}
        />
      )}

      {linkUploadItem && (
        <LinkOrUploadDocumentModal
          item={linkUploadItem}
          onClose={() => setLinkUploadItem(null)}
          onSuccess={() => {
            setLinkUploadItem(null);
            items.refetch?.();
          }}
        />
      )}

      {exportOpen && (
        <ExportOutstandingItemsModal
          items={filtered}
          propertiesById={propertiesById}
          propertyTitle={
            propertyFilter !== 'All'
              ? propertiesById.get(propertyFilter)?.fields.Title
              : undefined
          }
          onClose={() => setExportOpen(false)}
        />
      )}
    </div>
  );
}

// =============================================================================
// List view
// =============================================================================

function ListView({
  items,
  propertiesById,
  onRowClick,
  onLinkUpload,
  onAssigneeChange,
  onDueDateChange,
  updatingId,
}: {
  items: OutstandingItem[];
  propertiesById: Map<string, Property>;
  onRowClick: (id: string) => void;
  onLinkUpload: (item: OutstandingItem) => void;
  onAssigneeChange: (itemId: string, newAssignee: string) => void;
  onDueDateChange: (itemId: string, newDueDate: string) => void;
  updatingId: string | null;
}) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
          <tr>
            <th className="px-4 py-3 text-left">Title</th>
            <th className="px-4 py-3 text-left">Property</th>
            <th className="px-4 py-3 text-left">Status</th>
            <th className="px-4 py-3 text-left">Priority</th>
            <th className="px-4 py-3 text-left">Assignee</th>
            <th className="px-4 py-3 text-left">Due Date</th>
            <th className="px-4 py-3 text-left">Category</th>
            <th className="px-4 py-3 text-left">Document</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {items.map((item) => {
            const property = item.fields.PropertyLookupId
              ? propertiesById.get(String(item.fields.PropertyLookupId))
              : null;
            const overdue = isOverdue(item);
            const hasDoc = Boolean(item.fields.RelatedDocUrl);
            return (
              <tr
                key={item.id}
                className={`hover:bg-gray-50 transition-colors ${overdue ? 'bg-red-50' : ''}`}
              >
                <td
                  className="px-4 py-3 font-medium text-gray-900 cursor-pointer"
                  onClick={() => onRowClick(item.id)}
                >
                  {overdue && <span className="text-error mr-1">⚠</span>}
                  {item.fields.Title}
                </td>
                <td
                  className="px-4 py-3 text-xs text-gray-700 cursor-pointer"
                  onClick={() => onRowClick(item.id)}
                >
                  {property?.fields.Title ?? '—'}
                </td>
                <td
                  className="px-4 py-3 cursor-pointer"
                  onClick={() => onRowClick(item.id)}
                >
                  {item.fields.ItemStatus ? (
                    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[item.fields.ItemStatus]}`}>
                      {item.fields.ItemStatus}
                    </span>
                  ) : '—'}
                </td>
                <td
                  className="px-4 py-3 cursor-pointer"
                  onClick={() => onRowClick(item.id)}
                >
                  {item.fields.Priority ? (
                    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold border ${PRIORITY_STYLES[item.fields.Priority]}`}>
                      {item.fields.Priority}
                    </span>
                  ) : '—'}
                </td>
                <td
                  className="px-4 py-3"
                  onClick={(e) => e.stopPropagation()}
                >
                  <InlineAssigneeCell
                    value={item.fields.AssignedTo}
                    saving={updatingId === item.id}
                    onCommit={(v) => onAssigneeChange(item.id, v)}
                  />
                </td>
                <td
                  className="px-4 py-3"
                  onClick={(e) => e.stopPropagation()}
                >
                  <InlineDueDateCell
                    value={item.fields.DueDate}
                    saving={updatingId === item.id}
                    overdue={overdue}
                    onCommit={(v) => onDueDateChange(item.id, v)}
                  />
                </td>
                <td
                  className="px-4 py-3 text-xs text-gray-600 cursor-pointer"
                  onClick={() => onRowClick(item.id)}
                >
                  {item.fields.ItemCategory || '—'}
                </td>
                <td className="px-4 py-3 text-xs">
                  {hasDoc ? (
                    <a
                      href={item.fields.RelatedDocUrl!}
                      target="_blank"
                      rel="noopener noreferrer"
                      onClick={(e) => e.stopPropagation()}
                      className="text-success hover:underline inline-flex items-center gap-1 max-w-[180px] truncate"
                      title={item.fields.RelatedDocFilename}
                    >
                      <Icon name="check" size={11} />
                      <span className="truncate">{item.fields.RelatedDocFilename ?? 'View'}</span>
                    </a>
                  ) : (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onLinkUpload(item);
                      }}
                      className="text-gold-700 hover:text-gold-900 underline inline-flex items-center gap-1"
                    >
                      <Icon name="plus" size={11} />
                      Link / Upload
                    </button>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// =============================================================================
// KPI
// =============================================================================

function KPI({
  label,
  value,
  accent = 'default',
}: {
  label: string;
  value: number;
  accent?: 'default' | 'warning' | 'danger';
}) {
  const accentClass =
    accent === 'danger' ? 'text-error' : accent === 'warning' ? 'text-warning' : 'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}

/**
 * Inline assignee editor for the list view.
 *
 * Tracks a local draft so AssigneePicker's per-keystroke onChange doesn't hammer
 * the API. Saves on blur or Enter when the draft differs from the committed
 * value. Re-syncs the draft if the prop changes externally (e.g., after refetch).
 */
function InlineAssigneeCell({
  value,
  saving,
  onCommit,
}: {
  value: string | undefined;
  saving: boolean;
  onCommit: (newValue: string) => void;
}) {
  const [draft, setDraft] = useState(value ?? '');
  useEffect(() => {
    setDraft(value ?? '');
  }, [value]);

  const commit = () => {
    if (draft.trim() === (value ?? '').trim()) return;
    onCommit(draft);
  };

  return (
    <div className="relative min-w-[140px]">
      <AssigneePicker
        value={draft}
        onChange={(v) => setDraft(v)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            (e.currentTarget as HTMLInputElement).blur();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setDraft(value ?? '');
            (e.currentTarget as HTMLInputElement).blur();
          }
        }}
        disabled={saving}
        placeholder="Unassigned"
        className="w-full pl-2 pr-6 py-1 border border-gray-200 rounded text-xs bg-white focus:outline-none focus:border-teal-500 hover:border-gray-300 disabled:bg-gray-50"
      />
      {saving && (
        <div className="absolute right-1.5 top-1/2 -translate-y-1/2 pointer-events-none">
          <div className="w-2.5 h-2.5 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Calendar view — month grid keyed off DueDate, with assignee tally
// =============================================================================

const UNASSIGNED_KEY = '__unassigned__';

function startOfMonth(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function startOfWeek(d: Date): Date {
  // Sunday-anchored
  const out = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  out.setDate(out.getDate() - out.getDay());
  return out;
}

function formatYmd(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

type CalendarMode = 'month' | 'week';

function CalendarView({
  items,
  propertiesById,
  onItemClick,
  onDayClick,
}: {
  items: OutstandingItem[];
  propertiesById: Map<string, Property>;
  onItemClick: (id: string) => void;
  onDayClick: (dateYmd: string) => void;
}) {
  const [mode, setMode] = useState<CalendarMode>('month');
  // `anchor` is any date inside the visible month or week. Navigation shifts
  // by one month or one week depending on `mode`.
  const [anchor, setAnchor] = useState<Date>(() => new Date());

  // Bucket items by YYYY-MM-DD local date (drop items with no DueDate)
  const itemsByDate = useMemo(() => {
    const map = new Map<string, OutstandingItem[]>();
    for (const item of items) {
      const d = parseDateOnly(item.fields.DueDate);
      if (!d) continue;
      map.set(formatYmd(d), (map.get(formatYmd(d)) ?? []).concat(item));
    }
    return map;
  }, [items]);

  // Day grid + range bounds depend on mode
  const { days, rangeStart, rangeEnd, rangeLabel } = useMemo(() => {
    if (mode === 'week') {
      const start = startOfWeek(anchor);
      const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59);
      const days: Date[] = [];
      for (let i = 0; i < 7; i++) {
        days.push(new Date(start.getFullYear(), start.getMonth(), start.getDate() + i));
      }
      // Label: "Week of May 18 – 24, 2026" (collapse year if same)
      const lastDay = days[6];
      const sameMonth = start.getMonth() === lastDay.getMonth();
      const label = sameMonth
        ? `${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${lastDay.getDate()}, ${lastDay.getFullYear()}`
        : `${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${lastDay.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}, ${lastDay.getFullYear()}`;
      return { days, rangeStart: start, rangeEnd: end, rangeLabel: `Week of ${label}` };
    }
    // month
    const first = startOfMonth(anchor);
    const last = new Date(first.getFullYear(), first.getMonth() + 1, 0, 23, 59, 59);
    const gridStart = new Date(first.getFullYear(), first.getMonth(), 1 - first.getDay());
    const days: Date[] = [];
    for (let i = 0; i < 42; i++) {
      days.push(new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i));
    }
    return {
      days,
      rangeStart: first,
      rangeEnd: last,
      rangeLabel: anchor.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }),
    };
  }, [mode, anchor]);

  // Per-person tally scoped to the visible range
  const assigneeTally = useMemo(() => {
    const counts = new Map<string, number>();
    for (const item of items) {
      const d = parseDateOnly(item.fields.DueDate);
      if (!d) continue;
      if (d < rangeStart || d > rangeEnd) continue;
      const assignee = (item.fields.AssignedTo ?? '').trim() || UNASSIGNED_KEY;
      counts.set(assignee, (counts.get(assignee) ?? 0) + 1);
    }
    return Array.from(counts.entries()).sort((a, b) => {
      if (a[0] === UNASSIGNED_KEY) return 1;
      if (b[0] === UNASSIGNED_KEY) return -1;
      return b[1] - a[1];
    });
  }, [items, rangeStart, rangeEnd]);

  const today = new Date();
  const isSameDay = (a: Date, b: Date) =>
    a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();

  const goPrev = () => {
    if (mode === 'week') {
      setAnchor(new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - 7));
    } else {
      setAnchor(new Date(anchor.getFullYear(), anchor.getMonth() - 1, 1));
    }
  };
  const goNext = () => {
    if (mode === 'week') {
      setAnchor(new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() + 7));
    } else {
      setAnchor(new Date(anchor.getFullYear(), anchor.getMonth() + 1, 1));
    }
  };
  const goToday = () => setAnchor(new Date());

  // Render one day cell — items shown as one chip per item, color-coded by
  // status (yellow/blue/red/green). Clicking the day header (or the bare cell
  // area) drills into the List view filtered to that date.
  const renderDayCell = (day: Date) => {
    const inRange = day >= rangeStart && day <= rangeEnd;
    const key = formatYmd(day);
    const dayItems = itemsByDate.get(key) ?? [];
    const isToday = isSameDay(day, today);
    const visibleLimit = mode === 'week' ? 12 : 4;

    return (
      <div
        key={key}
        className={`border border-gray-100 p-1.5 text-xs flex flex-col ${
          mode === 'week' ? 'min-h-[260px]' : 'min-h-[100px]'
        } ${inRange ? 'bg-white' : 'bg-gray-50 text-gray-400'}`}
      >
        <button
          type="button"
          onClick={() => onDayClick(key)}
          className="flex items-center justify-between mb-1 hover:bg-gray-50 rounded -mx-1 px-1 py-0.5 text-left"
          title="Show this day's items in the list view"
        >
          <span
            className={
              isToday
                ? 'bg-teal-700 text-white px-1.5 py-0.5 rounded-full text-[10px] font-bold font-mono-data'
                : 'text-[11px] text-gray-600 font-mono-data'
            }
          >
            {mode === 'week'
              ? day.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' })
              : day.getDate()}
          </span>
          {dayItems.length > 0 && (
            <span className="text-[10px] text-gray-400 font-mono-data">{dayItems.length}</span>
          )}
        </button>
        <div className="flex flex-col gap-0.5 overflow-hidden">
          {dayItems.slice(0, visibleLimit).map((item) => {
            const status = normalizeStatus(item.fields.ItemStatus);
            const propertyTitle = item.fields.PropertyLookupId
              ? propertiesById.get(String(item.fields.PropertyLookupId))?.fields.Title
              : undefined;
            const assignee = (item.fields.AssignedTo ?? '').trim();
            return (
              <button
                key={item.id}
                onClick={() => onItemClick(item.id)}
                className={`text-left px-1.5 py-0.5 rounded text-[10px] font-medium truncate ${STATUS_CHIP_COLOR[status]} hover:opacity-80`}
                title={`${item.fields.Title}${propertyTitle ? ` — ${propertyTitle}` : ''}${assignee ? ` — ${assignee}` : ''} (${status})`}
              >
                {assignee ? <span className="font-semibold mr-1">{assignee}:</span> : <span className="font-semibold mr-1 text-gray-700">·</span>}
                {item.fields.Title}
              </button>
            );
          })}
          {dayItems.length > visibleLimit && (
            <button
              onClick={() => onDayClick(key)}
              className="text-[10px] text-gray-500 hover:text-teal-700 text-left"
            >
              +{dayItems.length - visibleLimit} more →
            </button>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      <div className="bg-white border border-gray-200 rounded-lg shadow-card p-3 flex items-center justify-between flex-wrap gap-3">
        <div className="flex items-center gap-2">
          <button
            onClick={goPrev}
            className="px-2 py-1 border border-gray-300 rounded text-xs hover:bg-gray-50"
            aria-label={`Previous ${mode}`}
          >
            ←
          </button>
          <button
            onClick={goToday}
            className="px-2.5 py-1 border border-gray-300 rounded text-xs font-medium hover:bg-gray-50"
          >
            Today
          </button>
          <button
            onClick={goNext}
            className="px-2 py-1 border border-gray-300 rounded text-xs hover:bg-gray-50"
            aria-label={`Next ${mode}`}
          >
            →
          </button>
          <h2 className="text-base font-semibold text-teal-700 ml-2">{rangeLabel}</h2>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          {/* Status color legend */}
          <div className="flex items-center gap-1.5 text-[11px] text-gray-600">
            {(['Not Started', 'In Progress', 'Blocked', 'Done'] as const).map((s) => (
              <span key={s} className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${STATUS_CHIP_COLOR[s]}`}>{s}</span>
            ))}
          </div>
          {/* Month / Week segmented toggle */}
          <div className="inline-flex bg-gray-100 rounded-md p-0.5">
            {(['month', 'week'] as const).map((m) => (
              <button
                key={m}
                onClick={() => setMode(m)}
                className={`px-2.5 py-1 text-xs font-medium rounded capitalize ${
                  mode === m ? 'bg-white shadow text-teal-700' : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                {m}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* By-assignee tally — just the count breakdown, no filter behavior */}
      {assigneeTally.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card p-3">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-2">
            Task count by assignee · this {mode}
          </div>
          <div className="flex flex-wrap gap-1.5">
            {assigneeTally.map(([assignee, count]) => (
              <span
                key={assignee}
                className="text-[11px] px-2 py-0.5 rounded font-medium inline-flex items-center gap-1 bg-gray-100 text-gray-800"
              >
                <span>{assignee === UNASSIGNED_KEY ? 'Unassigned' : assignee}</span>
                <span className="font-mono-data text-teal-700">{count}</span>
              </span>
            ))}
          </div>
        </div>
      )}

      {/* Day grid */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        {mode === 'month' && (
          <div className="grid grid-cols-7 bg-gray-50 border-b border-gray-200 text-[11px] font-semibold text-gray-600 uppercase tracking-wider">
            {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((d) => (
              <div key={d} className="px-2 py-2 text-center">{d}</div>
            ))}
          </div>
        )}
        <div className="grid grid-cols-7">
          {days.map(renderDayCell)}
        </div>
      </div>
    </div>
  );
}

/**
 * Inline due-date editor for the list view. Uses a native date input.
 * Commits on change (no need for blur — the picker closes after a date is chosen).
 */
function InlineDueDateCell({
  value,
  saving,
  overdue,
  onCommit,
}: {
  value: string | undefined;
  saving: boolean;
  overdue: boolean;
  onCommit: (newValue: string) => void;
}) {
  // Native date input wants YYYY-MM-DD; extract from the stored ISO/date-only string
  const initial = value ? String(value).slice(0, 10) : '';
  const [draft, setDraft] = useState(initial);
  useEffect(() => {
    setDraft(initial);
  }, [initial]);

  return (
    <div className="relative min-w-[120px]">
      <input
        type="date"
        value={draft}
        disabled={saving}
        onChange={(e) => {
          const next = e.target.value;
          setDraft(next);
          // Native picker fires on each selection; commit immediately on a real change
          if (next !== initial) onCommit(next);
        }}
        className={`w-full px-1.5 py-1 border border-gray-200 rounded font-mono-data text-xs bg-white hover:border-gray-300 focus:outline-none focus:border-teal-500 disabled:bg-gray-50 ${
          overdue ? 'text-error font-semibold' : 'text-gray-700'
        }`}
      />
      {saving && (
        <div className="absolute right-1.5 top-1/2 -translate-y-1/2 pointer-events-none">
          <div className="w-2.5 h-2.5 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
        </div>
      )}
    </div>
  );
}
