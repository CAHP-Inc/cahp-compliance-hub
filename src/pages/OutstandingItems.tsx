import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
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
import { NewOutstandingItemModal } from '../components/NewOutstandingItemModal';

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

// Map ItemStatus → kanban column. Legacy "Requested" / "Received" map to new columns
const KANBAN_COLUMNS: { id: ItemStatus; label: string; legacyAliases: ItemStatus[] }[] = [
  { id: 'Not Started', label: 'Not Started', legacyAliases: ['Requested'] },
  { id: 'In Progress', label: 'In Progress', legacyAliases: [] },
  { id: 'Blocked', label: 'Blocked', legacyAliases: [] },
  { id: 'Done', label: 'Done', legacyAliases: ['Received', 'Not Applicable'] },
];

function getKanbanColumn(status: ItemStatus | undefined): ItemStatus | null {
  if (!status) return 'Not Started';
  for (const col of KANBAN_COLUMNS) {
    if (col.id === status || col.legacyAliases.includes(status)) return col.id;
  }
  return null; // Overdue or unmapped — render in current column with overdue treatment
}

function isOverdue(item: OutstandingItem): boolean {
  if (!item.fields.DueDate) return false;
  const due = new Date(item.fields.DueDate);
  const closed = item.fields.ItemStatus === 'Done' || item.fields.ItemStatus === 'Received' || item.fields.ItemStatus === 'Not Applicable';
  return due.getTime() < Date.now() && !closed;
}

// =============================================================================
// Page
// =============================================================================

type ViewMode = 'kanban' | 'list';

export function OutstandingItems() {
  const navigate = useNavigate();
  const items = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [view, setView] = useState<ViewMode>('kanban');
  const [search, setSearch] = useState('');
  const [priorityFilter, setPriorityFilter] = useState<ItemPriority | 'All'>('All');
  const [assigneeFilter, setAssigneeFilter] = useState<string>('All');
  const [propertyFilter, setPropertyFilter] = useState<string>('All');
  const [showClosed, setShowClosed] = useState(false);
  const [updatingId, setUpdatingId] = useState<string | null>(null);
  const [newItemOpen, setNewItemOpen] = useState(false);

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
      return true;
    });
  }, [items.data, search, priorityFilter, assigneeFilter, propertyFilter, showClosed, propertiesById]);

  // Group by kanban column
  const byColumn = useMemo(() => {
    const map: Record<string, OutstandingItem[]> = {
      'Not Started': [],
      'In Progress': [],
      'Blocked': [],
      'Done': [],
    };
    filtered.forEach((item) => {
      const col = getKanbanColumn(item.fields.ItemStatus);
      if (col && map[col]) {
        map[col].push(item);
      } else {
        // Fallback: drop unmapped into Not Started
        map['Not Started'].push(item);
      }
    });
    // Sort each column: overdue first, then by priority, then by due date
    const priorityOrder: Record<string, number> = { Critical: 0, High: 1, Medium: 2, Low: 3 };
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => {
        const aOverdue = isOverdue(a) ? 0 : 1;
        const bOverdue = isOverdue(b) ? 0 : 1;
        if (aOverdue !== bOverdue) return aOverdue - bOverdue;
        const aP = priorityOrder[a.fields.Priority ?? 'Medium'] ?? 2;
        const bP = priorityOrder[b.fields.Priority ?? 'Medium'] ?? 2;
        if (aP !== bP) return aP - bP;
        const aDue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Infinity;
        const bDue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Infinity;
        return aDue - bDue;
      });
    });
    return map;
  }, [filtered]);

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

  // Quick status update from dropdown on card
  const handleStatusChange = async (itemId: string, newStatus: ItemStatus) => {
    setUpdatingId(itemId);
    try {
      const patch: Record<string, unknown> = { ItemStatus: newStatus };
      if (newStatus === 'Done' || newStatus === 'Received') {
        patch.DateReceivedItem = new Date().toISOString();
      }
      await updateListItem(LIST_NAMES.Outstanding, itemId, patch);
      await items.refetch?.();
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error('Status update failed:', e);
      alert(`Status update failed: ${e instanceof Error ? e.message : String(e)}`);
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
            onClick={() => setNewItemOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium flex items-center gap-1.5"
          >
            <Icon name="plus" size={14} />
            New Item
          </button>
          <button
            onClick={() => setView(view === 'kanban' ? 'list' : 'kanban')}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium flex items-center gap-1.5"
          >
            <Icon name={view === 'kanban' ? 'inbox' : 'check'} size={14} />
            {view === 'kanban' ? 'List View' : 'Kanban View'}
          </button>
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
            or log a DOR letter with a response deadline. They'll appear here in the kanban.
          </p>
        </div>
      ) : view === 'kanban' ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
          {KANBAN_COLUMNS.map((col) => (
            <KanbanColumn
              key={col.id}
              column={col}
              items={byColumn[col.id] ?? []}
              propertiesById={propertiesById}
              onCardClick={(itemId) => navigate(`/outstanding-items/${itemId}`)}
              onStatusChange={handleStatusChange}
              updatingId={updatingId}
            />
          ))}
        </div>
      ) : (
        <ListView
          items={filtered}
          propertiesById={propertiesById}
          onRowClick={(itemId) => navigate(`/outstanding-items/${itemId}`)}
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
    </div>
  );
}

// =============================================================================
// Kanban column + card
// =============================================================================

function KanbanColumn({
  column,
  items,
  propertiesById,
  onCardClick,
  onStatusChange,
  updatingId,
}: {
  column: { id: ItemStatus; label: string };
  items: OutstandingItem[];
  propertiesById: Map<string, Property>;
  onCardClick: (id: string) => void;
  onStatusChange: (id: string, status: ItemStatus) => void;
  updatingId: string | null;
}) {
  return (
    <div className="bg-gray-50 border border-gray-200 rounded-lg overflow-hidden">
      <div className="px-3 py-2 border-b border-gray-200 bg-white flex items-center justify-between">
        <h3 className="text-xs font-semibold text-gray-700 uppercase tracking-wider">{column.label}</h3>
        <span className="text-[10px] font-semibold text-gray-500 bg-gray-100 px-1.5 py-0.5 rounded-full font-mono-data">
          {items.length}
        </span>
      </div>
      <div className="p-2 space-y-2 min-h-[200px] max-h-[700px] overflow-y-auto">
        {items.length === 0 ? (
          <p className="text-xs text-gray-400 italic text-center py-6">No items</p>
        ) : (
          items.map((item) => (
            <KanbanCard
              key={item.id}
              item={item}
              property={item.fields.PropertyLookupId ? propertiesById.get(String(item.fields.PropertyLookupId)) : undefined}
              onClick={() => onCardClick(item.id)}
              onStatusChange={(status) => onStatusChange(item.id, status)}
              updating={updatingId === item.id}
            />
          ))
        )}
      </div>
    </div>
  );
}

function KanbanCard({
  item,
  property,
  onClick,
  onStatusChange,
  updating,
}: {
  item: OutstandingItem;
  property?: Property;
  onClick: () => void;
  onStatusChange: (status: ItemStatus) => void;
  updating: boolean;
}) {
  const overdue = isOverdue(item);
  const f = item.fields;

  return (
    <div
      onClick={onClick}
      className={`bg-white border rounded-md p-3 shadow-sm hover:shadow-card-hover cursor-pointer transition-shadow ${
        overdue ? 'border-red-300 ring-1 ring-red-200' : 'border-gray-200'
      } ${updating ? 'opacity-60' : ''}`}
    >
      <div className="flex items-start justify-between gap-2 mb-1">
        <h4 className="text-sm font-medium text-gray-900 line-clamp-2">{f.Title}</h4>
        {f.Priority && (
          <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded border flex-shrink-0 ${PRIORITY_STYLES[f.Priority]}`}>
            {f.Priority}
          </span>
        )}
      </div>
      {property && (
        <p className="text-xs text-gray-600 truncate mb-1.5">{property.fields.Title}</p>
      )}
      <div className="flex items-center justify-between gap-2 text-[11px]">
        <div className="text-gray-500 font-mono-data flex items-center gap-1.5 min-w-0">
          {f.DueDate ? (
            <span className={overdue ? 'text-error font-semibold' : ''}>
              {overdue && '⚠ '}
              {new Date(f.DueDate).toLocaleDateString()}
            </span>
          ) : (
            <span className="text-gray-400">no due date</span>
          )}
          {f.AssignedTo && <span className="text-gray-500 truncate">· {f.AssignedTo}</span>}
        </div>
      </div>
      {/* Status quick-move dropdown */}
      <select
        value={f.ItemStatus ?? 'Not Started'}
        onClick={(e) => e.stopPropagation()}
        onChange={(e) => onStatusChange(e.target.value as ItemStatus)}
        disabled={updating}
        className="mt-2 w-full text-[10px] uppercase tracking-wider font-semibold border border-gray-200 rounded px-1 py-0.5 bg-white text-gray-600 cursor-pointer hover:border-teal-500 focus:outline-none focus:border-teal-500"
      >
        <option value="Not Started">Move → Not Started</option>
        <option value="In Progress">Move → In Progress</option>
        <option value="Blocked">Move → Blocked</option>
        <option value="Done">Move → Done</option>
      </select>
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
}: {
  items: OutstandingItem[];
  propertiesById: Map<string, Property>;
  onRowClick: (id: string) => void;
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
            <th className="px-4 py-3 text-left">Assigned To</th>
            <th className="px-4 py-3 text-left">Due Date</th>
            <th className="px-4 py-3 text-left">Category</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {items.map((item) => {
            const property = item.fields.PropertyLookupId
              ? propertiesById.get(String(item.fields.PropertyLookupId))
              : null;
            const overdue = isOverdue(item);
            return (
              <tr
                key={item.id}
                onClick={() => onRowClick(item.id)}
                className={`hover:bg-gray-50 transition-colors cursor-pointer ${overdue ? 'bg-red-50' : ''}`}
              >
                <td className="px-4 py-3 font-medium text-gray-900">
                  {overdue && <span className="text-error mr-1">⚠</span>}
                  {item.fields.Title}
                </td>
                <td className="px-4 py-3 text-xs text-gray-700">{property?.fields.Title ?? '—'}</td>
                <td className="px-4 py-3">
                  {item.fields.ItemStatus ? (
                    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[item.fields.ItemStatus]}`}>
                      {item.fields.ItemStatus}
                    </span>
                  ) : '—'}
                </td>
                <td className="px-4 py-3">
                  {item.fields.Priority ? (
                    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold border ${PRIORITY_STYLES[item.fields.Priority]}`}>
                      {item.fields.Priority}
                    </span>
                  ) : '—'}
                </td>
                <td className="px-4 py-3 text-xs text-gray-700">{item.fields.AssignedTo || '—'}</td>
                <td className={`px-4 py-3 font-mono-data text-xs ${overdue ? 'text-error font-semibold' : 'text-gray-700'}`}>
                  {item.fields.DueDate ? new Date(item.fields.DueDate).toLocaleDateString() : '—'}
                </td>
                <td className="px-4 py-3 text-xs text-gray-600">{item.fields.ItemCategory || '—'}</td>
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
