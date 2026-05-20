import { useMemo, useState } from 'react';
import type { OutstandingItem, Property } from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { formatDateOnly } from '../lib/dates';

/**
 * Export Outstanding Items modal.
 *
 * Renders a formatted summary the user can paste into an email/Slack to ping
 * an assignee about their pending work. The export is scoped to one property
 * (when opened from PropertyDetail) or across the portfolio (when opened from
 * the global Outstanding Items page).
 *
 * Output formats:
 *   - "Copy as text"  — markdown-ish; readable in plain emails and chats
 *   - "Download CSV"  — for record-keeping or pasting into a spreadsheet
 */
export interface ExportOutstandingItemsModalProps {
  /** All open items in scope. Caller pre-filters by property/status if desired. */
  items: OutstandingItem[];
  /** Property lookup for resolving titles when items span multiple properties. */
  propertiesById?: Map<string, Property>;
  /** When showing items for a single property, the title appears in the heading. */
  propertyTitle?: string;
  onClose: () => void;
}

const UNASSIGNED = '(unassigned)';

export function ExportOutstandingItemsModal({
  items,
  propertiesById,
  propertyTitle,
  onClose,
}: ExportOutstandingItemsModalProps) {
  // Distinct assignees in the input set; sentinel "All" plus "(unassigned)" if any items lack one
  const assignees = useMemo(() => {
    const set = new Set<string>();
    for (const item of items) {
      const a = (item.fields.AssignedTo ?? '').trim();
      set.add(a || UNASSIGNED);
    }
    return ['All', ...Array.from(set).sort((a, b) => a.localeCompare(b))];
  }, [items]);

  const [assigneeFilter, setAssigneeFilter] = useState<string>('All');
  const [copied, setCopied] = useState(false);

  // Items to include after assignee filtering
  const filteredItems = useMemo(() => {
    if (assigneeFilter === 'All') return items;
    if (assigneeFilter === UNASSIGNED) {
      return items.filter((i) => !(i.fields.AssignedTo ?? '').trim());
    }
    return items.filter(
      (i) => (i.fields.AssignedTo ?? '').trim().toLowerCase() === assigneeFilter.toLowerCase(),
    );
  }, [items, assigneeFilter]);

  // Group by assignee → property for the rendered text
  const groups = useMemo(() => {
    type ItemRow = {
      title: string;
      due?: string;
      isOverdue: boolean;
      priority?: string;
      status?: string;
      category?: string;
    };
    type PropertyGroup = { propertyTitle: string; items: ItemRow[] };
    type AssigneeGroup = { assignee: string; properties: Map<string, PropertyGroup> };

    const out = new Map<string, AssigneeGroup>();
    const now = Date.now();
    for (const item of filteredItems) {
      const assignee = (item.fields.AssignedTo ?? '').trim() || UNASSIGNED;
      const pId = item.fields.PropertyLookupId ? String(item.fields.PropertyLookupId) : '__none__';
      const pTitle =
        propertyTitle ??
        (propertiesById?.get(pId)?.fields.Title) ??
        (pId === '__none__' ? '(no property linked)' : `Property #${pId}`);

      if (!out.has(assignee)) out.set(assignee, { assignee, properties: new Map() });
      const aGroup = out.get(assignee)!;
      if (!aGroup.properties.has(pId)) {
        aGroup.properties.set(pId, { propertyTitle: pTitle, items: [] });
      }
      const due = item.fields.DueDate ? new Date(item.fields.DueDate).getTime() : null;
      aGroup.properties.get(pId)!.items.push({
        title: item.fields.Title ?? '(untitled)',
        due: item.fields.DueDate,
        isOverdue: due != null && due < now,
        priority: item.fields.Priority,
        status: item.fields.ItemStatus,
        category: item.fields.ItemCategory,
      });
    }

    // Sort: assignees alphabetically (unassigned last), properties alphabetically,
    // items overdue-first then by due date.
    const sorted = Array.from(out.values()).sort((a, b) => {
      if (a.assignee === UNASSIGNED) return 1;
      if (b.assignee === UNASSIGNED) return -1;
      return a.assignee.localeCompare(b.assignee);
    });
    for (const g of sorted) {
      for (const pg of g.properties.values()) {
        pg.items.sort((a, b) => {
          if (a.isOverdue !== b.isOverdue) return a.isOverdue ? -1 : 1;
          const aD = a.due ? new Date(a.due).getTime() : Infinity;
          const bD = b.due ? new Date(b.due).getTime() : Infinity;
          return aD - bD;
        });
      }
    }
    return sorted;
  }, [filteredItems, propertiesById, propertyTitle]);

  // Render groups as plain text suitable for email/Slack
  const textOutput = useMemo(() => {
    if (filteredItems.length === 0) {
      return '(No open items match the current filter.)';
    }
    const lines: string[] = [];
    const today = new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
    lines.push(`CAHP Outstanding Items — as of ${today}`);
    if (propertyTitle) lines.push(`Property: ${propertyTitle}`);
    lines.push('');

    for (const g of groups) {
      lines.push(`▸ ${g.assignee}`);
      for (const pg of g.properties.values()) {
        // Only print property header when scope spans multiple properties
        if (!propertyTitle && g.properties.size > 1) {
          lines.push(`   ${pg.propertyTitle}`);
        } else if (!propertyTitle) {
          lines.push(`   ${pg.propertyTitle}`);
        }
        for (const i of pg.items) {
          const dueStr = i.due ? formatDateOnly(i.due) : 'no due date';
          const overdueTag = i.isOverdue ? ' [OVERDUE]' : '';
          const priTag = i.priority && i.priority !== 'Medium' ? ` (${i.priority})` : '';
          lines.push(`     • ${i.title} — due ${dueStr}${overdueTag}${priTag}`);
        }
      }
      lines.push('');
    }
    lines.push('— Sent from CAHP Compliance Hub');
    return lines.join('\n').trim();
  }, [groups, filteredItems.length, propertyTitle]);

  const csvOutput = useMemo(() => {
    const rows: string[][] = [
      ['Assignee', 'Property', 'Item', 'Category', 'Status', 'Priority', 'Due Date', 'Overdue'],
    ];
    const now = Date.now();
    for (const item of filteredItems) {
      const due = item.fields.DueDate ? new Date(item.fields.DueDate).getTime() : null;
      const isOverdue = due != null && due < now;
      const pId = item.fields.PropertyLookupId ? String(item.fields.PropertyLookupId) : '';
      const pTitle =
        propertyTitle ?? propertiesById?.get(pId)?.fields.Title ?? '';
      rows.push([
        (item.fields.AssignedTo ?? '').trim() || UNASSIGNED,
        pTitle,
        item.fields.Title ?? '',
        item.fields.ItemCategory ?? '',
        item.fields.ItemStatus ?? '',
        item.fields.Priority ?? '',
        item.fields.DueDate ? formatDateOnly(item.fields.DueDate) : '',
        isOverdue ? 'Yes' : 'No',
      ]);
    }
    return rows
      .map((r) => r.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(','))
      .join('\n');
  }, [filteredItems, propertiesById, propertyTitle]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(textOutput);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // If clipboard API is blocked (HTTP context, permissions), fall back to selecting the textarea
      const ta = document.getElementById('export-outstanding-textarea') as HTMLTextAreaElement | null;
      if (ta) {
        ta.select();
        document.execCommand('copy');
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      }
    }
  };

  const handleDownloadCsv = () => {
    const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const stamp = new Date().toISOString().slice(0, 10);
    const stem = propertyTitle ? propertyTitle.replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_') : 'Portfolio';
    a.href = url;
    a.download = `${stem}_OutstandingItems_${stamp}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] flex flex-col">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">Export Outstanding Items</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            {propertyTitle
              ? <>Generate a reminder for <strong>{propertyTitle}</strong>. Pick an assignee or export everything.</>
              : 'Generate a reminder across the portfolio. Pick an assignee to narrow it down.'}
          </p>
        </div>

        <div className="px-6 py-3 border-b border-gray-200 flex items-end gap-3 flex-wrap">
          <div>
            <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
              Assignee
            </label>
            <select
              value={assigneeFilter}
              onChange={(e) => setAssigneeFilter(e.target.value)}
              className="px-2 py-1 border border-gray-300 rounded text-sm bg-white focus:outline-none focus:border-teal-500"
            >
              {assignees.map((a) => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>
          <div className="text-xs text-gray-500 self-center">
            {filteredItems.length} item{filteredItems.length === 1 ? '' : 's'} in scope
          </div>
        </div>

        <div className="px-6 py-3 flex-1 overflow-y-auto">
          <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
            Preview
          </label>
          <textarea
            id="export-outstanding-textarea"
            readOnly
            value={textOutput}
            className="w-full h-72 px-3 py-2 border border-gray-300 rounded text-xs font-mono-data resize-none focus:outline-none focus:border-teal-500"
          />
          <p className="text-[11px] text-gray-500 mt-1">
            Copy and paste into an email, Slack, or Teams. Items already marked Done are excluded by the caller.
          </p>
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded-md"
          >
            Close
          </button>
          <button
            onClick={handleDownloadCsv}
            disabled={filteredItems.length === 0}
            className="px-3 py-1.5 text-sm border border-gray-300 bg-white hover:bg-gray-100 rounded-md disabled:opacity-50 inline-flex items-center gap-1.5"
          >
            <Icon name="file" size={12} />
            Download CSV
          </button>
          <button
            onClick={handleCopy}
            disabled={filteredItems.length === 0}
            className={`px-3 py-1.5 text-sm rounded-md font-medium inline-flex items-center gap-1.5 ${
              copied
                ? 'bg-success text-white'
                : 'bg-teal-700 hover:bg-teal-900 text-white disabled:bg-gray-300 disabled:cursor-not-allowed'
            }`}
          >
            <Icon name={copied ? 'check' : 'edit'} size={12} />
            {copied ? 'Copied' : 'Copy as text'}
          </button>
        </div>
      </div>
    </div>
  );
}
