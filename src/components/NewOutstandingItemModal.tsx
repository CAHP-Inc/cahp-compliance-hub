import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type ItemStatus,
  type ItemPriority,
  type ItemCategory,
} from '../lib/sharepoint';

const STATUSES: ItemStatus[] = ['Not Started', 'In Progress', 'Blocked', 'Done'];
const PRIORITIES: ItemPriority[] = ['Critical', 'High', 'Medium', 'Low'];
const CATEGORIES: ItemCategory[] = [
  'Operating Agreement',
  'LURA',
  'AMI Certification',
  'Org Chart',
  'Income Documentation',
  'Signed Submittal',
  'Determination Letter',
  'Other',
];

export interface NewOutstandingItemModalProps {
  onClose: () => void;
  onSuccess: () => void;
  /** Pre-fill + lock property selection (e.g., when launched from Property Detail) */
  defaultPropertyId?: string;
  /** When true, the property dropdown is hidden — used when scope is locked */
  hidePropertyPicker?: boolean;
}

export function NewOutstandingItemModal({
  onClose,
  onSuccess,
  defaultPropertyId,
  hidePropertyPicker,
}: NewOutstandingItemModalProps) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [title, setTitle] = useState('');
  const [propertyId, setPropertyId] = useState(defaultPropertyId ?? '');
  const [category, setCategory] = useState<ItemCategory>('Other');
  const [status, setStatus] = useState<ItemStatus>('Not Started');
  const [priority, setPriority] = useState<ItemPriority>('Medium');
  const [dueDate, setDueDate] = useState('');
  const [assignedTo, setAssignedTo] = useState('');
  const [notes, setNotes] = useState('');

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);

  const sortedProperties = useMemo(() => {
    if (!properties.data) return [];
    return [...properties.data].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [properties.data]);

  const handleSubmit = async () => {
    setValidationError(null);
    if (!title.trim()) {
      setValidationError('Title is required.');
      return;
    }

    setSaving(true);
    setError(null);
    try {
      const fields: Record<string, unknown> = {
        Title: title,
        ItemStatus: status,
        Priority: priority,
        ItemCategory: category,
        DateRequested: new Date().toISOString(),
      };
      if (propertyId) fields.PropertyLookupId = propertyId;
      if (dueDate) fields.DueDate = new Date(dueDate).toISOString();
      if (assignedTo) fields.AssignedTo = assignedTo;
      if (notes) fields.ItemNotes = notes;

      await createListItem(LIST_NAMES.Outstanding, fields);
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-lg w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">New Outstanding Item</h3>
        <p className="text-sm text-gray-600 mb-4">
          Manual task entry — for items the cascade workflows didn't catch.
          It'll show up in the Outstanding kanban and (if assigned) on the assignee's queue.
        </p>

        <div className="space-y-3">
          <Field label="Title" required>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="e.g., Get insurance certificate updated"
              className={inputClass}
              disabled={saving}
              autoFocus
            />
          </Field>

          {!hidePropertyPicker && (
            <Field label="Property">
              <select
                value={propertyId}
                onChange={(e) => setPropertyId(e.target.value)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="">— optional —</option>
                {sortedProperties.map((p) => (
                  <option key={p.id} value={p.id}>{p.fields.Title}</option>
                ))}
              </select>
            </Field>
          )}

          <div className="grid grid-cols-2 gap-3">
            <Field label="Status">
              <select
                value={status}
                onChange={(e) => setStatus(e.target.value as ItemStatus)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {STATUSES.map((s) => (<option key={s} value={s}>{s}</option>))}
              </select>
            </Field>

            <Field label="Priority">
              <select
                value={priority}
                onChange={(e) => setPriority(e.target.value as ItemPriority)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {PRIORITIES.map((p) => (<option key={p} value={p}>{p}</option>))}
              </select>
            </Field>

            <Field label="Category">
              <select
                value={category}
                onChange={(e) => setCategory(e.target.value as ItemCategory)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {CATEGORIES.map((c) => (<option key={c} value={c}>{c}</option>))}
              </select>
            </Field>

            <Field label="Due Date">
              <input
                type="date"
                value={dueDate}
                onChange={(e) => setDueDate(e.target.value)}
                className={`${inputClass} font-mono-data`}
                disabled={saving}
              />
            </Field>
          </div>

          <Field label="Assigned To">
            <input
              type="text"
              value={assignedTo}
              onChange={(e) => setAssignedTo(e.target.value)}
              placeholder="Name or email"
              className={inputClass}
              disabled={saving}
            />
          </Field>

          <Field label="Notes">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              rows={3}
              placeholder="Context, links, what 'done' looks like"
              className={`${inputClass} resize-y`}
              disabled={saving}
            />
          </Field>
        </div>

        {validationError && (
          <div className="mt-3 bg-yellow-50 border border-yellow-200 rounded-md p-2 text-xs text-yellow-800">
            {validationError}
          </div>
        )}
        {error && (
          <div className="mt-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
            {error}
          </div>
        )}

        <div className="flex items-center justify-end gap-2 mt-5">
          <button
            onClick={onClose}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleSubmit}
            disabled={saving}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Creating…' : 'Create Item'}
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}{required && <span className="text-error ml-0.5">*</span>}
      </label>
      {children}
    </div>
  );
}
