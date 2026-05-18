import { useState } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type DeadlineType,
  type DeadlineStatus,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { AssigneePicker } from './AssigneePicker';

const DEADLINE_TYPES: DeadlineType[] = [
  'IRS 990 Filing',
  'Annual Recertification',
  'Rent Roll Review',
  'AMI Cert Renewal',
  'State Compliance Report',
  'Property Tax Filing',
  'Operating Agreement Review',
  'Other',
];

export interface NewDeadlineModalProps {
  onClose: () => void;
  onSuccess: () => void;
  defaultPropertyId?: string;
}

export function NewDeadlineModal({ onClose, onSuccess, defaultPropertyId }: NewDeadlineModalProps) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [title, setTitle] = useState('');
  const [propertyId, setPropertyId] = useState<string>(defaultPropertyId ?? '');
  const [deadlineType, setDeadlineType] = useState<DeadlineType | ''>('');
  const [dueDate, setDueDate] = useState<string>('');
  const [assignedTo, setAssignedTo] = useState<string>('');
  const [notes, setNotes] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sortedProperties = (properties.data ?? [])
    .slice()
    .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));

  const handleSubmit = async () => {
    if (!title.trim()) {
      setError('Title is required');
      return;
    }
    if (!dueDate) {
      setError('Due date is required');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const property = propertyId
        ? sortedProperties.find((p) => String(p.id) === propertyId)
        : null;
      const fields: Record<string, unknown> = {
        Title: title.trim(),
        DueDate: new Date(dueDate).toISOString(),
        DeadlineStatus: 'Upcoming' as DeadlineStatus,
      };
      if (propertyId) fields.PropertyLookupId = propertyId;
      if (deadlineType) fields.DeadlineType = deadlineType;
      if (assignedTo) fields.AssignedTo = assignedTo;
      if (notes.trim()) fields.DeadlineNotes = notes.trim();
      if (property?.fields.cahpState) fields.cahpState = property.fields.cahpState;
      await createListItem(LIST_NAMES.ComplianceDeadlines, fields);
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-lg w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">New Compliance Deadline</h3>
        <p className="text-sm text-gray-600 mb-4">
          Manually add a compliance deadline. Most deadlines are auto-created (AMI Cert Renewal at property creation,
          DOR Response when a letter is logged). Use this for one-off events.
        </p>

        <div className="space-y-3 mb-4">
          <Field label="Title" required>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="e.g. 2026 IRS 990 Filing"
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
            />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Due Date" required>
              <input
                type="date"
                value={dueDate}
                onChange={(e) => setDueDate(e.target.value)}
                className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
              />
            </Field>
            <Field label="Deadline Type">
              <select
                value={deadlineType}
                onChange={(e) => setDeadlineType(e.target.value as DeadlineType | '')}
                className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              >
                <option value="">— select —</option>
                {DEADLINE_TYPES.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </Field>
          </div>
          <Field label="Property">
            <select
              value={propertyId}
              onChange={(e) => setPropertyId(e.target.value)}
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              disabled={Boolean(defaultPropertyId)}
            >
              <option value="">— portfolio-wide (no property) —</option>
              {sortedProperties.map((p) => (
                <option key={p.id} value={String(p.id)}>{p.fields.Title}</option>
              ))}
            </select>
          </Field>
          <Field label="Assigned To">
            <AssigneePicker
              value={assignedTo}
              onChange={setAssignedTo}
              disabled={saving}
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
            />
          </Field>
          <Field label="Notes">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              rows={3}
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
              placeholder="Any context or instructions"
            />
          </Field>
        </div>

        {error && (
          <div className="mb-3 bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">{error}</div>
        )}

        <div className="flex justify-end gap-2">
          <button
            onClick={onClose}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleSubmit}
            disabled={saving}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            <Icon name="plus" size={12} />
            Create Deadline
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({
  label,
  required,
  children,
}: {
  label: string;
  required?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div>
      <label className="block text-[11px] font-semibold text-gray-700 uppercase tracking-wider mb-1">
        {label}
        {required && <span className="text-error ml-0.5">*</span>}
      </label>
      {children}
    </div>
  );
}
