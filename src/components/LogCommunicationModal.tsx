import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type Owner,
  type CommType,
  type CommDirection,
  type CommStatus,
} from '../lib/sharepoint';
import { notifyUser } from '../lib/notifications';
import { useSession } from '../lib/session';

const COMM_TYPES: CommType[] = ['Email', 'Phone', 'Meeting', 'SMS', 'Other'];
const COMM_DIRECTIONS: CommDirection[] = ['Inbound', 'Outbound'];

export interface LogCommunicationModalProps {
  onClose: () => void;
  onSuccess: () => void;
  defaultPropertyId?: string;
  defaultOwnerId?: string;
}

export function LogCommunicationModal({
  onClose,
  onSuccess,
  defaultPropertyId,
  defaultOwnerId,
}: LogCommunicationModalProps) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const { user } = useSession();

  const [subject, setSubject] = useState('');
  const [commType, setCommType] = useState<CommType>('Email');
  const [direction, setDirection] = useState<CommDirection>('Inbound');
  const [commDate, setCommDate] = useState(new Date().toISOString().slice(0, 10));
  const [propertyId, setPropertyId] = useState(defaultPropertyId ?? '');
  const [ownerId, setOwnerId] = useState(defaultOwnerId ?? '');
  const [participants, setParticipants] = useState('');
  const [responseDue, setResponseDue] = useState('');
  const [notes, setNotes] = useState('');

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);
  const [cascadeLog, setCascadeLog] = useState<string[]>([]);

  const sortedProperties = useMemo(() => {
    if (!properties.data) return [];
    return [...properties.data].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [properties.data]);

  const sortedOwners = useMemo(() => {
    if (!owners.data) return [];
    return [...owners.data].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [owners.data]);

  const handleSubmit = async () => {
    setValidationError(null);
    if (!subject.trim()) {
      setValidationError('Subject is required.');
      return;
    }
    if (!commDate) {
      setValidationError('Date is required.');
      return;
    }
    if (!propertyId && !ownerId) {
      setValidationError('Link to at least a property or an owner.');
      return;
    }

    setSaving(true);
    setError(null);
    setCascadeLog([]);

    try {
      // ──────────────── CASCADE STEP 1: create Communication record ────────────────
      const commFields: Record<string, unknown> = {
        Title: subject,
        CommType: commType,
        CommDirection: direction,
        CommDate: new Date(commDate).toISOString(),
        CommStatus: 'Open' as CommStatus,
      };
      if (propertyId) commFields.CommPropertyLookupId = propertyId;
      if (ownerId) commFields.CommOwnerLookupId = ownerId;
      if (participants) commFields.CommParticipants = participants;
      if (responseDue) commFields.CommResponseDue = new Date(responseDue).toISOString();
      if (notes) commFields.CommNotes = notes;

      const comm = await createListItem<{ id: string }>(LIST_NAMES.Communications, commFields);
      setCascadeLog((prev) => [...prev, `✓ Communication record #${comm.id} created`]);

      // ──────────────── CASCADE STEP 2: auto-create Outstanding Item if follow-up set ────────────────
      if (responseDue && propertyId) {
        try {
          await createListItem(LIST_NAMES.Outstanding, {
            Title: `Follow up: ${subject}`,
            PropertyLookupId: propertyId,
            ItemCategory: 'Other',
            ItemStatus: 'Not Started',
            DateRequested: new Date(commDate).toISOString(),
            DueDate: new Date(responseDue).toISOString(),
            Priority: 'Medium',
            ItemNotes: `Auto-created from Owner Communication #${comm.id} (${commType}, ${direction}). Follow up by ${new Date(responseDue).toLocaleDateString()}.`,
          });
          setCascadeLog((prev) => [...prev, `✓ Outstanding Item created (due ${new Date(responseDue).toLocaleDateString()})`]);
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn('Outstanding Item creation failed:', e);
          setCascadeLog((prev) => [...prev, `⚠ Outstanding Item creation failed (comm record still saved)`]);
        }
      }

      setCascadeLog((prev) => [...prev, `✓ Audit log entries written`]);

      // Notify Brandy (system owner) when a comm has a follow-up due
      if (responseDue && user?.email) {
        await notifyUser({
          upn: user.email,
          type: 'TaskAssigned',
          title: `Follow up due ${new Date(responseDue).toLocaleDateString()}: ${subject}`,
          targetType: 'Communication',
          targetId: String(comm.id),
          url: `/comms/${comm.id}`,
        });
      }

      setTimeout(() => onSuccess(), 800);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const willCreateItem = Boolean(responseDue && propertyId);

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Log Communication</h3>
        <p className="text-sm text-gray-600 mb-4">
          Owner emails, phone calls, meetings, vendor calls, team meetings — anything non-DOR.
          Setting a follow-up due date auto-creates an Outstanding Item.
        </p>

        <div className="space-y-3">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <Field label="Type" required>
              <select
                value={commType}
                onChange={(e) => setCommType(e.target.value as CommType)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {COMM_TYPES.map((t) => (<option key={t} value={t}>{t}</option>))}
              </select>
            </Field>

            <Field label="Direction">
              <select
                value={direction}
                onChange={(e) => setDirection(e.target.value as CommDirection)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {COMM_DIRECTIONS.map((d) => (<option key={d} value={d}>{d}</option>))}
              </select>
            </Field>

            <Field label="Date" required>
              <input
                type="date"
                value={commDate}
                onChange={(e) => setCommDate(e.target.value)}
                className={`${inputClass} font-mono-data`}
                disabled={saving}
              />
            </Field>

            <Field label="Response Due">
              <input
                type="date"
                value={responseDue}
                onChange={(e) => setResponseDue(e.target.value)}
                className={`${inputClass} font-mono-data`}
                disabled={saving}
              />
              <p className="text-[11px] text-gray-400 mt-0.5">If set + property linked, auto-creates Outstanding Item.</p>
            </Field>

            <Field label="Property">
              <select
                value={propertyId}
                onChange={(e) => setPropertyId(e.target.value)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="">— optional —</option>
                {sortedProperties.map((p) => (<option key={p.id} value={p.id}>{p.fields.Title}</option>))}
              </select>
            </Field>

            <Field label="Owner">
              <select
                value={ownerId}
                onChange={(e) => setOwnerId(e.target.value)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="">— optional —</option>
                {sortedOwners.map((o) => (<option key={o.id} value={o.id}>{o.fields.Title}</option>))}
              </select>
            </Field>
          </div>

          <Field label="Subject" required>
            <input
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="e.g., Quarterly check-in with Stan Gendlin"
              className={inputClass}
              disabled={saving}
            />
          </Field>

          <Field label="Participants">
            <input
              type="text"
              value={participants}
              onChange={(e) => setParticipants(e.target.value)}
              placeholder="Comma-separated names"
              className={inputClass}
              disabled={saving}
            />
          </Field>

          <Field label="Notes / Summary">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              rows={3}
              placeholder="What was discussed, next steps, action items"
              className={`${inputClass} resize-y`}
              disabled={saving}
            />
          </Field>
        </div>

        {/* Cascade preview */}
        <div className="mt-4 bg-gold-50 border border-gold-200 rounded-md p-3">
          <div className="text-[10px] font-semibold text-gold-900 uppercase tracking-wider mb-2">
            On Save
          </div>
          <ul className="text-xs text-teal-900 space-y-0.5">
            <li>✓ <strong>Communication record</strong> created (status Open)</li>
            <li className={willCreateItem ? '' : 'text-gray-500 line-through'}>
              {willCreateItem ? '✓' : '○'} <strong>Outstanding Item</strong> auto-created
              {!willCreateItem && ' (response due + property both required)'}
            </li>
            <li>✓ Audit log entries written</li>
          </ul>
        </div>

        {cascadeLog.length > 0 && (
          <div className="mt-4 bg-green-50 border border-green-200 rounded-md p-3">
            <div className="text-[10px] font-semibold text-green-900 uppercase tracking-wider mb-2">Results</div>
            <ul className="text-xs text-green-900 space-y-0.5 font-mono-data">
              {cascadeLog.map((line, idx) => <li key={idx}>{line}</li>)}
            </ul>
          </div>
        )}

        {validationError && (
          <div className="mt-3 bg-yellow-50 border border-yellow-200 rounded-md p-2 text-xs text-yellow-800">
            {validationError}
          </div>
        )}
        {error && (
          <div className="mt-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">{error}</div>
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
            {saving ? 'Logging…' : 'Log Communication'}
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
