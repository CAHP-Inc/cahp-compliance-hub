import { useState, useEffect } from 'react';
import {
  updateListItem,
  getListItems,
  LIST_NAMES,
  type Property,
  type PropertyStatus,
  type OutstandingItem,
  type ComplianceDeadline,
  type Submittal,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

type DispositionType = Extract<PropertyStatus, 'Sold' | 'Withdrawn' | 'Removed from Program'>;

const DISPOSITION_OPTIONS: { value: DispositionType; label: string; description: string }[] = [
  {
    value: 'Sold',
    label: 'Sold',
    description: 'Property has been sold to a third party. Use for completed real-estate transactions.',
  },
  {
    value: 'Withdrawn',
    label: 'Withdrawn',
    description: 'Owner voluntarily withdrew the property from the CAHP program. Property still owned, just exited the program.',
  },
  {
    value: 'Removed from Program',
    label: 'Removed from Program',
    description: 'Property removed for non-compliance, lost AMI eligibility, or other involuntary reason.',
  },
];

interface DispositionModalProps {
  property: Property;
  onClose: () => void;
  onComplete: () => void;
}

const isClosedItemStatus = (s: string | undefined) =>
  s === 'Done' || s === 'Received' || s === 'Not Applicable';

export function DispositionModal({ property, onClose, onComplete }: DispositionModalProps) {
  const [type, setType] = useState<DispositionType | ''>('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [reason, setReason] = useState('');
  const [closeOpenItems, setCloseOpenItems] = useState(true);
  const [closeDeadlines, setCloseDeadlines] = useState(true);
  const [withdrawDraftSubmittals, setWithdrawDraftSubmittals] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<string | null>(null);

  // Preview counts — fetch open items/deadlines/submittals for this property
  const [openItemsCount, setOpenItemsCount] = useState<number | null>(null);
  const [openDeadlinesCount, setOpenDeadlinesCount] = useState<number | null>(null);
  const [draftSubmittalsCount, setDraftSubmittalsCount] = useState<number | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const [items, deadlines, submittals] = await Promise.all([
          getListItems<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 }),
          getListItems<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 }),
          getListItems<Submittal>(LIST_NAMES.Submittals, { top: 500 }),
        ]);
        if (cancelled) return;
        const pid = String(property.id);
        setOpenItemsCount(
          items.filter(
            (i) => String(i.fields.PropertyLookupId) === pid && !isClosedItemStatus(i.fields.ItemStatus)
          ).length
        );
        setOpenDeadlinesCount(
          deadlines.filter(
            (d) => String(d.fields.PropertyLookupId) === pid && d.fields.DeadlineStatus !== 'Completed'
          ).length
        );
        setDraftSubmittalsCount(
          submittals.filter(
            (s) =>
              String(s.fields.PropertyLookupId) === pid &&
              (s.fields.SubmittalStatus === 'Draft' || s.fields.SubmittalStatus === 'Package Mailed (NC)')
          ).length
        );
      } catch {
        // non-blocking — preview is just informational
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [property.id]);

  const canSubmit = type !== '' && date !== '' && reason.trim() !== '';

  const handleSubmit = async () => {
    if (!canSubmit) {
      setError('Disposition type, date, and reason are all required.');
      return;
    }
    setSaving(true);
    setError(null);
    setProgress(null);
    try {
      const formattedReason = `Disposed ${date} as ${type}.\nReason: ${reason.trim()}`;
      const dispositionNote = `Auto-closed on property disposition (${type}, ${date}).`;

      // 1. Update the property itself
      setProgress('Updating property status…');
      await updateListItem(LIST_NAMES.Properties, property.id, {
        PropertyStatus: type,
        RemovedReason: formattedReason,
      });

      const pid = String(property.id);

      // 2. Close open Outstanding Items
      if (closeOpenItems) {
        setProgress('Closing open outstanding items…');
        const items = await getListItems<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
        const open = items.filter(
          (i) => String(i.fields.PropertyLookupId) === pid && !isClosedItemStatus(i.fields.ItemStatus)
        );
        for (const i of open) {
          try {
            const existingNotes = i.fields.ItemNotes ?? '';
            await updateListItem(LIST_NAMES.Outstanding, i.id, {
              ItemStatus: 'Not Applicable',
              ItemNotes: existingNotes
                ? `${existingNotes}\n\n— ${dispositionNote}`
                : dispositionNote,
            });
          } catch (e) {
            // eslint-disable-next-line no-console
            console.warn(`Failed to close outstanding item ${i.id}:`, e);
          }
        }
      }

      // 3. Mark open Compliance Deadlines as Completed with disposition note
      if (closeDeadlines) {
        setProgress('Closing open compliance deadlines…');
        const deadlines = await getListItems<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });
        const open = deadlines.filter(
          (d) => String(d.fields.PropertyLookupId) === pid && d.fields.DeadlineStatus !== 'Completed'
        );
        for (const d of open) {
          try {
            const existingNotes = d.fields.DeadlineNotes ?? '';
            await updateListItem(LIST_NAMES.ComplianceDeadlines, d.id, {
              DeadlineStatus: 'Completed',
              CompletionDate: new Date().toISOString(),
              DeadlineNotes: existingNotes
                ? `${existingNotes}\n\n— ${dispositionNote}`
                : dispositionNote,
            });
          } catch (e) {
            // eslint-disable-next-line no-console
            console.warn(`Failed to close deadline ${d.id}:`, e);
          }
        }
      }

      // 4. Withdraw any open-state Submittals (Draft / Package Mailed)
      if (withdrawDraftSubmittals) {
        setProgress('Withdrawing in-progress submittals…');
        const submittals = await getListItems<Submittal>(LIST_NAMES.Submittals, { top: 500 });
        const draftLike = submittals.filter(
          (s) =>
            String(s.fields.PropertyLookupId) === pid &&
            (s.fields.SubmittalStatus === 'Draft' || s.fields.SubmittalStatus === 'Package Mailed (NC)')
        );
        for (const s of draftLike) {
          try {
            const existingNotes = s.fields.SubmittalNotes ?? '';
            await updateListItem(LIST_NAMES.Submittals, s.id, {
              SubmittalStatus: 'Withdrawn',
              WithdrawnReason: dispositionNote,
              SubmittalNotes: existingNotes
                ? `${existingNotes}\n\n— ${dispositionNote}`
                : dispositionNote,
            });
          } catch (e) {
            // eslint-disable-next-line no-console
            console.warn(`Failed to withdraw submittal ${s.id}:`, e);
          }
        }
      }

      onComplete();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
      setProgress(null);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-200 flex items-start justify-between">
          <div>
            <h2 className="text-lg font-bold text-teal-700">Dispose Property</h2>
            <p className="text-sm text-gray-500 mt-0.5">{property.fields.Title}</p>
          </div>
          <button
            onClick={onClose}
            disabled={saving}
            className="text-gray-400 hover:text-gray-600 disabled:opacity-50"
            aria-label="Close"
          >
            <Icon name="alert" size={20} />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-5">
          {/* Disposition type */}
          <div>
            <label className="block text-sm font-semibold text-gray-700 mb-2">
              Disposition Type <span className="text-error">*</span>
            </label>
            <div className="space-y-2">
              {DISPOSITION_OPTIONS.map((opt) => (
                <label
                  key={opt.value}
                  className={`block border rounded-lg p-3 cursor-pointer transition-colors ${
                    type === opt.value
                      ? 'border-teal-500 bg-teal-50'
                      : 'border-gray-200 hover:border-gray-300'
                  }`}
                >
                  <div className="flex items-start gap-3">
                    <input
                      type="radio"
                      name="disposition-type"
                      value={opt.value}
                      checked={type === opt.value}
                      onChange={() => setType(opt.value)}
                      disabled={saving}
                      className="mt-0.5"
                    />
                    <div className="flex-1">
                      <div className="font-semibold text-sm text-gray-900">{opt.label}</div>
                      <div className="text-xs text-gray-600 mt-0.5">{opt.description}</div>
                    </div>
                  </div>
                </label>
              ))}
            </div>
          </div>

          {/* Date */}
          <div>
            <label className="block text-sm font-semibold text-gray-700 mb-2">
              Disposition Date <span className="text-error">*</span>
            </label>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              disabled={saving}
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm font-mono-data focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
            />
          </div>

          {/* Reason */}
          <div>
            <label className="block text-sm font-semibold text-gray-700 mb-2">
              Reason / Notes <span className="text-error">*</span>
            </label>
            <textarea
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              disabled={saving}
              rows={4}
              placeholder={
                type === 'Sold'
                  ? 'Buyer, sale price, closing details…'
                  : type === 'Withdrawn'
                    ? 'Why the owner withdrew, who made the decision, etc.'
                    : 'Why the property was removed — compliance failure, lost AMI status, etc.'
              }
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 resize-y"
            />
            <p className="text-xs text-gray-500 mt-1">
              Written to the property's RemovedReason field with the disposition date prepended.
            </p>
          </div>

          {/* Cascade options — close related items */}
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3">
            <h4 className="text-xs font-semibold text-blue-900 uppercase tracking-wider mb-2 flex items-center gap-1.5">
              <Icon name="check" size={12} />
              Auto-close related records
            </h4>
            <p className="text-[11px] text-blue-800 mb-3">
              Don't leave open work hanging on a disposed property. These cascade actions can be unchecked
              if you want to handle them manually.
            </p>
            <div className="space-y-2">
              <label className="flex items-start gap-2 cursor-pointer text-xs text-blue-900">
                <input
                  type="checkbox"
                  checked={closeOpenItems}
                  onChange={(e) => setCloseOpenItems(e.target.checked)}
                  disabled={saving}
                  className="mt-0.5"
                />
                <span>
                  Close <strong>{openItemsCount ?? '…'}</strong> open outstanding item{openItemsCount === 1 ? '' : 's'}
                  {' '}as <strong>Not Applicable</strong> with disposition note appended.
                </span>
              </label>
              <label className="flex items-start gap-2 cursor-pointer text-xs text-blue-900">
                <input
                  type="checkbox"
                  checked={closeDeadlines}
                  onChange={(e) => setCloseDeadlines(e.target.checked)}
                  disabled={saving}
                  className="mt-0.5"
                />
                <span>
                  Mark <strong>{openDeadlinesCount ?? '…'}</strong> open compliance deadline{openDeadlinesCount === 1 ? '' : 's'}
                  {' '}as <strong>Completed</strong> with disposition note.
                </span>
              </label>
              <label className="flex items-start gap-2 cursor-pointer text-xs text-blue-900">
                <input
                  type="checkbox"
                  checked={withdrawDraftSubmittals}
                  onChange={(e) => setWithdrawDraftSubmittals(e.target.checked)}
                  disabled={saving}
                  className="mt-0.5"
                />
                <span>
                  Withdraw <strong>{draftSubmittalsCount ?? '…'}</strong> in-progress submittal{draftSubmittalsCount === 1 ? '' : 's'}
                  {' '}(Draft or Package Mailed status). Approved / Denied / Withdrawn submittals are preserved as-is.
                </span>
              </label>
            </div>
          </div>

          {/* Warning */}
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-yellow-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-yellow-900">
              This changes the property's status and is intended for properties leaving the CAHP program.
              The action is logged to the Audit Log. Reversible via the regular Edit button if needed.
            </p>
          </div>

          {progress && (
            <div className="bg-teal-50 border border-teal-200 rounded-lg p-3 flex items-start gap-2">
              <div className="w-3 h-3 rounded-full border-2 border-teal-500 border-r-transparent animate-spin flex-shrink-0 mt-0.5" />
              <p className="text-xs text-teal-900">{progress}</p>
            </div>
          )}

          {error && (
            <div className="bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <p className="text-sm text-error font-mono-data text-xs">{error}</p>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-200 flex items-center justify-end gap-2 bg-gray-50 rounded-b-lg">
          <button
            onClick={onClose}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-white rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleSubmit}
            disabled={saving || !canSubmit}
            className="px-4 py-1.5 bg-error hover:bg-red-700 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Disposing…' : 'Dispose Property'}
          </button>
        </div>
      </div>
    </div>
  );
}
