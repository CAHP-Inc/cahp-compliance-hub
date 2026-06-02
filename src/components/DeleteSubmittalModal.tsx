import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { deleteListItem, LIST_NAMES, type Submittal } from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { formatDateOnly } from '../lib/dates';

interface DeleteSubmittalModalProps {
  submittal: Submittal;
  propertyTitle?: string;
  onClose: () => void;
}

/**
 * Permanently delete a single submittal created in error.
 *
 * A submittal is a standalone record — its related correspondence and outstanding
 * items are property-level, shared across all submittals for that property, so they
 * are intentionally NOT cascaded. Only the submittal row itself is removed.
 *
 * Type 'DELETE' to confirm. NOT undoable — deleteListItem captures the full
 * pre-delete state to the AuditLog for forensics.
 *
 * A submittal that has already been filed with DOR (status beyond Draft) carries a
 * frozen org chart snapshot and confirmation #; deleting it loses that audit trail,
 * so we warn loudly. Withdrawing rather than deleting is usually the right move for
 * those — deletion is meant for records entered by mistake.
 */
export function DeleteSubmittalModal({ submittal, propertyTitle, onClose }: DeleteSubmittalModalProps) {
  const navigate = useNavigate();
  const [confirmText, setConfirmText] = useState('');
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const f = submittal.fields;
  const status = f.SubmittalStatus ?? 'Draft';
  const isFiled = status !== 'Draft' && status !== 'Withdrawn';

  const expectedConfirm = 'DELETE';
  const canDelete = confirmText.trim() === expectedConfirm && !deleting;

  const handleDelete = async () => {
    setDeleting(true);
    setError(null);
    try {
      await deleteListItem(LIST_NAMES.Submittals, submittal.id);
      navigate('/submittals');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setDeleting(false);
    }
  };

  // Block accidental ESC close mid-delete
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !deleting) onClose();
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [deleting, onClose]);

  return (
    <div
      className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !deleting) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-2xl max-w-lg w-full max-h-[90vh] overflow-y-auto">
        {/* Header */}
        <div className="px-6 py-4 border-b-2 border-red-200 bg-red-50">
          <div className="flex items-start gap-3">
            <div className="w-10 h-10 rounded-full bg-error/10 flex items-center justify-center flex-shrink-0">
              <Icon name="alert" size={18} className="text-error" />
            </div>
            <div className="flex-1">
              <h2 className="text-lg font-bold text-error">Permanently delete submittal</h2>
              <p className="text-sm text-gray-700 mt-0.5">
                This permanently removes <strong>{f.Title || '(untitled submittal)'}</strong>. This action
                cannot be undone.
              </p>
            </div>
            {!deleting && (
              <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl leading-none">
                ×
              </button>
            )}
          </div>
        </div>

        <div className="px-6 py-4">
          {/* Record summary */}
          <ul className="space-y-1.5 text-sm bg-gray-50 rounded-md p-3 border border-gray-200">
            <SummaryRow label="Title" value={f.Title || '(untitled)'} />
            <SummaryRow label="Property" value={propertyTitle || '(unlinked)'} />
            <SummaryRow label="Tax Year" value={f.cahpTaxYear ?? '—'} mono />
            <SummaryRow label="Filing Type" value={f.FilingType ?? '—'} />
            <SummaryRow label="Status" value={status} />
            {f.DateFiled && <SummaryRow label="Date Filed" value={formatDateOnly(f.DateFiled)} mono />}
            {f.ConfirmationNumber && <SummaryRow label="Confirmation #" value={f.ConfirmationNumber} mono />}
          </ul>

          {isFiled ? (
            <p className="text-xs text-red-800 mt-3 bg-red-50 border border-red-200 rounded p-2">
              <strong>This submittal has been filed with DOR.</strong> Deleting it discards the frozen org
              chart snapshot, confirmation number, and filing audit trail. If this filing is real but no longer
              active, <strong>Withdraw</strong> it instead — deletion is meant only for records created in error.
            </p>
          ) : (
            <p className="text-xs text-amber-800 mt-3 bg-amber-50 border border-amber-200 rounded p-2">
              <strong>Note:</strong> The property, its correspondence, and outstanding items are shared across
              all submittals and are <strong>not</strong> affected. Only this submittal record is removed.
            </p>
          )}

          {/* Confirmation input */}
          <div className="mt-4">
            <label className="block text-xs font-semibold text-gray-700 uppercase tracking-wider mb-1">
              To confirm, type:{' '}
              <code className="bg-gray-100 px-1.5 py-0.5 rounded text-error font-mono-data">{expectedConfirm}</code>
            </label>
            <input
              type="text"
              value={confirmText}
              onChange={(e) => setConfirmText(e.target.value)}
              disabled={deleting}
              placeholder={expectedConfirm}
              className="w-full px-3 py-2 border-2 border-gray-300 rounded text-sm focus:outline-none focus:border-error font-mono-data"
              autoFocus
            />
          </div>

          {error && (
            <div className="mt-3 bg-red-50 border border-red-200 rounded p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <div className="text-xs text-error">
                <div className="font-semibold mb-0.5">Delete failed</div>
                <div className="font-mono-data">{error}</div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={deleting}
            className="px-4 py-2 text-sm text-gray-700 hover:bg-gray-200 rounded-md disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleDelete}
            disabled={!canDelete}
            className="bg-error hover:bg-red-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-4 py-2 rounded-md text-sm font-bold inline-flex items-center gap-2"
          >
            {deleting ? (
              <>
                <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                Deleting…
              </>
            ) : (
              <>
                <Icon name="alert" size={14} />
                Permanently delete submittal
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}

function SummaryRow({ label, value, mono }: { label: string; value: string; mono?: boolean }) {
  return (
    <li className="flex justify-between items-center gap-4 text-gray-700">
      <span className="text-gray-500">{label}</span>
      <span className={`font-semibold text-gray-900 text-right truncate ${mono ? 'font-mono-data' : ''}`}>
        {value}
      </span>
    </li>
  );
}
