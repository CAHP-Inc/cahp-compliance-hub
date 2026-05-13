import { useState } from 'react';
import { updateListItem, LIST_NAMES, type Property, type PropertyStatus } from '../lib/sharepoint';
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

export function DispositionModal({ property, onClose, onComplete }: DispositionModalProps) {
  const [type, setType] = useState<DispositionType | ''>('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [reason, setReason] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const canSubmit = type !== '' && date !== '' && reason.trim() !== '';

  const handleSubmit = async () => {
    if (!canSubmit) {
      setError('Disposition type, date, and reason are all required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Build a structured RemovedReason that's readable in SharePoint and audit log
      const formattedReason = `Disposed ${date} as ${type}.\nReason: ${reason.trim()}`;

      await updateListItem(LIST_NAMES.Properties, property.id, {
        PropertyStatus: type,
        RemovedReason: formattedReason,
      });

      onComplete();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
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

          {/* Warning */}
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-yellow-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-yellow-900">
              This changes the property's status and is intended for properties leaving the CAHP program.
              The action is logged to the Audit Log. Reversible via the regular Edit button if needed.
            </p>
          </div>

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
