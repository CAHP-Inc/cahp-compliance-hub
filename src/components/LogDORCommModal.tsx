import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type CorrChannel,
  type CorrespondenceDirection,
  type CahpTaxYear,
  type CahpState,
} from '../lib/sharepoint';

/**
 * Log General DOR Communication.
 *
 * Distinct from the formal Letter modal: this is for phone calls, emails,
 * meetings, and policy inquiries that don't fit the LetterType taxonomy.
 * A general comm:
 *   - Always has a channel (Email / Phone / Meeting / Other)
 *   - Has no LetterType
 *   - Can link to many properties or none at all (general inquiry)
 *   - Does NOT auto-create Outstanding Items or move submittal status
 *     (those are letter-specific signals)
 *
 * Writes to the same DOR Correspondence Log list as LogLetterModal, with
 * CorrChannel set to distinguish.
 */

const CHANNELS: { value: CorrChannel; label: string }[] = [
  { value: 'Email',   label: 'Email' },
  { value: 'Phone',   label: 'Phone call' },
  { value: 'Meeting', label: 'Meeting' },
  { value: 'Other',   label: 'Other' },
];

const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];

export function LogDORCommModal({
  onClose,
  onSuccess,
  defaultPropertyId,
}: {
  onClose: () => void;
  onSuccess: () => void;
  defaultPropertyId?: string;
}) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [selectedPropertyIds, setSelectedPropertyIds] = useState<Set<string>>(
    new Set(defaultPropertyId ? [defaultPropertyId] : []),
  );
  const [propertySearch, setPropertySearch] = useState('');
  const [channel, setChannel] = useState<CorrChannel>('Email');
  const [direction, setDirection] = useState<CorrespondenceDirection>('Outbound (to DOR)');
  const [commDate, setCommDate] = useState<string>(new Date().toISOString().slice(0, 10));
  const [subject, setSubject] = useState<string>('');
  const [summary, setSummary] = useState<string>('');
  const [participants, setParticipants] = useState<string>('');
  const [taxYear, setTaxYear] = useState<CahpTaxYear | ''>('');
  const [state, setState] = useState<CahpState | ''>('');

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);

  const sortedProperties = useMemo(() => {
    if (!properties.data) return [];
    return [...properties.data].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [properties.data]);

  const filteredProperties = useMemo(() => {
    const q = propertySearch.trim().toLowerCase();
    if (!q) return sortedProperties;
    return sortedProperties.filter((p) => (p.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedProperties, propertySearch]);

  const toggleProperty = (id: string) =>
    setSelectedPropertyIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

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

    setSaving(true);
    setError(null);

    const propIds = Array.from(selectedPropertyIds);
    const primaryProp = propIds[0] ?? null;

    try {
      const corrFields: Record<string, unknown> = {
        // Subject ends up as the row's Title, prefixed with the channel for easy
        // scanning of the list.
        Title: subject,
        Direction: direction,
        CorrChannel: channel,
        DateReceived: new Date(commDate).toISOString(),
      };
      if (primaryProp) corrFields.PropertyLookupId = primaryProp;
      if (summary || participants) {
        const combined = [
          participants.trim() ? `Participants: ${participants.trim()}` : '',
          summary.trim() ? summary.trim() : '',
        ].filter(Boolean).join('\n\n');
        if (combined) corrFields.RequestSummary = combined;
      }
      if (taxYear) corrFields.cahpTaxYear = taxYear;
      if (state) corrFields.cahpState = state;

      const corr = await createListItem<{ id: string }>(LIST_NAMES.Correspondence, corrFields);

      // Junction rows for every linked property (silently skipped on failure)
      for (const pid of propIds) {
        try {
          await createListItem(LIST_NAMES.CorrespondencePropertyLinks, {
            Title: `Corr ${corr.id} ↔ Property ${pid}`,
            CorrLookupId: Number(corr.id),
            PropertyLookupId: Number(pid),
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn(`Failed to link DOR comm to property ${pid}:`, e);
        }
      }

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
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Log DOR Communication</h3>
        <p className="text-sm text-gray-600 mb-4">
          Phone calls, emails, meetings with DOR. Leave properties blank for general inquiries.
          For formal incoming letters with a response deadline use <strong>Log Letter</strong> instead — it cascades to Outstanding Items + submittal status.
        </p>

        <div className="space-y-3">
          <Field label={`Properties (${selectedPropertyIds.size} selected — optional)`}>
            <input
              type="text"
              value={propertySearch}
              onChange={(e) => setPropertySearch(e.target.value)}
              placeholder="Search properties…"
              disabled={saving}
              className={inputClass + ' mb-1'}
            />
            <div className="border border-gray-300 rounded max-h-40 overflow-y-auto bg-white">
              {filteredProperties.length === 0 ? (
                <div className="px-3 py-2 text-[11px] text-gray-500 italic">No properties match.</div>
              ) : (
                filteredProperties.map((p) => (
                  <label key={p.id} className="flex items-center gap-2 px-2 py-1.5 hover:bg-teal-50 cursor-pointer text-xs">
                    <input
                      type="checkbox"
                      checked={selectedPropertyIds.has(String(p.id))}
                      onChange={() => toggleProperty(String(p.id))}
                      disabled={saving}
                    />
                    <span className="flex-1 truncate">{p.fields.Title}</span>
                    {p.fields.cahpState && (
                      <span className="text-[10px] text-gray-500 flex-shrink-0">{p.fields.cahpState}</span>
                    )}
                  </label>
                ))
              )}
            </div>
            <p className="text-[11px] text-gray-500 mt-1">
              For a general inquiry that doesn't touch a specific property, leave blank.
            </p>
          </Field>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <Field label="Channel" required>
              <select
                value={channel}
                onChange={(e) => setChannel(e.target.value as CorrChannel)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                {CHANNELS.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
              </select>
            </Field>

            <Field label="Direction">
              <select
                value={direction}
                onChange={(e) => setDirection(e.target.value as CorrespondenceDirection)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="Outbound (to DOR)">Outbound (to DOR)</option>
                <option value="Inbound (from DOR)">Inbound (from DOR)</option>
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

            <Field label="Participants (optional)">
              <input
                type="text"
                value={participants}
                onChange={(e) => setParticipants(e.target.value)}
                placeholder="Names + roles (DOR rep, your team)"
                className={inputClass}
                disabled={saving}
              />
            </Field>

            <Field label="Tax Year">
              <select
                value={taxYear}
                onChange={(e) => setTaxYear(e.target.value as CahpTaxYear | '')}
                className={`${inputClass} bg-white font-mono-data`}
                disabled={saving}
              >
                <option value="">—</option>
                {TAX_YEARS.map((y) => (<option key={y} value={y}>{y}</option>))}
              </select>
            </Field>

            <Field label="State">
              <select
                value={state}
                onChange={(e) => setState(e.target.value as CahpState | '')}
                className={`${inputClass} bg-white font-mono-data`}
                disabled={saving}
              >
                <option value="">—</option>
                <option value="SC">SC</option>
                <option value="NC">NC</option>
              </select>
            </Field>
          </div>

          <Field label="Subject" required>
            <input
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="e.g., Called DOR about CAHP fee question"
              className={inputClass}
              disabled={saving}
            />
          </Field>

          <Field label="Summary / Notes">
            <textarea
              value={summary}
              onChange={(e) => setSummary(e.target.value)}
              rows={4}
              placeholder="What was discussed, decisions made, next steps"
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
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50"
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
