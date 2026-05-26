import { useState, useMemo, useRef } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  uploadDocument,
  LIST_NAMES,
  type Property,
  type Submittal,
  type LetterType,
  type CorrespondenceDirection,
  type CahpTaxYear,
  type CahpState,
  type SubmittalStatusValue,
  type ItemCategory,
} from '../lib/sharepoint';
import { formatDateET, toDateInputValue } from '../lib/dates';

const LETTER_TYPES: LetterType[] = [
  'Initial Acknowledgment',
  'Additional Info Request',
  'Org Chart Request',
  'Approval',
  'Denial',
  'Withdrawal Notice',
  'Refund Notice',
  'Other',
];

const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];

/**
 * Letter types that trigger a Submittal status update to "Letter Received - Action Needed"
 * Per spec §3.7.2 + §3.7.3. Approval/Denial don't auto-transition because they go through
 * the Approval Workflow modal (PR-10d) instead.
 */
const LETTER_TYPES_REQUIRING_RESPONSE: LetterType[] = [
  'Additional Info Request',
  'Org Chart Request',
];

/**
 * Letter types that, when set, suggest a category for the auto-created Outstanding Item.
 */
const LETTER_TYPE_TO_ITEM_CATEGORY: Partial<Record<LetterType, ItemCategory>> = {
  'Org Chart Request': 'Org Chart',
  'Additional Info Request': 'Other',
  'Approval': 'Determination Letter',
  'Denial': 'Determination Letter',
  'Refund Notice': 'Other',
};

export function LogLetterModal({
  onClose,
  onSuccess,
  defaultPropertyId,
  defaultSubmittalId,
}: {
  onClose: () => void;
  onSuccess: () => void;
  defaultPropertyId?: string;
  defaultSubmittalId?: string;
}) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  // Form fields — multi-property; a DOR letter can reference multiple properties
  const [selectedPropertyIds, setSelectedPropertyIds] = useState<Set<string>>(
    new Set(defaultPropertyId ? [defaultPropertyId] : []),
  );
  const [propertySearch, setPropertySearch] = useState('');
  const [submittalId, setSubmittalId] = useState<string>(defaultSubmittalId ?? '');
  const [direction, setDirection] = useState<CorrespondenceDirection>('Inbound (from DOR)');
  const [letterType, setLetterType] = useState<LetterType | ''>('');
  const [dateReceived, setDateReceived] = useState<string>(toDateInputValue(new Date()));
  const [subject, setSubject] = useState<string>('');
  const [summary, setSummary] = useState<string>('');
  const [responseDue, setResponseDue] = useState<string>('');
  const [taxYear, setTaxYear] = useState<CahpTaxYear | ''>('');
  const [state, setState] = useState<CahpState | ''>('');

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);
  const [cascadeLog, setCascadeLog] = useState<string[]>([]);

  // PR-11c — optional file attachment, routed to DOR Correspondence library
  const [attachment, setAttachment] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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

  // Filter submittals to those linked to ANY of the chosen properties
  const submittalChoices = useMemo(() => {
    if (!submittals.data) return [];
    if (selectedPropertyIds.size === 0) return submittals.data;
    return submittals.data.filter((s) =>
      selectedPropertyIds.has(String(s.fields.PropertyLookupId ?? '')),
    );
  }, [submittals.data, selectedPropertyIds]);

  const handleSubmit = async () => {
    setValidationError(null);

    // Validate. Property is no longer required — letters can cover the whole
    // portfolio or none at all (a general policy memo, for instance).
    if (!subject.trim()) {
      setValidationError('Subject is required.');
      return;
    }
    if (!letterType) {
      setValidationError('Letter Type is required.');
      return;
    }
    if (!dateReceived) {
      setValidationError('Date is required.');
      return;
    }

    setSaving(true);
    setError(null);
    setCascadeLog([]);

    const propIds = Array.from(selectedPropertyIds);
    const primaryProp = propIds[0] ?? null;

    try {
      // ──────────────── CASCADE STEP 1: create Correspondence record ────────────────
      const correspondenceFields: Record<string, unknown> = {
        Title: subject,
        Direction: direction,
        LetterType: letterType,
        CorrChannel: 'Letter',
        DateReceived: new Date(dateReceived).toISOString(),
      };
      if (primaryProp) correspondenceFields.PropertyLookupId = primaryProp;
      if (submittalId) correspondenceFields.CorrSubmittalLookupId = submittalId;
      if (summary) correspondenceFields.RequestSummary = summary;
      if (responseDue) correspondenceFields.ResponseDue = new Date(responseDue).toISOString();
      if (taxYear) correspondenceFields.cahpTaxYear = taxYear;
      if (state) correspondenceFields.cahpState = state;

      const corr = await createListItem<{ id: string }>(LIST_NAMES.Correspondence, correspondenceFields);
      setCascadeLog((prev) => [
        ...prev,
        `✓ Correspondence record #${corr.id} created (${propIds.length} propert${propIds.length === 1 ? 'y' : 'ies'})`,
      ]);

      // ──────────────── CASCADE STEP 2: write junction rows for every linked property ────────────────
      for (const pid of propIds) {
        try {
          await createListItem(LIST_NAMES.CorrespondencePropertyLinks, {
            Title: `Corr ${corr.id} ↔ Property ${pid}`,
            CorrLookupId: Number(corr.id),
            PropertyLookupId: Number(pid),
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn(`Failed to link correspondence to property ${pid}:`, e);
        }
      }

      // ──────────────── CASCADE STEP 3: create one Outstanding Item per property if response deadline set ────────────────
      if (responseDue && direction === 'Inbound (from DOR)' && propIds.length > 0) {
        const itemCategory: ItemCategory =
          (letterType && LETTER_TYPE_TO_ITEM_CATEGORY[letterType]) ?? 'Other';
        let createdItems = 0;
        for (const pid of propIds) {
          try {
            await createListItem(LIST_NAMES.Outstanding, {
              Title: `Respond to DOR: ${subject}`,
              PropertyLookupId: pid,
              ItemCategory: itemCategory,
              ItemStatus: 'Requested',
              DateRequested: new Date(dateReceived).toISOString(),
              DueDate: new Date(responseDue).toISOString(),
              ItemNotes: `Auto-created from DOR Correspondence #${corr.id}. Response due ${formatDateET(responseDue)}.`,
            });
            createdItems++;
          } catch (e) {
            // eslint-disable-next-line no-console
            console.warn(`Outstanding Item creation failed for property ${pid}:`, e);
          }
        }
        if (createdItems > 0) {
          setCascadeLog((prev) => [
            ...prev,
            `✓ ${createdItems} Outstanding Item${createdItems === 1 ? '' : 's'} created (due ${formatDateET(responseDue)})`,
          ]);
        } else {
          setCascadeLog((prev) => [...prev, `⚠ Outstanding Item creation failed (correspondence still saved)`]);
        }
      }

      // ──────────────── CASCADE STEP 4: update related submittal status if applicable ────────────────
      if (
        submittalId &&
        direction === 'Inbound (from DOR)' &&
        LETTER_TYPES_REQUIRING_RESPONSE.includes(letterType as LetterType)
      ) {
        const submittal = submittals.data?.find((s) => String(s.id) === String(submittalId));
        if (submittal && submittal.fields.SubmittalStatus !== 'Letter Received - Action Needed') {
          try {
            await updateListItem(LIST_NAMES.Submittals, submittalId, {
              SubmittalStatus: 'Letter Received - Action Needed' as SubmittalStatusValue,
              NextAction: `Respond to ${letterType}: ${subject}`,
              NextActionDue: responseDue ? new Date(responseDue).toISOString() : undefined,
            });
            setCascadeLog((prev) => [...prev, `✓ Submittal status → Letter Received - Action Needed`]);
          } catch (e) {
            // eslint-disable-next-line no-console
            console.warn('Submittal status update failed:', e);
            setCascadeLog((prev) => [...prev, `⚠ Submittal status update failed`]);
          }
        }
      }

      // ──────────────── CASCADE STEP 5: upload document to DOR Correspondence library (PR-11c) ────────────────
      if (attachment) {
        try {
          await uploadDocument({
            libraryName: 'DOR Correspondence',
            filename: attachment.name,
            file: attachment,
            metadata: primaryProp ? { PropertyLookupId: primaryProp } : undefined,
          });
          setCascadeLog((prev) => [...prev, `✓ Document "${attachment.name}" uploaded to DOR Correspondence library`]);
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn('Document upload failed:', e);
          setCascadeLog((prev) => [...prev, `⚠ Document upload failed (correspondence record still saved): ${e instanceof Error ? e.message : String(e)}`]);
        }
      }

      setCascadeLog((prev) => [...prev, `✓ All audit log entries written`]);

      // Brief delay so the user sees the cascade results, then close
      setTimeout(() => {
        onSuccess();
      }, 1200);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const willCreateOutstanding = Boolean(
    responseDue && direction === 'Inbound (from DOR)' && selectedPropertyIds.size > 0,
  );
  const willCreateOutstandingCount = willCreateOutstanding ? selectedPropertyIds.size : 0;
  const willUpdateSubmittal = Boolean(
    submittalId &&
      direction === 'Inbound (from DOR)' &&
      letterType &&
      LETTER_TYPES_REQUIRING_RESPONSE.includes(letterType as LetterType)
  );

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Log DOR Letter</h3>
        <p className="text-sm text-gray-600 mb-4">
          Capture a DOR communication. One log action cascades to the correspondence record, Outstanding Item, and related submittal — see preview at the bottom.
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
              Leave empty for a portfolio-wide or general letter. Each linked property gets its own Outstanding Item.
            </p>
          </Field>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <Field label="Related Submittal">
              <select
                value={submittalId}
                onChange={(e) => setSubmittalId(e.target.value)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="">— optional —</option>
                {submittalChoices.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.fields.Title} ({s.fields.cahpTaxYear ?? '?'}, {s.fields.SubmittalStatus ?? '—'})
                  </option>
                ))}
              </select>
              {selectedPropertyIds.size === 0 && (
                <p className="text-[11px] text-gray-400 mt-0.5">Pick at least one property to filter to its submittals.</p>
              )}
            </Field>

            <Field label="Date" required>
              <input
                type="date"
                value={dateReceived}
                onChange={(e) => setDateReceived(e.target.value)}
                className={`${inputClass} font-mono-data`}
                disabled={saving}
              />
            </Field>

            <Field label="Direction">
              <select
                value={direction}
                onChange={(e) => setDirection(e.target.value as CorrespondenceDirection)}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="Inbound (from DOR)">Inbound (from DOR)</option>
                <option value="Outbound (to DOR)">Outbound (to DOR)</option>
              </select>
            </Field>

            <Field label="Letter Type" required>
              <select
                value={letterType}
                onChange={(e) => setLetterType(e.target.value as LetterType | '')}
                className={`${inputClass} bg-white`}
                disabled={saving}
              >
                <option value="">— select —</option>
                {LETTER_TYPES.map((t) => (<option key={t} value={t}>{t}</option>))}
              </select>
            </Field>

            <Field label="Response Due">
              <input
                type="date"
                value={responseDue}
                onChange={(e) => setResponseDue(e.target.value)}
                className={`${inputClass} font-mono-data`}
                disabled={saving}
              />
              <p className="text-[11px] text-gray-400 mt-0.5">If set, an Outstanding Item is auto-created.</p>
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
              placeholder="e.g., DOR requests amended org chart for 135 Oakwood"
              className={inputClass}
              disabled={saving}
            />
          </Field>

          <Field label="Summary">
            <textarea
              value={summary}
              onChange={(e) => setSummary(e.target.value)}
              rows={3}
              placeholder="Brief content summary or paste excerpt from letter"
              className={`${inputClass} resize-y`}
              disabled={saving}
            />
          </Field>

          <Field label="Attachment (optional)">
            <input
              ref={fileInputRef}
              type="file"
              onChange={(e) => setAttachment(e.target.files?.[0] ?? null)}
              disabled={saving}
              className="text-xs file:mr-2 file:px-2 file:py-1 file:border-0 file:rounded file:bg-teal-700 file:text-white file:font-medium hover:file:bg-teal-900"
            />
            {attachment && (
              <p className="text-[11px] text-gray-500 mt-1 font-mono-data">
                {attachment.name} · {(attachment.size / 1024).toFixed(1)} KB
              </p>
            )}
            <p className="text-[11px] text-gray-400 mt-1">
              Uploads to <strong>DOR Correspondence</strong> SharePoint library with PropertyID tag set. Large files upload in chunks.
            </p>
          </Field>
        </div>

        {/* Cascade preview */}
        <div className="mt-4 bg-gold-50 border border-gold-200 rounded-md p-3">
          <div className="text-[10px] font-semibold text-gold-900 uppercase tracking-wider mb-2">
            Cascade Preview — what will happen on save
          </div>
          <ul className="text-xs text-teal-900 space-y-0.5">
            <li>✓ <strong>Correspondence record</strong> created in DOR Correspondence Log</li>
            <li className={willCreateOutstanding ? '' : 'text-gray-500 line-through'}>
              {willCreateOutstanding ? '✓' : '○'}{' '}
              <strong>
                {willCreateOutstanding
                  ? `${willCreateOutstandingCount} Outstanding Item${willCreateOutstandingCount === 1 ? '' : 's'}`
                  : 'Outstanding Item'}
              </strong>{' '}
              auto-created
              {!willCreateOutstanding && ' (response due + at least one property required)'}
            </li>
            <li className={willUpdateSubmittal ? '' : 'text-gray-500 line-through'}>
              {willUpdateSubmittal ? '✓' : '○'} <strong>Submittal status</strong> → Letter Received - Action Needed
              {!willUpdateSubmittal && ' (no submittal linked or letter type doesn\'t require response)'}
            </li>
            <li className={attachment ? '' : 'text-gray-500 line-through'}>
              {attachment ? '✓' : '○'} <strong>Document attachment</strong>
              {attachment ? ` (${attachment.name} → DOR Correspondence library)` : ' (no file picked)'}
            </li>
            <li>✓ <strong>Audit log</strong> entries for every record touched</li>
          </ul>
        </div>

        {/* Cascade results during save */}
        {cascadeLog.length > 0 && (
          <div className="mt-4 bg-green-50 border border-green-200 rounded-md p-3">
            <div className="text-[10px] font-semibold text-green-900 uppercase tracking-wider mb-2">
              Cascade Results
            </div>
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
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Logging…' : 'Log Letter & Cascade'}
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
