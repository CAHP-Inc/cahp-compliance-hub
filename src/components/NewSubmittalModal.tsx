import { useState, useMemo, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type Submittal,
  type TaxMapID,
  type SubmittalFilingType,
  type CahpTaxYear,
  type CahpState,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

const FILING_TYPES: SubmittalFilingType[] = ['Initial', 'Annual', 'Amendment'];
const TAX_YEARS: CahpTaxYear[] = ['2024', '2025', '2026', '2027', '2028'];

// ---------------------------------------------------------------------------
// Single-submittal modal
// ---------------------------------------------------------------------------
interface NewSubmittalModalProps {
  /** If provided, locks the property and only shows tax map IDs for this property. */
  fixedPropertyId?: string;
  onClose: () => void;
  onCreated: (submittalId: string) => void;
}

export function NewSubmittalModal({ fixedPropertyId, onClose, onCreated }: NewSubmittalModalProps) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 1000 });

  const [propertyId, setPropertyId] = useState(fixedPropertyId ?? '');
  const [taxMapID, setTaxMapID] = useState('');
  const [taxYear, setTaxYear] = useState<CahpTaxYear>('2026');
  const [filingType, setFilingType] = useState<SubmittalFilingType>('Initial');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const selectedProperty = useMemo(
    () => properties.data?.find((p) => String(p.id) === String(propertyId)),
    [properties.data, propertyId]
  );

  const availableParcels = useMemo(() => {
    if (!propertyId) return [];
    return (taxMapIDs.data ?? []).filter(
      (t) => String(t.fields.LinkedPropertyLookupId ?? '') === String(propertyId)
    );
  }, [taxMapIDs.data, propertyId]);

  const sortedProperties = useMemo(() => {
    return [...(properties.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [properties.data]);

  // Reset tax map ID if property changes
  useEffect(() => {
    if (taxMapID && !availableParcels.find((p) => String(p.id) === taxMapID)) {
      setTaxMapID('');
    }
  }, [taxMapID, availableParcels]);

  // Check for duplicate
  const duplicate = useMemo(() => {
    if (!propertyId || !taxYear || !filingType) return null;
    return (submittals.data ?? []).find(
      (s) =>
        String(s.fields.PropertyLookupId ?? '') === String(propertyId) &&
        String(s.fields.TaxMapIDLookupId ?? '') === String(taxMapID) &&
        s.fields.cahpTaxYear === taxYear &&
        s.fields.FilingType === filingType
    );
  }, [submittals.data, propertyId, taxMapID, taxYear, filingType]);

  const handleSave = async () => {
    if (!propertyId) {
      setError('Property is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const parcel = taxMapID
        ? availableParcels.find((p) => String(p.id) === taxMapID)
        : null;
      const parcelSuffix = parcel ? ` — ${parcel.fields.Title}` : '';
      const submittalTitle = `${selectedProperty?.fields.Title ?? 'Property'}${parcelSuffix} — ${filingType} ${taxYear}`;
      const payload: Record<string, unknown> = {
        Title: submittalTitle,
        PropertyLookupId: Number(propertyId),
        cahpTaxYear: taxYear,
        cahpState: selectedProperty?.fields.cahpState,
        SubmittalStatus: 'Draft',
        FilingType: filingType,
      };
      if (taxMapID) {
        payload.TaxMapIDLookupId = Number(taxMapID);
      }
      const created = await createListItem<{ id: string }>(LIST_NAMES.Submittals, payload);
      onCreated(String(created.id));
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
      <div className="bg-white rounded-lg shadow-xl max-w-lg w-full">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">New Submittal</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Create a new DOR submittal record. One per tax map ID is required.
          </p>
        </div>

        <div className="px-6 py-4 space-y-4">
          <Row label="Property *">
            <select
              value={propertyId}
              onChange={(e) => setPropertyId(e.target.value)}
              disabled={saving || !!fixedPropertyId}
              className={INPUT + ' bg-white disabled:bg-gray-50'}
            >
              <option value="">— Select a property —</option>
              {sortedProperties.map((p) => (
                <option key={p.id} value={String(p.id)}>{p.fields.Title}</option>
              ))}
            </select>
          </Row>

          <Row label="Tax Map ID">
            <select
              value={taxMapID}
              onChange={(e) => setTaxMapID(e.target.value)}
              disabled={saving || !propertyId}
              className={INPUT + ' bg-white font-mono-data disabled:bg-gray-50'}
            >
              <option value="">— Unassigned (assign later) —</option>
              {availableParcels.map((t) => (
                <option key={t.id} value={String(t.id)}>
                  {t.fields.Title}
                  {t.fields.ParcelAddress ? ` — ${t.fields.ParcelAddress}` : ''}
                </option>
              ))}
            </select>
            {propertyId && availableParcels.length === 0 && (
              <p className="text-[11px] text-amber-700 mt-1 italic">
                This property has no tax map IDs yet. Add them on the property's Overview tab before creating per-parcel submittals.
              </p>
            )}
          </Row>

          <div className="grid grid-cols-2 gap-4">
            <Row label="Tax Year *">
              <select
                value={taxYear}
                onChange={(e) => setTaxYear(e.target.value as CahpTaxYear)}
                disabled={saving}
                className={INPUT + ' bg-white font-mono-data'}
              >
                {TAX_YEARS.map((y) => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
            </Row>
            <Row label="Filing Type *">
              <select
                value={filingType}
                onChange={(e) => setFilingType(e.target.value as SubmittalFilingType)}
                disabled={saving}
                className={INPUT + ' bg-white'}
              >
                {FILING_TYPES.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </Row>
          </div>

          {duplicate && (
            <div className="bg-amber-50 border border-amber-200 rounded p-3 text-xs text-amber-800 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-amber-700 flex-shrink-0 mt-0.5" />
              <div>
                A {filingType} submittal already exists for this {taxMapID ? 'tax map ID' : 'property'} in tax year {taxYear}.
                Saving will create a duplicate.
              </div>
            </div>
          )}

          {error && (
            <div className="bg-red-50 border border-red-200 rounded p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <p className="text-xs text-error">{error}</p>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={saving}
            className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-100 rounded-md"
          >
            Cancel
          </button>
          <button
            onClick={handleSave}
            disabled={saving || !propertyId}
            className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            {saving ? (
              <>
                <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                Creating…
              </>
            ) : (
              'Create Submittal'
            )}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Bulk-create modal — one submittal per tax map ID for a property
// ---------------------------------------------------------------------------
interface BulkCreateSubmittalsModalProps {
  propertyId: string;
  propertyTitle: string;
  propertyState?: CahpState;
  onClose: () => void;
  onCreated: (count: number) => void;
}

export function BulkCreateSubmittalsModal({
  propertyId,
  propertyTitle,
  propertyState,
  onClose,
  onCreated,
}: BulkCreateSubmittalsModalProps) {
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 1000 });

  const [taxYear, setTaxYear] = useState<CahpTaxYear>('2026');
  const [filingType, setFilingType] = useState<SubmittalFilingType>('Initial');
  const [creating, setCreating] = useState(false);
  const [progress, setProgress] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [skipExisting, setSkipExisting] = useState(true);

  const propertyParcels = useMemo(() => {
    return (taxMapIDs.data ?? []).filter(
      (t) => String(t.fields.LinkedPropertyLookupId ?? '') === String(propertyId)
    );
  }, [taxMapIDs.data, propertyId]);

  // For each parcel, check whether a submittal already exists for the given year + type
  const parcelStatus = useMemo(() => {
    return propertyParcels.map((p) => {
      const existing = (submittals.data ?? []).find(
        (s) =>
          String(s.fields.PropertyLookupId ?? '') === String(propertyId) &&
          String(s.fields.TaxMapIDLookupId ?? '') === String(p.id) &&
          s.fields.cahpTaxYear === taxYear &&
          s.fields.FilingType === filingType
      );
      return { parcel: p, existing };
    });
  }, [propertyParcels, submittals.data, propertyId, taxYear, filingType]);

  const toCreate = skipExisting
    ? parcelStatus.filter((s) => !s.existing)
    : parcelStatus;
  const alreadyExistCount = parcelStatus.filter((s) => s.existing).length;

  const handleCreate = async () => {
    if (toCreate.length === 0) {
      setError('Nothing to create.');
      return;
    }
    setCreating(true);
    setError(null);
    let created = 0;
    try {
      for (const { parcel } of toCreate) {
        setProgress(`Creating submittal ${created + 1} of ${toCreate.length}…`);
        const title = `${propertyTitle} — ${parcel.fields.Title} — ${filingType} ${taxYear}`;
        await createListItem(LIST_NAMES.Submittals, {
          Title: title,
          PropertyLookupId: Number(propertyId),
          TaxMapIDLookupId: Number(parcel.id),
          cahpTaxYear: taxYear,
          cahpState: propertyState,
          SubmittalStatus: 'Draft',
          FilingType: filingType,
        });
        created++;
      }
      setProgress(`Done — created ${created} submittal${created === 1 ? '' : 's'}.`);
      setTimeout(() => {
        onCreated(created);
        onClose();
      }, 800);
    } catch (err) {
      setError(`Created ${created} of ${toCreate.length} before failing: ${err instanceof Error ? err.message : String(err)}`);
      setCreating(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !creating) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">Bulk Create Submittals</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Generate one submittal per tax map ID for <strong>{propertyTitle}</strong>.
          </p>
        </div>

        <div className="px-6 py-4 space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <Row label="Tax Year">
              <select
                value={taxYear}
                onChange={(e) => setTaxYear(e.target.value as CahpTaxYear)}
                disabled={creating}
                className={INPUT + ' bg-white font-mono-data'}
              >
                {TAX_YEARS.map((y) => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
            </Row>
            <Row label="Filing Type">
              <select
                value={filingType}
                onChange={(e) => setFilingType(e.target.value as SubmittalFilingType)}
                disabled={creating}
                className={INPUT + ' bg-white'}
              >
                {FILING_TYPES.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </Row>
          </div>

          <label className="flex items-start gap-2 text-xs cursor-pointer">
            <input
              type="checkbox"
              checked={skipExisting}
              onChange={(e) => setSkipExisting(e.target.checked)}
              disabled={creating}
              className="mt-0.5"
            />
            <span>
              Skip parcels that already have a {filingType} submittal for tax year {taxYear}{' '}
              <span className="text-gray-500">({alreadyExistCount} existing)</span>
            </span>
          </label>

          {/* Preview */}
          <div className="bg-gray-50 border border-gray-200 rounded-lg p-3">
            <div className="text-xs font-semibold text-gray-700 uppercase tracking-wider mb-2">
              Will create {toCreate.length} submittal{toCreate.length === 1 ? '' : 's'}:
            </div>
            <div className="max-h-64 overflow-y-auto bg-white border border-gray-200 rounded">
              {parcelStatus.length === 0 ? (
                <div className="p-3 text-xs text-gray-500 italic">
                  This property has no tax map IDs. Add them on the Overview tab first.
                </div>
              ) : (
                <table className="w-full text-xs">
                  <tbody className="divide-y divide-gray-100">
                    {parcelStatus.map(({ parcel, existing }) => {
                      const willCreate = skipExisting ? !existing : true;
                      return (
                        <tr key={parcel.id} className={existing && skipExisting ? 'opacity-50' : ''}>
                          <td className="px-3 py-1.5 font-mono-data text-gray-800">
                            {parcel.fields.Title}
                            {parcel.fields.ParcelAddress && (
                              <span className="text-gray-500 font-sans ml-2">{parcel.fields.ParcelAddress}</span>
                            )}
                          </td>
                          <td className="px-3 py-1.5 text-right">
                            {existing ? (
                              <span className="text-amber-700 italic font-sans">already filed</span>
                            ) : willCreate ? (
                              <span className="text-green-700 font-sans">will create</span>
                            ) : (
                              <span className="text-gray-400 italic font-sans">skip</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>

          {progress && (
            <div className="bg-blue-50 border border-blue-200 rounded p-2 text-xs text-blue-800 flex items-center gap-2">
              {creating && (
                <div className="w-3 h-3 rounded-full border-2 border-blue-500 border-r-transparent animate-spin" />
              )}
              {progress}
            </div>
          )}

          {error && (
            <div className="bg-red-50 border border-red-200 rounded p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <p className="text-xs text-error">{error}</p>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={creating}
            className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-100 rounded-md"
          >
            Cancel
          </button>
          <button
            onClick={handleCreate}
            disabled={creating || toCreate.length === 0}
            className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            {creating ? 'Creating…' : `Create ${toCreate.length} submittal${toCreate.length === 1 ? '' : 's'}`}
          </button>
        </div>
      </div>
    </div>
  );
}

const INPUT = 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}
      </label>
      {children}
    </div>
  );
}

// Re-export navigate for callers that want to jump to the new submittal
export { useNavigate };
