import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type TaxMapID,
  type Submittal,
  type ParcelStatus,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

interface TaxMapIDsSectionProps {
  propertyId: string;
  propertyTitle: string;
}

const STATUS_STYLES: Record<ParcelStatus, string> = {
  Active: 'bg-green-100 text-green-800',
  Inactive: 'bg-gray-100 text-gray-600',
  Split: 'bg-amber-100 text-amber-800',
  Merged: 'bg-purple-100 text-purple-800',
};

/**
 * Tax Map IDs section — surfaces all parcels for the property and lets the user
 * add/edit/delete them. Each parcel can have its own submittals (DOR requires
 * one submission per tax map ID for properties spanning multiple parcels).
 */
export function TaxMapIDsSection({ propertyId, propertyTitle }: TaxMapIDsSectionProps) {
  const taxMapIds = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const [addOpen, setAddOpen] = useState(false);
  const [editingParcelId, setEditingParcelId] = useState<string | null>(null);

  const linkedParcels = useMemo(() => {
    return (taxMapIds.data ?? []).filter(
      (t) => String(t.fields.LinkedPropertyLookupId ?? '') === String(propertyId)
    );
  }, [taxMapIds.data, propertyId]);

  // For each parcel, count submittals
  const submittalCountByParcel = useMemo(() => {
    const map = new Map<string, number>();
    if (!submittals.data) return map;
    submittals.data.forEach((s) => {
      const pid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
      if (pid) {
        map.set(pid, (map.get(pid) ?? 0) + 1);
      }
    });
    return map;
  }, [submittals.data]);

  if (taxMapIds.loading) {
    return (
      <div className="text-sm text-gray-500 py-4">Loading tax map IDs…</div>
    );
  }

  return (
    <div className="bg-white border-l-4 border-blue-500 border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 bg-blue-50 flex items-center justify-between">
        <div>
          <h3 className="text-sm font-semibold text-teal-900">Tax Map IDs</h3>
          <p className="text-xs text-gray-600 mt-0.5">
            Per-parcel tracking. SCDOR requires one submittal per tax map ID for
            properties spanning multiple parcels.
          </p>
        </div>
        <button
          onClick={() => setAddOpen(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
        >
          <Icon name="plus" size={12} />
          Add Tax Map ID
        </button>
      </div>

      {linkedParcels.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500">
          No tax map IDs recorded for this property yet. Add at least one to track filings.
        </div>
      ) : (
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Tax Map ID</th>
              <th className="px-4 py-3 text-left">County</th>
              <th className="px-4 py-3 text-right">Acreage</th>
              <th className="px-4 py-3 text-left">Status</th>
              <th className="px-4 py-3 text-right">Submittals</th>
              <th className="px-4 py-3 text-right w-20"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {linkedParcels.map((p) => (
              <tr key={p.id} className="hover:bg-gray-50">
                <td className="px-4 py-3 font-mono-data text-xs font-medium text-gray-900">
                  {p.fields.Title}
                </td>
                <td className="px-4 py-3 text-xs text-gray-700">{p.fields.County || '—'}</td>
                <td className="px-4 py-3 text-right font-mono-data text-xs text-gray-700">
                  {p.fields.Acreage != null ? p.fields.Acreage.toFixed(2) : '—'}
                </td>
                <td className="px-4 py-3">
                  {p.fields.ParcelStatus && (
                    <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${STATUS_STYLES[p.fields.ParcelStatus]}`}>
                      {p.fields.ParcelStatus}
                    </span>
                  )}
                </td>
                <td className="px-4 py-3 text-right font-mono-data text-xs">
                  {submittalCountByParcel.get(p.id) ?? 0}
                </td>
                <td className="px-4 py-3 text-right">
                  <button
                    onClick={() => setEditingParcelId(p.id)}
                    className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50"
                  >
                    Edit
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {addOpen && (
        <TaxMapIDModal
          propertyId={propertyId}
          propertyTitle={propertyTitle}
          onClose={() => setAddOpen(false)}
          onSaved={() => taxMapIds.refetch?.()}
        />
      )}

      {editingParcelId && (
        <TaxMapIDModal
          parcelId={editingParcelId}
          propertyId={propertyId}
          propertyTitle={propertyTitle}
          onClose={() => setEditingParcelId(null)}
          onSaved={() => taxMapIds.refetch?.()}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Add/Edit Tax Map ID modal
// ---------------------------------------------------------------------------
interface TaxMapIDModalProps {
  parcelId?: string;
  propertyId: string;
  propertyTitle: string;
  onClose: () => void;
  onSaved: () => void;
}

function TaxMapIDModal({ parcelId, propertyId, propertyTitle, onClose, onSaved }: TaxMapIDModalProps) {
  const taxMapIds = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const existing = parcelId ? taxMapIds.data?.find((t) => t.id === parcelId) : undefined;

  const [taxMapID, setTaxMapID] = useState(existing?.fields.Title ?? '');
  const [county, setCounty] = useState(existing?.fields.County ?? '');
  const [acreage, setAcreage] = useState<string>(existing?.fields.Acreage?.toString() ?? '');
  const [legalDesc, setLegalDesc] = useState(existing?.fields.LegalDescription ?? '');
  const [status, setStatus] = useState<ParcelStatus>(existing?.fields.ParcelStatus ?? 'Active');
  const [notes, setNotes] = useState(existing?.fields.ParcelNotes ?? '');
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSave = async () => {
    if (!taxMapID.trim()) {
      setError('Tax Map ID is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const payload = {
        Title: taxMapID.trim(),
        LinkedPropertyLookupId: Number(propertyId),
        County: county.trim() || undefined,
        Acreage: acreage ? Number(acreage) : undefined,
        LegalDescription: legalDesc.trim() || undefined,
        ParcelStatus: status,
        ParcelNotes: notes.trim() || undefined,
      };
      if (parcelId) {
        await updateListItem(LIST_NAMES.TaxMapIDs, parcelId, payload);
      } else {
        await createListItem(LIST_NAMES.TaxMapIDs, payload);
      }
      onSaved();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!parcelId) return;
    const ok = window.confirm(
      'Delete this Tax Map ID?\n\nAny submittals linked to this parcel will lose their parcel reference but will NOT be deleted. You can reassign them afterwards.'
    );
    if (!ok) return;
    setDeleting(true);
    try {
      await deleteListItem(LIST_NAMES.TaxMapIDs, parcelId);
      onSaved();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setDeleting(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving && !deleting) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-lg w-full">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">
            {parcelId ? 'Edit Tax Map ID' : 'Add Tax Map ID'}
          </h2>
          <p className="text-xs text-gray-500 mt-0.5">For property: {propertyTitle}</p>
        </div>

        <div className="px-6 py-4 space-y-4">
          <Row label="Tax Map ID *">
            <input
              type="text"
              value={taxMapID}
              onChange={(e) => setTaxMapID(e.target.value)}
              disabled={saving}
              placeholder="e.g., 0123-45-678.000"
              className={INPUT}
              autoFocus
            />
          </Row>

          <div className="grid grid-cols-2 gap-4">
            <Row label="County">
              <input
                type="text"
                value={county}
                onChange={(e) => setCounty(e.target.value)}
                disabled={saving}
                placeholder="Greenville"
                className={INPUT}
              />
            </Row>
            <Row label="Acreage">
              <input
                type="number"
                step="0.01"
                value={acreage}
                onChange={(e) => setAcreage(e.target.value)}
                disabled={saving}
                placeholder="0.00"
                className={INPUT}
              />
            </Row>
          </div>

          <Row label="Parcel Status">
            <select
              value={status}
              onChange={(e) => setStatus(e.target.value as ParcelStatus)}
              disabled={saving}
              className={INPUT + ' bg-white'}
            >
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
              <option value="Split">Split (legacy — replaced by other parcels)</option>
              <option value="Merged">Merged (legacy — combined into another parcel)</option>
            </select>
          </Row>

          <Row label="Legal Description">
            <textarea
              value={legalDesc}
              onChange={(e) => setLegalDesc(e.target.value)}
              disabled={saving}
              rows={3}
              placeholder="As it appears on the deed (optional)"
              className={INPUT + ' resize-y'}
            />
          </Row>

          <Row label="Notes">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              disabled={saving}
              rows={2}
              placeholder="Anything else worth recording about this parcel"
              className={INPUT + ' resize-y'}
            />
          </Row>

          {error && (
            <div className="bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <p className="text-xs text-error">{error}</p>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-between gap-2">
          {parcelId ? (
            <button
              onClick={handleDelete}
              disabled={saving || deleting}
              className="text-xs text-error hover:text-red-700 font-medium px-3 py-1.5 rounded hover:bg-red-50 disabled:opacity-50"
            >
              {deleting ? 'Deleting…' : 'Delete'}
            </button>
          ) : <div />}
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              disabled={saving || deleting}
              className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-100 rounded-md"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={saving || deleting || !taxMapID.trim()}
              className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
            >
              {saving ? (
                <>
                  <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                  Saving…
                </>
              ) : (
                parcelId ? 'Save' : 'Add'
              )}
            </button>
          </div>
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
