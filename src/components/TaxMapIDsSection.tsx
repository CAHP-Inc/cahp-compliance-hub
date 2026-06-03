import { useState, useMemo, useEffect } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type TaxMapID,
  type Submittal,
  type Property,
  type Deed,
  type DeedParcelLink,
  type ParcelStatus,
  type CahpState,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { formatDateOnly } from '../lib/dates';
import { DeedModal } from './DeedsSection';
import { CURRENT_FILING_YEAR } from './layout/FilingFreezeBanner';

interface TaxMapIDsSectionProps {
  propertyId: string;
  propertyTitle: string;
  propertyState?: CahpState;
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
export function TaxMapIDsSection({ propertyId, propertyTitle, propertyState }: TaxMapIDsSectionProps) {
  const taxMapIds = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const deeds = useSharePointList<Deed>(LIST_NAMES.Deeds, { top: 500 });
  const links = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 2000 });

  const [addOpen, setAddOpen] = useState(false);
  const [editingParcelId, setEditingParcelId] = useState<string | null>(null);
  const [linkingDeedForParcelId, setLinkingDeedForParcelId] = useState<string | null>(null);

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

  // For each parcel, list of linked deeds (ordered by DateRecorded DESC = most recent first)
  const deedsByParcel = useMemo(() => {
    const map = new Map<string, Deed[]>();
    if (!links.data || !deeds.data) return map;
    const deedsById = new Map((deeds.data ?? []).map((d) => [String(d.id), d]));
    links.data.forEach((link) => {
      const parcelId = link.fields.TaxMapIDLookupId ? String(link.fields.TaxMapIDLookupId) : '';
      const deedId = link.fields.DeedLookupId ? String(link.fields.DeedLookupId) : '';
      if (!parcelId || !deedId) return;
      const deed = deedsById.get(deedId);
      if (!deed) return;
      if (!map.has(parcelId)) map.set(parcelId, []);
      map.get(parcelId)!.push(deed);
    });
    // Sort each parcel's deeds by date desc
    map.forEach((deedList) => {
      deedList.sort((a, b) => (b.fields.DateRecorded ?? '').localeCompare(a.fields.DateRecorded ?? ''));
    });
    return map;
  }, [links.data, deeds.data]);

  // Helper for refreshes
  const refetchAll = () => {
    taxMapIds.refetch?.();
    deeds.refetch?.();
    links.refetch?.();
  };

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
            properties spanning multiple parcels. The only per-parcel check is deed coverage —
            the rest of the filing checklist lives at the property/entity level.
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

      {/* Deed coverage summary */}
      {linkedParcels.length > 0 && (() => {
        const withDeed = linkedParcels.filter((p) => (deedsByParcel.get(p.id)?.length ?? 0) > 0).length;
        const without = linkedParcels.length - withDeed;
        if (without === 0) {
          return (
            <div className="px-4 py-2 bg-green-50 border-b border-green-200 text-xs text-green-800 flex items-center gap-2">
              <Icon name="check" size={12} className="text-green-700" />
              All {linkedParcels.length} parcels have at least one deed linked.
            </div>
          );
        }
        return (
          <div className="px-4 py-2 bg-amber-50 border-b border-amber-200 text-xs text-amber-800 flex items-center gap-2">
            <Icon name="alert" size={12} className="text-amber-700" />
            <strong>{without} of {linkedParcels.length}</strong> parcels need a deed linked or obtained.
          </div>
        );
      })()}

      {linkedParcels.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500">
          No tax map IDs recorded for this property yet. Add at least one to track filings.
        </div>
      ) : (
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Tax Map ID</th>
              <th className="px-4 py-3 text-left">Address</th>
              <th className="px-4 py-3 text-left">County</th>
              <th className="px-4 py-3 text-right">Acreage</th>
              <th className="px-4 py-3 text-left">Status</th>
              <th className="px-4 py-3 text-left">Deed</th>
              <th className="px-4 py-3 text-right">Submittals</th>
              <th className="px-4 py-3 text-right w-24"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {linkedParcels.map((p) => {
              const parcelDeeds = deedsByParcel.get(p.id) ?? [];
              const hasDeed = parcelDeeds.length > 0;
              return (
              <tr key={p.id} className="hover:bg-gray-50">
                <td className="px-4 py-3 font-mono-data text-xs font-medium text-gray-900">
                  {p.fields.Title}
                </td>
                <td className="px-4 py-3 text-xs text-gray-700">{p.fields.ParcelAddress || '—'}</td>
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
                <td className="px-4 py-3 text-xs">
                  {hasDeed ? (
                    <div className="flex flex-col gap-0.5">
                      {parcelDeeds.slice(0, 2).map((d) => (
                        <span key={d.id} className="inline-flex items-center gap-1">
                          <Icon name="check" size={10} className="text-success" />
                          <span className="font-mono-data">{d.fields.Title}</span>
                          {d.fields.DateRecorded && (
                            <span className="text-gray-500">· {formatDateOnly(d.fields.DateRecorded)}</span>
                          )}
                        </span>
                      ))}
                      {parcelDeeds.length > 2 && (
                        <span className="text-gray-500 italic">+ {parcelDeeds.length - 2} more</span>
                      )}
                    </div>
                  ) : (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-semibold">
                      <Icon name="alert" size={10} />
                      No deed
                    </span>
                  )}
                </td>
                <td className="px-4 py-3 text-right font-mono-data text-xs">
                  {submittalCountByParcel.get(p.id) ?? 0}
                </td>
                <td className="px-4 py-3 text-right">
                  <div className="flex items-center justify-end gap-1">
                    <button
                      onClick={() => setLinkingDeedForParcelId(p.id)}
                      className="text-[11px] text-amber-700 hover:text-amber-900 font-medium px-2 py-1 rounded hover:bg-amber-50"
                      title="Link an existing deed to this parcel"
                    >
                      Link Deed
                    </button>
                    <button
                      onClick={() => setEditingParcelId(p.id)}
                      className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50"
                    >
                      Edit
                    </button>
                  </div>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      )}

      {addOpen && (
        <TaxMapIDModal
          propertyId={propertyId}
          propertyTitle={propertyTitle}
          propertyState={propertyState}
          onClose={() => setAddOpen(false)}
          onSaved={() => {
            taxMapIds.refetch?.();
            submittals.refetch?.();
          }}
        />
      )}

      {editingParcelId && (
        <TaxMapIDModal
          parcelId={editingParcelId}
          propertyId={propertyId}
          propertyTitle={propertyTitle}
          propertyState={propertyState}
          onClose={() => setEditingParcelId(null)}
          onSaved={() => {
            taxMapIds.refetch?.();
            submittals.refetch?.();
          }}
        />
      )}

      {linkingDeedForParcelId && (
        <LinkDeedToParcelModal
          parcelId={linkingDeedForParcelId}
          parcelTitle={linkedParcels.find((p) => p.id === linkingDeedForParcelId)?.fields.Title ?? ''}
          existingDeedIds={new Set((deedsByParcel.get(linkingDeedForParcelId) ?? []).map((d) => String(d.id)))}
          onClose={() => setLinkingDeedForParcelId(null)}
          onSaved={refetchAll}
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
  propertyState?: CahpState;
  onClose: () => void;
  onSaved: () => void;
}

function TaxMapIDModal({ parcelId, propertyId, propertyTitle, propertyState, onClose, onSaved }: TaxMapIDModalProps) {
  const taxMapIds = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const existing = parcelId ? taxMapIds.data?.find((t) => t.id === parcelId) : undefined;

  const sortedProperties = useMemo(
    () => [...(properties.data ?? [])].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')),
    [properties.data],
  );

  const [taxMapID, setTaxMapID] = useState(existing?.fields.Title ?? '');
  const [selectedPropertyId, setSelectedPropertyId] = useState<string>(
    String(existing?.fields.LinkedPropertyLookupId ?? propertyId),
  );
  const [parcelAddress, setParcelAddress] = useState(existing?.fields.ParcelAddress ?? '');
  const [county, setCounty] = useState(existing?.fields.County ?? '');
  const [acreage, setAcreage] = useState<string>(existing?.fields.Acreage?.toString() ?? '');
  const [legalDesc, setLegalDesc] = useState(existing?.fields.LegalDescription ?? '');
  const [status, setStatus] = useState<ParcelStatus>(existing?.fields.ParcelStatus ?? 'Active');
  const [notes, setNotes] = useState(existing?.fields.ParcelNotes ?? '');
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [hydrated, setHydrated] = useState(!parcelId); // create mode = already "hydrated"

  // When editing an existing parcel, hydrate the form once data arrives from SharePoint.
  // useState only captures the FIRST render's value — at that point taxMapIds.data is
  // typically still undefined, so without this effect the form would stay blank.
  useEffect(() => {
    if (parcelId && existing && !hydrated) {
      setTaxMapID(existing.fields.Title ?? '');
      setParcelAddress(existing.fields.ParcelAddress ?? '');
      setCounty(existing.fields.County ?? '');
      setAcreage(existing.fields.Acreage?.toString() ?? '');
      setLegalDesc(existing.fields.LegalDescription ?? '');
      setStatus(existing.fields.ParcelStatus ?? 'Active');
      setNotes(existing.fields.ParcelNotes ?? '');
      setSelectedPropertyId(String(existing.fields.LinkedPropertyLookupId ?? propertyId));
      setHydrated(true);
    }
  }, [parcelId, existing, hydrated]);

  const handleSave = async () => {
    if (!taxMapID.trim()) {
      setError('Tax Map ID is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // In edit mode the property is selectable (to fix mis-filed parcels);
      // in create mode it's always the property whose page we're on.
      const targetPropertyId = parcelId ? (selectedPropertyId || String(propertyId)) : String(propertyId);
      const payload = {
        Title: taxMapID.trim(),
        LinkedPropertyLookupId: Number(targetPropertyId),
        ParcelAddress: parcelAddress.trim() || undefined,
        County: county.trim() || undefined,
        Acreage: acreage ? Number(acreage) : undefined,
        LegalDescription: legalDesc.trim() || undefined,
        ParcelStatus: status,
        ParcelNotes: notes.trim() || undefined,
      };
      if (parcelId) {
        await updateListItem(LIST_NAMES.TaxMapIDs, parcelId, payload);
        // If the parcel moved to a different property, its submittals must
        // follow — each submittal carries both TaxMapIDLookupId and
        // PropertyLookupId, and the SC freeze tracker groups by property.
        const previousPropertyId = String(existing?.fields.LinkedPropertyLookupId ?? propertyId);
        if (String(targetPropertyId) !== previousPropertyId) {
          const linkedSubmittals = (submittals.data ?? []).filter(
            (s) => String(s.fields.TaxMapIDLookupId ?? '') === String(parcelId),
          );
          for (const s of linkedSubmittals) {
            await updateListItem(LIST_NAMES.Submittals, s.id, {
              PropertyLookupId: Number(targetPropertyId),
            });
          }
        }
      } else {
        const newParcel = await createListItem<{ id: string }>(LIST_NAMES.TaxMapIDs, payload);
        // Auto-start a Draft submittal for the current filing year so every new
        // parcel shows up in the SC freeze tracker without manual setup.
        try {
          await createListItem(LIST_NAMES.Submittals, {
            Title: `${propertyTitle} — ${taxMapID.trim()} — Initial ${CURRENT_FILING_YEAR}`,
            PropertyLookupId: Number(propertyId),
            TaxMapIDLookupId: Number(newParcel.id),
            cahpTaxYear: CURRENT_FILING_YEAR,
            cahpState: propertyState,
            SubmittalStatus: 'Draft',
            FilingType: 'Initial',
          });
        } catch (draftErr) {
          // Don't fail the parcel save if the auto-draft fails — surface a soft warning.
          console.warn('Auto-create Draft submittal failed for new parcel', draftErr);
        }
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
          {parcelId && !hydrated ? (
            <div className="py-8 text-center text-sm text-gray-500">Loading parcel details…</div>
          ) : (
          <>
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

          {/* Reassign — only when editing an existing parcel. Lets you fix a
              parcel filed under the wrong property without re-entering it. */}
          {parcelId && (
            <Row label="Property">
              <select
                value={selectedPropertyId}
                onChange={(e) => setSelectedPropertyId(e.target.value)}
                disabled={saving || properties.loading}
                className={INPUT + ' bg-white'}
              >
                {sortedProperties.map((p) => (
                  <option key={p.id} value={String(p.id)}>{p.fields.Title}</option>
                ))}
              </select>
              {selectedPropertyId !== String(existing?.fields.LinkedPropertyLookupId ?? propertyId) && (
                <p className="text-[11px] text-amber-700 mt-1">
                  Reassigning this parcel will also move its linked submittals to the new property.
                </p>
              )}
            </Row>
          )}

          <Row label="Physical Address">
            <input
              type="text"
              value={parcelAddress}
              onChange={(e) => setParcelAddress(e.target.value)}
              disabled={saving}
              placeholder="e.g., 310 Walker Ave, Greenwood, SC 29649"
              className={INPUT}
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
          </>
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

// ---------------------------------------------------------------------------
// Link existing deed(s) to a parcel — multi-select to handle multi-deed coverage
// ---------------------------------------------------------------------------
interface LinkDeedToParcelModalProps {
  parcelId: string;
  parcelTitle: string;
  existingDeedIds: Set<string>;
  onClose: () => void;
  onSaved: () => void;
}

function LinkDeedToParcelModal({
  parcelId,
  parcelTitle,
  existingDeedIds,
  onClose,
  onSaved,
}: LinkDeedToParcelModalProps) {
  const deeds = useSharePointList<Deed>(LIST_NAMES.Deeds, { top: 500 });
  const links = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 2000 });
  const [selected, setSelected] = useState<Set<string>>(new Set(existingDeedIds));
  const [search, setSearch] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [creatingDeed, setCreatingDeed] = useState(false);

  // Live set of deeds currently linked to this parcel — derived from refetched
  // junction data. The `existingDeedIds` prop only seeds initial state; once
  // the user uploads a new deed in-flow we recompute from `links.data` so the
  // save-time diff doesn't blow away the link that the DeedModal just created.
  const liveLinkedDeedIds = useMemo(() => {
    const s = new Set<string>();
    (links.data ?? []).forEach((l) => {
      if (String(l.fields.TaxMapIDLookupId ?? '') === String(parcelId)) {
        s.add(String(l.fields.DeedLookupId ?? ''));
      }
    });
    return s;
  }, [links.data, parcelId]);

  const sortedDeeds = useMemo(() => {
    const all = [...(deeds.data ?? [])].sort((a, b) =>
      (b.fields.DateRecorded ?? '').localeCompare(a.fields.DateRecorded ?? '')
    );
    if (!search.trim()) return all;
    const q = search.toLowerCase();
    return all.filter((d) =>
      (d.fields.Title ?? '').toLowerCase().includes(q) ||
      (d.fields.FileLeafRef ?? '').toLowerCase().includes(q) ||
      (d.fields.BookPage ?? '').toLowerCase().includes(q)
    );
  }, [deeds.data, search]);

  const toggle = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleSave = async () => {
    setSaving(true);
    setError(null);
    try {
      // Diff against the LIVE set (not the stale prop) so links created
      // mid-flow by the inline DeedModal aren't deleted on save.
      const toAdd = [...selected].filter((id) => !liveLinkedDeedIds.has(id));
      const toRemove = [...liveLinkedDeedIds].filter((id) => !selected.has(id));
      // Find junction rows touching this parcel
      const parcelLinks = (links.data ?? []).filter(
        (l) => String(l.fields.TaxMapIDLookupId ?? '') === String(parcelId)
      );
      for (const deedId of toAdd) {
        await createListItem(LIST_NAMES.DeedParcelLinks, {
          Title: `Deed ${deedId} ↔ Parcel ${parcelId}`,
          DeedLookupId: Number(deedId),
          TaxMapIDLookupId: Number(parcelId),
        });
      }
      for (const deedId of toRemove) {
        const row = parcelLinks.find((l) => String(l.fields.DeedLookupId ?? '') === String(deedId));
        if (row) {
          await deleteListItem(LIST_NAMES.DeedParcelLinks, row.id);
        }
      }
      onSaved();
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
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-hidden flex flex-col">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">Link Deed(s) to Parcel</h2>
          <p className="text-xs text-gray-500 mt-0.5 font-mono-data">{parcelTitle}</p>
        </div>

        <div className="px-6 py-3 border-b border-gray-200">
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search deeds by label, grantor, or book/page…"
            className="w-full px-3 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
          />
        </div>

        <div className="flex-1 overflow-y-auto">
          {deeds.loading ? (
            <div className="p-6 text-center text-sm text-gray-500">Loading deeds…</div>
          ) : sortedDeeds.length === 0 ? (
            <div className="p-6 text-center text-sm text-gray-500">
              {search ? 'No deeds match your search.' : (
                <>
                  <p className="mb-3">No deed PDFs uploaded yet.</p>
                  <button
                    type="button"
                    onClick={() => setCreatingDeed(true)}
                    className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
                  >
                    <Icon name="plus" size={14} />
                    Upload a new deed
                  </button>
                  <p className="mt-3 text-xs text-gray-500">
                    Drag-drop the PDF, fill in book/page + date recorded, and we'll link it to this parcel automatically.
                  </p>
                </>
              )}
            </div>
          ) : (
            <table className="w-full text-sm">
              <tbody className="divide-y divide-gray-100">
                {sortedDeeds.map((d) => (
                  <tr
                    key={d.id}
                    onClick={() => toggle(String(d.id))}
                    className={`cursor-pointer hover:bg-teal-50 ${selected.has(String(d.id)) ? 'bg-teal-50' : ''}`}
                  >
                    <td className="px-4 py-2 w-10">
                      <input
                        type="checkbox"
                        checked={selected.has(String(d.id))}
                        onChange={() => toggle(String(d.id))}
                        disabled={saving}
                      />
                    </td>
                    <td className="px-4 py-2">
                      <div className="font-medium text-gray-900">{d.fields.Title || d.fields.FileLeafRef}</div>
                      <div className="text-[11px] text-gray-500 mt-0.5">
                        {d.fields.BookPage && <span className="font-mono-data">{d.fields.BookPage}</span>}
                        {d.fields.BookPage && d.fields.DateRecorded && <span> · </span>}
                        {d.fields.DateRecorded && <span>{formatDateOnly(d.fields.DateRecorded)}</span>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {error && (
          <div className="px-6 py-3 border-t border-red-200 bg-red-50">
            <p className="text-xs text-error">{error}</p>
          </div>
        )}

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 flex items-center justify-between gap-2">
          <div className="flex items-center gap-3">
            <span className="text-xs text-gray-600">
              {selected.size} selected · {liveLinkedDeedIds.size} currently linked
            </span>
            <button
              type="button"
              onClick={() => setCreatingDeed(true)}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50 inline-flex items-center gap-1"
              title="Upload a new deed PDF and link it to this parcel"
            >
              <Icon name="plus" size={10} />
              Upload new PDF
            </button>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              disabled={saving}
              className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded-md"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={saving}
              className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
            >
              {saving ? (
                <>
                  <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                  Saving…
                </>
              ) : (
                'Save Links'
              )}
            </button>
          </div>
        </div>
      </div>

      {creatingDeed && (
        <DeedModal
          preCheckedParcelIds={new Set([parcelId])}
          onClose={() => setCreatingDeed(false)}
          onSaved={(newDeedId) => {
            // Refetch so the new deed appears in the list and its junction row is visible
            deeds.refetch?.();
            links.refetch?.();
            // Auto-check the freshly uploaded deed so it stays linked after the user hits Save Links
            if (newDeedId) {
              setSelected((prev) => new Set([...prev, newDeedId]));
            }
          }}
        />
      )}
    </div>
  );
}
