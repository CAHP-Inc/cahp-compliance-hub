import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Deed,
  type DeedParcelLink,
  type DeedType,
  type TaxMapID,
  type Property,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { formatDateOnly, toDateOnlyISO } from '../lib/dates';

interface DeedsSectionProps {
  /** Filter deeds by grantee owner. If both ownerId AND propertyId omitted, shows all. */
  ownerId?: string;
  ownerTitle?: string;
  /** Filter deeds by any tax map IDs linked to this property */
  propertyId?: string;
  propertyTitle?: string;
}

const DEED_TYPES: DeedType[] = [
  'Warranty Deed',
  'Special Warranty Deed',
  'Limited Warranty Deed',
  'Quitclaim Deed',
  "Trustee's Deed",
  'Tax Deed',
  'Other',
];

const DEED_TYPE_STYLES: Record<DeedType, string> = {
  'Warranty Deed': 'bg-green-100 text-green-800',
  'Special Warranty Deed': 'bg-emerald-100 text-emerald-800',
  'Limited Warranty Deed': 'bg-teal-100 text-teal-800',
  'Quitclaim Deed': 'bg-amber-100 text-amber-800',
  "Trustee's Deed": 'bg-purple-100 text-purple-800',
  'Tax Deed': 'bg-red-100 text-red-800',
  'Other': 'bg-gray-100 text-gray-800',
};

/**
 * SharePoint URL fields are stored as { Url, Description }. Normalize to string.
 */
function getUrlValue(v: unknown): string {
  if (!v) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    return (v as { Url: string }).Url ?? '';
  }
  return '';
}

export function DeedsSection({ ownerId, ownerTitle, propertyId, propertyTitle }: DeedsSectionProps) {
  const deeds = useSharePointList<Deed>(LIST_NAMES.Deeds, { top: 500 });
  const links = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 1000 });
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });

  const [addOpen, setAddOpen] = useState(false);
  const [editingDeedId, setEditingDeedId] = useState<string | null>(null);

  // Index parcels by id for fast lookup
  const parcelsById = useMemo(() => {
    const m = new Map<string, TaxMapID>();
    (taxMapIDs.data ?? []).forEach((t) => m.set(String(t.id), t));
    return m;
  }, [taxMapIDs.data]);

  // For property filtering: identify which parcels belong to this property
  const propertyParcelIds = useMemo(() => {
    if (!propertyId) return null;
    const ids = new Set<string>();
    (taxMapIDs.data ?? []).forEach((t) => {
      if (String(t.fields.LinkedPropertyLookupId ?? '') === String(propertyId)) {
        ids.add(String(t.id));
      }
    });
    return ids;
  }, [taxMapIDs.data, propertyId]);

  // Index links by deed id (deed → set of parcel ids)
  const parcelsByDeed = useMemo(() => {
    const m = new Map<string, Set<string>>();
    (links.data ?? []).forEach((l) => {
      const dId = l.fields.DeedLookupId ? String(l.fields.DeedLookupId) : '';
      const tId = l.fields.TaxMapIDLookupId ? String(l.fields.TaxMapIDLookupId) : '';
      if (!dId || !tId) return;
      if (!m.has(dId)) m.set(dId, new Set());
      m.get(dId)!.add(tId);
    });
    return m;
  }, [links.data]);

  // Filter deeds based on owner / property context
  const filteredDeeds = useMemo(() => {
    let result = deeds.data ?? [];
    if (ownerId) {
      result = result.filter(
        (d) => String(d.fields.GranteeOwnerLookupId ?? '') === String(ownerId)
      );
    }
    if (propertyParcelIds) {
      result = result.filter((d) => {
        const linkedParcels = parcelsByDeed.get(String(d.id));
        if (!linkedParcels) return false;
        for (const pid of linkedParcels) {
          if (propertyParcelIds.has(pid)) return true;
        }
        return false;
      });
    }
    // Sort: most recent deed first
    return [...result].sort((a, b) => {
      const aD = a.fields.DateRecorded ?? '';
      const bD = b.fields.DateRecorded ?? '';
      return bD.localeCompare(aD);
    });
  }, [deeds.data, ownerId, propertyParcelIds, parcelsByDeed]);

  const isLoading = deeds.loading || links.loading || taxMapIDs.loading;

  if (isLoading) {
    return <div className="text-sm text-gray-500 py-4">Loading deeds…</div>;
  }

  const subtitle = ownerId
    ? `Deeds where ${ownerTitle ?? 'this entity'} is the grantee. A deed can cover multiple tax map IDs.`
    : propertyId
      ? `Deeds touching any tax map ID for ${propertyTitle ?? 'this property'}.`
      : 'All deeds in the system.';

  return (
    <div className="bg-white border-l-4 border-amber-500 border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 bg-amber-50 flex items-center justify-between">
        <div>
          <h3 className="text-sm font-semibold text-teal-900">Deeds</h3>
          <p className="text-xs text-gray-600 mt-0.5">{subtitle}</p>
        </div>
        {ownerId && (
          <button
            onClick={() => setAddOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={12} />
            Add Deed
          </button>
        )}
      </div>

      {filteredDeeds.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500">
          {ownerId
            ? 'No deeds recorded for this entity yet. Add one to capture the chain of title.'
            : 'No deeds touch this property yet.'}
        </div>
      ) : (
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Deed</th>
              <th className="px-4 py-3 text-left">Type</th>
              <th className="px-4 py-3 text-left">Grantor</th>
              <th className="px-4 py-3 text-left">Recorded</th>
              <th className="px-4 py-3 text-left">Book/Page</th>
              <th className="px-4 py-3 text-left">Tax Map IDs</th>
              <th className="px-4 py-3 text-right w-20"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filteredDeeds.map((d) => {
              const linkedParcelIds = parcelsByDeed.get(String(d.id)) ?? new Set<string>();
              const linkedParcels = [...linkedParcelIds]
                .map((id) => parcelsById.get(id))
                .filter((t): t is TaxMapID => !!t);
              const url = getUrlValue(d.fields.DocumentURL);
              return (
                <tr key={d.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">
                    {url ? (
                      <a
                        href={url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-teal-700 hover:text-teal-900 underline"
                        onClick={(e) => e.stopPropagation()}
                      >
                        {d.fields.Title}
                      </a>
                    ) : (
                      d.fields.Title
                    )}
                  </td>
                  <td className="px-4 py-3">
                    {d.fields.DeedType && (
                      <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${DEED_TYPE_STYLES[d.fields.DeedType]}`}>
                        {d.fields.DeedType}
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-xs text-gray-700">{d.fields.GrantorName || '—'}</td>
                  <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">
                    {formatDateOnly(d.fields.DateRecorded)}
                  </td>
                  <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">{d.fields.BookPage || '—'}</td>
                  <td className="px-4 py-3 text-xs">
                    {linkedParcels.length === 0 ? (
                      <span className="text-gray-400 italic">None linked</span>
                    ) : (
                      <div className="flex flex-col gap-0.5">
                        {linkedParcels.map((t) => (
                          <span key={t.id} className="font-mono-data text-[11px] text-gray-700">
                            {t.fields.Title}
                            {t.fields.ParcelAddress && (
                              <span className="text-gray-500 font-sans"> · {t.fields.ParcelAddress}</span>
                            )}
                          </span>
                        ))}
                      </div>
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <button
                      onClick={() => setEditingDeedId(d.id)}
                      className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50"
                    >
                      Edit
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}

      {addOpen && ownerId && (
        <DeedModal
          granteeOwnerId={ownerId}
          granteeOwnerTitle={ownerTitle ?? ''}
          onClose={() => setAddOpen(false)}
          onSaved={() => {
            deeds.refetch?.();
            links.refetch?.();
          }}
        />
      )}

      {editingDeedId && (
        <DeedModal
          deedId={editingDeedId}
          granteeOwnerId={ownerId ?? ''}
          granteeOwnerTitle={ownerTitle ?? ''}
          existingDeed={deeds.data?.find((d) => d.id === editingDeedId)}
          existingLinkedParcelIds={parcelsByDeed.get(editingDeedId) ?? new Set()}
          onClose={() => setEditingDeedId(null)}
          onSaved={() => {
            deeds.refetch?.();
            links.refetch?.();
          }}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Add / Edit Deed modal
// ---------------------------------------------------------------------------
interface DeedModalProps {
  deedId?: string;
  granteeOwnerId: string;
  granteeOwnerTitle: string;
  existingDeed?: Deed;
  existingLinkedParcelIds?: Set<string>;
  onClose: () => void;
  onSaved: () => void;
}

function DeedModal({
  deedId,
  granteeOwnerId,
  granteeOwnerTitle,
  existingDeed,
  existingLinkedParcelIds,
  onClose,
  onSaved,
}: DeedModalProps) {
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const allLinks = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 1000 });

  const [title, setTitle] = useState(existingDeed?.fields.Title ?? '');
  const [grantor, setGrantor] = useState(existingDeed?.fields.GrantorName ?? '');
  const [deedType, setDeedType] = useState<DeedType>(existingDeed?.fields.DeedType ?? 'Warranty Deed');
  const [dateRecorded, setDateRecorded] = useState(
    existingDeed?.fields.DateRecorded
      ? String(existingDeed.fields.DateRecorded).slice(0, 10)
      : ''
  );
  const [bookPage, setBookPage] = useState(existingDeed?.fields.BookPage ?? '');
  const [county, setCounty] = useState(existingDeed?.fields.RecordingCounty ?? '');
  const [consideration, setConsideration] = useState<string>(
    existingDeed?.fields.ConsiderationAmount?.toString() ?? ''
  );
  const [docUrl, setDocUrl] = useState(getUrlValue(existingDeed?.fields.DocumentURL));
  const [notes, setNotes] = useState(existingDeed?.fields.DeedNotes ?? '');
  const [selectedParcelIds, setSelectedParcelIds] = useState<Set<string>>(
    new Set(existingLinkedParcelIds ?? [])
  );
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Group parcels by property for display
  const parcelsGrouped = useMemo(() => {
    const groups = new Map<string, { propertyTitle: string; parcels: TaxMapID[] }>();
    (taxMapIDs.data ?? []).forEach((t) => {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : 'unlinked';
      const property = pid !== 'unlinked'
        ? properties.data?.find((p) => String(p.id) === pid)
        : null;
      const propTitle = property?.fields.Title ?? 'Unlinked parcels';
      if (!groups.has(pid)) {
        groups.set(pid, { propertyTitle: propTitle, parcels: [] });
      }
      groups.get(pid)!.parcels.push(t);
    });
    return Array.from(groups.entries())
      .map(([pid, group]) => ({ propertyId: pid, ...group }))
      .sort((a, b) => a.propertyTitle.localeCompare(b.propertyTitle));
  }, [taxMapIDs.data, properties.data]);

  const toggleParcel = (parcelId: string) => {
    setSelectedParcelIds((prev) => {
      const next = new Set(prev);
      if (next.has(parcelId)) next.delete(parcelId);
      else next.add(parcelId);
      return next;
    });
  };

  const handleSave = async () => {
    if (!title.trim()) {
      setError('Deed label is required.');
      return;
    }
    if (selectedParcelIds.size === 0) {
      setError('Select at least one tax map ID this deed covers.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Build deed payload
      const deedPayload: Record<string, unknown> = {
        Title: title.trim(),
        GranteeOwnerLookupId: Number(granteeOwnerId),
        GrantorName: grantor.trim() || undefined,
        DeedType: deedType,
        DateRecorded: toDateOnlyISO(dateRecorded),
        BookPage: bookPage.trim() || undefined,
        RecordingCounty: county.trim() || undefined,
        ConsiderationAmount: consideration ? Number(consideration) : undefined,
        DeedNotes: notes.trim() || undefined,
      };
      // SharePoint URL fields take object form
      if (docUrl.trim()) {
        deedPayload.DocumentURL = { Url: docUrl.trim(), Description: docUrl.trim() };
      } else {
        deedPayload.DocumentURL = null;
      }

      let savedDeedId: string;
      if (deedId) {
        await updateListItem(LIST_NAMES.Deeds, deedId, deedPayload);
        savedDeedId = deedId;
      } else {
        const created = await createListItem<{ id: string }>(LIST_NAMES.Deeds, deedPayload);
        savedDeedId = String(created.id);
      }

      // Diff junction rows: add new, remove unselected
      const previouslyLinked = existingLinkedParcelIds ?? new Set<string>();
      const toAdd = [...selectedParcelIds].filter((id) => !previouslyLinked.has(id));
      const toRemove = [...previouslyLinked].filter((id) => !selectedParcelIds.has(id));

      // Find existing junction rows for this deed (to know which to delete)
      const existingJunctionRows = (allLinks.data ?? []).filter(
        (l) => String(l.fields.DeedLookupId ?? '') === String(savedDeedId)
      );

      for (const parcelId of toAdd) {
        await createListItem(LIST_NAMES.DeedParcelLinks, {
          Title: `Deed ${savedDeedId} ↔ Parcel ${parcelId}`,
          DeedLookupId: Number(savedDeedId),
          TaxMapIDLookupId: Number(parcelId),
        });
      }
      for (const parcelId of toRemove) {
        const row = existingJunctionRows.find(
          (r) => String(r.fields.TaxMapIDLookupId ?? '') === String(parcelId)
        );
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

  const handleDelete = async () => {
    if (!deedId) return;
    const ok = window.confirm(
      'Delete this deed and all its parcel links?\n\nThis does not delete the underlying PDF document, only the structured deed record.'
    );
    if (!ok) return;
    setDeleting(true);
    try {
      // Delete junction rows first
      const existingJunctionRows = (allLinks.data ?? []).filter(
        (l) => String(l.fields.DeedLookupId ?? '') === String(deedId)
      );
      for (const row of existingJunctionRows) {
        await deleteListItem(LIST_NAMES.DeedParcelLinks, row.id);
      }
      // Then the deed itself
      await deleteListItem(LIST_NAMES.Deeds, deedId);
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
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">
            {deedId ? 'Edit Deed' : 'Add Deed'}
          </h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Grantee entity: <strong>{granteeOwnerTitle}</strong>
          </p>
        </div>

        <div className="px-6 py-4 space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <Row label="Deed Label / Instrument # *">
              <input
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                disabled={saving}
                placeholder='e.g., "WD 2024-08-15" or "Inst. 2024-12345"'
                className={INPUT}
                autoFocus
              />
            </Row>
            <Row label="Deed Type">
              <select
                value={deedType}
                onChange={(e) => setDeedType(e.target.value as DeedType)}
                disabled={saving}
                className={INPUT + ' bg-white'}
              >
                {DEED_TYPES.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </Row>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <Row label="Grantor (from)">
              <input
                type="text"
                value={grantor}
                onChange={(e) => setGrantor(e.target.value)}
                disabled={saving}
                placeholder='e.g., "John Doe" or "ABC Holdings LLC"'
                className={INPUT}
              />
            </Row>
            <Row label="Date Recorded">
              <input
                type="date"
                value={dateRecorded}
                onChange={(e) => setDateRecorded(e.target.value)}
                disabled={saving}
                className={INPUT}
              />
            </Row>
          </div>

          <div className="grid grid-cols-3 gap-4">
            <Row label="Book / Page or Instrument #">
              <input
                type="text"
                value={bookPage}
                onChange={(e) => setBookPage(e.target.value)}
                disabled={saving}
                placeholder='Book 1234 Pg 567'
                className={INPUT}
              />
            </Row>
            <Row label="Recording County">
              <input
                type="text"
                value={county}
                onChange={(e) => setCounty(e.target.value)}
                disabled={saving}
                placeholder="Greenville"
                className={INPUT}
              />
            </Row>
            <Row label="Consideration Amount">
              <input
                type="number"
                step="0.01"
                value={consideration}
                onChange={(e) => setConsideration(e.target.value)}
                disabled={saving}
                placeholder="0.00"
                className={INPUT}
              />
            </Row>
          </div>

          <Row label="Document URL (link to deed PDF in SharePoint)">
            <input
              type="url"
              value={docUrl}
              onChange={(e) => setDocUrl(e.target.value)}
              disabled={saving}
              placeholder="https://vanrockre.sharepoint.com/sites/..."
              className={INPUT}
            />
          </Row>

          <Row label="Notes">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              disabled={saving}
              rows={2}
              className={INPUT + ' resize-y'}
            />
          </Row>

          {/* Tax Map IDs multi-select */}
          <div>
            <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
              Tax Map IDs this deed conveys * <span className="text-gray-400 normal-case font-normal">({selectedParcelIds.size} selected)</span>
            </label>
            <div className="border border-gray-300 rounded max-h-64 overflow-y-auto bg-white">
              {parcelsGrouped.length === 0 ? (
                <div className="p-3 text-xs text-gray-500 italic">
                  No tax map IDs in the system yet. Add some on a property's Overview tab first.
                </div>
              ) : (
                parcelsGrouped.map((group) => (
                  <div key={group.propertyId} className="border-b border-gray-100 last:border-b-0">
                    <div className="px-3 py-1.5 bg-gray-50 text-[11px] font-semibold text-gray-700 uppercase tracking-wider">
                      {group.propertyTitle}
                    </div>
                    {group.parcels.map((p) => (
                      <label
                        key={p.id}
                        className="flex items-start gap-2 px-3 py-2 hover:bg-teal-50 cursor-pointer text-xs"
                      >
                        <input
                          type="checkbox"
                          checked={selectedParcelIds.has(p.id)}
                          onChange={() => toggleParcel(p.id)}
                          disabled={saving}
                          className="mt-0.5"
                        />
                        <div className="flex-1 min-w-0">
                          <div className="font-mono-data font-medium">{p.fields.Title}</div>
                          {p.fields.ParcelAddress && (
                            <div className="text-gray-500 text-[11px]">{p.fields.ParcelAddress}</div>
                          )}
                        </div>
                      </label>
                    ))}
                  </div>
                ))
              )}
            </div>
          </div>

          {error && (
            <div className="bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <p className="text-xs text-error">{error}</p>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-between gap-2">
          {deedId ? (
            <button
              onClick={handleDelete}
              disabled={saving || deleting}
              className="text-xs text-error hover:text-red-700 font-medium px-3 py-1.5 rounded hover:bg-red-50 disabled:opacity-50"
            >
              {deleting ? 'Deleting…' : 'Delete deed'}
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
              disabled={saving || deleting || !title.trim() || selectedParcelIds.size === 0}
              className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
            >
              {saving ? (
                <>
                  <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                  Saving…
                </>
              ) : (
                deedId ? 'Save' : 'Add Deed'
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
