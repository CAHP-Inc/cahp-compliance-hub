import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Deed,
  type DeedParcelLink,
  type TaxMapID,
  type Property,
  type Owner,
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

export function DeedsSection({ ownerId, ownerTitle, propertyId, propertyTitle }: DeedsSectionProps) {
  const deeds = useSharePointList<Deed>(LIST_NAMES.Deeds, { top: 500 });
  const links = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 1000 });
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });

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
          <a
            href="https://vanrockre.sharepoint.com/sites/CAHPComplianceHub/Property%20Deeds/Forms/AllItems.aspx"
            target="_blank"
            rel="noopener noreferrer"
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
            title="Upload a new deed PDF to the Property Deeds library, then come back here to fill in metadata."
          >
            <Icon name="plus" size={12} />
            Upload Deed PDF →
          </a>
        )}
      </div>

      {filteredDeeds.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500">
          {ownerId ? (
            <>
              <p className="mb-2">No deeds yet for this entity.</p>
              <p className="text-xs">
                Upload a deed PDF to the{' '}
                <a
                  href="https://vanrockre.sharepoint.com/sites/CAHPComplianceHub/Property%20Deeds/Forms/AllItems.aspx"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-teal-700 hover:text-teal-900 underline"
                >
                  Property Deeds library
                </a>
                , then refresh this page to fill in metadata and link parcels.
              </p>
            </>
          ) : 'No deeds touch this property yet.'}
        </div>
      ) : (
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Deed</th>
              <th className="px-4 py-3 text-left">Book/Page</th>
              <th className="px-4 py-3 text-left">Recorded</th>
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
              const url = d.webUrl ?? '';
              const displayLabel = d.fields.Title || d.fields.FileLeafRef || '(untitled)';
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
                        {displayLabel}
                      </a>
                    ) : (
                      displayLabel
                    )}
                  </td>
                  <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">{d.fields.BookPage || '—'}</td>
                  <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">
                    {formatDateOnly(d.fields.DateRecorded)}
                  </td>
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
// Add / Edit Deed modal (exported so other sections can open it)
// ---------------------------------------------------------------------------
export interface DeedModalProps {
  deedId?: string;
  /** If set, locks the grantee to this owner (no picker). Used from Owner page. */
  granteeOwnerId?: string;
  granteeOwnerTitle?: string;
  existingDeed?: Deed;
  existingLinkedParcelIds?: Set<string>;
  /** When opening for new deed, pre-check these parcels in the multi-select. */
  preCheckedParcelIds?: Set<string>;
  onClose: () => void;
  onSaved: () => void;
}

export function DeedModal({
  deedId,
  granteeOwnerId: fixedGranteeOwnerId,
  granteeOwnerTitle,
  existingDeed,
  existingLinkedParcelIds,
  preCheckedParcelIds,
  onClose,
  onSaved,
}: DeedModalProps) {
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const allLinks = useSharePointList<DeedParcelLink>(LIST_NAMES.DeedParcelLinks, { top: 1000 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  // Grantee can either be locked (passed in) or picked here
  const initialGranteeId =
    fixedGranteeOwnerId ||
    (existingDeed?.fields.GranteeOwnerLookupId ? String(existingDeed.fields.GranteeOwnerLookupId) : '');
  const [granteeOwnerId, setGranteeOwnerId] = useState(initialGranteeId);

  const [title, setTitle] = useState(existingDeed?.fields.Title ?? '');
  const [dateRecorded, setDateRecorded] = useState(
    existingDeed?.fields.DateRecorded
      ? String(existingDeed.fields.DateRecorded).slice(0, 10)
      : ''
  );
  const [bookPage, setBookPage] = useState(existingDeed?.fields.BookPage ?? '');
  const [selectedParcelIds, setSelectedParcelIds] = useState<Set<string>>(
    new Set([...(existingLinkedParcelIds ?? []), ...(preCheckedParcelIds ?? [])])
  );
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Lookup title of selected grantee for display
  const selectedGranteeTitle =
    granteeOwnerTitle ||
    owners.data?.find((o) => String(o.id) === String(granteeOwnerId))?.fields.Title ||
    '';

  const sortedOwners = useMemo(() => {
    return [...(owners.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [owners.data]);

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

  const [parcelSearch, setParcelSearch] = useState('');

  /** Check every parcel in a group. */
  const selectAllInGroup = (parcelIds: string[]) => {
    setSelectedParcelIds((prev) => {
      const next = new Set(prev);
      parcelIds.forEach((id) => next.add(id));
      return next;
    });
  };

  /** Uncheck every parcel in a group. */
  const clearGroup = (parcelIds: string[]) => {
    setSelectedParcelIds((prev) => {
      const next = new Set(prev);
      parcelIds.forEach((id) => next.delete(id));
      return next;
    });
  };

  /** Are all parcels in a group already checked? */
  const isGroupAllSelected = (parcelIds: string[]) =>
    parcelIds.length > 0 && parcelIds.every((id) => selectedParcelIds.has(id));

  /** Clear every selection. */
  const clearAllParcels = () => setSelectedParcelIds(new Set());

  const handleSave = async () => {
    if (!title.trim()) {
      setError('Deed label is required.');
      return;
    }
    if (!granteeOwnerId) {
      setError('Grantee entity is required.');
      return;
    }
    if (selectedParcelIds.size === 0) {
      setError('Select at least one tax map ID this deed covers.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Build deed payload (library item metadata)
      const deedPayload: Record<string, unknown> = {
        Title: title.trim(),
        GranteeOwnerLookupId: Number(granteeOwnerId),
        BookPage: bookPage.trim() || undefined,
        DateRecorded: toDateOnlyISO(dateRecorded),
      };

      // Library items are created by uploading a PDF to SharePoint, never
      // by app-side metadata create. If we don't have an existing deedId,
      // this is a bug — the modal shouldn't allow save in that case.
      if (!deedId) {
        setError('To create a new deed, first upload the PDF to the Property Deeds library in SharePoint, then come back and edit its metadata.');
        setSaving(false);
        return;
      }
      await updateListItem(LIST_NAMES.Deeds, deedId, deedPayload);
      const savedDeedId = deedId;

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
      'Unlink this deed from all its parcels?\n\n' +
      'This removes the parcel links and clears the deed metadata, but does NOT delete the PDF file from the Property Deeds library. ' +
      'To delete the PDF itself, remove it directly from SharePoint.'
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
      // Clear the metadata on the library item (don't delete the file)
      await updateListItem(LIST_NAMES.Deeds, deedId, {
        GranteeOwnerLookupId: null,
        BookPage: null,
        DateRecorded: null,
      });
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
          {fixedGranteeOwnerId && (
            <p className="text-xs text-gray-500 mt-0.5">
              Grantee entity: <strong>{selectedGranteeTitle || '—'}</strong>
            </p>
          )}
        </div>

        <div className="px-6 py-4 space-y-4">
          {!fixedGranteeOwnerId && (
            <Row label="Grantee Entity *">
              <select
                value={granteeOwnerId}
                onChange={(e) => setGranteeOwnerId(e.target.value)}
                disabled={saving}
                className={INPUT + ' bg-white'}
              >
                <option value="">— Select the receiving entity —</option>
                {sortedOwners.map((o) => (
                  <option key={o.id} value={String(o.id)}>{o.fields.Title}</option>
                ))}
              </select>
              <p className="text-[11px] text-gray-500 mt-1">
                The entity that received this deed. Usually the LLC that holds title.
              </p>
            </Row>
          )}
          <div className="grid grid-cols-2 gap-4">
            <Row label="Deed Label *">
              <input
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                disabled={saving}
                placeholder='e.g., "Townes at Converse - Book 151R Pg 575"'
                className={INPUT}
                autoFocus
              />
            </Row>
            <Row label="Book/Page *">
              <input
                type="text"
                value={bookPage}
                onChange={(e) => setBookPage(e.target.value)}
                disabled={saving}
                placeholder='e.g., "151R 575"'
                className={INPUT}
              />
            </Row>
          </div>

          <Row label="Date Recorded">
            <input
              type="date"
              value={dateRecorded}
              onChange={(e) => setDateRecorded(e.target.value)}
              disabled={saving}
              className={INPUT}
            />
          </Row>

          {existingDeed && (
            <Row label="PDF File">
              <div className="px-3 py-2 bg-gray-50 border border-gray-200 rounded text-xs flex items-center justify-between gap-2">
                <span className="font-mono-data text-gray-700 truncate">
                  {existingDeed.fields.FileLeafRef || '(filename unknown)'}
                </span>
                {existingDeed.webUrl && (
                  <a
                    href={existingDeed.webUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-teal-700 hover:text-teal-900 underline whitespace-nowrap"
                  >
                    Open PDF →
                  </a>
                )}
              </div>
              <p className="text-[11px] text-gray-500 mt-1">
                To replace the PDF, upload a new version directly in the{' '}
                <a
                  href="https://vanrockre.sharepoint.com/sites/CAHPComplianceHub/Property%20Deeds/Forms/AllItems.aspx"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-teal-700 hover:text-teal-900 underline"
                >
                  Property Deeds library
                </a>.
              </p>
            </Row>
          )}

          {/* Tax Map IDs multi-select */}
          <div>
            <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
              Tax Map IDs this deed conveys * <span className="text-gray-400 normal-case font-normal">({selectedParcelIds.size} selected)</span>
            </label>
            <div className="flex items-center gap-2 mb-1">
              <input
                type="text"
                value={parcelSearch}
                onChange={(e) => setParcelSearch(e.target.value)}
                placeholder="Search parcels by tax map ID or address…"
                className="flex-1 px-2 py-1 border border-gray-200 rounded text-xs focus:outline-none focus:border-teal-500"
              />
              {selectedParcelIds.size > 0 && (
                <button
                  type="button"
                  onClick={clearAllParcels}
                  className="text-[11px] text-gray-600 hover:text-red-700 font-medium px-2 py-1 rounded hover:bg-red-50 whitespace-nowrap"
                >
                  Clear all ({selectedParcelIds.size})
                </button>
              )}
            </div>
            <div className="border border-gray-300 rounded max-h-64 overflow-y-auto bg-white">
              {parcelsGrouped.length === 0 ? (
                <div className="p-3 text-xs text-gray-500 italic">
                  No tax map IDs in the system yet. Add some on a property's Overview tab first.
                </div>
              ) : (
                parcelsGrouped.map((group) => {
                  // Apply search filter to this group's parcels
                  const q = parcelSearch.toLowerCase().trim();
                  const visibleParcels = q
                    ? group.parcels.filter((p) =>
                        (p.fields.Title ?? '').toLowerCase().includes(q) ||
                        (p.fields.ParcelAddress ?? '').toLowerCase().includes(q)
                      )
                    : group.parcels;
                  if (visibleParcels.length === 0) return null;

                  const allParcelIds = group.parcels.map((p) => p.id);
                  const allSelected = isGroupAllSelected(allParcelIds);
                  return (
                    <div key={group.propertyId} className="border-b border-gray-100 last:border-b-0">
                      <div className="px-3 py-1.5 bg-gray-50 flex items-center justify-between gap-2 sticky top-0 z-10">
                        <span className="text-[11px] font-semibold text-gray-700 uppercase tracking-wider">
                          {group.propertyTitle}
                          <span className="ml-1.5 text-gray-400 normal-case font-normal">({group.parcels.length})</span>
                        </span>
                        <button
                          type="button"
                          onClick={() =>
                            allSelected ? clearGroup(allParcelIds) : selectAllInGroup(allParcelIds)
                          }
                          className="text-[10px] font-medium text-teal-700 hover:text-teal-900 px-2 py-0.5 rounded hover:bg-teal-50 whitespace-nowrap"
                        >
                          {allSelected ? 'Clear this property' : 'Select all under this property'}
                        </button>
                      </div>
                      {visibleParcels.map((p) => (
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
                  );
                })
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
              {deleting ? 'Unlinking…' : 'Unlink + clear metadata'}
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
              disabled={saving || deleting || !title.trim() || !granteeOwnerId || selectedParcelIds.size === 0}
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
