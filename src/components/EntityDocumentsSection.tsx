import { useMemo, useState } from 'react';
import { useSharePointList, updateListItem, deleteListItem, LIST_NAMES, type Owner } from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY, UploadDocumentModal } from './UploadDocumentModal';
import { Icon } from './ui/Icon';
import { formatDateET } from '../lib/dates';

interface DocItemRaw {
  id: string;
  webUrl?: string;
  lastModifiedDateTime: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    OwnerLookupId?: string;
    PropertyLookupId?: string;
    Modified?: string;
    Editor?: { LookupValue?: string };
    File_x0020_Size?: string;
  };
}

interface AggregatedDoc {
  id: string;
  itemId: string;
  library: string;
  filename: string;
  webUrl?: string;
  uploadDate?: string;
  uploader?: string;
  size?: number;
  currentOwnerTag?: string;  // For in-app tagging — the OwnerLookupId on this doc, if any
}

export interface EntityDocumentsSectionProps {
  /** Owner IDs whose documents should be shown (used when reading from tagged libraries) */
  ownerIds: string[];
  /** Optional owner title for the upload modal scope label */
  primaryOwnerTitle?: string;
  /** Optional title — defaults to "Documents" */
  title?: string;
  /** Optional subtitle */
  subtitle?: string;
  /** When provided, enables an Upload button that tags to this owner */
  uploadOwnerId?: string;
  /** Visual variant — 'inline' for embedding inside a tab, 'card' for standalone */
  variant?: 'inline' | 'card';
  /**
   * When true, reads ALL files from the dedicated CAHP Entity Documents library
   * regardless of OwnerLookupId. Used for CAHP entity contexts where the library
   * membership itself is the filter. ownerIds is ignored in this mode.
   */
  useCahpEntityLibrary?: boolean;
  /**
   * When true, only docs tagged to one of the ownerIds are shown. Untagged docs
   * are excluded. Use this on entity detail pages where users expect a strict
   * per-entity view. Default false (untagged docs treated as shared).
   */
  strictEntityFilter?: boolean;
}

export function EntityDocumentsSection({
  ownerIds,
  primaryOwnerTitle,
  title = 'Documents',
  subtitle,
  uploadOwnerId,
  variant = 'card',
  useCahpEntityLibrary = false,
  strictEntityFilter = false,
}: EntityDocumentsSectionProps) {
  // Fetch the 8 property libraries (for OwnerLookupId-tagged docs)
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });
  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];

  // Also fetch the dedicated CAHP Entity Documents library
  const cahpLib = useSharePointList<DocItemRaw>(CAHP_ENTITY_LIBRARY, { top: 500 });

  const [uploadOpen, setUploadOpen] = useState(false);
  const [manageMode, setManageMode] = useState(false);
  const [savingTagId, setSavingTagId] = useState<string | null>(null);

  // Fetch all owners (regardless of in-scope) to populate the tagging dropdown in manage mode
  const allOwners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const cahpOwnerOptions = useMemo(() => {
    return (allOwners.data ?? [])
      .filter((o) => {
        const t = (o.fields.Title ?? '').toLowerCase();
        return t.includes('cahp') || t.includes('carolina affordable housing project');
      })
      .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [allOwners.data]);

  const ownerIdSet = useMemo(() => new Set(ownerIds.map(String)), [ownerIds]);

  const docs = useMemo(() => {
    const collected: AggregatedDoc[] = [];

    // Multi-value Lookup columns surface as either OwnerLookupId (single ID),
    // OwnerLookupId@odata.type with arrays, or as an array under the raw column
    // name. Normalize all shapes into an array of string IDs. Empty array = untagged.
    const extractTagIds = (fields: Record<string, unknown> | undefined, key: 'Owner' | 'Property'): string[] => {
      if (!fields) return [];
      const lookupKey = `${key}LookupId`;
      const direct = (fields as Record<string, unknown>)[lookupKey];
      if (Array.isArray(direct)) {
        return direct.map((v) => String(v)).filter(Boolean);
      }
      if (direct !== undefined && direct !== null && direct !== '') {
        return [String(direct)];
      }
      // Some multi-value lookups expose the array under the raw name, with each
      // entry as { LookupId, LookupValue }
      const rawArr = (fields as Record<string, unknown>)[key];
      if (Array.isArray(rawArr)) {
        return rawArr
          .map((v) => (typeof v === 'object' && v !== null && 'LookupId' in v ? String((v as { LookupId: unknown }).LookupId) : null))
          .filter((id): id is string => !!id);
      }
      return [];
    };

    // 1. Files from the dedicated CAHP Entity Documents library
    // Filter by OwnerLookupId match. Untagged docs are treated as "shared" (visible everywhere)
    // UNLESS strictEntityFilter is true, in which case untagged docs are hidden.
    if (useCahpEntityLibrary && cahpLib.data) {
      cahpLib.data.forEach((item) => {
        const ownerTags = extractTagIds(item.fields as unknown as Record<string, unknown>, 'Owner');
        const hasAnyTag = ownerTags.length > 0;
        const matchesScope = ownerTags.some((id) => ownerIdSet.has(id));
        // In manage mode, surface ALL docs (including untagged + out-of-scope) so they can be tagged
        const include = manageMode
          ? true
          : ownerIds.length === 0 ||                            // no scope restriction → show all
            matchesScope ||                                     // tagged to ANY in-scope entity
            (!hasAnyTag && !strictEntityFilter);                // untagged: include unless strict mode
        if (!include) return;
        collected.push({
          id: `${CAHP_ENTITY_LIBRARY}:${item.id}`,
          itemId: item.id,
          library: CAHP_ENTITY_LIBRARY,
          filename: item.fields.FileLeafRef || item.fields.Title || '(unnamed)',
          webUrl: item.webUrl,
          uploadDate: item.fields.Modified || item.lastModifiedDateTime,
          uploader: item.fields.Editor?.LookupValue,
          size: item.fields.File_x0020_Size ? parseInt(item.fields.File_x0020_Size, 10) : undefined,
          currentOwnerTag: ownerTags[0],
        });
      });
    }

    // 2. Owner-tagged files from the 8 property-linked libraries (multi-value-aware)
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((item) => {
        const ownerTags = extractTagIds(item.fields as unknown as Record<string, unknown>, 'Owner');
        if (ownerTags.length === 0) return;
        if (!ownerTags.some((id) => ownerIdSet.has(id))) return;
        // If the doc is also property-scoped, it belongs on a property page, not
        // an owner page. Without this guard, the scrub-script Owner back-fill
        // (which sets Owner from each property's primary direct owner) causes
        // every sibling entity under a shared parent to surface every other
        // sibling's property-specific docs as if they were owner-level files.
        const propTags = extractTagIds(item.fields as unknown as Record<string, unknown>, 'Property');
        if (propTags.length > 0) return;
        collected.push({
          id: `${libraryName}:${item.id}`,
          itemId: item.id,
          library: libraryName,
          filename: item.fields.FileLeafRef || item.fields.Title || '(unnamed)',
          webUrl: item.webUrl,
          uploadDate: item.fields.Modified || item.lastModifiedDateTime,
          uploader: item.fields.Editor?.LookupValue,
          size: item.fields.File_x0020_Size ? parseInt(item.fields.File_x0020_Size, 10) : undefined,
        });
      });
    });

    // Sort by library then by upload date desc
    return collected.sort((a, b) => {
      if (a.library !== b.library) return a.library.localeCompare(b.library);
      const da = a.uploadDate ? new Date(a.uploadDate).getTime() : 0;
      const db = b.uploadDate ? new Date(b.uploadDate).getTime() : 0;
      return db - da;
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data,
    cahpLib.data, ownerIdSet, useCahpEntityLibrary, strictEntityFilter, manageMode,
  ]);

  // Count of untagged docs in the CAHP library — surface as a hint on entity pages in strict mode
  const untaggedCahpCount = useMemo(() => {
    if (!useCahpEntityLibrary || !cahpLib.data) return 0;
    return cahpLib.data.filter((d) => !d.fields.OwnerLookupId).length;
  }, [useCahpEntityLibrary, cahpLib.data]);

  const loading = libraries.some((l) => l.loading) || (useCahpEntityLibrary && cahpLib.loading);
  const refetchAll = () => {
    libraries.forEach((l) => l.refetch?.());
    cahpLib.refetch?.();
  };

  const handleTagChange = async (doc: AggregatedDoc, newOwnerId: string) => {
    setSavingTagId(doc.id);
    try {
      const numericId = newOwnerId ? Number(newOwnerId) : null;
      if (newOwnerId && (numericId === null || Number.isNaN(numericId))) {
        throw new Error(`Invalid owner ID: ${newOwnerId}`);
      }

      // Try the standard write, then verify it persisted by re-reading.
      // If it didn't persist, try alternate field names. This handles cases where
      // SharePoint named the lookup ID column differently than expected.
      const candidates = ['OwnerLookupId', 'Owner', 'OwnerId'];
      let succeededWith: string | null = null;
      let lastError: unknown = null;

      for (const fieldName of candidates) {
        try {
          // eslint-disable-next-line no-console
          console.log(`[CAHP tag] PATCH attempt with field="${fieldName}" value=${numericId}`);
          await updateListItem(CAHP_ENTITY_LIBRARY, doc.itemId, {
            [fieldName]: numericId,
          });

          // Re-read to verify
          const verify = await import('../lib/sharepoint/client').then((m) =>
            m.getListItem<{ fields: Record<string, unknown> }>(CAHP_ENTITY_LIBRARY, doc.itemId)
          );
          const readBack =
            verify.fields.OwnerLookupId ?? verify.fields.OwnerId ?? verify.fields.Owner;
          // eslint-disable-next-line no-console
          console.log(`[CAHP tag] After PATCH read-back:`, { OwnerLookupId: verify.fields.OwnerLookupId, OwnerId: verify.fields.OwnerId, Owner: verify.fields.Owner });

          const persisted =
            numericId === null
              ? !readBack
              : String(readBack ?? '') === String(numericId);
          if (persisted) {
            succeededWith = fieldName;
            break;
          }
          // eslint-disable-next-line no-console
          console.warn(`[CAHP tag] PATCH with "${fieldName}" returned OK but value did NOT persist`);
        } catch (e) {
          lastError = e;
          // eslint-disable-next-line no-console
          console.warn(`[CAHP tag] PATCH with "${fieldName}" threw:`, e);
        }
      }

      if (!succeededWith) {
        const detail = lastError instanceof Error ? lastError.message : 'No field name accepted the write';
        throw new Error(
          `Tag did not persist. Tried: ${candidates.join(', ')}. ${detail}. ` +
          `Open browser DevTools → Console to see what SharePoint returned, ` +
          `then send the log lines to Brandy's developer.`
        );
      }

      // eslint-disable-next-line no-console
      console.log(`[CAHP tag] Persisted via field="${succeededWith}"`);
      await cahpLib.refetch?.();
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error('Failed to update tag:', e);
      alert('Failed to save tag: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setSavingTagId(null);
    }
  };

  // Group by library for display
  const byLibrary = useMemo(() => {
    const map = new Map<string, AggregatedDoc[]>();
    docs.forEach((d) => {
      const list = map.get(d.library) ?? [];
      list.push(d);
      map.set(d.library, list);
    });
    return map;
  }, [docs]);

  const containerClass =
    variant === 'card'
      ? 'bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden mb-6'
      : 'bg-gold-50 border border-gold-200 rounded-lg mb-6 overflow-hidden';

  return (
    <div className={containerClass}>
      <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between flex-wrap gap-2">
        <div>
          <h3 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
            <Icon name="folder" size={14} className="text-teal-700" />
            {title}
            <span className="text-[10px] font-mono-data bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded-full">
              {docs.length}
            </span>
          </h3>
          {subtitle && <p className="text-xs text-gray-500 mt-0.5">{subtitle}</p>}
        </div>
        <div className="flex items-center gap-2">
          {useCahpEntityLibrary && (
            <button
              onClick={() => setManageMode((v) => !v)}
              className={`px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5 border ${
                manageMode
                  ? 'bg-amber-100 text-amber-900 border-amber-300 hover:bg-amber-200'
                  : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'
              }`}
            >
              <Icon name={manageMode ? 'check' : 'file'} size={12} />
              {manageMode ? 'Done tagging' : 'Manage tags'}
            </button>
          )}
          {uploadOwnerId && (
            <button
              onClick={() => setUploadOpen(true)}
              className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
            >
              <Icon name="plus" size={12} />
              Upload
            </button>
          )}
        </div>
      </div>

      {loading ? (
        <div className="px-4 py-6 text-center text-xs text-gray-500">Loading documents…</div>
      ) : (
        <>
          {manageMode && (
            <div className="px-4 py-2 bg-amber-50 border-t border-amber-200 text-xs text-amber-900">
              <strong>Manage tags mode.</strong> All docs in the CAHP Entity Documents library are listed below
              (including untagged + out-of-scope). Pick an entity from the dropdown on each row to tag it.
              Saves auto-commit to SharePoint.
            </div>
          )}
          {strictEntityFilter && untaggedCahpCount > 0 && (
            <div className="px-4 py-2 bg-amber-50 border-t border-amber-200 text-xs text-amber-900 flex items-start gap-2">
              <Icon name="alert" size={12} className="text-amber-700 flex-shrink-0 mt-0.5" />
              <span>
                <strong>{untaggedCahpCount}</strong> doc{untaggedCahpCount === 1 ? '' : 's'} in the CAHP Entity Documents library
                {' '}{untaggedCahpCount === 1 ? "isn't" : "aren't"} tagged to a specific entity yet.
                Open the library in SharePoint and set the <strong>Owner</strong> column on each file so it surfaces here.
              </span>
            </div>
          )}
          {docs.length === 0 ? (
        <div className="px-4 py-6 text-center">
          <p className="text-xs text-gray-500 mb-2">
            No documents tagged to {primaryOwnerTitle ?? 'this entity'} yet.
          </p>
          {uploadOwnerId && (
            <button
              onClick={() => setUploadOpen(true)}
              className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
            >
              <Icon name="plus" size={12} />
              Upload First Document
            </button>
          )}
        </div>
      ) : (
        <div className="divide-y divide-gray-100">
          {Array.from(byLibrary.entries()).map(([library, libraryDocs]) => (
            <div key={library} className="px-4 py-2">
              <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider mb-1">
                {library} ({libraryDocs.length})
              </div>
              <ul className="space-y-1">
                {libraryDocs.map((d) => {
                  const isCahpLibDoc = d.library === CAHP_ENTITY_LIBRARY;
                  const showTagger = manageMode && isCahpLibDoc;
                  const handleDelete = async (e: React.MouseEvent) => {
                    e.preventDefault();
                    e.stopPropagation();
                    const ok = window.confirm(
                      `Permanently delete "${d.filename}"?\n\nThis removes the file from SharePoint and cannot be undone. Action is logged.`
                    );
                    if (!ok) return;
                    try {
                      await deleteListItem(d.library as typeof PROPERTY_LINKED_LIBRARIES[number] | typeof CAHP_ENTITY_LIBRARY, d.itemId);
                      // Refetch the appropriate library
                      if (d.library === CAHP_ENTITY_LIBRARY) {
                        await cahpLib.refetch?.();
                      } else {
                        const idx = PROPERTY_LINKED_LIBRARIES.indexOf(d.library as typeof PROPERTY_LINKED_LIBRARIES[number]);
                        if (idx >= 0) await libraries[idx]?.refetch?.();
                      }
                    } catch (err) {
                      alert('Failed to delete: ' + (err instanceof Error ? err.message : String(err)));
                    }
                  };
                  return (
                  <li key={d.id} className="text-xs flex items-center justify-between gap-2 py-1">
                    <a
                      href={d.webUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="flex-1 min-w-0 text-teal-700 hover:text-teal-900 hover:underline truncate flex items-center gap-1.5"
                    >
                      <Icon name="file" size={11} />
                      <span className="truncate">{d.filename}</span>
                    </a>
                    {showTagger ? (
                      <div className="flex items-center gap-1.5 flex-shrink-0">
                        <select
                          value={d.currentOwnerTag ?? ''}
                          onChange={(e) => handleTagChange(d, e.target.value)}
                          disabled={savingTagId === d.id}
                          className="px-2 py-1 border border-gray-300 rounded text-[11px] bg-white focus:outline-none focus:border-teal-500 disabled:opacity-50 max-w-[200px]"
                        >
                          <option value="">— Untagged —</option>
                          {cahpOwnerOptions.map((o) => (
                            <option key={o.id} value={String(o.id)}>
                              {o.fields.Title}
                            </option>
                          ))}
                        </select>
                        {savingTagId === d.id && (
                          <div className="w-3 h-3 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
                        )}
                      </div>
                    ) : (
                    <span className="text-gray-400 font-mono-data text-[10px] flex-shrink-0">
                      {d.uploadDate ? formatDateET(d.uploadDate) : ''}
                    </span>
                    )}
                    <button
                      onClick={handleDelete}
                      className="text-gray-400 hover:text-error flex-shrink-0 p-1 rounded hover:bg-red-50"
                      title="Delete this document permanently"
                    >
                      <Icon name="alert" size={11} />
                    </button>
                  </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </div>
      )}
        </>
      )}

      {uploadOpen && uploadOwnerId && (
        <UploadDocumentModal
          scope={{ type: 'owner', ownerId: uploadOwnerId, ownerTitle: primaryOwnerTitle }}
          onSuccess={() => {
            setUploadOpen(false);
            refetchAll();
          }}
          onClose={() => setUploadOpen(false)}
        />
      )}
    </div>
  );
}
