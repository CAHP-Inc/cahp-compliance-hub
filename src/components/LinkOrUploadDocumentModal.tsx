import { useState, useMemo } from 'react';
import {
  useSharePointList,
  useSharePointItem,
  updateListItem,
  LIST_NAMES,
  getUpstreamOwnerIds,
  type Owner,
  type Ownership,
  type Property,
  type OutstandingItem,
  type ItemStatus,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY, UploadDocumentModal } from './UploadDocumentModal';
import type { PropertyLinkedLibrary } from './UploadDocumentModal';
import { Icon } from './ui/Icon';
import { formatDateET } from '../lib/dates';

interface DocItemRaw {
  id: string;
  webUrl?: string;
  lastModifiedDateTime: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    PropertyLookupId?: string;
    OwnerLookupId?: string;
    Modified?: string;
    Editor?: { LookupValue?: string };
    File_x0020_Size?: string;
  };
}

interface AvailableDoc {
  id: string;             // library:itemId
  library: PropertyLinkedLibrary | typeof CAHP_ENTITY_LIBRARY;
  filename: string;
  webUrl: string;
  uploadDate?: string;
  /**
   * Where this document came from relative to the outstanding item:
   *  - 'this-property' : tagged directly to the item's property
   *  - 'upstream-owner': tagged to any owner anywhere in the property's
   *                      ownership chain (direct LLC, parent nonprofit, etc.)
   *  - 'cahp-entity'   : in the dedicated CAHP Entity Documents library
   *  - 'owner'         : tagged to some other (unrelated) owner — shown as a
   *                      catch-all for cross-property reference docs
   */
  scope: 'this-property' | 'upstream-owner' | 'cahp-entity' | 'owner';
  scopeLabel: string;
}

/**
 * Category → suggested library mapping. The Link picker auto-filters when an
 * item has a category set, and Upload pre-selects the right library.
 */
const CATEGORY_TO_LIBRARY: Partial<Record<string, PropertyLinkedLibrary>> = {
  'Operating Agreement': 'Operating Agreements',
  'Articles of Incorporation': 'Supporting Documentation',
  'EIN Confirmation': 'Supporting Documentation',
  'Certificate of Existence': 'Supporting Documentation',
  'Certificate of Authorization': 'Supporting Documentation',
  '501(c)(3) Determination': 'Supporting Documentation',
  'Determination Letter': 'Supporting Documentation',
  'Deed': 'Property Deeds',
  'Rent Roll': 'Supporting Documentation',
  'LURA': 'Land Use Restriction Agreements',
  'AMI Certification': 'AMI Certification Letters',
  'Org Chart': 'Org Charts',
  'Income Documentation': 'Supporting Documentation',
  'Signed Submittal': 'DOR Submittal Packages',
};

export interface LinkOrUploadDocumentModalProps {
  item: OutstandingItem;
  onClose: () => void;
  onSuccess: () => void;
}

export function LinkOrUploadDocumentModal({ item, onClose, onSuccess }: LinkOrUploadDocumentModalProps) {
  const [mode, setMode] = useState<'link' | 'upload'>('link');
  const [uploadSubmodalOpen, setUploadSubmodalOpen] = useState(false);
  const [filterLibrary, setFilterLibrary] = useState<PropertyLinkedLibrary | typeof CAHP_ENTITY_LIBRARY | 'All'>('All');
  const [search, setSearch] = useState('');
  const [linking, setLinking] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [autoMarkReceived, setAutoMarkReceived] = useState(true);

  const propertyId = item.fields.PropertyLookupId ? String(item.fields.PropertyLookupId) : null;
  const itemCategory = item.fields.ItemCategory;
  const suggestedLibrary = itemCategory ? CATEGORY_TO_LIBRARY[itemCategory] : undefined;
  // Fetch the property only to validate that propertyId resolves to a real record;
  // the upstream-owner walk handles CAHP entity filtering without needing state.
  useSharePointItem<Property>(LIST_NAMES.Properties, propertyId ?? undefined);

  // Resolve every owner upstream of this property via the ownership chain.
  // No state heuristic — we just walk the actual relationships, so a
  // property's NC subsidiary AND its SC parent nonprofit both qualify.
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const upstreamOwnerIds = useMemo(() => {
    if (!propertyId || !ownership.data) return new Set<string>();
    return getUpstreamOwnerIds(propertyId, ownership.data);
  }, [propertyId, ownership.data]);
  // Also keep a CAHP-name detector for the dedicated CAHP Entity Documents
  // library — that library holds *shared reference* material and isn't keyed
  // to any specific property, so we still surface untagged docs from it.
  const isCahpEntityName = (title: string | undefined) => {
    const t = (title ?? '').toLowerCase();
    return t.includes('cahp') || t.includes('carolina affordable housing project');
  };

  // Multi-value Lookup readback: Graph returns the value under the base name
  // ('Property', 'Owner') as an array of {LookupId, LookupValue}, or (legacy
  // single-value) under the *LookupId suffix as a scalar. Both shapes coexist
  // because the 8 document libraries were migrated mid-flight; older docs
  // still surface via the scalar. Without this, multi-value-tagged docs look
  // untagged and the picker shows only the CAHP shared docs.
  const extractTagIds = (
    fields: Record<string, unknown> | undefined,
    base: 'Owner' | 'Property',
  ): string[] => {
    if (!fields) return [];
    const ids: string[] = [];
    const scalar = fields[`${base}LookupId`];
    if (typeof scalar === 'string' || typeof scalar === 'number') {
      const s = String(scalar);
      if (s && s !== '0') ids.push(s);
    } else if (Array.isArray(scalar)) {
      for (const v of scalar) ids.push(String(v));
    }
    const arr = fields[base];
    if (Array.isArray(arr)) {
      for (const entry of arr) {
        if (entry && typeof entry === 'object' && 'LookupId' in entry) {
          ids.push(String((entry as { LookupId: unknown }).LookupId));
        }
      }
    }
    return ids;
  };

  // Fetch all 8 libraries
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

  const loading = libraries.some((l) => l.loading) || owners.loading || ownership.loading || cahpLib.loading;

  // Aggregate documents relevant to this item:
  // - tagged to this property
  // - tagged to any owner upstream of this property in the ownership chain
  //   (direct LLC + every parent / nonprofit above it)
  // - in the dedicated CAHP Entity Documents library (shared reference docs)
  const availableDocs = useMemo(() => {
    const docs: AvailableDoc[] = [];
    const ownerNameById = new Map<string, string>();
    (owners.data ?? []).forEach((o) => {
      if (o.fields.Title) ownerNameById.set(String(o.id), o.fields.Title);
    });

    // 1. CAHP Entity Documents library — show untagged shared docs +
    //    any docs tagged to an upstream owner OR a CAHP-named entity.
    if (cahpLib.data) {
      cahpLib.data.forEach((doc) => {
        if (!doc.webUrl) return;
        const ownerIds = extractTagIds(doc.fields as unknown as Record<string, unknown>, 'Owner');
        let scope: AvailableDoc['scope'] = 'cahp-entity';
        let scopeLabel = 'CAHP Entity (shared)';
        if (ownerIds.length > 0) {
          const upstreamMatch = ownerIds.find((id) => upstreamOwnerIds.has(id));
          if (upstreamMatch) {
            scope = 'upstream-owner';
            scopeLabel = ownerNameById.get(upstreamMatch) ?? 'Upstream owner';
          } else {
            // Not an upstream owner — only include if any tag looks like a CAHP entity
            const cahpMatch = ownerIds.find((id) => isCahpEntityName(ownerNameById.get(id)));
            if (!cahpMatch) return;
            scopeLabel = ownerNameById.get(cahpMatch) ?? 'CAHP Entity';
          }
        }
        docs.push({
          id: `${CAHP_ENTITY_LIBRARY}:${doc.id}`,
          library: CAHP_ENTITY_LIBRARY,
          filename: doc.fields.FileLeafRef || doc.fields.Title || '(unnamed)',
          webUrl: doc.webUrl,
          uploadDate: doc.fields.Modified || doc.lastModifiedDateTime,
          scope,
          scopeLabel,
        });
      });
    }

    // 2. The 8 property-linked libraries (scoped by property or owner tag)
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((doc) => {
        const propIds = extractTagIds(doc.fields as unknown as Record<string, unknown>, 'Property');
        const ownerIds = extractTagIds(doc.fields as unknown as Record<string, unknown>, 'Owner');

        let scope: AvailableDoc['scope'] | null = null;
        let scopeLabel = '';

        if (propertyId && propIds.includes(propertyId)) {
          scope = 'this-property';
          scopeLabel = 'This Property';
        } else if (propIds.length > 0) {
          // Doc is explicitly scoped to one or more other properties — don't
          // fall through to the owner-chain bucket, otherwise sibling
          // properties under the same parent (e.g. all IV Fund Global SFRs)
          // would pull each other's property-specific docs into this picker.
          return;
        } else if (ownerIds.length > 0) {
          const upstreamMatch = ownerIds.find((id) => upstreamOwnerIds.has(id));
          const cahpMatch = !upstreamMatch
            ? ownerIds.find((id) => isCahpEntityName(ownerNameById.get(id)))
            : undefined;
          if (upstreamMatch) {
            scope = 'upstream-owner';
            scopeLabel = ownerNameById.get(upstreamMatch) ?? 'Upstream owner';
          } else if (cahpMatch) {
            // Catch CAHP-named entities even if they're not in this property's chain
            // (helps when ownership rows haven't been fully wired up yet).
            scope = 'cahp-entity';
            scopeLabel = ownerNameById.get(cahpMatch) ?? 'CAHP Entity';
          } else {
            scope = 'owner';
            scopeLabel = ownerNameById.get(ownerIds[0]) ?? 'Owner';
          }
        } else {
          return; // untagged — don't surface in link picker
        }

        if (!doc.webUrl) return;
        docs.push({
          id: `${libraryName}:${doc.id}`,
          library: libraryName,
          filename: doc.fields.FileLeafRef || doc.fields.Title || '(unnamed)',
          webUrl: doc.webUrl,
          uploadDate: doc.fields.Modified || doc.lastModifiedDateTime,
          scope,
          scopeLabel,
        });
      });
    });
    return docs.sort((a, b) => {
      // This property first, then upstream chain, then CAHP shared, then other owners;
      // within each group, newest first.
      const scopeOrder = { 'this-property': 0, 'upstream-owner': 1, 'cahp-entity': 2, owner: 3 };
      if (a.scope !== b.scope) return scopeOrder[a.scope] - scopeOrder[b.scope];
      const da = a.uploadDate ? new Date(a.uploadDate).getTime() : 0;
      const db = b.uploadDate ? new Date(b.uploadDate).getTime() : 0;
      return db - da;
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data, cahpLib.data, propertyId, upstreamOwnerIds, owners.data]);

  const filtered = useMemo(() => {
    let docs = availableDocs;
    if (filterLibrary !== 'All') {
      // Always include CAHP Entity Documents (reference material) alongside any selected library
      docs = docs.filter((d) => d.library === filterLibrary || d.library === CAHP_ENTITY_LIBRARY);
    } else if (suggestedLibrary) {
      // Soft-filter to suggested library when no explicit filter set
      // CAHP entity docs always show alongside the suggested library
      const suggested = docs.filter(
        (d) => d.library === suggestedLibrary || d.library === CAHP_ENTITY_LIBRARY
      );
      if (suggested.length > 0) docs = suggested;
    }
    if (search) {
      const q = search.toLowerCase();
      docs = docs.filter((d) => d.filename.toLowerCase().includes(q));
    }
    return docs;
  }, [availableDocs, filterLibrary, search, suggestedLibrary]);

  const linkDocument = async (doc: AvailableDoc) => {
    setLinking(true);
    setError(null);
    try {
      const updates: Record<string, unknown> = {
        RelatedDocUrl: doc.webUrl,
        RelatedDocFilename: doc.filename,
        RelatedDocLibrary: doc.library,
      };
      if (autoMarkReceived) {
        updates.ItemStatus = 'Received' as ItemStatus;
        updates.DateReceivedItem = new Date().toISOString();
      }
      await updateListItem(LIST_NAMES.Outstanding, item.id, updates);
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setLinking(false);
    }
  };

  const handleUploadSuccess = async (result: { filename: string; webUrl: string; library: string }) => {
    setUploadSubmodalOpen(false);
    setLinking(true);
    setError(null);
    try {
      const updates: Record<string, unknown> = {
        RelatedDocUrl: result.webUrl,
        RelatedDocFilename: result.filename,
        RelatedDocLibrary: result.library,
      };
      if (autoMarkReceived) {
        updates.ItemStatus = 'Received' as ItemStatus;
        updates.DateReceivedItem = new Date().toISOString();
      }
      await updateListItem(LIST_NAMES.Outstanding, item.id, updates);
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setLinking(false);
    }
  };

  if (uploadSubmodalOpen && propertyId) {
    return (
      <UploadDocumentModal
        scope={{ type: 'property', propertyId }}
        defaultLibrary={suggestedLibrary}
        onSuccess={handleUploadSuccess}
        onClose={() => setUploadSubmodalOpen(false)}
      />
    );
  }

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Link or Upload Document</h3>
        <p className="text-sm text-gray-600 mb-4">
          Fulfilling <strong>{item.fields.Title}</strong>
          {itemCategory && <> · category <strong>{itemCategory}</strong></>}.
          Either link an existing SharePoint document or upload a new file.
        </p>

        {/* Mode toggle */}
        <div className="inline-flex bg-gray-100 rounded-md p-0.5 mb-4">
          <button
            onClick={() => setMode('link')}
            className={`px-3 py-1.5 text-xs font-medium rounded ${mode === 'link' ? 'bg-white shadow text-teal-700' : 'text-gray-600 hover:text-gray-900'}`}
          >
            <Icon name="folder" size={12} className="inline mr-1" />
            Link Existing
          </button>
          <button
            onClick={() => setMode('upload')}
            className={`px-3 py-1.5 text-xs font-medium rounded ${mode === 'upload' ? 'bg-white shadow text-teal-700' : 'text-gray-600 hover:text-gray-900'}`}
          >
            <Icon name="plus" size={12} className="inline mr-1" />
            Upload New
          </button>
        </div>

        {mode === 'link' && (
          <>
            <div className="flex items-center gap-2 mb-3 flex-wrap">
              <div className="relative flex-1 min-w-[200px]">
                <Icon name="search" size={12} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search filename…"
                  className="w-full pl-7 pr-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
                />
              </div>
              <select
                value={filterLibrary}
                onChange={(e) => setFilterLibrary(e.target.value as PropertyLinkedLibrary | typeof CAHP_ENTITY_LIBRARY | 'All')}
                className="px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              >
                <option value="All">
                  {suggestedLibrary ? `Suggested: ${suggestedLibrary} + CAHP` : 'All libraries'}
                </option>
                {PROPERTY_LINKED_LIBRARIES.map((lib) => (
                  <option key={lib} value={lib}>{lib}</option>
                ))}
                <option value={CAHP_ENTITY_LIBRARY}>{CAHP_ENTITY_LIBRARY} (entity reference)</option>
              </select>
            </div>

            <div className="border border-gray-200 rounded-md max-h-96 overflow-y-auto">
              {loading ? (
                <div className="p-4 text-center text-xs text-gray-500">Loading documents…</div>
              ) : filtered.length === 0 ? (
                <div className="p-6 text-center">
                  <p className="text-sm text-gray-500 mb-2">
                    No matching documents.
                    {suggestedLibrary && filterLibrary === 'All' && (
                      <> Try <button onClick={() => setFilterLibrary('All')} className="text-teal-700 hover:underline">all libraries</button>, or upload a new file.</>
                    )}
                  </p>
                  <button
                    onClick={() => setMode('upload')}
                    className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
                  >
                    <Icon name="plus" size={12} />
                    Upload Instead
                  </button>
                </div>
              ) : (
                <ul className="divide-y divide-gray-100">
                  {filtered.map((doc) => {
                    const scopeBadgeClass =
                      doc.scope === 'this-property' ? 'bg-teal-100 text-teal-800' :
                      doc.scope === 'upstream-owner' ? 'bg-blue-100 text-blue-800' :
                      doc.scope === 'cahp-entity' ? 'bg-gold-100 text-gold-900' :
                      'bg-gray-100 text-gray-700';
                    return (
                      <li key={doc.id} className="px-3 py-2 hover:bg-gray-50 flex items-center gap-3">
                        <Icon name="file" size={12} className="text-gray-400 flex-shrink-0" />
                        <div className="flex-1 min-w-0">
                          <div className="text-sm text-gray-900 truncate">{doc.filename}</div>
                          <div className="text-[11px] text-gray-500 flex items-center gap-2 mt-0.5">
                            <span className={`px-1.5 py-0.5 rounded text-[9px] font-semibold ${scopeBadgeClass}`}>
                              {doc.scopeLabel}
                            </span>
                            <span>{doc.library}</span>
                            {doc.uploadDate && <span className="font-mono-data">{formatDateET(doc.uploadDate)}</span>}
                          </div>
                        </div>
                        <button
                          onClick={() => linkDocument(doc)}
                          disabled={linking}
                          className="px-2 py-1 bg-teal-700 hover:bg-teal-900 text-white rounded text-xs font-medium disabled:opacity-50 flex-shrink-0"
                        >
                          Link
                        </button>
                      </li>
                    );
                  })}
                </ul>
              )}
            </div>
          </>
        )}

        {mode === 'upload' && (
          <div className="bg-gray-50 border border-gray-200 rounded-md p-4 text-center">
            <p className="text-sm text-gray-700 mb-3">
              Upload a new file. It'll go to <strong>{suggestedLibrary ?? 'Supporting Documentation'}</strong>
              {' '}and be tagged to this property, then automatically linked to this checklist item.
            </p>
            <button
              onClick={() => setUploadSubmodalOpen(true)}
              disabled={!propertyId}
              className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md text-sm font-medium inline-flex items-center gap-2 disabled:opacity-50"
            >
              <Icon name="plus" size={14} />
              Choose File
            </button>
            {!propertyId && (
              <p className="text-xs text-red-700 mt-2">
                This item isn't linked to a property — link it to a property first to enable upload.
              </p>
            )}
          </div>
        )}

        {/* Auto-mark toggle */}
        <label className="flex items-center gap-2 text-xs text-gray-700 mt-4 cursor-pointer">
          <input
            type="checkbox"
            checked={autoMarkReceived}
            onChange={(e) => setAutoMarkReceived(e.target.checked)}
          />
          Mark item as <strong>Received</strong> when linked/uploaded
        </label>

        {error && (
          <div className="mt-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
            {error}
          </div>
        )}

        <div className="flex items-center justify-end gap-2 mt-5">
          <button
            onClick={onClose}
            disabled={linking}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
