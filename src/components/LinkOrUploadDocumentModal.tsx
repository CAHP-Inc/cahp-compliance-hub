import { useState, useMemo } from 'react';
import {
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type Owner,
  type OutstandingItem,
  type ItemStatus,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY, UploadDocumentModal } from './UploadDocumentModal';
import type { PropertyLinkedLibrary } from './UploadDocumentModal';
import { Icon } from './ui/Icon';

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
  library: PropertyLinkedLibrary;
  filename: string;
  webUrl: string;
  uploadDate?: string;
  scope: 'this-property' | 'cahp-entity' | 'owner';
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
  const [filterLibrary, setFilterLibrary] = useState<PropertyLinkedLibrary | 'All'>('All');
  const [search, setSearch] = useState('');
  const [linking, setLinking] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [autoMarkReceived, setAutoMarkReceived] = useState(true);

  const propertyId = item.fields.PropertyLookupId ? String(item.fields.PropertyLookupId) : null;
  const itemCategory = item.fields.ItemCategory;
  const suggestedLibrary = itemCategory ? CATEGORY_TO_LIBRARY[itemCategory] : undefined;

  // Identify CAHP entities for filtering
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const cahpOwnerIds = useMemo(() => {
    if (!owners.data) return new Set<string>();
    return new Set(
      owners.data
        .filter((o) => {
          const t = (o.fields.Title ?? '').toLowerCase();
          return t.includes('cahp') || t.includes('carolina affordable housing project');
        })
        .map((o) => String(o.id))
    );
  }, [owners.data]);

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

  const loading = libraries.some((l) => l.loading) || owners.loading || cahpLib.loading;

  // Aggregate documents relevant to this item:
  // - tagged to this property
  // - tagged to any CAHP entity (reference docs)
  // - in the dedicated CAHP Entity Documents library
  const availableDocs = useMemo(() => {
    const docs: AvailableDoc[] = [];

    // 1. CAHP Entity Documents library — every file is a CAHP reference doc
    if (cahpLib.data) {
      cahpLib.data.forEach((doc) => {
        if (!doc.webUrl) return;
        docs.push({
          id: `${CAHP_ENTITY_LIBRARY}:${doc.id}`,
          library: CAHP_ENTITY_LIBRARY as 'Supporting Documentation', // type narrow — treated like Supporting Documentation for downstream
          filename: doc.fields.FileLeafRef || doc.fields.Title || '(unnamed)',
          webUrl: doc.webUrl,
          uploadDate: doc.fields.Modified || doc.lastModifiedDateTime,
          scope: 'cahp-entity',
          scopeLabel: 'CAHP Entity',
        });
      });
    }

    // 2. The 8 property-linked libraries (scoped by property or owner tag)
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((doc) => {
        const propTag = doc.fields.PropertyLookupId ? String(doc.fields.PropertyLookupId) : null;
        const ownerTag = doc.fields.OwnerLookupId ? String(doc.fields.OwnerLookupId) : null;

        let scope: AvailableDoc['scope'] | null = null;
        let scopeLabel = '';

        if (propertyId && propTag === propertyId) {
          scope = 'this-property';
          scopeLabel = 'This Property';
        } else if (ownerTag && cahpOwnerIds.has(ownerTag)) {
          scope = 'cahp-entity';
          const ownerName = owners.data?.find((o) => String(o.id) === ownerTag)?.fields.Title;
          scopeLabel = ownerName ?? 'CAHP Entity';
        } else if (ownerTag) {
          scope = 'owner';
          const ownerName = owners.data?.find((o) => String(o.id) === ownerTag)?.fields.Title;
          scopeLabel = ownerName ?? 'Owner';
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
      // This property first, then CAHP, then owner; within group, newest first
      const scopeOrder = { 'this-property': 0, 'cahp-entity': 1, owner: 2 };
      if (a.scope !== b.scope) return scopeOrder[a.scope] - scopeOrder[b.scope];
      const da = a.uploadDate ? new Date(a.uploadDate).getTime() : 0;
      const db = b.uploadDate ? new Date(b.uploadDate).getTime() : 0;
      return db - da;
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data, cahpLib.data, propertyId, cahpOwnerIds, owners.data]);

  const filtered = useMemo(() => {
    let docs = availableDocs;
    if (filterLibrary !== 'All') docs = docs.filter((d) => d.library === filterLibrary);
    else if (suggestedLibrary) {
      // Soft-filter to suggested library when no explicit filter set
      const suggested = docs.filter((d) => d.library === suggestedLibrary);
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
                onChange={(e) => setFilterLibrary(e.target.value as PropertyLinkedLibrary | 'All')}
                className="px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              >
                <option value="All">
                  {suggestedLibrary ? `Suggested: ${suggestedLibrary}` : 'All libraries'}
                </option>
                {PROPERTY_LINKED_LIBRARIES.map((lib) => (
                  <option key={lib} value={lib}>{lib}</option>
                ))}
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
                            {doc.uploadDate && <span className="font-mono-data">{new Date(doc.uploadDate).toLocaleDateString()}</span>}
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
