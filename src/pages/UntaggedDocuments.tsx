import { useState, useMemo } from 'react';
import {
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type Property,
  type Owner,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { PROPERTY_LINKED_LIBRARIES } from '../components/UploadDocumentModal';

interface UntaggedDoc {
  id: string;             // Unique key: library:itemId
  itemId: string;
  library: string;
  filename: string;
  webUrl?: string;
  uploadDate?: string;
  uploader?: string;
  size?: number;
}

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

type TagMode = 'property' | 'owner';

export function UntaggedDocuments() {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  // Fetch all 8 libraries in parallel
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });

  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];
  const loading = libraries.some((l) => l.loading) || properties.loading || owners.loading;
  const errors = libraries.filter((l) => l.error).map((l) => l.error!.message);

  // Filters / state
  const [search, setSearch] = useState('');
  const [libraryFilter, setLibraryFilter] = useState<string>('All');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [tagMode, setTagMode] = useState<TagMode>('property');
  const [bulkPropertyId, setBulkPropertyId] = useState<string>('');
  const [bulkOwnerId, setBulkOwnerId] = useState<string>('');
  const [applying, setApplying] = useState(false);
  const [applyError, setApplyError] = useState<string | null>(null);
  const [appliedCount, setAppliedCount] = useState<number | null>(null);

  const untagged = useMemo(() => {
    const docs: UntaggedDoc[] = [];
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((item) => {
        const propertyTag = item.fields.PropertyLookupId;
        const ownerTag = item.fields.OwnerLookupId;
        const noProperty = !propertyTag || propertyTag === '' || propertyTag === '0';
        const noOwner = !ownerTag || ownerTag === '' || ownerTag === '0';
        if (!noProperty || !noOwner) return; // tagged to at least one — skip
        docs.push({
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
    return docs.sort((a, b) => {
      const da = a.uploadDate ? new Date(a.uploadDate).getTime() : 0;
      const db = b.uploadDate ? new Date(b.uploadDate).getTime() : 0;
      return db - da;
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data]);

  const filtered = useMemo(() => {
    return untagged.filter((doc) => {
      if (search && !doc.filename.toLowerCase().includes(search.toLowerCase())) return false;
      if (libraryFilter !== 'All' && doc.library !== libraryFilter) return false;
      return true;
    });
  }, [untagged, search, libraryFilter]);

  // Stats per library
  const perLibraryCounts = useMemo(() => {
    const counts: Record<string, number> = {};
    untagged.forEach((d) => {
      counts[d.library] = (counts[d.library] ?? 0) + 1;
    });
    return counts;
  }, [untagged]);

  const sortedProperties = useMemo(() => {
    if (!properties.data) return [];
    return [...properties.data].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [properties.data]);

  const sortedOwners = useMemo(() => {
    if (!owners.data) return [];
    return [...owners.data].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [owners.data]);

  const refetchAll = () => libraries.forEach((l) => l.refetch?.());

  // Selection
  const toggleSelect = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAllFiltered = () => {
    setSelectedIds(new Set(filtered.map((d) => d.id)));
  };

  const clearSelection = () => {
    setSelectedIds(new Set());
  };

  // Bulk-tag — works in either mode
  const applyBulkTag = async () => {
    const targetId = tagMode === 'property' ? bulkPropertyId : bulkOwnerId;
    if (!targetId || selectedIds.size === 0) return;

    setApplying(true);
    setApplyError(null);
    setAppliedCount(null);

    const itemsToUpdate = filtered.filter((d) => selectedIds.has(d.id));
    let succeeded = 0;
    const failures: string[] = [];
    const queue = [...itemsToUpdate];
    const concurrency = 5;

    async function worker() {
      while (queue.length > 0) {
        const doc = queue.shift();
        if (!doc) break;
        try {
          const metadata =
            tagMode === 'property'
              ? { PropertyLookupId: targetId }
              : { OwnerLookupId: targetId };
          await updateListItem(doc.library, doc.itemId, metadata);
          succeeded++;
        } catch (e) {
          failures.push(`${doc.filename}: ${e instanceof Error ? e.message : String(e)}`);
        }
      }
    }

    await Promise.all(Array.from({ length: concurrency }, () => worker()));

    setApplying(false);
    setAppliedCount(succeeded);
    if (failures.length > 0) {
      setApplyError(`${failures.length} item${failures.length === 1 ? '' : 's'} failed: ${failures.slice(0, 3).join('; ')}${failures.length > 3 ? '…' : ''}`);
    }
    clearSelection();
    setBulkPropertyId('');
    setBulkOwnerId('');
    await refetchAll();
  };

  // Single-row tag
  const applySingleTag = async (doc: UntaggedDoc, propertyId: string) => {
    setApplying(true);
    setApplyError(null);
    try {
      await updateListItem(doc.library, doc.itemId, { PropertyLookupId: propertyId });
      await refetchAll();
    } catch (e) {
      setApplyError(e instanceof Error ? e.message : String(e));
    } finally {
      setApplying(false);
    }
  };

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Untagged Documents</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Scanning {PROPERTY_LINKED_LIBRARIES.length} libraries for untagged files…</span>
          </div>
        </div>
      </div>
    );
  }

  if (errors.length > 0) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">Untagged Documents</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to scan one or more libraries</div>
          <ul className="text-sm text-red-700 font-mono-data text-xs space-y-1">
            {errors.map((e, i) => <li key={i}>{e}</li>)}
          </ul>
        </div>
      </div>
    );
  }

  const allFilteredSelected = filtered.length > 0 && filtered.every((d) => selectedIds.has(d.id));

  return (
    <div>
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Untagged Documents</h1>
          <p className="text-sm text-gray-500 mt-1">
            Files in SharePoint libraries that lack the PropertyID metadata tag.
            Typically from bulk imports or direct SharePoint uploads outside the app. Tag them to associate with a property.
          </p>
        </div>
      </div>

      {/* KPI cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Total Untagged" value={untagged.length} accent={untagged.length > 0 ? 'warning' : 'default'} />
        <KPI label="Libraries Affected" value={Object.keys(perLibraryCounts).length} />
        <KPI label="Filtered" value={filtered.length} />
        <KPI label="Selected" value={selectedIds.size} accent={selectedIds.size > 0 ? 'success' : 'default'} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search filename…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={libraryFilter}
          onChange={(e) => setLibraryFilter(e.target.value)}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All libraries</option>
          {PROPERTY_LINKED_LIBRARIES.map((lib) => (
            <option key={lib} value={lib}>
              {lib} ({perLibraryCounts[lib] ?? 0})
            </option>
          ))}
        </select>
        {filtered.length !== untagged.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {untagged.length}</span>
        )}
      </div>

      {/* Result banners */}
      {appliedCount !== null && (
        <div className="mb-4 bg-green-50 border border-green-200 rounded-md p-3 flex items-start gap-2">
          <Icon name="check" size={14} className="text-success flex-shrink-0 mt-0.5" />
          <p className="text-sm text-green-900">
            Tagged <strong>{appliedCount}</strong> document{appliedCount === 1 ? '' : 's'}. List refreshed.
          </p>
        </div>
      )}
      {applyError && (
        <div className="mb-4 bg-red-50 border border-red-200 rounded-md p-3 text-sm text-red-800">
          {applyError}
        </div>
      )}

      {/* Bulk-action bar */}
      {selectedIds.size > 0 && (
        <div className="mb-4 bg-teal-50 border border-teal-200 rounded-lg p-3 shadow-card">
          <div className="flex items-center gap-3 flex-wrap">
            <span className="text-sm font-medium text-teal-900">
              {selectedIds.size} selected
            </span>
            <div className="inline-flex bg-white border border-teal-300 rounded overflow-hidden">
              <button
                onClick={() => setTagMode('property')}
                disabled={applying}
                className={`px-2.5 py-1 text-xs font-medium ${tagMode === 'property' ? 'bg-teal-700 text-white' : 'text-teal-700 hover:bg-teal-50'} disabled:opacity-50`}
              >
                Tag to Property
              </button>
              <button
                onClick={() => setTagMode('owner')}
                disabled={applying}
                className={`px-2.5 py-1 text-xs font-medium border-l border-teal-300 ${tagMode === 'owner' ? 'bg-teal-700 text-white' : 'text-teal-700 hover:bg-teal-50'} disabled:opacity-50`}
              >
                Tag to Owner / Entity
              </button>
            </div>
            {tagMode === 'property' ? (
              <select
                value={bulkPropertyId}
                onChange={(e) => setBulkPropertyId(e.target.value)}
                disabled={applying}
                className="flex-1 min-w-[200px] px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              >
                <option value="">— pick property —</option>
                {sortedProperties.map((p) => (
                  <option key={p.id} value={p.id}>{p.fields.Title}</option>
                ))}
              </select>
            ) : (
              <select
                value={bulkOwnerId}
                onChange={(e) => setBulkOwnerId(e.target.value)}
                disabled={applying}
                className="flex-1 min-w-[200px] px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
              >
                <option value="">— pick owner / entity —</option>
                {sortedOwners.map((o) => (
                  <option key={o.id} value={o.id}>{o.fields.Title}</option>
                ))}
              </select>
            )}
            <button
              onClick={applyBulkTag}
              disabled={applying || (tagMode === 'property' ? !bulkPropertyId : !bulkOwnerId)}
              className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {applying && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
              {applying ? `Tagging ${selectedIds.size}…` : `Tag ${selectedIds.size}`}
            </button>
            <button
              onClick={clearSelection}
              disabled={applying}
              className="text-xs text-gray-600 hover:text-gray-900 underline disabled:opacity-50"
            >
              Clear
            </button>
          </div>
        </div>
      )}

      {/* Empty state */}
      {untagged.length === 0 ? (
        <div className="bg-green-50 border border-green-200 rounded-lg p-8 text-center">
          <Icon name="check" size={32} className="text-success mx-auto mb-2" />
          <p className="text-base font-semibold text-green-900 mb-1">All documents tagged</p>
          <p className="text-sm text-green-800">
            Every file across the {PROPERTY_LINKED_LIBRARIES.length} property-linked libraries has a PropertyID or OwnerID set. Nothing to clean up.
          </p>
        </div>
      ) : filtered.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500">No untagged documents match your filters.</p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-3 py-3 w-10">
                  <input
                    type="checkbox"
                    checked={allFilteredSelected}
                    onChange={(e) => {
                      if (e.target.checked) selectAllFiltered();
                      else clearSelection();
                    }}
                  />
                </th>
                <th className="px-4 py-3 text-left">Filename</th>
                <th className="px-4 py-3 text-left">Library</th>
                <th className="px-4 py-3 text-left">Modified</th>
                <th className="px-4 py-3 text-left">Editor</th>
                <th className="px-4 py-3 text-left">Tag as property…</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((doc) => (
                <tr key={doc.id} className={selectedIds.has(doc.id) ? 'bg-teal-50' : 'hover:bg-gray-50'}>
                  <td className="px-3 py-3">
                    <input
                      type="checkbox"
                      checked={selectedIds.has(doc.id)}
                      onChange={() => toggleSelect(doc.id)}
                    />
                  </td>
                  <td className="px-4 py-3 font-medium text-gray-900">
                    {doc.webUrl ? (
                      <a
                        href={doc.webUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-teal-700 hover:text-teal-900 underline"
                      >
                        {doc.filename}
                      </a>
                    ) : doc.filename}
                    {doc.size != null && (
                      <span className="ml-2 text-[10px] text-gray-400 font-mono-data">{formatFileSize(doc.size)}</span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-xs text-gray-700">{doc.library}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                    {doc.uploadDate ? new Date(doc.uploadDate).toLocaleDateString() : '—'}
                  </td>
                  <td className="px-4 py-3 text-xs text-gray-600">{doc.uploader ?? '—'}</td>
                  <td className="px-4 py-3">
                    <select
                      onChange={(e) => {
                        if (e.target.value) applySingleTag(doc, e.target.value);
                        e.target.value = '';  // reset selector after triggering
                      }}
                      disabled={applying}
                      className="px-2 py-1 border border-gray-200 rounded text-xs bg-white max-w-[180px] focus:outline-none focus:border-teal-500 disabled:opacity-50"
                    >
                      <option value="">— pick —</option>
                      {sortedProperties.map((p) => (
                        <option key={p.id} value={p.id}>{p.fields.Title}</option>
                      ))}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function KPI({
  label,
  value,
  accent = 'default',
}: {
  label: string;
  value: number;
  accent?: 'default' | 'warning' | 'danger' | 'success';
}) {
  const accentClass =
    accent === 'danger' ? 'text-error' :
    accent === 'warning' ? 'text-warning' :
    accent === 'success' ? 'text-success' :
    'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}
