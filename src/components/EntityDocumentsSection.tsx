import { useMemo, useState } from 'react';
import { useSharePointList } from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY, UploadDocumentModal } from './UploadDocumentModal';
import { Icon } from './ui/Icon';

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
}

export function EntityDocumentsSection({
  ownerIds,
  primaryOwnerTitle,
  title = 'Documents',
  subtitle,
  uploadOwnerId,
  variant = 'card',
  useCahpEntityLibrary = false,
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

  const ownerIdSet = useMemo(() => new Set(ownerIds.map(String)), [ownerIds]);

  const docs = useMemo(() => {
    const collected: AggregatedDoc[] = [];

    // 1. Files from the dedicated CAHP Entity Documents library
    // (only when useCahpEntityLibrary is true — every file here is a CAHP entity doc by virtue of being in this library)
    if (useCahpEntityLibrary && cahpLib.data) {
      cahpLib.data.forEach((item) => {
        collected.push({
          id: `${CAHP_ENTITY_LIBRARY}:${item.id}`,
          itemId: item.id,
          library: CAHP_ENTITY_LIBRARY,
          filename: item.fields.FileLeafRef || item.fields.Title || '(unnamed)',
          webUrl: item.webUrl,
          uploadDate: item.fields.Modified || item.lastModifiedDateTime,
          uploader: item.fields.Editor?.LookupValue,
          size: item.fields.File_x0020_Size ? parseInt(item.fields.File_x0020_Size, 10) : undefined,
        });
      });
    }

    // 2. OwnerLookupId-tagged files from the 8 property-linked libraries (for non-CAHP entities like LLCs, trusts)
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((item) => {
        const ownerTag = item.fields.OwnerLookupId;
        if (!ownerTag || !ownerIdSet.has(String(ownerTag))) return;
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
    cahpLib.data, ownerIdSet, useCahpEntityLibrary,
  ]);

  const loading = libraries.some((l) => l.loading) || (useCahpEntityLibrary && cahpLib.loading);
  const refetchAll = () => {
    libraries.forEach((l) => l.refetch?.());
    cahpLib.refetch?.();
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

      {loading ? (
        <div className="px-4 py-6 text-center text-xs text-gray-500">Loading documents…</div>
      ) : docs.length === 0 ? (
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
                {libraryDocs.map((d) => (
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
                    <span className="text-gray-400 font-mono-data text-[10px] flex-shrink-0">
                      {d.uploadDate ? new Date(d.uploadDate).toLocaleDateString() : ''}
                    </span>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
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
