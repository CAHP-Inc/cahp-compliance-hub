import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  deleteListItem,
  LIST_NAMES,
  type Ownership,
  type Submittal,
  type ComplianceDeadline,
  type OutstandingItem,
  type PropertyNote,
  type Correspondence,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { PROPERTY_LINKED_LIBRARIES } from './UploadDocumentModal';

interface DeletePropertyModalProps {
  propertyId: string;
  propertyTitle: string;
  onClose: () => void;
}

/**
 * Full property delete with cascade visibility.
 *
 * Type 'DELETE {property name}' to confirm. Shows counts of every related record
 * type that will be deleted. Walks structured lists (Ownership, Submittals, etc.)
 * plus the 8 property-linked document libraries.
 *
 * NOT undoable. Audit log captures each deletion.
 */
export function DeletePropertyModal({ propertyId, propertyTitle, onClose }: DeletePropertyModalProps) {
  const navigate = useNavigate();
  const [confirmText, setConfirmText] = useState('');
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<string | null>(null);

  // Pull all linked records to show counts
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const deadlines = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const notes = useSharePointList<PropertyNote>(LIST_NAMES.PropertyNotes, { top: 500 });
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });

  // Document libraries — query each with PropertyLookupId filter (lazy: we fetch all and filter client-side)
  const lib0 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<{ fields: { PropertyLookupId?: string }; id: string }>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });

  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];

  const filterByProperty = <T extends { fields: { PropertyLookupId?: string }; id: string }>(
    list: T[] | null | undefined
  ): T[] => {
    if (!list) return [];
    return list.filter((item) => String(item.fields.PropertyLookupId ?? '') === String(propertyId));
  };

  const linkedOwnership = filterByProperty(
    ownership.data?.map((o) => ({ id: o.id, fields: { PropertyLookupId: o.fields.LinkedPropertyLookupId } }))
  );
  const linkedSubmittals = filterByProperty(submittals.data);
  const linkedDeadlines = filterByProperty(deadlines.data);
  const linkedOutstanding = filterByProperty(outstanding.data);
  const linkedNotes = filterByProperty(notes.data);
  const linkedCorrespondence = filterByProperty(correspondence.data);

  const linkedDocs = libraries.map((lib, idx) => ({
    libraryName: PROPERTY_LINKED_LIBRARIES[idx],
    items: filterByProperty(lib.data),
  }));

  const totalDocs = linkedDocs.reduce((sum, l) => sum + l.items.length, 0);
  const totalDeletions =
    1 + // the property itself
    linkedOwnership.length +
    linkedSubmittals.length +
    linkedDeadlines.length +
    linkedOutstanding.length +
    linkedNotes.length +
    linkedCorrespondence.length +
    totalDocs;

  const expectedConfirm = `DELETE ${propertyTitle}`;
  const canDelete = confirmText.trim() === expectedConfirm && !deleting;

  const isLoading =
    ownership.loading || submittals.loading || deadlines.loading ||
    outstanding.loading || notes.loading || correspondence.loading ||
    libraries.some((l) => l.loading);

  const handleDelete = async () => {
    setDeleting(true);
    setError(null);

    try {
      // Delete child records first, then the property record itself
      // (avoids any FK-style constraint issues SharePoint may enforce)

      let done = 0;
      const total = totalDeletions;
      const tick = (label: string) => {
        done++;
        setProgress(`${label} (${done}/${total})`);
      };

      // Documents in each library
      for (const { libraryName, items } of linkedDocs) {
        for (const item of items) {
          await deleteListItem(libraryName, item.id);
          tick(`Deleting documents in ${libraryName}`);
        }
      }

      // Structured lists
      for (const o of linkedOwnership) {
        await deleteListItem(LIST_NAMES.Ownership, o.id);
        tick('Deleting ownership records');
      }
      for (const s of linkedSubmittals) {
        await deleteListItem(LIST_NAMES.Submittals, s.id);
        tick('Deleting submittals');
      }
      for (const d of linkedDeadlines) {
        await deleteListItem(LIST_NAMES.ComplianceDeadlines, d.id);
        tick('Deleting deadlines');
      }
      for (const oi of linkedOutstanding) {
        await deleteListItem(LIST_NAMES.Outstanding, oi.id);
        tick('Deleting outstanding items');
      }
      for (const n of linkedNotes) {
        await deleteListItem(LIST_NAMES.PropertyNotes, n.id);
        tick('Deleting notes');
      }
      for (const c of linkedCorrespondence) {
        await deleteListItem(LIST_NAMES.Correspondence, c.id);
        tick('Deleting correspondence');
      }

      // Finally the property record itself
      setProgress('Deleting property record…');
      await deleteListItem(LIST_NAMES.Properties, propertyId);

      setProgress('Done. Redirecting…');
      // Navigate to properties list after success
      setTimeout(() => navigate('/properties'), 500);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setDeleting(false);
    }
  };

  // Block accidental ESC close mid-delete
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !deleting) onClose();
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [deleting, onClose]);

  return (
    <div
      className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !deleting) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        {/* Header */}
        <div className="px-6 py-4 border-b-2 border-red-200 bg-red-50">
          <div className="flex items-start gap-3">
            <div className="w-10 h-10 rounded-full bg-error/10 flex items-center justify-center flex-shrink-0">
              <Icon name="alert" size={18} className="text-error" />
            </div>
            <div className="flex-1">
              <h2 className="text-lg font-bold text-error">Permanently delete property</h2>
              <p className="text-sm text-gray-700 mt-0.5">
                This will permanently delete <strong>{propertyTitle}</strong> and ALL related records.
                This action cannot be undone.
              </p>
            </div>
            {!deleting && (
              <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl leading-none">
                ×
              </button>
            )}
          </div>
        </div>

        <div className="px-6 py-4">
          <p className="text-xs text-gray-700 mb-3 font-semibold uppercase tracking-wider">
            What will be deleted:
          </p>

          {isLoading ? (
            <div className="text-sm text-gray-500 py-4 text-center">Counting linked records…</div>
          ) : (
            <ul className="space-y-1.5 text-sm bg-gray-50 rounded-md p-3 border border-gray-200">
              <CascadeRow label="Property record" count={1} />
              <CascadeRow label="Ownership records" count={linkedOwnership.length} />
              <CascadeRow label="Submittals" count={linkedSubmittals.length} />
              <CascadeRow label="Compliance deadlines" count={linkedDeadlines.length} />
              <CascadeRow label="Outstanding items" count={linkedOutstanding.length} />
              <CascadeRow label="Property notes" count={linkedNotes.length} />
              <CascadeRow label="Correspondence letters" count={linkedCorrespondence.length} />
              <li className="pt-2 border-t border-gray-200 mt-2">
                <div className="text-xs text-gray-600 font-semibold mb-1">Documents:</div>
                {linkedDocs.every((l) => l.items.length === 0) ? (
                  <div className="text-xs text-gray-500 italic pl-2">— No documents tagged to this property</div>
                ) : (
                  linkedDocs.map(({ libraryName, items }) => (
                    items.length > 0 && (
                      <div key={libraryName} className="text-xs text-gray-700 pl-2 flex justify-between">
                        <span>{libraryName}</span>
                        <span className="font-mono-data font-semibold">{items.length}</span>
                      </div>
                    )
                  ))
                )}
              </li>
              <li className="pt-2 border-t border-gray-300 mt-2 flex justify-between font-bold">
                <span className="text-error">TOTAL RECORDS TO DELETE</span>
                <span className="font-mono-data text-error">{totalDeletions}</span>
              </li>
            </ul>
          )}

          <p className="text-xs text-amber-800 mt-3 bg-amber-50 border border-amber-200 rounded p-2">
            <strong>Note:</strong> Owners (entities like CAHP SC LLC) are NOT deleted — they may be linked to other properties.
            CAHP entity documents are also preserved. Only records tied specifically to this property are removed.
          </p>

          {/* Confirmation input */}
          <div className="mt-4">
            <label className="block text-xs font-semibold text-gray-700 uppercase tracking-wider mb-1">
              To confirm, type: <code className="bg-gray-100 px-1.5 py-0.5 rounded text-error font-mono-data">{expectedConfirm}</code>
            </label>
            <input
              type="text"
              value={confirmText}
              onChange={(e) => setConfirmText(e.target.value)}
              disabled={deleting}
              placeholder={expectedConfirm}
              className="w-full px-3 py-2 border-2 border-gray-300 rounded text-sm focus:outline-none focus:border-error font-mono-data"
              autoFocus
            />
          </div>

          {progress && (
            <div className="mt-3 bg-blue-50 border border-blue-200 rounded p-2 text-xs text-blue-800 flex items-center gap-2">
              <div className="w-3 h-3 rounded-full border-2 border-blue-500 border-r-transparent animate-spin" />
              {progress}
            </div>
          )}

          {error && (
            <div className="mt-3 bg-red-50 border border-red-200 rounded p-3 flex items-start gap-2">
              <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
              <div className="text-xs text-error">
                <div className="font-semibold mb-0.5">Delete failed mid-process</div>
                <div className="font-mono-data">{error}</div>
                <div className="mt-1 text-gray-700">Some records may have already been deleted. You can retry or check SharePoint directly.</div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={deleting}
            className="px-4 py-2 text-sm text-gray-700 hover:bg-gray-200 rounded-md disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleDelete}
            disabled={!canDelete || isLoading}
            className="bg-error hover:bg-red-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-4 py-2 rounded-md text-sm font-bold inline-flex items-center gap-2"
          >
            {deleting ? (
              <>
                <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                Deleting…
              </>
            ) : (
              <>
                <Icon name="alert" size={14} />
                Permanently delete {totalDeletions} records
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}

function CascadeRow({ label, count }: { label: string; count: number }) {
  return (
    <li className="flex justify-between items-center text-gray-700">
      <span>{label}</span>
      <span className={`font-mono-data font-semibold ${count > 0 ? 'text-gray-900' : 'text-gray-400'}`}>
        {count}
      </span>
    </li>
  );
}
