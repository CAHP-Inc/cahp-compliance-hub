import { useState, useMemo } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Owner,
  type Ownership,
  type Submittal,
  type ItemCategory,
  type ItemStatus,
  type ItemPriority,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES } from './UploadDocumentModal';
import type { PropertyLinkedLibrary } from './UploadDocumentModal';
import { DOR_FILING_CHECKLIST } from '../lib/filing-checklist';
import type { FilingChecklistItem } from '../lib/filing-checklist';
import { Icon } from './ui/Icon';

interface DocItemRaw {
  id: string;
  webUrl?: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    PropertyLookupId?: string;
    OwnerLookupId?: string;
  };
}

const CATEGORY_TO_LIBRARY: Partial<Record<ItemCategory, PropertyLinkedLibrary>> = {
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

interface PreviewItem {
  template: FilingChecklistItem;
  matched: boolean;
  matchedDocFilename?: string;
  matchedDocUrl?: string;
  matchedDocLibrary?: string;
}

export interface FilingChecklistGeneratorProps {
  /** Submittal context — when provided, items link back via RelatedSubmittalLookupId */
  submittal?: Submittal;
  /** Property ID — required. Derived from submittal if a submittal is provided. */
  propertyId: string;
  /** Display title for the property — for context in the modal */
  propertyTitle?: string;
  onClose: () => void;
  onSuccess: (createdCount: number, matchedCount: number) => void;
}

export function FilingChecklistGenerator({ submittal, propertyId, propertyTitle, onClose, onSuccess }: FilingChecklistGeneratorProps) {

  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });

  // Fetch all 8 libraries to auto-match
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });
  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];

  const loading = libraries.some((l) => l.loading) || owners.loading || ownership.loading;

  const [creating, setCreating] = useState(false);
  const [createOnlyUnmatched, setCreateOnlyUnmatched] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Find CAHP entity IDs
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

  // Find property's direct-owner entity IDs (LLCs that own this property)
  const propertyOwnerIds = useMemo(() => {
    if (!propertyId || !ownership.data) return new Set<string>();
    return new Set(
      ownership.data
        .filter((rel) => String(rel.fields.LinkedPropertyLookupId) === propertyId)
        .map((rel) => rel.fields.OwnerLookupId ? String(rel.fields.OwnerLookupId) : null)
        .filter((id): id is string => id !== null)
    );
  }, [propertyId, ownership.data]);

  // Build preview by walking the template and looking for matches
  const preview = useMemo(() => {
    const items: PreviewItem[] = [];
    DOR_FILING_CHECKLIST.forEach((template) => {
      const targetLibrary = CATEGORY_TO_LIBRARY[template.category];
      let matched = false;
      let matchedDoc: { filename: string; url: string; library: string } | undefined;

      if (targetLibrary) {
        const libIdx = PROPERTY_LINKED_LIBRARIES.indexOf(targetLibrary);
        const lib = libraries[libIdx];
        if (lib?.data) {
          // Filter scope
          const candidates = lib.data.filter((doc) => {
            const propTag = doc.fields.PropertyLookupId ? String(doc.fields.PropertyLookupId) : null;
            const ownerTag = doc.fields.OwnerLookupId ? String(doc.fields.OwnerLookupId) : null;
            if (template.scope === 'property') return propTag === propertyId;
            if (template.scope === 'cahp') return ownerTag !== null && cahpOwnerIds.has(ownerTag);
            if (template.scope === 'owner') return ownerTag !== null && propertyOwnerIds.has(ownerTag);
            return false;
          });
          if (candidates.length > 0 && candidates[0].webUrl) {
            matched = true;
            matchedDoc = {
              filename: candidates[0].fields.FileLeafRef || candidates[0].fields.Title || '(unnamed)',
              url: candidates[0].webUrl,
              library: targetLibrary,
            };
          }
        }
      }

      items.push({
        template,
        matched,
        matchedDocFilename: matchedDoc?.filename,
        matchedDocUrl: matchedDoc?.url,
        matchedDocLibrary: matchedDoc?.library,
      });
    });
    return items;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data, propertyId, cahpOwnerIds, propertyOwnerIds]);

  const matchedCount = preview.filter((p) => p.matched).length;
  const unmatchedCount = preview.filter((p) => !p.matched).length;

  const handleGenerate = async () => {
    if (!propertyId) {
      setError("No property context — can't create checklist items.");
      return;
    }
    setCreating(true);
    setError(null);

    const itemsToCreate = createOnlyUnmatched ? preview.filter((p) => !p.matched) : preview;
    let createdSuccess = 0;
    let matchedSuccess = 0;

    try {
      for (const p of itemsToCreate) {
        const fields: Record<string, unknown> = {
          Title: p.template.title,
          PropertyLookupId: propertyId,
          ItemCategory: p.template.category,
          ItemStatus: (p.matched ? 'Received' : 'Not Started') as ItemStatus,
          Priority: 'High' as ItemPriority,
          DateRequested: new Date().toISOString(),
          ItemNotes: p.template.notes ?? '',
        };
        if (submittal) {
          fields.RelatedSubmittalLookupId = submittal.id;
        }
        if (p.matched) {
          fields.RelatedDocUrl = p.matchedDocUrl;
          fields.RelatedDocFilename = p.matchedDocFilename;
          fields.RelatedDocLibrary = p.matchedDocLibrary;
          fields.DateReceivedItem = new Date().toISOString();
          matchedSuccess++;
        }
        await createListItem(LIST_NAMES.Outstanding, fields);
        createdSuccess++;
      }
      // Also count auto-matches that we skipped (when createOnlyUnmatched=true)
      if (createOnlyUnmatched) {
        // matchedCount items were skipped because they already exist
        onSuccess(createdSuccess, matchedCount);
      } else {
        onSuccess(createdSuccess, matchedSuccess);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setCreating(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Generate Filing Checklist</h3>
        <p className="text-sm text-gray-600 mb-4">
          {submittal
            ? <>Creates Outstanding Items for the 12 documents DOR requires for this submittal{propertyTitle && <> on <strong>{propertyTitle}</strong></>}.</>
            : <>Backfills the 12 DOR-required documents as Outstanding Items{propertyTitle && <> for <strong>{propertyTitle}</strong></>}. Use this to clean up properties that don't have an active submittal yet.</>
          } Already-uploaded docs at the CAHP entity, property-owner entity, and this property are auto-matched and pre-linked.
        </p>

        {loading ? (
          <div className="py-8 text-center text-sm text-gray-500">
            <div className="inline-flex items-center gap-2">
              <div className="w-3 h-3 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
              Scanning libraries for existing docs…
            </div>
          </div>
        ) : (
          <>
            {/* Summary bar */}
            <div className="grid grid-cols-3 gap-3 mb-4">
              <div className="bg-green-50 border border-green-200 rounded p-3 text-center">
                <div className="text-[10px] font-semibold text-green-900 uppercase tracking-wider">Auto-matched</div>
                <div className="text-2xl font-bold text-success mt-1">{matchedCount}</div>
              </div>
              <div className="bg-amber-50 border border-amber-200 rounded p-3 text-center">
                <div className="text-[10px] font-semibold text-amber-900 uppercase tracking-wider">Need to obtain</div>
                <div className="text-2xl font-bold text-warning mt-1">{unmatchedCount}</div>
              </div>
              <div className="bg-gray-50 border border-gray-200 rounded p-3 text-center">
                <div className="text-[10px] font-semibold text-gray-700 uppercase tracking-wider">Total in template</div>
                <div className="text-2xl font-bold text-gray-700 mt-1">{preview.length}</div>
              </div>
            </div>

            {/* Preview list */}
            <div className="border border-gray-200 rounded-md max-h-96 overflow-y-auto mb-4">
              <table className="w-full text-xs">
                <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-semibold text-gray-600 uppercase tracking-wider sticky top-0">
                  <tr>
                    <th className="px-3 py-2 text-left w-8"></th>
                    <th className="px-3 py-2 text-left">Required Document</th>
                    <th className="px-3 py-2 text-left">Scope</th>
                    <th className="px-3 py-2 text-left">Match</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {preview.map((p, idx) => {
                    const scopeBadge =
                      p.template.scope === 'cahp' ? 'bg-gold-100 text-gold-900' :
                      p.template.scope === 'owner' ? 'bg-blue-100 text-blue-800' :
                      'bg-teal-100 text-teal-800';
                    const scopeLabel =
                      p.template.scope === 'cahp' ? 'CAHP Entity' :
                      p.template.scope === 'owner' ? 'Property Owner' :
                      'Property';
                    return (
                      <tr key={idx} className={p.matched ? 'bg-green-50/40' : ''}>
                        <td className="px-3 py-2 text-center">
                          {p.matched ? (
                            <Icon name="check" size={12} className="text-success" />
                          ) : (
                            <Icon name="alert" size={12} className="text-warning" />
                          )}
                        </td>
                        <td className="px-3 py-2">
                          <div className="font-medium text-gray-900">{p.template.title}</div>
                          <div className="text-[10px] text-gray-500">{p.template.category}</div>
                        </td>
                        <td className="px-3 py-2">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold ${scopeBadge}`}>
                            {scopeLabel}
                          </span>
                        </td>
                        <td className="px-3 py-2 text-[10px] text-gray-600">
                          {p.matched ? (
                            <span className="text-success truncate" title={p.matchedDocFilename}>
                              {p.matchedDocFilename}
                            </span>
                          ) : (
                            <span className="text-warning italic">missing</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <label className="flex items-center gap-2 text-xs text-gray-700 mb-3 cursor-pointer">
              <input
                type="checkbox"
                checked={createOnlyUnmatched}
                onChange={(e) => setCreateOnlyUnmatched(e.target.checked)}
              />
              Only create items for missing documents ({unmatchedCount} items). Uncheck to also create
              already-fulfilled tracking items ({preview.length} total).
            </label>

            <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900 mb-3">
              <div className="font-semibold mb-1">On Generate:</div>
              <ul className="space-y-0.5 list-disc list-inside">
                <li>{createOnlyUnmatched ? unmatchedCount : preview.length} Outstanding Items created, all linked to this submittal</li>
                <li>Auto-matched items pre-linked to their fulfilling docs and marked Received</li>
                <li>Unmatched items show on the property's Outstanding tab with Link/Upload action ready</li>
                <li>All actions audit-logged</li>
              </ul>
            </div>
          </>
        )}

        {error && (
          <div className="mb-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">{error}</div>
        )}

        <div className="flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={creating}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleGenerate}
            disabled={creating || loading}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
          >
            {creating && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {creating ? 'Generating…' : `Generate ${createOnlyUnmatched ? unmatchedCount : preview.length} Items`}
          </button>
        </div>
      </div>
    </div>
  );
}
