import { useState, useMemo, useEffect } from 'react';
import {
  useSharePointList,
  useSharePointItem,
  createListItem,
  updateListItem,
  LIST_NAMES,
  getUpstreamOwnerIds,
  type Owner,
  type Ownership,
  type Submittal,
  type Property,
  type ItemCategory,
  type ItemStatus,
  type ItemPriority,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY } from './UploadDocumentModal';
import type { PropertyLinkedLibrary } from './UploadDocumentModal';
import { useChecklistTemplates } from '../lib/filing-checklist';
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
  /** SharePoint item ID of the matched doc — needed to PATCH its tag when generating */
  matchedDocId?: string;
  /** Whether the matched doc already has a tag (PropertyLookupId / OwnerLookupId) that lines up */
  matchedDocAlreadyTagged?: boolean;
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
  const propertyItem = useSharePointItem<Property>(LIST_NAMES.Properties, propertyId);
  const propertyState = propertyItem.data?.fields.cahpState;

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
  const cahpLib = useSharePointList<DocItemRaw>(CAHP_ENTITY_LIBRARY, { top: 500 });

  const loading = libraries.some((l) => l.loading) || owners.loading || ownership.loading || cahpLib.loading;

  const [creating, setCreating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Per-row selection: titles are unique within a template so we key by title.
  const [selectedTitles, setSelectedTitles] = useState<Set<string>>(new Set());
  // Whether to also pull matched docs into the entity by tagging them (when
  // they're missing the proper Owner/Property tag). Defaults on.
  const [tagMatchedDocs, setTagMatchedDocs] = useState(true);

  // Identify CAHP-scope owners: every entity upstream of this property in the
  // ownership chain (direct LLC, parent nonprofit, etc.), plus any CAHP-named
  // entity that hasn't been wired into the chain yet. This replaces the older
  // state-based heuristic which incorrectly hid the SC parent nonprofit from
  // NC properties (and vice versa).
  const upstreamOwnerIds = useMemo(() => {
    if (!propertyId || !ownership.data) return new Set<string>();
    return getUpstreamOwnerIds(propertyId, ownership.data);
  }, [propertyId, ownership.data]);

  const cahpOwnerIds = useMemo(() => {
    if (!owners.data) return new Set<string>();
    const ids = new Set<string>(upstreamOwnerIds);
    // Belt-and-suspenders: also include any CAHP-named entity even if it's
    // not yet linked through Ownership Structure rows. State filter dropped.
    for (const o of owners.data) {
      const t = (o.fields.Title ?? '').toLowerCase();
      if (t.includes('cahp') || t.includes('carolina affordable housing project')) {
        ids.add(String(o.id));
      }
    }
    return ids;
  }, [owners.data, upstreamOwnerIds]);

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

  // Build preview by walking the template and looking for matches.
  // Templates are loaded from the shared SharePoint list (Settings → Checklist
  // Templates), with the hardcoded DOR list as fallback when the list is
  // empty or unprovisioned. State-scoped items are filtered out when the
  // property's state doesn't match — items with no state apply to every property.
  const { templates: allTemplates } = useChecklistTemplates();
  const activeChecklist = useMemo(() => {
    if (!propertyState) return allTemplates;
    return allTemplates.filter((t) => !t.state || t.state === propertyState);
  }, [allTemplates, propertyState]);
  const preview = useMemo(() => {
    const items: PreviewItem[] = [];
    activeChecklist.forEach((template) => {
      // Per-item library override wins over the category default.
      const targetLibrary =
        (template.library as PropertyLinkedLibrary | undefined) ??
        CATEGORY_TO_LIBRARY[template.category];
      let matched = false;
      let matchedDoc: { filename: string; url: string; library: string } | undefined;

      // CAHP-scoped items: check the dedicated CAHP Entity Documents library first by filename match
      // Filter by OwnerLookupId state-scope; untagged docs count as shared (visible everywhere)
      let candidateDoc: DocItemRaw | undefined;
      if (template.scope === 'cahp' && cahpLib.data && cahpLib.data.length > 0) {
        const keywords = extractKeywords(template.title);
        candidateDoc = cahpLib.data.find((doc) => {
          const ownerTag = doc.fields.OwnerLookupId ? String(doc.fields.OwnerLookupId) : null;
          // Skip docs tagged to entities NOT in scope (e.g. NC LLC docs when filing an SC property)
          if (ownerTag && !cahpOwnerIds.has(ownerTag)) return false;
          const filename = (doc.fields.FileLeafRef || doc.fields.Title || '').toLowerCase();
          return keywords.some((kw) => filename.includes(kw));
        });
        if (candidateDoc && candidateDoc.webUrl) {
          matched = true;
          matchedDoc = {
            filename: candidateDoc.fields.FileLeafRef || candidateDoc.fields.Title || '(unnamed)',
            url: candidateDoc.webUrl,
            library: CAHP_ENTITY_LIBRARY,
          };
        }
      }

      // Fall through to the 8 property-linked libraries (uses OwnerLookupId / PropertyLookupId)
      if (!matched && targetLibrary) {
        const libIdx = PROPERTY_LINKED_LIBRARIES.indexOf(targetLibrary);
        const lib = libraries[libIdx];
        if (lib?.data) {
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
            candidateDoc = candidates[0];
            matchedDoc = {
              filename: candidates[0].fields.FileLeafRef || candidates[0].fields.Title || '(unnamed)',
              url: candidates[0].webUrl,
              library: targetLibrary,
            };
          }
        }
      }

      // Determine whether the matched doc is already correctly tagged for this filing.
      // - property scope: needs PropertyLookupId === propertyId
      // - owner scope:    needs OwnerLookupId in propertyOwnerIds
      // - cahp scope:     needs OwnerLookupId in cahpOwnerIds (or untagged = shared)
      let alreadyTagged = false;
      if (candidateDoc) {
        const propTag = candidateDoc.fields.PropertyLookupId ? String(candidateDoc.fields.PropertyLookupId) : null;
        const ownerTag = candidateDoc.fields.OwnerLookupId ? String(candidateDoc.fields.OwnerLookupId) : null;
        if (template.scope === 'property') alreadyTagged = propTag === propertyId;
        else if (template.scope === 'owner') alreadyTagged = ownerTag !== null && propertyOwnerIds.has(ownerTag);
        else if (template.scope === 'cahp') alreadyTagged = ownerTag === null || cahpOwnerIds.has(ownerTag);
      }

      items.push({
        template,
        matched,
        matchedDocFilename: matchedDoc?.filename,
        matchedDocUrl: matchedDoc?.url,
        matchedDocLibrary: matchedDoc?.library,
        matchedDocId: candidateDoc?.id,
        matchedDocAlreadyTagged: alreadyTagged,
      });
    });
    return items;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data, cahpLib.data, propertyId, cahpOwnerIds, propertyOwnerIds, activeChecklist]);

  const matchedCount = preview.filter((p) => p.matched).length;
  const unmatchedCount = preview.filter((p) => !p.matched).length;

  // Seed selection with every preview row the first time the preview lands.
  // Re-runs only if titles set changes (e.g. user toggled state filters).
  useEffect(() => {
    if (preview.length === 0) return;
    setSelectedTitles((prev) => {
      if (prev.size > 0) return prev;
      return new Set(preview.map((p) => p.template.title));
    });
  }, [preview]);

  const selectedCount = preview.filter((p) => selectedTitles.has(p.template.title)).length;
  const toggleOne = (title: string) =>
    setSelectedTitles((prev) => {
      const next = new Set(prev);
      if (next.has(title)) next.delete(title);
      else next.add(title);
      return next;
    });
  const selectAll = () => setSelectedTitles(new Set(preview.map((p) => p.template.title)));
  const selectNone = () => setSelectedTitles(new Set());
  const selectUnmatchedOnly = () =>
    setSelectedTitles(new Set(preview.filter((p) => !p.matched).map((p) => p.template.title)));

  const handleGenerate = async () => {
    if (!propertyId) {
      setError("No property context — can't create checklist items.");
      return;
    }
    setCreating(true);
    setError(null);

    // Pick a target owner tag for owner-scope items: prefer the property's
    // primary direct-owner LLC (largest %). Falls back to the first direct
    // owner if percentages aren't set.
    const ownerForOwnerScope = (() => {
      if (!ownership.data) return undefined;
      const rows = ownership.data.filter(
        (rel) => String(rel.fields.LinkedPropertyLookupId) === propertyId && rel.fields.OwnerLookupId,
      );
      if (rows.length === 0) return undefined;
      rows.sort((a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0));
      return String(rows[0].fields.OwnerLookupId);
    })();

    const itemsToCreate = preview.filter((p) => selectedTitles.has(p.template.title));
    let createdSuccess = 0;
    let matchedSuccess = 0;
    let taggedSuccess = 0;

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

        // Pull the matched doc into the entity's Documents view by tagging it.
        // Only touches docs that don't already have the right tag (we never
        // overwrite a tag that's already set to a different entity).
        if (
          tagMatchedDocs &&
          p.matched &&
          p.matchedDocId &&
          p.matchedDocLibrary &&
          !p.matchedDocAlreadyTagged
        ) {
          const patch: Record<string, unknown> = {};
          if (p.template.scope === 'property') {
            patch.PropertyLookupId = propertyId;
          } else if (p.template.scope === 'owner' && ownerForOwnerScope) {
            patch.OwnerLookupId = ownerForOwnerScope;
          }
          // CAHP-scope docs stay untagged so they remain shared across filings.
          if (Object.keys(patch).length > 0) {
            try {
              await updateListItem(p.matchedDocLibrary, p.matchedDocId, patch, {
                reason: 'Tagged via Filing Checklist generator so this doc appears in the entity Documents view',
              });
              taggedSuccess++;
            } catch (e) {
              // Don't fail the whole generation on a single tag failure — log only.
              // eslint-disable-next-line no-console
              console.warn(`Failed to tag '${p.matchedDocFilename}' in '${p.matchedDocLibrary}':`, e);
            }
          }
        }
      }
      // Report counts back to the caller (existing handler shows a toast).
      onSuccess(createdSuccess, matchedSuccess);
      // Surface the tag count in console for now — caller's signature is fixed.
      if (taggedSuccess > 0) {
        // eslint-disable-next-line no-console
        console.log(`[FilingChecklist] Tagged ${taggedSuccess} matched doc(s) into the entity Documents view.`);
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

            {/* Bulk selection toolbar */}
            <div className="flex items-center justify-between gap-2 mb-2 text-xs">
              <div className="text-gray-600">
                <span className="font-semibold">{selectedCount}</span> of {preview.length} selected for generation
              </div>
              <div className="flex items-center gap-1">
                <button
                  onClick={selectAll}
                  className="px-2 py-1 rounded border border-gray-200 text-gray-700 hover:bg-gray-50"
                  title="Include every template row in the generation"
                >
                  All
                </button>
                <button
                  onClick={selectUnmatchedOnly}
                  className="px-2 py-1 rounded border border-gray-200 text-gray-700 hover:bg-gray-50"
                  title="Only generate items for documents not yet on file"
                >
                  Unmatched only
                </button>
                <button
                  onClick={selectNone}
                  className="px-2 py-1 rounded border border-gray-200 text-gray-700 hover:bg-gray-50"
                  title="Clear selection"
                >
                  None
                </button>
              </div>
            </div>

            {/* Preview list */}
            <div className="border border-gray-200 rounded-md max-h-96 overflow-y-auto mb-4">
              <table className="w-full text-xs">
                <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-semibold text-gray-600 uppercase tracking-wider sticky top-0">
                  <tr>
                    <th className="px-3 py-2 text-center w-8"></th>
                    <th className="px-3 py-2 text-left w-6"></th>
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
                    const isSelected = selectedTitles.has(p.template.title);
                    return (
                      <tr key={idx} className={`${p.matched ? 'bg-green-50/40' : ''} ${isSelected ? '' : 'opacity-50'}`}>
                        <td className="px-3 py-2 text-center">
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => toggleOne(p.template.title)}
                            title={isSelected ? 'Uncheck to skip this item when generating' : 'Check to include this item in generation'}
                          />
                        </td>
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
                            <div className="flex flex-col">
                              <span className="text-success truncate" title={p.matchedDocFilename}>
                                {p.matchedDocFilename}
                              </span>
                              {tagMatchedDocs && isSelected && !p.matchedDocAlreadyTagged && p.template.scope !== 'cahp' && (
                                <span className="text-[10px] text-teal-700 italic">will tag into entity docs</span>
                              )}
                            </div>
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

            <label className="flex items-start gap-2 text-xs text-gray-700 mb-3 cursor-pointer">
              <input
                type="checkbox"
                checked={tagMatchedDocs}
                onChange={(e) => setTagMatchedDocs(e.target.checked)}
                className="mt-0.5"
              />
              <span>
                For matched docs, also tag them with the right Owner/Property ID so they show up in the
                entity's <strong>Documents</strong> view on the property page. Only touches docs that
                aren't already tagged to a specific entity (never overwrites an existing tag).
              </span>
            </label>

            <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900 mb-3">
              <div className="font-semibold mb-1">On Generate:</div>
              <ul className="space-y-0.5 list-disc list-inside">
                <li>{selectedCount} Outstanding Item{selectedCount === 1 ? '' : 's'} created, linked to this submittal</li>
                <li>Auto-matched selections pre-linked to their fulfilling docs and marked Received</li>
                <li>Unmatched selections show on the Outstanding tab with Link/Upload ready</li>
                <li>Unchecked rows are skipped entirely (use this for items that don't apply, e.g., Reassignment of Interest when not re-filing)</li>
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
            disabled={creating || loading || selectedCount === 0}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {creating && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {creating ? 'Generating…' : `Generate ${selectedCount} Item${selectedCount === 1 ? '' : 's'}`}
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * Extract matchable keywords from a checklist item title for loose filename matching
 * against docs in the CAHP Entity Documents library.
 *
 * Example: "CAHP 501(c)(3) Determination Letter" → ["501", "determination", "501c3"]
 */
function extractKeywords(title: string): string[] {
  const lower = title.toLowerCase();
  const map: { keyword: string; aliases: string[] }[] = [
    { keyword: 'operating agreement', aliases: ['operating-agreement', 'operating_agreement', 'op agreement', 'oa'] },
    { keyword: '501(c)(3)', aliases: ['501c3', 'determination letter', 'irs determination'] },
    { keyword: 'ein', aliases: ['ein confirmation', 'employer identification', 'cp575'] },
    { keyword: 'articles of incorporation', aliases: ['articles', 'incorporation'] },
    { keyword: 'articles of organization', aliases: ['articles', 'organization'] },
    { keyword: 'certificate of existence', aliases: ['coe', 'good standing'] },
    { keyword: 'certificate of authorization', aliases: ['cert of auth', 'authority'] },
    { keyword: 'bylaws', aliases: ['by-laws'] },
  ];
  const matches: string[] = [];
  for (const { keyword, aliases } of map) {
    if (lower.includes(keyword)) {
      matches.push(keyword);
      matches.push(...aliases);
    }
  }
  return matches.length > 0 ? matches : lower.split(/\s+/).filter((w) => w.length > 3);
}
