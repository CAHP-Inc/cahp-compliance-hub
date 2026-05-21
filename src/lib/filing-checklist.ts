import type { ItemCategory, CahpState } from './sharepoint';

/**
 * Filing Checklist Template — based on the DOR Townes at Converse submission.
 *
 * Each item maps to a category from ItemCategory. Scope hints tell the
 * auto-match logic where to look:
 *  - 'cahp'      — tagged to a CAHP entity owner
 *  - 'owner'     — tagged to the property's owner entity
 *  - 'property'  — tagged directly to the property
 *
 * The active list is editable via Settings → Checklist Templates and lives in
 * localStorage. We seed from DOR_FILING_CHECKLIST below the first time it's
 * read, so a fresh browser sees the same defaults as before.
 *
 * Priority is High by default because these are filing-blocking docs.
 */

export type FilingChecklistScope = 'cahp' | 'owner' | 'property';

export interface FilingChecklistItem {
  title: string;
  category: ItemCategory;
  scope: FilingChecklistScope;
  notes?: string;
  /** Optional override of the SharePoint library this item maps to.
   *  When unset, the FilingChecklistGenerator falls back to its category→library map. */
  library?: string;
  /** Optional state filter — when set, this item is only added for properties in that state.
   *  Undefined = applies to every property regardless of state. */
  state?: CahpState;
}

export const DOR_FILING_CHECKLIST: FilingChecklistItem[] = [
  // ─── CAHP / Nonprofit entity docs (5) — one-time setup, reused across all filings ───
  {
    title: 'CAHP Operating Agreement (Non Profit OA)',
    category: 'Operating Agreement',
    scope: 'cahp',
    notes: "CAHP SC LLC Operating Agreement. Lives at the CAHP entity level — should be reusable across all filings.",
  },
  {
    title: 'CAHP 501(c)(3) Determination Letter',
    category: '501(c)(3) Determination',
    scope: 'cahp',
    notes: 'IRS determination letter for Carolina Affordable Housing Project (501(c)(3) status).',
  },
  {
    title: 'CAHP EIN Confirmation',
    category: 'EIN Confirmation',
    scope: 'cahp',
    notes: 'IRS EIN letter for the nonprofit.',
  },
  {
    title: 'CAHP Articles of Incorporation',
    category: 'Articles of Incorporation',
    scope: 'cahp',
    notes: 'Nonprofit Articles of Incorporation.',
  },
  {
    title: 'CAHP Certificate of Existence (COE)',
    category: 'Certificate of Existence',
    scope: 'cahp',
    notes: 'State-issued Certificate of Existence / Good Standing for the nonprofit.',
  },

  // ─── Property-Owner Entity docs (4) — uploaded once per LLC, reused per filing ───
  {
    title: 'Entity Certification Letter (Cert of Authorization)',
    category: 'Certificate of Authorization',
    scope: 'owner',
    notes: "State-issued Cert of Authorization for the property-owning LLC.",
  },
  {
    title: 'Entity EIN Confirmation',
    category: 'EIN Confirmation',
    scope: 'owner',
    notes: 'IRS EIN letter for the property-owning LLC.',
  },
  {
    title: 'Entity Operating Agreement',
    category: 'Operating Agreement',
    scope: 'owner',
    notes: 'Operating Agreement for the property-owning LLC.',
  },
  {
    title: 'Entity Articles of Organization',
    category: 'Articles of Incorporation',
    scope: 'owner',
    notes: 'Articles of Organization for the property-owning LLC.',
  },

  // ─── Property-specific docs (4) — needed per filing ───
  {
    title: 'Property Deed(s)',
    category: 'Deed',
    scope: 'property',
    notes: 'Recorded property deed(s). Multiple parcels = multiple deeds.',
  },
  {
    title: 'Rent Roll (current year)',
    category: 'Rent Roll',
    scope: 'property',
    notes: 'Current-year rent roll showing tenant income qualification.',
  },
  {
    title: 'IRS Determination Letter (property-specific, if applicable)',
    category: 'Determination Letter',
    scope: 'property',
    notes: "Property-specific IRS determination, if one exists for this filing.",
  },
];

// =============================================================================
// Local override store — Settings → Checklist Templates edits land here.
//
// localStorage keeps changes per-browser. For shared editing across teammates
// the user can export the JSON from Settings and import it elsewhere; we can
// move this to a SharePoint list when multi-user sync becomes important.
// =============================================================================

const STORAGE_KEY = 'cahp.checklistTemplates.v1';

/** Read the active checklist — user-overridden if present, else hardcoded defaults. */
export function getFilingChecklist(): FilingChecklistItem[] {
  if (typeof localStorage === 'undefined') return DOR_FILING_CHECKLIST;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DOR_FILING_CHECKLIST;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return DOR_FILING_CHECKLIST;
    // Light schema validation — skip rows missing required fields rather than throwing.
    // `state` and `library` are optional and pass through unchanged when present.
    return parsed.filter(
      (item): item is FilingChecklistItem =>
        item &&
        typeof item.title === 'string' &&
        typeof item.category === 'string' &&
        typeof item.scope === 'string',
    );
  } catch {
    return DOR_FILING_CHECKLIST;
  }
}

/** Persist a custom checklist. Pass an empty array to keep an explicitly-empty list. */
export function saveFilingChecklist(items: FilingChecklistItem[]): void {
  if (typeof localStorage === 'undefined') return;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
}

/** Clear the override and revert to the hardcoded defaults. */
export function resetFilingChecklist(): void {
  if (typeof localStorage === 'undefined') return;
  localStorage.removeItem(STORAGE_KEY);
}

/** True if the user has saved any custom overrides (vs. running on defaults). */
export function hasCustomFilingChecklist(): boolean {
  if (typeof localStorage === 'undefined') return false;
  return localStorage.getItem(STORAGE_KEY) !== null;
}
