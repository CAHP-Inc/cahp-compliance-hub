import { useMemo } from 'react';
import {
  useSharePointList,
  LIST_NAMES,
  type ItemCategory,
  type CahpState,
  type ChecklistTemplate,
} from './sharepoint';

/**
 * Filing Checklist Template — drives both the Filing Checklist Generator
 * (existing property → checklist button) and the New Property wizard's
 * Step-5 outstanding-item creation.
 *
 * Source of truth: the `Checklist Templates` SharePoint list. Edits in
 * Settings → Checklist Templates write to that list, so everyone on the
 * team sees the same configuration regardless of browser or device.
 *
 * Fallback: if the SharePoint list hasn't been provisioned yet OR is empty,
 * the hook returns the hardcoded DOR_FILING_CHECKLIST below so the app
 * keeps working without manual setup.
 *
 * Legacy localStorage helpers are kept for one-time import — the Settings
 * editor offers a button to push your previously-saved local config up to
 * SharePoint so you don't lose any customizations.
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
// SharePoint <-> FilingChecklistItem conversion helpers
// =============================================================================

/** Map a SharePoint row to the in-app shape. */
export function templateRowToItem(row: ChecklistTemplate): FilingChecklistItem | null {
  const t = row.fields;
  if (!t.Title || !t.TemplateCategory || !t.TemplateScope) return null;
  return {
    title: t.Title,
    category: t.TemplateCategory as ItemCategory,
    scope: t.TemplateScope,
    notes: t.TemplateNotes || undefined,
    library: t.TemplateLibrary || undefined,
    state: t.TemplateState,
  };
}

/** Build a SharePoint field payload from an in-app item + sort position. */
export function itemToTemplateFields(
  item: FilingChecklistItem,
  sortOrder: number,
): Record<string, unknown> {
  return {
    Title: item.title,
    TemplateCategory: item.category,
    TemplateScope: item.scope,
    TemplateState: item.state ?? null,
    TemplateLibrary: item.library ?? null,
    TemplateNotes: item.notes ?? null,
    TemplateSortOrder: sortOrder,
  };
}

// =============================================================================
// React hook — primary read path
// =============================================================================

export interface UseChecklistTemplatesResult {
  /** Resolved checklist items in display order (SharePoint if available, else defaults). */
  templates: FilingChecklistItem[];
  /** Raw SharePoint rows — null if the list doesn't exist yet OR isn't loaded. */
  rawRows: ChecklistTemplate[] | null;
  /** True while we're still fetching; consumers can show their default-or-loading UI. */
  loading: boolean;
  /** Hard fetch error (e.g., list doesn't exist). The hook still returns hardcoded defaults so callers degrade gracefully. */
  error: Error | null;
  /** True when we're serving the hardcoded fallback because SP returned no rows. */
  usingFallback: boolean;
  refetch?: () => void;
}

/**
 * Read-side hook. Fetches the SharePoint Checklist Templates list and maps it
 * to FilingChecklistItem[]. Falls back to DOR_FILING_CHECKLIST when the list
 * doesn't exist (not provisioned) or has zero rows.
 */
export function useChecklistTemplates(): UseChecklistTemplatesResult {
  const list = useSharePointList<ChecklistTemplate>(LIST_NAMES.ChecklistTemplates, { top: 500 });

  const templates = useMemo<FilingChecklistItem[]>(() => {
    if (!list.data || list.data.length === 0) return DOR_FILING_CHECKLIST;
    // Sort by TemplateSortOrder ascending; missing values sort to end
    const sorted = [...list.data].sort((a, b) => {
      const aOrder = a.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
      const bOrder = b.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
      return aOrder - bOrder;
    });
    const items = sorted.map(templateRowToItem).filter((i): i is FilingChecklistItem => i !== null);
    // Fallback if everything was malformed
    return items.length > 0 ? items : DOR_FILING_CHECKLIST;
  }, [list.data]);

  const usingFallback = !list.loading && (!list.data || list.data.length === 0);

  return {
    templates,
    rawRows: list.data ?? null,
    loading: list.loading,
    error: list.error,
    usingFallback,
    refetch: list.refetch,
  };
}

// =============================================================================
// Legacy localStorage helpers — kept for one-time migration from the old
// per-browser config into the shared SharePoint list. Settings → Checklist
// Templates exposes an "Import from this browser" button that reads via
// readLocalChecklistOverride() and writes the result to SharePoint.
// =============================================================================

const STORAGE_KEY = 'cahp.checklistTemplates.v1';

/** Returns the per-browser overrides (if any) for one-time migration. */
export function readLocalChecklistOverride(): FilingChecklistItem[] | null {
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return null;
    const items = parsed.filter(
      (item): item is FilingChecklistItem =>
        item &&
        typeof item.title === 'string' &&
        typeof item.category === 'string' &&
        typeof item.scope === 'string',
    );
    return items.length > 0 ? items : null;
  } catch {
    return null;
  }
}

/** Clear the localStorage override after a successful import. */
export function clearLocalChecklistOverride(): void {
  if (typeof localStorage === 'undefined') return;
  localStorage.removeItem(STORAGE_KEY);
}
