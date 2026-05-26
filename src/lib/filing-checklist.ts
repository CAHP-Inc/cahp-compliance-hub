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

/**
 * SC PT-401-O exemption application checklist. Mirrors the official SC DOR
 * filing requirements as of 2026:
 *
 *   For the applicant + property-owning entity:
 *     - PT-401-O exemption application (the form itself)
 *     - Recorded deed
 *     - Partnership agreement (or Operating Agreement for LLCs)
 *     - Organizational structure chart
 *     - Stamped SC Articles of Organization
 *     - Rent rolls and/or restrictive covenants w/ SC State Housing
 *     - Compliance certificate from SC State Housing (if applicable)
 *
 *   For the wholly-owned LLC affiliate of the nonprofit (e.g. CAHP SC LLC):
 *     - Operating Agreement between the LLC and its sole member
 *     - Stamped SC Articles of Organization
 *
 *   For the nonprofit housing corporation (e.g. CAHP Inc):
 *     - IRS 501(c)(3) determination letter
 *     - Bylaws
 *     - Stamped SC Articles of Incorporation
 *
 *   Conditional:
 *     - Reassignment of interest sign-off (when refiling under a new nonprofit)
 *
 * NC filings have a different document set; items here are tagged `state: 'SC'`
 * so they only auto-create for SC properties. Add NC-specific items when
 * NC requirements are confirmed.
 */
export const DOR_FILING_CHECKLIST: FilingChecklistItem[] = [
  // ─── CAHP nonprofit corporation docs (3) — reused across every SC filing ───
  {
    title: 'CAHP 501(c)(3) Determination Letter',
    category: '501(c)(3) Determination',
    scope: 'cahp',
    state: 'SC',
    notes: 'IRS exempt determination letter for Carolina Affordable Housing Project.',
  },
  {
    title: 'CAHP Bylaws',
    category: 'Bylaws',
    scope: 'cahp',
    state: 'SC',
    notes: 'Bylaws of the nonprofit housing corporation.',
  },
  {
    title: 'CAHP Articles of Incorporation (Stamped SC)',
    category: 'Articles of Incorporation',
    scope: 'cahp',
    state: 'SC',
    notes: 'Stamped SC Articles of Incorporation for the nonprofit corporation.',
  },

  // ─── CAHP SC LLC (wholly-owned affiliate of the nonprofit) docs (2) ───
  {
    title: 'CAHP SC LLC Operating Agreement (LLC ↔ sole member)',
    category: 'Operating Agreement',
    scope: 'cahp',
    state: 'SC',
    notes: 'Operating agreement between CAHP SC LLC and its sole member (the nonprofit). Demonstrates wholly-owned-affiliate relationship.',
  },
  {
    title: 'CAHP SC LLC Articles of Organization (Stamped SC)',
    category: 'Articles of Incorporation',
    scope: 'cahp',
    state: 'SC',
    notes: 'Stamped SC Articles of Organization for the wholly-owned LLC affiliate.',
  },

  // ─── Property-Owner Entity docs (3) — one upload per LLC, reused per filing ───
  {
    title: 'Entity Partnership Agreement / Operating Agreement',
    category: 'Partnership Agreement',
    scope: 'owner',
    state: 'SC',
    notes: 'Partnership agreement, or Operating Agreement if the property-owning entity is an LLC.',
  },
  {
    title: 'Entity Articles of Organization (Stamped SC)',
    category: 'Articles of Incorporation',
    scope: 'owner',
    state: 'SC',
    notes: 'Stamped SC Articles of Organization for the property-owning entity.',
  },
  {
    title: 'Organizational Structure Chart',
    category: 'Org Chart',
    scope: 'owner',
    state: 'SC',
    notes: 'Org chart showing the property-owning entity, its members, and the chain up to the CAHP nonprofit. The app can export this from the property detail page.',
  },

  // ─── Per-filing property docs (5) ───
  {
    title: 'PT-401-O Exemption Application (completed)',
    category: 'Exemption Application',
    scope: 'property',
    state: 'SC',
    notes: 'Completed SC PT-401-O exemption application form.',
  },
  {
    title: 'Recorded Property Deed(s)',
    category: 'Deed',
    scope: 'property',
    state: 'SC',
    notes: 'Recorded property deed(s) from the county. Multiple parcels = multiple deeds.',
  },
  {
    title: 'Rent Roll (current year)',
    category: 'Rent Roll',
    scope: 'property',
    state: 'SC',
    notes: 'Current-year rent roll. Required either as standalone or paired with the recorded restrictive covenants below.',
  },
  {
    title: 'Recorded Restrictive Covenants (SC State Housing)',
    category: 'Restrictive Covenants',
    scope: 'property',
    state: 'SC',
    notes: 'Recorded restrictive covenants filed with SC State Housing. Pair with the rent roll above (at least one of the two is required).',
  },
  {
    title: 'Most Recent SC State Housing Compliance Certificate (if applicable)',
    category: 'Compliance Certificate',
    scope: 'property',
    state: 'SC',
    notes: 'Optional but recommended when one exists — supports the rent-roll/covenants documentation.',
  },

  // ─── Conditional (re-filings only) ───
  // Not auto-added; surfaced as an empty placeholder so the user remembers to
  // attach it when a property is being re-filed under a new nonprofit.
  {
    title: 'Reassignment of Interest Sign-Off (if re-filing under new nonprofit)',
    category: 'Reassignment of Interest',
    scope: 'property',
    state: 'SC',
    notes: 'Only required when the property was previously filed under a different nonprofit and is being re-assigned. Skip otherwise.',
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
