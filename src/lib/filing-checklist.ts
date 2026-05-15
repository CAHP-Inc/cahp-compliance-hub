import type { ItemCategory } from './sharepoint';

/**
 * Filing Checklist Template — based on the DOR Townes at Converse submission.
 *
 * Each item maps to a category from ItemCategory. Scope hints tell the
 * auto-match logic where to look:
 *  - 'cahp'      — tagged to a CAHP entity owner
 *  - 'owner'     — tagged to the property's owner entity
 *  - 'property'  — tagged directly to the property
 *
 * Priority is High by default because these are filing-blocking docs.
 */

export type FilingChecklistScope = 'cahp' | 'owner' | 'property';

export interface FilingChecklistItem {
  title: string;
  category: ItemCategory;
  scope: FilingChecklistScope;
  notes?: string;
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
