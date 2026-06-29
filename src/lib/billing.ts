/**
 * Shared CAHP fee invoicing logic — used by both the per-submittal Generate
 * Invoice buttons (SubmittalDetail) and the "To Invoice" queue (BillingPage),
 * so the two never drift apart.
 *
 * The CAHP fee is what the OWNER pays CAHP — there is no owner disbursement.
 * Two distinct invoice kinds:
 *   - Filing Fee         : flat one-time charge per property (the Initial filing).
 *   - Percent of Savings : a % of that tax year's DOR-approved savings, billed
 *                          per tax year (the Initial year and each Annual after).
 */

import {
  createListItem,
  updateListItem,
  LIST_NAMES,
  type Billing,
  type Submittal,
  type Property,
  type BillingStatusValue,
  type SubmittalStatusValue,
} from './sharepoint';

/** Default fee terms presented in the UI — editable per invoice. */
export const DEFAULT_FILING_FEE = 3500;
export const DEFAULT_FEE_PERCENT = 20;

/** Statuses in the post-approval billing phase (DOR has approved the filing). */
export const BILLING_PHASE_STATUSES: SubmittalStatusValue[] = ['Approved', 'Invoiced', 'Paid'];

/** A Billing row is a % of savings invoice (treat untyped legacy rows as such). */
export function isPercentInvoice(b: Billing): boolean {
  return b.fields.BillingType === 'Percent of Savings' || b.fields.BillingType == null;
}

/** A Billing row is a flat filing-fee invoice. */
export function isFilingFeeInvoice(b: Billing): boolean {
  return b.fields.BillingType === 'Filing Fee';
}

/**
 * The % of savings invoice already on file for a submittal, if any. Prefers the
 * explicit submittal link; falls back to property + tax year for invoices made
 * before the BillSubmittal/BillingType columns existed, so we never double-bill.
 */
export function findPercentInvoiceForSubmittal(
  submittal: Submittal,
  billings: Billing[],
): Billing | null {
  const sid = String(submittal.id);
  const byLink = billings.find(
    (b) => String(b.fields.BillSubmittalLookupId ?? '') === sid && isPercentInvoice(b),
  );
  if (byLink) return byLink;
  const pid = submittal.fields.PropertyLookupId;
  const yr = submittal.fields.cahpTaxYear;
  if (!pid || !yr) return null;
  return (
    billings.find(
      (b) =>
        !b.fields.BillSubmittalLookupId &&
        isPercentInvoice(b) &&
        String(b.fields.PropertyLookupId ?? '') === String(pid) &&
        b.fields.cahpTaxYear === yr,
    ) ?? null
  );
}

/** The one-time filing-fee invoice for a property, if any. */
export function findFilingFeeInvoiceForProperty(
  propertyId: string,
  billings: Billing[],
): Billing | null {
  return (
    billings.find(
      (b) => isFilingFeeInvoice(b) && String(b.fields.PropertyLookupId ?? '') === String(propertyId),
    ) ?? null
  );
}

/** Move an Approved submittal into the Invoiced phase (never downgrade Paid). */
async function bumpToInvoiced(submittal: Submittal, extra?: Record<string, unknown>): Promise<void> {
  const updates: Record<string, unknown> = { ...extra };
  if (submittal.fields.SubmittalStatus === 'Approved') updates.SubmittalStatus = 'Invoiced';
  if (Object.keys(updates).length > 0) {
    await updateListItem(LIST_NAMES.Submittals, submittal.id, updates);
  }
}

/**
 * Create the % of savings invoice for a submittal (the contingency fee for that
 * tax year). Bumps the submittal to Invoiced and persists the confirmed savings.
 */
export async function generatePercentInvoice(opts: {
  submittal: Submittal;
  property?: Property | null;
  taxSavings: number;
  feePercent: number;
  letterRef?: string;
}): Promise<void> {
  const { submittal, property, taxSavings, feePercent, letterRef } = opts;
  const f = submittal.fields;
  if (!f.PropertyLookupId) throw new Error("Submittal isn't linked to a property — can't create the invoice.");

  const amount = (taxSavings * feePercent) / 100;
  const yr = f.cahpTaxYear ?? '';
  const propName = property?.fields.Title ?? 'Property';
  const ref = letterRef?.trim() ? ` Approval letter: ${letterRef.trim()}.` : '';

  await bumpToInvoiced(
    submittal,
    taxSavings !== (f.ApprovedAbatement ?? null) ? { ApprovedAbatement: taxSavings } : undefined,
  );

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} ${yr} CAHP Fee (% of Savings)`.trim(),
    PropertyLookupId: f.PropertyLookupId,
    BillSubmittalLookupId: submittal.id,
    BillingType: 'Percent of Savings',
    cahpTaxYear: f.cahpTaxYear,
    AmountBilled: amount,
    BillApprovedAbatement: taxSavings,
    CAHPFeePercent: feePercent,
    BillingStatus: 'Ready to Invoice' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `${feePercent}% of $${taxSavings.toLocaleString()} tax savings (TY ${yr}).${ref}`,
  });
}

/**
 * Create the flat one-time filing-fee invoice, anchored to the (Initial)
 * submittal. Bumps the submittal to Invoiced.
 */
export async function generateFilingFeeInvoice(opts: {
  submittal: Submittal;
  property?: Property | null;
  amount: number;
  letterRef?: string;
}): Promise<void> {
  const { submittal, property, amount, letterRef } = opts;
  const f = submittal.fields;
  if (!f.PropertyLookupId) throw new Error("Submittal isn't linked to a property — can't create the invoice.");

  const propName = property?.fields.Title ?? 'Property';
  const ref = letterRef?.trim() ? ` Approval letter: ${letterRef.trim()}.` : '';

  await bumpToInvoiced(submittal);

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} CAHP Filing Fee`.trim(),
    PropertyLookupId: f.PropertyLookupId,
    BillSubmittalLookupId: submittal.id,
    BillingType: 'Filing Fee',
    cahpTaxYear: f.cahpTaxYear,
    AmountBilled: amount,
    CAHPFilingFee: amount,
    BillingStatus: 'Ready to Invoice' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `One-time filing fee.${ref}`,
  });
}

// =============================================================================
// Queue computation — what still needs invoicing
// =============================================================================

export interface PercentQueueItem {
  submittal: Submittal;
  property?: Property | null;
  taxYear?: string;
  taxSavings: number;
  feePercent: number;
  amount: number;
}

export interface FilingFeeQueueItem {
  submittal: Submittal;        // the Initial filing the fee is anchored to
  property?: Property | null;
  amount: number;
}

export interface InvoiceQueues {
  percentItems: PercentQueueItem[];
  filingItems: FilingFeeQueueItem[];
}

/**
 * Build the "to invoice" work lists from the current data:
 *   - percentItems : approved submittals (any tax year) missing their % invoice
 *   - filingItems  : properties with an approved Initial filing but no filing fee
 */
export function computeInvoiceQueues(
  submittals: Submittal[],
  billings: Billing[],
  properties: Property[],
): InvoiceQueues {
  const propById = new Map(properties.map((p) => [String(p.id), p]));
  const inBillingPhase = (s: Submittal) =>
    s.fields.SubmittalStatus != null && BILLING_PHASE_STATUSES.includes(s.fields.SubmittalStatus);

  // % of savings — one per approved submittal that has a savings figure and no
  // % invoice yet, newest tax year first.
  const percentItems: PercentQueueItem[] = submittals
    .filter((s) => inBillingPhase(s))
    .filter((s) => (s.fields.ApprovedAbatement ?? 0) > 0)
    .filter((s) => !findPercentInvoiceForSubmittal(s, billings))
    .map((s) => {
      const taxSavings = s.fields.ApprovedAbatement ?? 0;
      return {
        submittal: s,
        property: s.fields.PropertyLookupId ? propById.get(String(s.fields.PropertyLookupId)) ?? null : null,
        taxYear: s.fields.cahpTaxYear,
        taxSavings,
        feePercent: DEFAULT_FEE_PERCENT,
        amount: (taxSavings * DEFAULT_FEE_PERCENT) / 100,
      };
    })
    .sort((a, b) => String(b.taxYear ?? '').localeCompare(String(a.taxYear ?? '')));

  // Filing fee — one per property that has an approved Initial filing and no
  // filing-fee invoice. Keyed by property so we never list it twice.
  const filingByProperty = new Map<string, FilingFeeQueueItem>();
  for (const s of submittals) {
    if (!inBillingPhase(s)) continue;
    if (s.fields.FilingType !== 'Initial') continue;
    const pid = s.fields.PropertyLookupId;
    if (!pid) continue;
    if (filingByProperty.has(String(pid))) continue;
    if (findFilingFeeInvoiceForProperty(String(pid), billings)) continue;
    filingByProperty.set(String(pid), {
      submittal: s,
      property: propById.get(String(pid)) ?? null,
      amount: DEFAULT_FILING_FEE,
    });
  }

  return { percentItems, filingItems: Array.from(filingByProperty.values()) };
}
