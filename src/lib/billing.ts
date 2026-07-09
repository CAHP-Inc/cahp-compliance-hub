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
  type CahpTaxYear,
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
 * A filing-fee row that records "we are NOT charging this property an initial
 * filing fee" — a $0 placeholder, status N/A. It still satisfies
 * findFilingFeeInvoiceForProperty, so the property drops out of the to-invoice
 * queue exactly like a real fee would.
 */
export function isNAFilingFee(b: Billing): boolean {
  return isFilingFeeInvoice(b) && b.fields.BillingStatus === 'N/A';
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

/**
 * The % of savings invoice for a property in a given tax year, if any. Used by
 * the annual (per-year) billing flow, which records each year's CAHP % directly
 * against the property — there is no per-year submittal to key off.
 */
export function findPercentInvoiceForPropertyYear(
  propertyId: string,
  taxYear: string,
  billings: Billing[],
  taxMapId?: string | null,
): Billing | null {
  // A % invoice is scoped either to the whole property (no TMID) or to a single
  // parcel; dedupe matches the same scope so each TMID can be billed separately.
  const tmid = taxMapId ? String(taxMapId) : '';
  return (
    billings.find(
      (b) =>
        isPercentInvoice(b) &&
        String(b.fields.PropertyLookupId ?? '') === String(propertyId) &&
        b.fields.cahpTaxYear === taxYear &&
        String(b.fields.BillTaxMapIDLookupId ?? '') === tmid,
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

/**
 * Record a CAHP "% of Annual Savings" invoice for a property in a single tax
 * year, decoupled from any submittal. This is the recurring annual billing:
 * after the Initial filing there is no per-year recertification submittal, so
 * each year's tax savings is entered directly and the % fee billed against the
 * property. Optionally links to the Initial submittal for traceability, but
 * never changes its status.
 */
export async function recordAnnualPercentInvoice(opts: {
  property: Property;
  taxYear: CahpTaxYear;
  lastFullTaxBill: number;
  mostRecentTaxBill: number;
  feePercent: number;
  taxMapId?: string | null;
  initialSubmittal?: Submittal | null;
  letterRef?: string;
}): Promise<void> {
  const { property, taxYear, lastFullTaxBill, mostRecentTaxBill, feePercent, taxMapId, initialSubmittal, letterRef } = opts;
  // Savings = the abatement: full (pre-abatement) bill minus the most recent (abated) bill.
  const taxSavings = Math.max(0, lastFullTaxBill - mostRecentTaxBill);
  const amount = (taxSavings * feePercent) / 100;
  const propName = property.fields.Title ?? 'Property';
  const ref = letterRef?.trim() ? ` Approval letter: ${letterRef.trim()}.` : '';

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} ${taxYear} CAHP Fee (% of Savings)`.trim(),
    PropertyLookupId: String(property.id),
    ...(initialSubmittal ? { BillSubmittalLookupId: String(initialSubmittal.id) } : {}),
    ...(taxMapId ? { BillTaxMapIDLookupId: String(taxMapId) } : {}),
    BillingType: 'Percent of Savings',
    cahpTaxYear: taxYear,
    LastFullTaxBill: lastFullTaxBill,
    MostRecentTaxBill: mostRecentTaxBill,
    AmountBilled: amount,
    BillApprovedAbatement: taxSavings,
    CAHPFeePercent: feePercent,
    BillingStatus: 'Ready to Invoice' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `${feePercent}% of $${taxSavings.toLocaleString()} savings (full $${lastFullTaxBill.toLocaleString()} − recent $${mostRecentTaxBill.toLocaleString()}), TY ${taxYear}.${ref}`,
  });
}

/** A % of savings row that records "we are NOT claiming the % for this year". */
export function isNAPercent(b: Billing): boolean {
  return isPercentInvoice(b) && b.fields.BillingStatus === 'N/A';
}

/**
 * Mark a property's % of Annual Savings as N/A for a tax year (e.g. an abatement
 * already obtained under another program where CAHP isn't claiming a fee). A $0
 * placeholder so the year drops out of the roll-forward, with an audit trail.
 */
export async function markPercentNA(opts: {
  property: Property;
  taxYear: CahpTaxYear;
  taxMapId?: string | null;
  initialSubmittal?: Submittal | null;
  note?: string;
}): Promise<void> {
  const { property, taxYear, taxMapId, initialSubmittal, note } = opts;
  const propName = property.fields.Title ?? 'Property';
  const reason = note?.trim() ? ` ${note.trim()}` : '';

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} ${taxYear} CAHP Fee (% of Savings) — N/A`.trim(),
    PropertyLookupId: String(property.id),
    ...(initialSubmittal ? { BillSubmittalLookupId: String(initialSubmittal.id) } : {}),
    ...(taxMapId ? { BillTaxMapIDLookupId: String(taxMapId) } : {}),
    BillingType: 'Percent of Savings',
    cahpTaxYear: taxYear,
    AmountBilled: 0,
    BillApprovedAbatement: 0,
    BillingStatus: 'N/A' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `% of savings not claimed for TY ${taxYear} (N/A).${reason}`,
  });
}

/**
 * Mark a property's initial filing fee as N/A (not charged) by writing a $0
 * placeholder Filing Fee row. Anchored to the Initial submittal when one exists.
 * Does NOT touch the submittal status (nothing is being invoiced).
 */
export async function markFilingFeeNA(opts: {
  property: Property;
  submittal?: Submittal | null;
  note?: string;
}): Promise<void> {
  const { property, submittal, note } = opts;
  const propName = property.fields.Title ?? 'Property';
  const reason = note?.trim() ? ` ${note.trim()}` : '';

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} CAHP Filing Fee — N/A`.trim(),
    PropertyLookupId: String(property.id),
    ...(submittal ? { BillSubmittalLookupId: String(submittal.id), cahpTaxYear: submittal.fields.cahpTaxYear } : {}),
    BillingType: 'Filing Fee',
    AmountBilled: 0,
    CAHPFilingFee: 0,
    BillingStatus: 'N/A' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `Initial filing fee not charged (N/A).${reason}`,
  });
}

/**
 * Record a full-bill BASELINE for a property/scope that has no abatement yet
 * (the most recent bill is still the full bill, so there are no savings to bill).
 * Stored as a $0, YEAR-LESS % row carrying the full bill, so it pulls forward as
 * the Last Full Tax Bill but never occupies — or blocks — a billable year. Once
 * the abated bill arrives, that year is billed normally with the full pulled in.
 */
export async function recordFullBillBaseline(opts: {
  property: Property;
  lastFullTaxBill: number;
  taxMapId?: string | null;
  note?: string;
}): Promise<void> {
  const { property, lastFullTaxBill, taxMapId, note } = opts;
  const propName = property.fields.Title ?? 'Property';
  const reason = note?.trim() ? ` ${note.trim()}` : '';

  await createListItem(LIST_NAMES.Billing, {
    Title: `${propName} CAHP % Baseline (full bill)`.trim(),
    PropertyLookupId: String(property.id),
    ...(taxMapId ? { BillTaxMapIDLookupId: String(taxMapId) } : {}),
    BillingType: 'Percent of Savings',
    LastFullTaxBill: lastFullTaxBill,
    AmountBilled: 0,
    BillApprovedAbatement: 0,
    BillingStatus: 'N/A' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    BillingNotes: `Baseline full tax bill $${lastFullTaxBill.toLocaleString()} (no abatement yet); pulled forward.${reason}`,
  });
}

/** A year-less % row that carries a full bill = a pull-forward baseline. */
export function isBaselineRow(b: Billing): boolean {
  return isPercentInvoice(b) && !b.fields.cahpTaxYear && typeof b.fields.LastFullTaxBill === 'number';
}

/** The full-bill baseline on file for a property/scope, if any. */
export function findBaselineForScope(
  propertyId: string,
  billings: Billing[],
  taxMapId?: string | null,
): Billing | null {
  const tmid = taxMapId ? String(taxMapId) : '';
  return (
    billings.find(
      (b) =>
        isBaselineRow(b) &&
        String(b.fields.PropertyLookupId ?? '') === String(propertyId) &&
        String(b.fields.BillTaxMapIDLookupId ?? '') === tmid,
    ) ?? null
  );
}

// =============================================================================
// Monthly billing — prorate the annual CAHP fee into a monthly amount
// (mirrors the "CAHP Bill Backs" spreadsheet: Months, Monthly, Invoice text)
// =============================================================================

const fmtUSD = (n: number): string =>
  n.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 });

/**
 * Number of months to bill in the (first) year, matching the spreadsheet:
 *   previously abated              -> 12 (full year)
 *   newly abated, start date known -> remaining months from the start month,
 *                                     counting the start month only if >15 days
 *                                     of it remain (EOMONTH(date)-date > 15)
 *   newly abated, no start date    -> 12 (assume full year)
 */
export function computeBillingMonths(previouslyAbated: boolean, billStartDate?: string | null): number {
  if (previouslyAbated) return 12;
  if (!billStartDate) return 12;
  const iso = billStartDate.slice(0, 10);
  const [y, m, day] = iso.split('-').map(Number);
  if (!y || !m || !day) return 12;
  const lastDayOfMonth = new Date(Date.UTC(y, m, 0)).getUTCDate(); // day 0 of next month
  const daysRemainingInMonth = lastDayOfMonth - day;
  const countStartMonth = daysRemainingInMonth > 15 ? 1 : 0;
  return (12 - m) + countStartMonth;
}

/** Monthly CAHP = annual fee / months (guards divide-by-zero). */
export function monthlyFee(annualFee: number, months: number): number {
  return months > 0 ? annualFee / months : 0;
}

/**
 * The invoice-description narrative, matching the spreadsheet's Invoice
 * Description column. `previouslyAbated` picks the short vs. full savings clause.
 */
export function buildMonthlyInvoiceDescription(opts: {
  owner: string;
  taxYear: string | number;
  previouslyAbated: boolean;
  lastFullTaxBill: number;
  mostRecentTaxBill: number;
  totalSavings: number;
  feePercent: number;
  annualFee: number;
  months: number;
  monthly: number;
}): string {
  const { owner, taxYear, previouslyAbated, lastFullTaxBill, mostRecentTaxBill, totalSavings, feePercent, annualFee, months, monthly } = opts;
  const savingsClause = previouslyAbated
    ? `Total Savings ${fmtUSD(totalSavings)}`
    : `Last Full Tax Bill ${fmtUSD(lastFullTaxBill)} less estimated County fees of ${fmtUSD(mostRecentTaxBill)} due to County = Total Savings ${fmtUSD(totalSavings)}`;
  return (
    `Monthly CAHP Bill — ${owner} (Tax Year ${taxYear}): ` +
    `${savingsClause} × ${feePercent}% CAHP = ${fmtUSD(annualFee)} annual ÷ ${months} months = ${fmtUSD(monthly)}/month`
  );
}

/**
 * Create-or-update the billing record that backs one approved-abatement row
 * (keyed to a submittal + parcel). Savings come from the two tax bills when both
 * are given, otherwise from the submittal's approved abatement (fallbackSavings).
 * Any subset of the editable inputs may be passed; only those are written.
 */
export async function upsertAbatementRecord(opts: {
  billingId?: string | null;
  propertyId: string;
  submittalId?: string | null;
  taxYear?: string;
  taxMapId?: string | null;
  lastFullTaxBill?: number | null;
  mostRecentTaxBill?: number | null;
  feePercent?: number;
  fallbackSavings?: number;
  previouslyAbated?: boolean;
  billStartDate?: string | null;
  refundStatus?: string;
}): Promise<void> {
  const hasBills = typeof opts.lastFullTaxBill === 'number' && typeof opts.mostRecentTaxBill === 'number';
  const feePercent = opts.feePercent ?? DEFAULT_FEE_PERCENT;
  const savings = hasBills
    ? Math.max(0, (opts.lastFullTaxBill as number) - (opts.mostRecentTaxBill as number))
    : (opts.fallbackSavings ?? 0);
  const amount = (savings * feePercent) / 100;

  const fields: Record<string, unknown> = {};
  if (hasBills) { fields.LastFullTaxBill = opts.lastFullTaxBill; fields.MostRecentTaxBill = opts.mostRecentTaxBill; }
  if (opts.feePercent !== undefined) fields.CAHPFeePercent = feePercent;
  if (hasBills || opts.feePercent !== undefined) { fields.BillApprovedAbatement = savings; fields.AmountBilled = amount; }
  if (opts.previouslyAbated !== undefined) fields.PreviouslyAbated = opts.previouslyAbated;
  if (opts.billStartDate !== undefined) fields.BillStartDate = opts.billStartDate;
  if (opts.refundStatus !== undefined) fields.RefundStatus = opts.refundStatus;

  if (opts.billingId) {
    if (Object.keys(fields).length > 0) await updateListItem(LIST_NAMES.Billing, opts.billingId, fields);
    return;
  }
  await createListItem(LIST_NAMES.Billing, {
    Title: `TY ${opts.taxYear ?? ''} CAHP Fee (% of Savings)`.trim(),
    PropertyLookupId: opts.propertyId,
    ...(opts.submittalId ? { BillSubmittalLookupId: opts.submittalId } : {}),
    ...(opts.taxMapId ? { BillTaxMapIDLookupId: opts.taxMapId } : {}),
    BillingType: 'Percent of Savings',
    cahpTaxYear: opts.taxYear as CahpTaxYear | undefined,
    CAHPFeePercent: feePercent,
    BillApprovedAbatement: savings,
    AmountBilled: amount,
    BillingStatus: 'Ready to Invoice' as BillingStatusValue,
    QBSyncStatus: 'Not Synced',
    ...fields,
  });
}

/** Update an abatement record's tax bills / fee %, recomputing savings + annual amount. */
export async function updateBillingRecord(
  billingId: string,
  values: { lastFullTaxBill: number; mostRecentTaxBill: number; feePercent: number },
): Promise<void> {
  const savings = Math.max(0, values.lastFullTaxBill - values.mostRecentTaxBill);
  await updateListItem(LIST_NAMES.Billing, billingId, {
    LastFullTaxBill: values.lastFullTaxBill,
    MostRecentTaxBill: values.mostRecentTaxBill,
    BillApprovedAbatement: savings,
    CAHPFeePercent: values.feePercent,
    AmountBilled: (savings * values.feePercent) / 100,
  });
}

/** Set the previously-paid-tax refund status on a Billing row. */
export async function updateRefundStatus(billingId: string, status: string): Promise<void> {
  await updateListItem(LIST_NAMES.Billing, billingId, { RefundStatus: status });
}

/** Persist the two monthly-proration inputs onto a Billing row. */
export async function updateBillingMonthlyInputs(
  billingId: string,
  values: { previouslyAbated?: boolean; billStartDate?: string | null },
): Promise<void> {
  const updates: Record<string, unknown> = {};
  if (values.previouslyAbated !== undefined) updates.PreviouslyAbated = values.previouslyAbated;
  if (values.billStartDate !== undefined) updates.BillStartDate = values.billStartDate;
  if (Object.keys(updates).length > 0) await updateListItem(LIST_NAMES.Billing, billingId, updates);
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
