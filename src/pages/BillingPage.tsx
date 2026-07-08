import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  deleteListItem,
  LIST_NAMES,
  type Billing,
  type Disbursement,
  type Property,
  type Submittal,
  type BillingStatusValue,
  type QBSyncStatus,
  type DisbursementStatus,
  type CahpTaxYear,
  type TaxMapID,
} from '../lib/sharepoint';
import {
  computeInvoiceQueues,
  generatePercentInvoice,
  generateFilingFeeInvoice,
  markFilingFeeNA,
  isNAFilingFee,
  isPercentInvoice,
  isNAPercent,
  isBaselineRow,
  recordAnnualPercentInvoice,
  markPercentNA,
  findPercentInvoiceForPropertyYear,
  computeBillingMonths,
  monthlyFee,
  buildMonthlyInvoiceDescription,
  updateBillingMonthlyInputs,
  DEFAULT_FEE_PERCENT,
  type PercentQueueItem,
  type FilingFeeQueueItem,
} from '../lib/billing';
import { Icon } from '../components/ui/Icon';

const BILLING_STATUS_STYLES: Record<BillingStatusValue, string> = {
  'Pending Approval': 'bg-gray-100 text-gray-800',
  'Ready to Invoice': 'bg-amber-100 text-amber-800',
  'Invoiced': 'bg-blue-100 text-blue-800',
  'Paid': 'bg-green-100 text-green-800',
  'Disputed': 'bg-red-100 text-red-800',
  'N/A': 'bg-gray-100 text-gray-500',
};

const QB_SYNC_STYLES: Record<QBSyncStatus, string> = {
  'Not Synced': 'bg-gray-100 text-gray-600',
  'Synced': 'bg-green-100 text-green-800',
  'Discrepancy': 'bg-red-100 text-red-800',
};

const DISB_STATUS_STYLES: Record<DisbursementStatus, string> = {
  'Pending': 'bg-amber-100 text-amber-800',
  'Issued': 'bg-blue-100 text-blue-800',
  'Cleared': 'bg-green-100 text-green-800',
  'Voided': 'bg-gray-100 text-gray-500',
};

type Tab = 'queue' | 'invoices' | 'monthly' | 'disbursements' | 'reconciliation';

export function BillingPage() {
  const [tab, setTab] = useState<Tab>('queue');

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Billing</h1>
        <p className="text-sm text-gray-500 mt-1">
          CAHP fee invoices — the filing fee (one-time) and the % of savings (per tax year) that owners pay CAHP. Generated from an Approved submittal once accounting confirms the numbers.
        </p>
      </div>

      {/* Tab nav */}
      <div className="border-b border-gray-200 mb-4 flex gap-1 overflow-x-auto">
        <TabButton active={tab === 'queue'} onClick={() => setTab('queue')}>
          To Invoice
        </TabButton>
        <TabButton active={tab === 'invoices'} onClick={() => setTab('invoices')}>
          CAHP Fee Invoices
        </TabButton>
        <TabButton active={tab === 'monthly'} onClick={() => setTab('monthly')}>
          Monthly Billing
        </TabButton>
        <TabButton active={tab === 'disbursements'} onClick={() => setTab('disbursements')}>
          Refund Disbursements
        </TabButton>
        <TabButton active={tab === 'reconciliation'} onClick={() => setTab('reconciliation')}>
          Reconciliation
        </TabButton>
      </div>

      {tab === 'queue' && <ToInvoiceTab />}
      {tab === 'invoices' && <InvoicesTab />}
      {tab === 'monthly' && <MonthlyBillingTab />}
      {tab === 'disbursements' && <DisbursementsTab />}
      {tab === 'reconciliation' && <ReconciliationTab />}
    </div>
  );
}

// =============================================================================
// To Invoice Tab — what still needs invoicing, derived from submittals' tax years
// =============================================================================

function ToInvoiceTab() {
  const navigate = useNavigate();
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const billings = useSharePointList<Billing>(LIST_NAMES.Billing, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxmaps = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });

  // id of the item currently being generated, or 'all' during a batch run
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  // Per-row overrides — defaults are pre-filled, but each property may sit under
  // a different agreement, so the fee % and filing-fee amount are editable here.
  const [pctOverrides, setPctOverrides] = useState<Record<string, string>>({});
  const [feeOverrides, setFeeOverrides] = useState<Record<string, string>>({});
  // Annual % of savings roll-forward: target year + per-property tax-bill entry.
  const [rollYear, setRollYear] = useState<CahpTaxYear | ''>('');
  const [rollFull, setRollFull] = useState<Record<string, string>>({});       // last full tax bill, by propertyId
  const [rollRecent, setRollRecent] = useState<Record<string, string>>({});    // most recent (abated) bill, by propertyId

  const queues = useMemo(() => {
    if (!submittals.data || !billings.data || !properties.data) return null;
    return computeInvoiceQueues(submittals.data, billings.data, properties.data);
  }, [submittals.data, billings.data, properties.data]);

  // Properties enrolled in recurring "% of Annual Savings" billing (i.e. they
  // already have at least one % invoice), for the roll-forward panel.
  const ROLL_TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];
  type RollUnit = { key: string; pid: string; tmid: string; property: Property; tmidLabel: string | null; lastYear: number | null; lastPct: number; lastFull: number | null };
  const rollForward = useMemo(() => {
    if (!billings.data || !properties.data) return { enrolled: [] as RollUnit[], suggestedYear: String(new Date().getFullYear()) as CahpTaxYear };
    const propById = new Map(properties.data.map((p) => [String(p.id), p]));
    const tmidById = new Map((taxmaps.data ?? []).map((t) => [String(t.id), t]));
    const pctRows = billings.data.filter(isPercentInvoice);
    // A "unit" is the scope a % was billed at: the whole property, or one TMID.
    const byUnit = new Map<string, { pid: string; tmid: string; rows: Billing[] }>();
    for (const b of pctRows) {
      const pid = b.fields.PropertyLookupId ? String(b.fields.PropertyLookupId) : '';
      if (!pid) continue;
      const tmid = b.fields.BillTaxMapIDLookupId ? String(b.fields.BillTaxMapIDLookupId) : '';
      const key = `${pid}|${tmid}`;
      const u = byUnit.get(key) ?? { pid, tmid, rows: [] };
      u.rows.push(b);
      byUnit.set(key, u);
    }
    const allYears = pctRows.map((b) => Number(b.fields.cahpTaxYear)).filter((n) => !Number.isNaN(n));
    const nextNum = allYears.length ? Math.max(...allYears) + 1 : new Date().getFullYear();
    const suggestedYear = (ROLL_TAX_YEARS.includes(String(nextNum) as CahpTaxYear) ? String(nextNum) : ROLL_TAX_YEARS[ROLL_TAX_YEARS.length - 1]) as CahpTaxYear;
    const enrolled = [...byUnit.values()]
      .map((u): RollUnit | null => {
        const property = propById.get(u.pid);
        if (!property) return null;
        const years = u.rows.map((r) => Number(r.fields.cahpTaxYear)).filter((n) => !Number.isNaN(n));
        const lastYear = years.length ? Math.max(...years) : null;
        const last = u.rows.find((r) => Number(r.fields.cahpTaxYear) === lastYear);
        const tmidLabel = u.tmid ? (tmidById.get(u.tmid)?.fields.Title ?? `TMID ${u.tmid}`) : null;
        // Pull the full (pre-abatement) bill forward from the most recent prior bill.
        const fullRow = [...u.rows].filter((r) => typeof r.fields.LastFullTaxBill === 'number').sort((a, b) => (Number(b.fields.cahpTaxYear) || -Infinity) - (Number(a.fields.cahpTaxYear) || -Infinity))[0];
        const lastFull = fullRow ? (fullRow.fields.LastFullTaxBill as number) : null;
        return { key: `${u.pid}|${u.tmid}`, pid: u.pid, tmid: u.tmid, property, tmidLabel, lastYear, lastPct: last?.fields.CAHPFeePercent ?? DEFAULT_FEE_PERCENT, lastFull };
      })
      .filter((e): e is RollUnit => e !== null)
      .sort((a, b) => (a.property.fields.Title ?? '').localeCompare(b.property.fields.Title ?? '') || (a.tmidLabel ?? '').localeCompare(b.tmidLabel ?? ''));
    return { enrolled, suggestedYear };
  }, [billings.data, properties.data, taxmaps.data]); // eslint-disable-line react-hooks/exhaustive-deps

  // Properties whose initial filing fee is marked N/A (not charged).
  const naFilingFees = useMemo(() => {
    if (!billings.data) return [];
    const propById = new Map((properties.data ?? []).map((p) => [String(p.id), p]));
    return billings.data
      .filter(isNAFilingFee)
      .map((b) => ({
        billing: b,
        property: b.fields.PropertyLookupId ? propById.get(String(b.fields.PropertyLookupId)) ?? null : null,
      }))
      .sort((a, b) => (a.property?.fields.Title ?? '').localeCompare(b.property?.fields.Title ?? ''));
  }, [billings.data, properties.data]);

  // Current (possibly-overridden) values for a row.
  const pctFor = (item: PercentQueueItem) =>
    pctOverrides[String(item.submittal.id)] ?? String(item.feePercent);
  const pctAmountFor = (item: PercentQueueItem) => {
    const p = parseFloat(pctFor(item));
    return isNaN(p) ? 0 : (item.taxSavings * p) / 100;
  };
  const feeFor = (item: FilingFeeQueueItem) =>
    feeOverrides[String(item.submittal.id)] ?? String(item.amount);

  const refetchAll = async () => {
    await Promise.all([submittals.refetch(), billings.refetch(), properties.refetch()]);
  };

  const runOnePercent = async (item: PercentQueueItem) => {
    const pct = parseFloat(pctFor(item));
    if (isNaN(pct) || pct <= 0 || pct > 100) {
      setError(`Enter a fee % between 0 and 100 for ${item.property?.fields.Title ?? 'this filing'}.`);
      return;
    }
    setError(null);
    setBusy(String(item.submittal.id));
    try {
      await generatePercentInvoice({
        submittal: item.submittal,
        property: item.property,
        taxSavings: item.taxSavings,
        feePercent: pct,
      });
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const runOneFiling = async (item: FilingFeeQueueItem) => {
    const amount = parseFloat(feeFor(item));
    if (isNaN(amount) || amount <= 0) {
      setError(`Enter a filing fee amount for ${item.property?.fields.Title ?? 'this property'}.`);
      return;
    }
    setError(null);
    setBusy(`filing-${item.submittal.id}`);
    try {
      await generateFilingFeeInvoice({ submittal: item.submittal, property: item.property, amount });
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  // Mark a property's initial filing fee as N/A (writes a $0 placeholder row so
  // the property leaves this queue and is recorded as intentionally not charged).
  const markOneNA = async (item: FilingFeeQueueItem) => {
    if (!item.property) {
      setError("Can't mark N/A — this filing isn't linked to a property.");
      return;
    }
    setError(null);
    setBusy(`na-${item.submittal.id}`);
    try {
      await markFilingFeeNA({ property: item.property, submittal: item.submittal });
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  // Undo an N/A by deleting its placeholder row (the property returns to the queue).
  const undoNA = async (billing: Billing) => {
    setError(null);
    setBusy(`undo-${billing.id}`);
    try {
      await deleteListItem(LIST_NAMES.Billing, String(billing.id));
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const clearRollEntry = (key: string) => {
    setRollFull((p) => { const n = { ...p }; delete n[key]; return n; });
    setRollRecent((p) => { const n = { ...p }; delete n[key]; return n; });
  };
  const rollScope = (e: RollUnit) => e.tmidLabel ? `${e.property.fields.Title ?? 'property'} · ${e.tmidLabel}` : (e.property.fields.Title ?? 'this property');

  // Generate a unit's % of savings invoice for the roll-forward target year.
  const rollFullValue = (e: RollUnit) => rollFull[e.key] ?? (e.lastFull != null ? String(e.lastFull) : '');

  const runRollForward = async (entry: RollUnit, year: CahpTaxYear) => {
    const full = parseFloat(rollFullValue(entry));
    const recent = parseFloat(rollRecent[entry.key] ?? '');
    const name = rollScope(entry);
    if (isNaN(full) || isNaN(recent)) { setError(`Enter both tax bills for ${name}.`); return; }
    if (full - recent <= 0) { setError(`Last full bill must exceed the most recent bill for ${name}.`); return; }
    if (billings.data && findPercentInvoiceForPropertyYear(entry.pid, year, billings.data, entry.tmid)) {
      setError(`${name} already has a TY ${year} % entry.`);
      return;
    }
    setError(null);
    setBusy(`roll-${entry.key}`);
    try {
      await recordAnnualPercentInvoice({ property: entry.property, taxYear: year, lastFullTaxBill: full, mostRecentTaxBill: recent, feePercent: entry.lastPct, taxMapId: entry.tmid || null });
      clearRollEntry(entry.key);
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  // Mark a unit's % as N/A for the roll-forward target year (not claimed).
  const markRollNA = async (entry: RollUnit, year: CahpTaxYear) => {
    if (billings.data && findPercentInvoiceForPropertyYear(entry.pid, year, billings.data, entry.tmid)) {
      setError(`${rollScope(entry)} already has a TY ${year} % entry.`);
      return;
    }
    setError(null);
    setBusy(`rollna-${entry.key}`);
    try {
      await markPercentNA({ property: entry.property, taxYear: year, taxMapId: entry.tmid || null });
      clearRollEntry(entry.key);
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const runAllPercent = async () => {
    if (!queues) return;
    // Validate every row's % up front so a bad value doesn't half-run the batch.
    for (const item of queues.percentItems) {
      const pct = parseFloat(pctFor(item));
      if (isNaN(pct) || pct <= 0 || pct > 100) {
        setError(`Fix the fee % for ${item.property?.fields.Title ?? 'a filing'} before generating all (must be 0–100).`);
        return;
      }
    }
    setError(null);
    setBusy('all');
    try {
      for (const item of queues.percentItems) {
        await generatePercentInvoice({
          submittal: item.submittal,
          property: item.property,
          taxSavings: item.taxSavings,
          feePercent: parseFloat(pctFor(item)),
        });
      }
      await refetchAll();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  if (submittals.loading || billings.loading || properties.loading) {
    return <Loading label="the invoicing queue" />;
  }
  if (submittals.error) return <ErrorBanner error={submittals.error} />;
  if (!queues) return null;

  const { percentItems, filingItems } = queues;
  const percentTotal = percentItems.reduce((sum, i) => sum + pctAmountFor(i), 0);
  const filingTotal = filingItems.reduce((sum, i) => sum + (parseFloat(feeFor(i)) || 0), 0);
  const nothingDue = percentItems.length === 0 && filingItems.length === 0;
  const rowInputClass =
    'w-24 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 disabled:opacity-50';

  return (
    <div>
      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Filing Fees Due" value={filingItems.length} accent={filingItems.length > 0 ? 'warning' : 'default'} />
        <KPI label="Filing Fees $" value={`$${filingTotal.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono />
        <KPI label="% of Savings Due" value={percentItems.length} accent={percentItems.length > 0 ? 'warning' : 'default'} />
        <KPI label="% of Savings $" value={`$${percentTotal.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono />
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">{error}</div>
      )}

      {nothingDue && (
        <div className="bg-green-50 border border-green-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-green-900 mb-1">All caught up</p>
          <p className="text-sm text-green-800">
            Every approved filing has its % of savings invoice, and every property with an initial filing has its filing fee invoiced.
          </p>
        </div>
      )}

      {/* Filing fees */}
      {filingItems.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden mb-6">
          <div className="px-4 py-3 border-b border-gray-200 flex items-center justify-between gap-2">
            <div>
              <h3 className="text-sm font-bold text-teal-700">Filing fees to invoice</h3>
              <p className="text-xs text-gray-500">One-time per property — properties with an approved initial filing but no filing-fee invoice.</p>
            </div>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Initial Filing</th>
                <th className="px-4 py-3 text-right">Filing Fee</th>
                <th className="px-4 py-3 text-right">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filingItems.map((item) => (
                <tr key={`filing-${item.submittal.id}`} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">{item.property?.fields.Title ?? '(unlinked)'}</td>
                  <td className="px-4 py-3 text-teal-700">
                    <button onClick={() => navigate(`/submittals/${item.submittal.id}`)} className="hover:underline">
                      {item.submittal.fields.Title ?? `Submittal ${item.submittal.id}`}
                    </button>
                  </td>
                  <td className="px-4 py-3 text-right">
                    <span className="text-gray-400 font-mono-data mr-0.5">$</span>
                    <input
                      type="number"
                      min="0"
                      step="0.01"
                      value={feeFor(item)}
                      onChange={(e) =>
                        setFeeOverrides((prev) => ({ ...prev, [String(item.submittal.id)]: e.target.value }))
                      }
                      disabled={busy !== null}
                      className={rowInputClass}
                    />
                  </td>
                  <td className="px-4 py-3 text-right whitespace-nowrap">
                    <button
                      onClick={() => markOneNA(item)}
                      disabled={busy !== null}
                      title="Not charging an initial filing fee for this property"
                      className="px-3 py-1 mr-1.5 rounded-md text-xs font-medium border border-gray-300 text-gray-600 hover:bg-gray-50 disabled:opacity-50"
                    >
                      {busy === `na-${item.submittal.id}` ? 'Marking…' : 'Mark N/A'}
                    </button>
                    <button
                      onClick={() => runOneFiling(item)}
                      disabled={busy !== null}
                      className="px-3 py-1 rounded-md text-xs font-medium bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50"
                    >
                      {busy === `filing-${item.submittal.id}` ? 'Generating…' : 'Generate'}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Filing fees marked N/A — not charged (with undo) */}
      {naFilingFees.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden mb-6">
          <div className="px-4 py-3 border-b border-gray-200">
            <h3 className="text-sm font-bold text-gray-600">Initial filing fee — N/A (not charged)</h3>
            <p className="text-xs text-gray-500">Properties intentionally not billed an initial filing fee. Undo to put one back in the queue above.</p>
          </div>
          <table className="w-full text-sm">
            <tbody className="divide-y divide-gray-100">
              {naFilingFees.map(({ billing, property }) => (
                <tr key={`na-${billing.id}`} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">{property?.fields.Title ?? '(unlinked)'}</td>
                  <td className="px-4 py-3">
                    <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-gray-100 text-gray-500">N/A</span>
                  </td>
                  <td className="px-4 py-3 text-right">
                    <button
                      onClick={() => undoNA(billing)}
                      disabled={busy !== null}
                      className="px-3 py-1 rounded-md text-xs font-medium border border-gray-300 text-gray-600 hover:bg-gray-50 disabled:opacity-50"
                    >
                      {busy === `undo-${billing.id}` ? 'Undoing…' : 'Undo'}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Annual % of savings — roll forward to the next tax year */}
      {rollForward.enrolled.length > 0 && (() => {
        const year = (rollYear || rollForward.suggestedYear) as CahpTaxYear;
        const rows = rollForward.enrolled.filter(
          (e) => !(billings.data && findPercentInvoiceForPropertyYear(e.pid, year, billings.data, e.tmid)),
        );
        return (
          <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden mb-6">
            <div className="px-4 py-3 border-b border-gray-200 flex items-center justify-between gap-2 flex-wrap">
              <div>
                <h3 className="text-sm font-bold text-teal-700">Annual % of savings — roll forward</h3>
                <p className="text-xs text-gray-500">Bill the CAHP % of each property's annual tax savings. Enter the year's savings, then Generate. Lists properties already on a % arrangement.</p>
              </div>
              <label className="flex items-center gap-1 text-xs text-gray-600">
                Tax year
                <select value={year} onChange={(e) => setRollYear(e.target.value as CahpTaxYear)} className="border border-gray-300 rounded px-2 py-1 font-mono-data bg-white">
                  {ROLL_TAX_YEARS.map((y) => <option key={y} value={y}>{y}</option>)}
                </select>
              </label>
            </div>
            {rows.length === 0 ? (
              <div className="px-4 py-4 text-xs text-gray-500">Every enrolled property already has a TY {year} % invoice.</div>
            ) : (
              <table className="w-full text-sm">
                <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
                  <tr>
                    <th className="px-4 py-3 text-left">Property</th>
                    <th className="px-4 py-3 text-left">Tax Map ID</th>
                    <th className="px-4 py-3 text-left">Last billed</th>
                    <th className="px-4 py-3 text-right">Fee %</th>
                    <th className="px-4 py-3 text-right">Last full bill</th>
                    <th className="px-4 py-3 text-right">Most recent bill</th>
                    <th className="px-4 py-3 text-right">CAHP fee</th>
                    <th className="px-4 py-3 text-right">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {rows.map((e) => {
                    const full = parseFloat(rollFullValue(e));
                    const recent = parseFloat(rollRecent[e.key] ?? '');
                    const fee = !isNaN(full) && !isNaN(recent) && full - recent > 0 ? ((full - recent) * e.lastPct) / 100 : null;
                    return (
                      <tr key={`roll-${e.key}`} className="hover:bg-gray-50">
                        <td className="px-4 py-3 font-medium text-gray-900">{e.property.fields.Title ?? '(unlinked)'}</td>
                        <td className="px-4 py-3 text-gray-600 text-xs">{e.tmidLabel ?? <span className="text-gray-400">Whole property</span>}</td>
                        <td className="px-4 py-3 text-gray-500 font-mono-data text-xs">{e.lastYear ?? '—'}</td>
                        <td className="px-4 py-3 text-right font-mono-data text-xs">{e.lastPct}%</td>
                        <td className="px-4 py-3 text-right">
                          <span className="text-gray-400 font-mono-data mr-0.5">$</span>
                          <input
                            type="number" min="0" step="0.01"
                            value={rollFullValue(e)}
                            onChange={(ev) => setRollFull((prev) => ({ ...prev, [e.key]: ev.target.value }))}
                            disabled={busy !== null}
                            className={rowInputClass}
                          />
                        </td>
                        <td className="px-4 py-3 text-right">
                          <span className="text-gray-400 font-mono-data mr-0.5">$</span>
                          <input
                            type="number" min="0" step="0.01"
                            value={rollRecent[e.key] ?? ''}
                            onChange={(ev) => setRollRecent((prev) => ({ ...prev, [e.key]: ev.target.value }))}
                            disabled={busy !== null}
                            className={rowInputClass}
                          />
                        </td>
                        <td className="px-4 py-3 text-right font-mono-data text-xs text-teal-700">
                          {fee != null ? `$${fee.toLocaleString(undefined, { maximumFractionDigits: 0 })}` : '—'}
                        </td>
                        <td className="px-4 py-3 text-right whitespace-nowrap">
                          <button
                            onClick={() => markRollNA(e, year)}
                            disabled={busy !== null}
                            title="Not claiming the CAHP % for this year"
                            className="px-3 py-1 mr-1.5 rounded-md text-xs font-medium border border-gray-300 text-gray-600 hover:bg-gray-50 disabled:opacity-50"
                          >
                            {busy === `rollna-${e.key}` ? 'Marking…' : 'N/A'}
                          </button>
                          <button
                            onClick={() => runRollForward(e, year)}
                            disabled={busy !== null}
                            className="px-3 py-1 rounded-md text-xs font-medium bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50"
                          >
                            {busy === `roll-${e.key}` ? 'Generating…' : `Bill ${year}`}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        );
      })()}

      {/* % of savings */}
      {percentItems.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-200 flex items-center justify-between gap-2 flex-wrap">
            <div>
              <h3 className="text-sm font-bold text-teal-700">% of savings to invoice</h3>
              <p className="text-xs text-gray-500">Per tax year — approved filings with no % of savings invoice yet. Defaults to {percentItems[0]?.feePercent ?? 20}% of the tax savings on file.</p>
            </div>
            <button
              onClick={runAllPercent}
              disabled={busy !== null}
              className="px-3 py-1.5 rounded-md text-xs font-semibold bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50"
            >
              {busy === 'all' ? 'Generating…' : `Generate all (${percentItems.length})`}
            </button>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Filing</th>
                <th className="px-4 py-3 text-left">Tax Year</th>
                <th className="px-4 py-3 text-right">Tax Savings</th>
                <th className="px-4 py-3 text-right">Fee %</th>
                <th className="px-4 py-3 text-right">Invoice</th>
                <th className="px-4 py-3 text-right">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {percentItems.map((item) => (
                <tr key={item.submittal.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">{item.property?.fields.Title ?? '(unlinked)'}</td>
                  <td className="px-4 py-3 text-teal-700">
                    <button onClick={() => navigate(`/submittals/${item.submittal.id}`)} className="hover:underline">
                      {item.submittal.fields.Title ?? `Submittal ${item.submittal.id}`}
                    </button>
                  </td>
                  <td className="px-4 py-3 font-mono-data text-xs text-gray-700">{item.taxYear ?? '—'}</td>
                  <td className="px-4 py-3 text-right font-mono-data text-gray-700">${item.taxSavings.toLocaleString()}</td>
                  <td className="px-4 py-3 text-right">
                    <input
                      type="number"
                      min="0"
                      max="100"
                      step="0.5"
                      value={pctFor(item)}
                      onChange={(e) =>
                        setPctOverrides((prev) => ({ ...prev, [String(item.submittal.id)]: e.target.value }))
                      }
                      disabled={busy !== null}
                      className={`${rowInputClass} w-16`}
                    />
                    <span className="text-gray-400 font-mono-data ml-0.5">%</span>
                  </td>
                  <td className="px-4 py-3 text-right font-mono-data font-semibold text-teal-700">
                    ${pctAmountFor(item).toLocaleString(undefined, { maximumFractionDigits: 2 })}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <button
                      onClick={() => runOnePercent(item)}
                      disabled={busy !== null}
                      className="px-3 py-1 rounded-md text-xs font-medium bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50"
                    >
                      {busy === String(item.submittal.id) ? 'Generating…' : 'Generate'}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Monthly Billing Tab — prorate each annual % invoice into a monthly CAHP bill
// (replicates the "CAHP Bill Backs" spreadsheet: Months, Monthly, invoice text,
//  monthly total, and the 2027+ ongoing projection)
// =============================================================================

function MonthlyBillingTab() {
  const billings = useSharePointList<Billing>(LIST_NAMES.Billing, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxmaps = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const [savingId, setSavingId] = useState<string | null>(null);
  const [copied, setCopied] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [showOngoing, setShowOngoing] = useState(false);

  const propById = useMemo(() => new Map((properties.data ?? []).map((p) => [String(p.id), p])), [properties.data]);
  const tmidById = useMemo(() => new Map((taxmaps.data ?? []).map((t) => [String(t.id), t])), [taxmaps.data]);

  type Row = {
    id: string; owner: string; property: string; tmidLabel: string | null; taxYear: string;
    pid: string; tmid: string; lastFull: number; mostRecent: number; savings: number;
    feePercent: number; annual: number; previouslyAbated: boolean; billStartDate: string;
    months: number; monthly: number; description: string;
  };

  const rows: Row[] = useMemo(() => {
    if (!billings.data) return [];
    return billings.data
      .filter((b) => isPercentInvoice(b) && !isNAPercent(b) && !isBaselineRow(b))
      .filter((b) => (b.fields.BillApprovedAbatement ?? 0) > 0 || (b.fields.AmountBilled ?? 0) > 0)
      .map((b): Row => {
        const f = b.fields;
        const pid = f.PropertyLookupId ? String(f.PropertyLookupId) : '';
        const tmid = f.BillTaxMapIDLookupId ? String(f.BillTaxMapIDLookupId) : '';
        const property = propById.get(pid);
        const owner = property?.fields.LegalEntity || property?.fields.Title || 'Owner';
        const lastFull = f.LastFullTaxBill ?? 0;
        const mostRecent = f.MostRecentTaxBill ?? 0;
        const savings = f.BillApprovedAbatement ?? Math.max(0, lastFull - mostRecent);
        const feePercent = f.CAHPFeePercent ?? DEFAULT_FEE_PERCENT;
        const annual = f.AmountBilled ?? (savings * feePercent) / 100;
        const previouslyAbated = f.PreviouslyAbated ?? false;
        const billStartDate = f.BillStartDate ? String(f.BillStartDate).slice(0, 10) : '';
        const months = computeBillingMonths(previouslyAbated, billStartDate);
        const monthly = monthlyFee(annual, months);
        const description = buildMonthlyInvoiceDescription({
          owner, taxYear: f.cahpTaxYear ?? '', previouslyAbated, lastFullTaxBill: lastFull,
          mostRecentTaxBill: mostRecent, totalSavings: savings, feePercent, annualFee: annual, months, monthly,
        });
        return {
          id: String(b.id), owner, property: property?.fields.Title ?? '—',
          tmidLabel: tmid ? (tmidById.get(tmid)?.fields.Title ?? `TMID ${tmid}`) : null,
          taxYear: String(f.cahpTaxYear ?? ''), pid, tmid, lastFull, mostRecent, savings,
          feePercent, annual, previouslyAbated, billStartDate, months, monthly, description,
        };
      })
      .sort((a, b) => a.owner.localeCompare(b.owner) || a.property.localeCompare(b.property) || b.taxYear.localeCompare(a.taxYear));
  }, [billings.data, propById, tmidById]);

  const monthlyTotal = rows.reduce((s, r) => s + r.monthly, 0);
  const annualTotal = rows.reduce((s, r) => s + r.annual, 0);

  // Ongoing 2027+: each parcel's latest tax year, billed the full 12 months.
  const ongoing = useMemo(() => {
    const byParcel = new Map<string, Row>();
    for (const r of rows) {
      const key = `${r.pid}|${r.tmid}`;
      const cur = byParcel.get(key);
      if (!cur || Number(r.taxYear) > Number(cur.taxYear)) byParcel.set(key, r);
    }
    const list = [...byParcel.values()].sort((a, b) => a.owner.localeCompare(b.owner) || a.property.localeCompare(b.property));
    const estAnnual = list.reduce((s, r) => s + r.annual, 0);
    return { list, estAnnual, estMonthly: estAnnual / 12 };
  }, [rows]);

  const usd = (n: number) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 });

  const saveRow = async (id: string, patch: { previouslyAbated?: boolean; billStartDate?: string }) => {
    setError(null); setSavingId(id);
    try {
      await updateBillingMonthlyInputs(id, {
        previouslyAbated: patch.previouslyAbated,
        billStartDate:
          patch.billStartDate !== undefined
            ? (patch.billStartDate ? new Date(patch.billStartDate + 'T00:00:00Z').toISOString() : null)
            : undefined,
      });
      billings.refetch();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSavingId(null);
    }
  };

  const copy = async (id: string, text: string) => {
    try { await navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1500); } catch { /* ignore */ }
  };

  if (billings.loading || properties.loading) return <div className="text-sm text-gray-500 p-4">Loading…</div>;

  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
        <KPI label="Billable Rows" value={rows.length} />
        <KPI label="Monthly Total" value={usd(monthlyTotal)} mono />
        <KPI label="Annual Total" value={usd(annualTotal)} mono />
        <KPI label="Ongoing Monthly (2027+)" value={usd(ongoing.estMonthly)} mono />
      </div>

      {error && <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">{error}</div>}

      <p className="text-sm text-gray-500 mb-3">
        Each approved % of savings invoice, prorated into a monthly CAHP bill. Set <strong>Prev. Abated</strong> (bill 12 months) or a <strong>Start Date</strong> (prorate the first year) per row — amounts come from the invoice already on file.
      </p>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto mb-6">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
            <tr>
              <th className="text-left px-3 py-2">Owner / Property</th>
              <th className="text-left px-3 py-2">Tax Yr</th>
              <th className="text-right px-3 py-2">Total Savings</th>
              <th className="text-right px-3 py-2">CAHP %</th>
              <th className="text-right px-3 py-2">Annual</th>
              <th className="text-center px-3 py-2">Prev. Abated</th>
              <th className="text-left px-3 py-2">Start Date</th>
              <th className="text-right px-3 py-2">Months</th>
              <th className="text-right px-3 py-2">Monthly</th>
              <th className="px-3 py-2"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {rows.map((r) => (
              <tr key={r.id} className={savingId === r.id ? 'opacity-60' : ''}>
                <td className="px-3 py-2">
                  <div className="font-medium text-gray-800">{r.owner}</div>
                  <div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div>
                </td>
                <td className="px-3 py-2">{r.taxYear}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.savings)}</td>
                <td className="px-3 py-2 text-right">{r.feePercent}%</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                <td className="px-3 py-2 text-center">
                  <input type="checkbox" checked={r.previouslyAbated} disabled={savingId === r.id}
                    onChange={(e) => saveRow(r.id, { previouslyAbated: e.target.checked })} />
                </td>
                <td className="px-3 py-2">
                  <input type="date" value={r.billStartDate} disabled={r.previouslyAbated || savingId === r.id}
                    className="border border-gray-300 rounded px-2 py-1 text-xs disabled:bg-gray-100"
                    onChange={(e) => saveRow(r.id, { billStartDate: e.target.value })} />
                </td>
                <td className="px-3 py-2 text-right">{r.months}</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.monthly)}</td>
                <td className="px-3 py-2 text-right whitespace-nowrap">
                  <button className="text-xs text-teal-700 hover:underline" onClick={() => copy(r.id, r.description)}>
                    {copied === r.id ? 'Copied!' : 'Copy invoice text'}
                  </button>
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={10} className="px-3 py-8 text-center text-gray-400">No % of savings invoices yet — generate them in the “To Invoice” tab.</td></tr>
            )}
          </tbody>
          {rows.length > 0 && (
            <tfoot className="bg-gray-50 font-semibold text-gray-800">
              <tr>
                <td className="px-3 py-2" colSpan={8}>Monthly Total</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(monthlyTotal)}</td>
                <td></td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      <button className="text-sm text-teal-700 hover:underline mb-3" onClick={() => setShowOngoing((v) => !v)}>
        {showOngoing ? '▾ Hide' : '▸ Show'} Estimated Ongoing Monthly Billing (2027+)
      </button>
      {showOngoing && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
          <div className="px-4 py-2 text-sm text-gray-600 border-b border-gray-200">
            Full 12-month billing from each parcel’s latest tax year — est. annual {usd(ongoing.estAnnual)} · monthly {usd(ongoing.estMonthly)}.
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
              <tr>
                <th className="text-left px-3 py-2">Owner / Property</th>
                <th className="text-left px-3 py-2">Basis Yr</th>
                <th className="text-right px-3 py-2">Est. Annual CAHP</th>
                <th className="text-right px-3 py-2">Est. Monthly</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {ongoing.list.map((r) => (
                <tr key={`ong-${r.id}`}>
                  <td className="px-3 py-2">
                    <div className="font-medium text-gray-800">{r.owner}</div>
                    <div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div>
                  </td>
                  <td className="px-3 py-2">{r.taxYear}</td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual / 12)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function TabButton({
  active,
  onClick,
  children,
}: {
  active: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      className={`px-4 py-2 text-sm font-medium border-b-2 transition-colors whitespace-nowrap ${
        active
          ? 'text-teal-700 border-teal-700'
          : 'text-gray-600 hover:text-teal-700 border-transparent hover:border-gray-300'
      }`}
    >
      {children}
    </button>
  );
}

// =============================================================================
// CAHP Fee Invoices Tab
// =============================================================================

function InvoicesTab() {
  const navigate = useNavigate();
  const billings = useSharePointList<Billing>(LIST_NAMES.Billing, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<BillingStatusValue | 'All'>('All');

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const filtered = useMemo(() => {
    if (!billings.data) return [];
    return billings.data
      .filter((b) => {
        const f = b.fields;
        if (search) {
          const propName = f.PropertyLookupId
            ? propertiesById.get(String(f.PropertyLookupId))?.fields.Title ?? ''
            : '';
          const hay = `${f.Title ?? ''} ${propName} ${f.InvoiceNumber ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (statusFilter !== 'All' && f.BillingStatus !== statusFilter) return false;
        return true;
      })
      .sort((a, b) => {
        const da = a.fields.InvoiceDate ? new Date(a.fields.InvoiceDate).getTime() : 0;
        const db = b.fields.InvoiceDate ? new Date(b.fields.InvoiceDate).getTime() : 0;
        return db - da;
      });
  }, [billings.data, search, statusFilter, propertiesById]);

  const stats = useMemo(() => {
    if (!billings.data) return null;
    const thisYear = new Date().getFullYear();
    const ytdBilled = billings.data
      .filter((b) => {
        if (!b.fields.InvoiceDate) return false;
        return new Date(b.fields.InvoiceDate).getFullYear() === thisYear;
      })
      .reduce((sum, b) => sum + (b.fields.AmountBilled ?? 0), 0);
    const sumForStatus = (status: BillingStatusValue) =>
      billings.data!
        .filter((b) => b.fields.BillingStatus === status)
        .reduce((sum, b) => sum + (b.fields.AmountBilled ?? 0), 0);
    const readyCount = billings.data.filter((b) => b.fields.BillingStatus === 'Ready to Invoice').length;
    const invoicedCount = billings.data.filter((b) => b.fields.BillingStatus === 'Invoiced').length;
    const readyDollars = sumForStatus('Ready to Invoice');
    const pendingApprovalDollars = sumForStatus('Pending Approval');
    const outstanding = sumForStatus('Invoiced');
    return { ytdBilled, readyCount, invoicedCount, readyDollars, pendingApprovalDollars, outstanding };
  }, [billings.data]);

  if (billings.loading || properties.loading) {
    return <Loading label="invoices" />;
  }
  if (billings.error) {
    return <ErrorBanner error={billings.error} />;
  }
  if (!billings.data || !stats) return null;

  return (
    <div>
      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-6">
        <KPI label="YTD Billed" value={`$${stats.ytdBilled.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono />
        <KPI label="Pending Approval" value={`$${stats.pendingApprovalDollars.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono accent={stats.pendingApprovalDollars > 0 ? 'warning' : 'default'} />
        <KPI label="Ready to Invoice" value={`$${stats.readyDollars.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono accent={stats.readyDollars > 0 ? 'warning' : 'default'} />
        <KPI label="Ready to Invoice (count)" value={stats.readyCount} />
        <KPI label="Invoiced (count)" value={stats.invoicedCount} />
        <KPI label="Outstanding Receivable" value={`$${stats.outstanding.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono accent={stats.outstanding > 0 ? 'warning' : 'default'} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search property, invoice #…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as BillingStatusValue | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All statuses</option>
          {(Object.keys(BILLING_STATUS_STYLES) as BillingStatusValue[]).map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </select>
        {filtered.length !== billings.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {billings.data.length}</span>
        )}
      </div>

      {/* Table */}
      {billings.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No invoices yet</p>
          <p className="text-sm text-blue-800">
            CAHP fee invoices are created from an Approved submittal — open the submittal and click <strong>Generate Invoice</strong> once accounting has confirmed the numbers.
          </p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Year</th>
                <th className="px-4 py-3 text-right">Tax Savings</th>
                <th className="px-4 py-3 text-right">Fee %</th>
                <th className="px-4 py-3 text-right">CAHP Fee</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">QB</th>
                <th className="px-4 py-3 text-left">Invoice #</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((b) => {
                const property = b.fields.PropertyLookupId
                  ? propertiesById.get(String(b.fields.PropertyLookupId))
                  : null;
                return (
                  <tr
                    key={b.id}
                    onClick={() => navigate(`/billing/invoices/${b.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">
                      {property?.fields.Title ?? <span className="text-gray-400 italic">(unlinked)</span>}
                    </td>
                    <td className="px-4 py-3 text-xs">
                      {b.fields.BillingType === 'Filing Fee' ? (
                        <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-purple-100 text-purple-800">Filing Fee</span>
                      ) : b.fields.BillingType === 'Percent of Savings' ? (
                        <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-teal-100 text-teal-800">% of Savings</span>
                      ) : (
                        <span className="text-gray-400">—</span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{b.fields.cahpTaxYear ?? '—'}</td>
                    <td className="px-4 py-3 text-right font-mono-data">
                      {b.fields.BillApprovedAbatement ? `$${b.fields.BillApprovedAbatement.toLocaleString()}` : '—'}
                    </td>
                    <td className="px-4 py-3 text-right font-mono-data text-xs">
                      {b.fields.CAHPFeePercent != null ? `${b.fields.CAHPFeePercent}%` : '—'}
                    </td>
                    <td className="px-4 py-3 text-right font-mono-data font-semibold text-teal-700">
                      {b.fields.AmountBilled ? `$${b.fields.AmountBilled.toLocaleString()}` : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {b.fields.BillingStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${BILLING_STATUS_STYLES[b.fields.BillingStatus]}`}>
                          {b.fields.BillingStatus}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {b.fields.QBSyncStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${QB_SYNC_STYLES[b.fields.QBSyncStatus]}`}>
                          {b.fields.QBSyncStatus}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{b.fields.InvoiceNumber || '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Refund Disbursements Tab
// =============================================================================

function DisbursementsTab() {
  const navigate = useNavigate();
  const disbursements = useSharePointList<Disbursement>(LIST_NAMES.Disbursements, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<DisbursementStatus | 'All'>('All');

  const propertiesById = useMemo(() => {
    if (!properties.data) return new Map<string, Property>();
    return new Map(properties.data.map((p) => [String(p.id), p]));
  }, [properties.data]);

  const filtered = useMemo(() => {
    if (!disbursements.data) return [];
    return disbursements.data
      .filter((d) => {
        const f = d.fields;
        if (search) {
          const propName = f.DisbPropertyLookupId
            ? propertiesById.get(String(f.DisbPropertyLookupId))?.fields.Title ?? ''
            : '';
          const hay = `${f.Title ?? ''} ${propName} ${f.DisbCheckNum ?? ''}`.toLowerCase();
          if (!hay.includes(search.toLowerCase())) return false;
        }
        if (statusFilter !== 'All' && f.DisbStatus !== statusFilter) return false;
        return true;
      })
      .sort((a, b) => {
        const da = a.fields.DisbIssueDate ? new Date(a.fields.DisbIssueDate).getTime() : 0;
        const db = b.fields.DisbIssueDate ? new Date(b.fields.DisbIssueDate).getTime() : 0;
        return db - da;
      });
  }, [disbursements.data, search, statusFilter, propertiesById]);

  const stats = useMemo(() => {
    if (!disbursements.data) return null;
    const pending = disbursements.data.filter((d) => d.fields.DisbStatus === 'Pending');
    const issued = disbursements.data.filter((d) => d.fields.DisbStatus === 'Issued');
    const cleared = disbursements.data.filter((d) => d.fields.DisbStatus === 'Cleared');
    const owedToOwners = pending.reduce((sum, d) => sum + (d.fields.DisbAmount ?? 0), 0);
    const thisYear = new Date().getFullYear();
    const ytdPaid = [...issued, ...cleared]
      .filter((d) => d.fields.DisbIssueDate && new Date(d.fields.DisbIssueDate).getFullYear() === thisYear)
      .reduce((sum, d) => sum + (d.fields.DisbAmount ?? 0), 0);
    return {
      owedToOwners,
      pendingCount: pending.length,
      issuedCount: issued.length,
      ytdPaid,
    };
  }, [disbursements.data]);

  if (disbursements.loading || properties.loading) {
    return <Loading label="disbursements" />;
  }
  if (disbursements.error) {
    return <ErrorBanner error={disbursements.error} />;
  }
  if (!disbursements.data || !stats) return null;

  return (
    <div>
      {/* KPIs */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI
          label="Owed to Owners"
          value={`$${stats.owedToOwners.toLocaleString(undefined, { maximumFractionDigits: 0 })}`}
          mono
          accent={stats.owedToOwners > 0 ? 'warning' : 'default'}
        />
        <KPI label="Pending" value={stats.pendingCount} accent={stats.pendingCount > 0 ? 'warning' : 'default'} />
        <KPI label="Issued" value={stats.issuedCount} />
        <KPI label="YTD Paid" value={`$${stats.ytdPaid.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search property, check #…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as DisbursementStatus | 'All')}
          className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
        >
          <option value="All">All statuses</option>
          {(Object.keys(DISB_STATUS_STYLES) as DisbursementStatus[]).map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </select>
        {filtered.length !== disbursements.data.length && (
          <span className="text-xs text-gray-500 px-1">{filtered.length} of {disbursements.data.length}</span>
        )}
      </div>

      {/* Table */}
      {disbursements.data.length === 0 ? (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No disbursements yet</p>
          <p className="text-sm text-blue-800">
            Disbursement records are created alongside the invoice when you click Generate Invoice on an Approved submittal.
          </p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-right">Amount</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Issue Date</th>
                <th className="px-4 py-3 text-left">Clear Date</th>
                <th className="px-4 py-3 text-left">Check #</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((d) => {
                const property = d.fields.DisbPropertyLookupId
                  ? propertiesById.get(String(d.fields.DisbPropertyLookupId))
                  : null;
                return (
                  <tr
                    key={d.id}
                    onClick={() => navigate(`/billing/disbursements/${d.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">
                      {property?.fields.Title ?? <span className="text-gray-400 italic">(unlinked)</span>}
                    </td>
                    <td className="px-4 py-3 text-right font-mono-data font-semibold text-teal-700">
                      {d.fields.DisbAmount ? `$${d.fields.DisbAmount.toLocaleString()}` : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {d.fields.DisbStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${DISB_STATUS_STYLES[d.fields.DisbStatus]}`}>
                          {d.fields.DisbStatus}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {d.fields.DisbIssueDate ? new Date(d.fields.DisbIssueDate).toLocaleDateString() : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {d.fields.DisbClearDate ? new Date(d.fields.DisbClearDate).toLocaleDateString() : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{d.fields.DisbCheckNum || '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Reconciliation Tab — QuickBooks integration deferred per Phase 3 scope decision
// =============================================================================

function ReconciliationTab() {
  return (
    <div>
      <div className="bg-gold-50 border-2 border-gold-200 rounded-lg p-6">
        <div className="flex items-start gap-3">
          <Icon name="alert" size={20} className="text-gold-700 flex-shrink-0 mt-0.5" />
          <div>
            <h3 className="text-base font-bold text-teal-900 mb-1">QuickBooks integration deferred</h3>
            <p className="text-sm text-gray-700 mb-3">
              Per the Phase 3 scope decision, QuickBooks API integration (push CAHP fee invoices, sync payment status,
              discrepancy detection) is on hold until QB Online + developer account credentials are sorted.
            </p>
            <p className="text-sm text-gray-700 mb-3">
              <strong>Manual workflow for now:</strong> When a CAHP fee invoice shows status <em>Ready to Invoice</em>,
              push it to QuickBooks manually. Then come back here and update the Invoice # and flip QB status to <em>Synced</em>
              on the invoice detail page.
            </p>
            <p className="text-xs text-gray-500 mt-4">
              Spec ref: §7.5. Reactivate this tab in a future PR when QB credentials are ready — the reconciliation
              dashboard will compare CAHP-side invoice totals against QB-reported AR balances and surface discrepancies.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

// =============================================================================
// Shared helpers
// =============================================================================

function Loading({ label }: { label: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
      <div className="inline-flex items-center gap-3 text-gray-500">
        <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
        <span className="text-sm">Loading {label}…</span>
      </div>
    </div>
  );
}

function ErrorBanner({ error }: { error: Error }) {
  return (
    <div className="bg-red-50 border border-red-200 rounded-lg p-6">
      <div className="font-semibold text-error mb-2">Failed to load</div>
      <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
    </div>
  );
}

function KPI({
  label,
  value,
  accent = 'default',
  mono = false,
}: {
  label: string;
  value: string | number;
  accent?: 'default' | 'warning' | 'danger';
  mono?: boolean;
}) {
  const accentClass =
    accent === 'danger' ? 'text-error' :
    accent === 'warning' ? 'text-warning' :
    'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass} ${mono ? 'font-mono-data text-2xl' : ''}`}>
        {value}
      </div>
    </div>
  );
}
