import { useState, useMemo, type ReactNode } from 'react';
import {
  useSharePointList,
  LIST_NAMES,
  type Billing,
  type Property,
  type Submittal,
  type TaxMapID,
  type CahpTaxYear,
  type RefundStatusValue,
} from '../lib/sharepoint';
import {
  isPercentInvoice,
  isNAPercent,
  isBaselineRow,
  computeInvoiceQueues,
  recordAnnualPercentInvoice,
  computeBillingMonths,
  monthlyFee,
  buildMonthlyInvoiceDescription,
  updateBillingMonthlyInputs,
  updateBillingRecord,
  updateRefundStatus,
  DEFAULT_FEE_PERCENT,
} from '../lib/billing';

// Tracking-only: billing is done in QuickBooks. This page records what each
// owner's CAHP fee SHOULD be (annual + amortized monthly) and tracks the
// previously-paid-tax refund status. No invoices are generated here.

const REFUND_STATES: RefundStatusValue[] = ['Needed', 'Requested', 'Approved & Sent', 'No Request Needed'];
const REFUND_STYLES: Record<string, string> = {
  '': 'bg-gray-100 text-gray-500',
  'Needed': 'bg-amber-100 text-amber-800',
  'Requested': 'bg-blue-100 text-blue-800',
  'Approved & Sent': 'bg-green-100 text-green-800',
  'No Request Needed': 'bg-gray-100 text-gray-600',
};
const usd = (n: number) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 });
const usd0 = (n: number) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });

type Rec = {
  id: string; pid: string; tmid: string; owner: string; property: string; tmidLabel: string | null;
  taxYear: string; lastFull: number; mostRecent: number; savings: number; feePercent: number;
  annual: number; previouslyAbated: boolean; billStartDate: string; months: number; monthly: number;
  refundStatus: RefundStatusValue | ''; description: string;
};

type Tab = 'abatements' | 'annual' | 'monthly' | 'ongoing';
const NEXT_TAX_YEAR = '2027';

/** Shared data + the unified list of recorded abatements every tab reads from. */
function useBillingData() {
  const billings = useSharePointList<Billing>(LIST_NAMES.Billing, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxmaps = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const propById = useMemo(() => new Map((properties.data ?? []).map((p) => [String(p.id), p])), [properties.data]);
  const tmidById = useMemo(() => new Map((taxmaps.data ?? []).map((t) => [String(t.id), t])), [taxmaps.data]);

  const records = useMemo<Rec[]>(() => {
    if (!billings.data) return [];
    return billings.data
      .filter((b) => isPercentInvoice(b) && !isNAPercent(b) && !isBaselineRow(b))
      .filter((b) => (b.fields.BillApprovedAbatement ?? 0) > 0 || (b.fields.AmountBilled ?? 0) > 0)
      .map((b): Rec => {
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
          id: String(b.id), pid, tmid, owner, property: property?.fields.Title ?? '—',
          tmidLabel: tmid ? (tmidById.get(tmid)?.fields.Title ?? `TMID ${tmid}`) : null,
          taxYear: String(f.cahpTaxYear ?? ''), lastFull, mostRecent, savings, feePercent, annual,
          previouslyAbated, billStartDate, months, monthly,
          refundStatus: (f.RefundStatus ?? '') as RefundStatusValue | '', description,
        };
      })
      .sort((a, b) => a.owner.localeCompare(b.owner) || a.property.localeCompare(b.property) || b.taxYear.localeCompare(a.taxYear));
  }, [billings.data, propById, tmidById]);

  return { billings, properties, taxmaps, submittals, propById, records };
}

export function BillingPage() {
  const [tab, setTab] = useState<Tab>('abatements');
  const data = useBillingData();

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">CAHP Fee Tracking</h1>
        <p className="text-sm text-gray-500 mt-1">
          What each owner’s CAHP fee should be (annual + amortized monthly) and the status of any previously-paid-tax refund. Billing is done in QuickBooks — this is for tracking and invoice text only.
        </p>
      </div>

      <div className="border-b border-gray-200 mb-4 flex gap-1 overflow-x-auto">
        <TabButton active={tab === 'abatements'} onClick={() => setTab('abatements')}>Abatements (Data)</TabButton>
        <TabButton active={tab === 'annual'} onClick={() => setTab('annual')}>Annual</TabButton>
        <TabButton active={tab === 'monthly'} onClick={() => setTab('monthly')}>Monthly</TabButton>
        <TabButton active={tab === 'ongoing'} onClick={() => setTab('ongoing')}>Estimated Ongoing</TabButton>
      </div>

      {(data.billings.loading || data.properties.loading) ? (
        <div className="text-sm text-gray-500 p-4">Loading…</div>
      ) : (
        <>
          {tab === 'abatements' && <AbatementsTab data={data} />}
          {tab === 'annual' && <AnnualTab data={data} />}
          {tab === 'monthly' && <MonthlyTab data={data} />}
          {tab === 'ongoing' && <OngoingTab data={data} />}
        </>
      )}
    </div>
  );
}

type DataCtx = ReturnType<typeof useBillingData>;

// =============================================================================
// Abatements (Data) — enter/edit the original info; record new approved abatements
// =============================================================================
function AbatementsTab({ data }: { data: DataCtx }) {
  const { billings, properties, submittals, records } = data;
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [copied, setCopied] = useState<string | null>(null);
  const [draft, setDraft] = useState<Record<string, { lastFull: string; mostRecent: string; fee: string }>>({});
  // New-record entry, keyed by submittal id.
  const [newFull, setNewFull] = useState<Record<string, string>>({});
  const [newRecent, setNewRecent] = useState<Record<string, string>>({});
  const [newFee, setNewFee] = useState<Record<string, string>>({});

  const toRecord = useMemo(() => {
    if (!submittals.data || !billings.data || !properties.data) return [];
    return computeInvoiceQueues(submittals.data, billings.data, properties.data).percentItems;
  }, [submittals.data, billings.data, properties.data]);

  const draftFor = (r: Rec) => draft[r.id] ?? { lastFull: String(r.lastFull || ''), mostRecent: String(r.mostRecent || ''), fee: String(r.feePercent) };
  const setD = (id: string, patch: Partial<{ lastFull: string; mostRecent: string; fee: string }>) =>
    setDraft((p) => ({ ...p, [id]: { ...(p[id] ?? draftFor(records.find((x) => x.id === id)!)), ...patch } }));

  const saveNumbers = async (r: Rec) => {
    const d = draftFor(r);
    const lastFull = parseFloat(d.lastFull), mostRecent = parseFloat(d.mostRecent), fee = parseFloat(d.fee);
    if (isNaN(lastFull) || isNaN(mostRecent) || isNaN(fee)) { setError('Enter valid numbers.'); return; }
    setError(null); setBusy(r.id);
    try {
      await updateBillingRecord(r.id, { lastFullTaxBill: lastFull, mostRecentTaxBill: mostRecent, feePercent: fee });
      setDraft((p) => { const n = { ...p }; delete n[r.id]; return n; });
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };

  const saveInput = async (id: string, patch: { previouslyAbated?: boolean; billStartDate?: string; refundStatus?: string }) => {
    setError(null); setBusy(id);
    try {
      if (patch.previouslyAbated !== undefined || patch.billStartDate !== undefined) {
        await updateBillingMonthlyInputs(id, {
          previouslyAbated: patch.previouslyAbated,
          billStartDate: patch.billStartDate !== undefined ? (patch.billStartDate ? new Date(patch.billStartDate + 'T00:00:00Z').toISOString() : null) : undefined,
        });
      }
      if (patch.refundStatus !== undefined) await updateRefundStatus(id, patch.refundStatus);
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };

  const recordNew = async (item: typeof toRecord[number]) => {
    const sid = String(item.submittal.id);
    const full = parseFloat(newFull[sid] ?? ''), recent = parseFloat(newRecent[sid] ?? '');
    const fee = parseFloat(newFee[sid] ?? String(DEFAULT_FEE_PERCENT));
    if (!item.property) { setError("This abatement isn't linked to a property."); return; }
    if (isNaN(full) || isNaN(recent)) { setError('Enter both tax bills.'); return; }
    if (full - recent <= 0) { setError('Last full bill must exceed the most recent bill.'); return; }
    setError(null); setBusy(`new-${sid}`);
    try {
      await recordAnnualPercentInvoice({
        property: item.property, taxYear: (item.taxYear ?? '') as CahpTaxYear,
        lastFullTaxBill: full, mostRecentTaxBill: recent, feePercent: isNaN(fee) ? DEFAULT_FEE_PERCENT : fee, taxMapId: null,
      });
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };

  const copy = async (id: string, text: string) => {
    try { await navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1500); } catch { /* ignore */ }
  };

  const inputCls = 'w-24 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data disabled:opacity-50';

  return (
    <div>
      {error && <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">{error}</div>}
      <p className="text-sm text-gray-500 mb-3">
        Enter the original numbers per abatement — <strong>Last Full</strong> and <strong>Most Recent</strong> tax bills (savings = the difference), the <strong>CAHP %</strong>, whether it was <strong>Previously Abated</strong> (or a <strong>Start Date</strong> to prorate), and the <strong>Refund</strong> status. Copy the invoice text to share with accounting.
      </p>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto mb-6">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
            <tr>
              <th className="text-left px-3 py-2">Owner / Property</th>
              <th className="text-left px-3 py-2">Yr</th>
              <th className="text-right px-3 py-2">Last Full</th>
              <th className="text-right px-3 py-2">Most Recent</th>
              <th className="text-right px-3 py-2">CAHP %</th>
              <th className="text-right px-3 py-2">Savings</th>
              <th className="text-right px-3 py-2">Annual</th>
              <th className="text-center px-3 py-2">Prev.</th>
              <th className="text-left px-3 py-2">Start</th>
              <th className="text-left px-3 py-2">Refund</th>
              <th className="px-3 py-2"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {records.map((r) => {
              const d = draftFor(r);
              const dirty = draft[r.id] !== undefined;
              return (
                <tr key={r.id} className={busy === r.id ? 'opacity-60' : ''}>
                  <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div></td>
                  <td className="px-3 py-2">{r.taxYear}</td>
                  <td className="px-3 py-2 text-right"><input type="number" className={inputCls} value={d.lastFull} disabled={busy === r.id} onChange={(e) => setD(r.id, { lastFull: e.target.value })} /></td>
                  <td className="px-3 py-2 text-right"><input type="number" className={inputCls} value={d.mostRecent} disabled={busy === r.id} onChange={(e) => setD(r.id, { mostRecent: e.target.value })} /></td>
                  <td className="px-3 py-2 text-right"><input type="number" className="w-14 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data" value={d.fee} disabled={busy === r.id} onChange={(e) => setD(r.id, { fee: e.target.value })} />%</td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd0(r.savings)}</td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd0(r.annual)}</td>
                  <td className="px-3 py-2 text-center"><input type="checkbox" checked={r.previouslyAbated} disabled={busy === r.id} onChange={(e) => saveInput(r.id, { previouslyAbated: e.target.checked })} /></td>
                  <td className="px-3 py-2"><input type="date" className="border border-gray-300 rounded px-1 py-1 text-xs disabled:bg-gray-100" value={r.billStartDate} disabled={r.previouslyAbated || busy === r.id} onChange={(e) => saveInput(r.id, { billStartDate: e.target.value })} /></td>
                  <td className="px-3 py-2">
                    <select className={`border border-gray-300 rounded px-1 py-1 text-xs ${REFUND_STYLES[r.refundStatus]}`} value={r.refundStatus} disabled={busy === r.id} onChange={(e) => saveInput(r.id, { refundStatus: e.target.value })}>
                      <option value="">—</option>
                      {REFUND_STATES.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                  <td className="px-3 py-2 text-right whitespace-nowrap">
                    {dirty && <button className="text-xs bg-teal-700 text-white rounded px-2 py-1 mr-1" disabled={busy === r.id} onClick={() => saveNumbers(r)}>Save</button>}
                    <button className="text-xs text-teal-700 hover:underline" onClick={() => copy(r.id, r.description)}>{copied === r.id ? 'Copied!' : 'Invoice text'}</button>
                  </td>
                </tr>
              );
            })}
            {records.length === 0 && <tr><td colSpan={11} className="px-3 py-6 text-center text-gray-400">No abatements recorded yet — record one below.</td></tr>}
          </tbody>
        </table>
      </div>

      {toRecord.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
          <div className="px-4 py-2 text-sm font-semibold text-teal-700 border-b border-gray-200">Approved abatements to record ({toRecord.length})</div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
              <tr>
                <th className="text-left px-3 py-2">Property</th>
                <th className="text-left px-3 py-2">Yr</th>
                <th className="text-right px-3 py-2">Last Full</th>
                <th className="text-right px-3 py-2">Most Recent</th>
                <th className="text-right px-3 py-2">CAHP %</th>
                <th className="px-3 py-2"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {toRecord.map((item) => {
                const sid = String(item.submittal.id);
                return (
                  <tr key={sid} className={busy === `new-${sid}` ? 'opacity-60' : ''}>
                    <td className="px-3 py-2 font-medium text-gray-800">{item.property?.fields.Title ?? '(unlinked)'}</td>
                    <td className="px-3 py-2">{item.taxYear ?? '—'}</td>
                    <td className="px-3 py-2 text-right"><input type="number" className={inputCls} value={newFull[sid] ?? ''} onChange={(e) => setNewFull((p) => ({ ...p, [sid]: e.target.value }))} /></td>
                    <td className="px-3 py-2 text-right"><input type="number" className={inputCls} value={newRecent[sid] ?? ''} onChange={(e) => setNewRecent((p) => ({ ...p, [sid]: e.target.value }))} /></td>
                    <td className="px-3 py-2 text-right"><input type="number" className="w-14 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data" value={newFee[sid] ?? String(DEFAULT_FEE_PERCENT)} onChange={(e) => setNewFee((p) => ({ ...p, [sid]: e.target.value }))} />%</td>
                    <td className="px-3 py-2 text-right"><button className="text-xs bg-teal-700 text-white rounded px-3 py-1" disabled={busy !== null} onClick={() => recordNew(item)}>{busy === `new-${sid}` ? 'Recording…' : 'Record'}</button></td>
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
// Annual — the annual CAHP amount per abatement + refund status
// =============================================================================
function AnnualTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const total = records.reduce((s, r) => s + r.annual, 0);
  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        <KPI label="Abatements" value={records.length} />
        <KPI label="Annual CAHP Total" value={usd(total)} mono />
        <KPI label="Refunds Outstanding" value={records.filter((r) => r.refundStatus === 'Needed' || r.refundStatus === 'Requested').length} accent="warning" />
      </div>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
            <tr>
              <th className="text-left px-3 py-2">Owner / Property</th>
              <th className="text-left px-3 py-2">Tax Yr</th>
              <th className="text-right px-3 py-2">Total Savings</th>
              <th className="text-right px-3 py-2">CAHP %</th>
              <th className="text-right px-3 py-2">Annual CAHP</th>
              <th className="text-left px-3 py-2">Refund</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {records.map((r) => (
              <tr key={r.id}>
                <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div></td>
                <td className="px-3 py-2">{r.taxYear}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.savings)}</td>
                <td className="px-3 py-2 text-right">{r.feePercent}%</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.annual)}</td>
                <td className="px-3 py-2"><span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${REFUND_STYLES[r.refundStatus]}`}>{r.refundStatus || '—'}</span></td>
              </tr>
            ))}
          </tbody>
          {records.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={4}>Annual Total</td><td className="px-3 py-2 text-right font-mono-data">{usd(total)}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

// =============================================================================
// Monthly — the amortized monthly CAHP amount (prorated) + invoice text
// =============================================================================
function MonthlyTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const [copied, setCopied] = useState<string | null>(null);
  const total = records.reduce((s, r) => s + r.monthly, 0);
  const copy = async (id: string, text: string) => {
    try { await navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1500); } catch { /* ignore */ }
  };
  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        <KPI label="Abatements" value={records.length} />
        <KPI label="Monthly CAHP Total" value={usd(total)} mono />
        <KPI label="Annualized" value={usd(total * 12)} mono />
      </div>
      <p className="text-sm text-gray-500 mb-3">Annual CAHP amortized into a monthly bill. Set Previously Abated / Start Date on the <strong>Abatements</strong> tab to control proration.</p>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
            <tr>
              <th className="text-left px-3 py-2">Owner / Property</th>
              <th className="text-left px-3 py-2">Tax Yr</th>
              <th className="text-right px-3 py-2">Annual</th>
              <th className="text-right px-3 py-2">Months</th>
              <th className="text-right px-3 py-2">Monthly</th>
              <th className="px-3 py-2"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {records.map((r) => (
              <tr key={r.id}>
                <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div></td>
                <td className="px-3 py-2">{r.taxYear}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                <td className="px-3 py-2 text-right">{r.months}</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.monthly)}</td>
                <td className="px-3 py-2 text-right"><button className="text-xs text-teal-700 hover:underline" onClick={() => copy(r.id, r.description)}>{copied === r.id ? 'Copied!' : 'Invoice text'}</button></td>
              </tr>
            ))}
          </tbody>
          {records.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={4}>Monthly Total</td><td className="px-3 py-2 text-right font-mono-data">{usd(total)}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

// =============================================================================
// Estimated Ongoing — next tax year (2027+), most-recent data, full 12 months
// =============================================================================
function OngoingTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const list = useMemo(() => {
    const byParcel = new Map<string, Rec>();
    for (const r of records) {
      const key = `${r.pid}|${r.tmid}`;
      const cur = byParcel.get(key);
      if (!cur || Number(r.taxYear) > Number(cur.taxYear)) byParcel.set(key, r);
    }
    return [...byParcel.values()].sort((a, b) => a.owner.localeCompare(b.owner) || a.property.localeCompare(b.property));
  }, [records]);
  const estAnnual = list.reduce((s, r) => s + r.annual, 0);
  const estMonthly = estAnnual / 12;
  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        <KPI label={`Parcels (TY ${NEXT_TAX_YEAR}+)`} value={list.length} />
        <KPI label="Est. Annual CAHP" value={usd(estAnnual)} mono />
        <KPI label="Est. Monthly CAHP" value={usd(estMonthly)} mono />
      </div>
      <p className="text-sm text-gray-500 mb-3">Projected ongoing billing starting <strong>TY {NEXT_TAX_YEAR}</strong>, using each parcel’s most recent basis year at the full 12 months (no proration).</p>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
            <tr>
              <th className="text-left px-3 py-2">Owner / Property</th>
              <th className="text-left px-3 py-2">Basis Yr</th>
              <th className="text-right px-3 py-2">Total Savings</th>
              <th className="text-right px-3 py-2">Est. Annual CAHP</th>
              <th className="text-right px-3 py-2">Est. Monthly</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {list.map((r) => (
              <tr key={`ong-${r.id}`}>
                <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}{r.tmidLabel ? ` · ${r.tmidLabel}` : ''}</div></td>
                <td className="px-3 py-2">{r.taxYear}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.savings)}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.annual / 12)}</td>
              </tr>
            ))}
          </tbody>
          {list.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={3}>Totals</td><td className="px-3 py-2 text-right font-mono-data">{usd(estAnnual)}</td><td className="px-3 py-2 text-right font-mono-data">{usd(estMonthly)}</td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: ReactNode }) {
  return (
    <button
      onClick={onClick}
      className={`px-4 py-2 text-sm font-medium border-b-2 whitespace-nowrap -mb-px ${active ? 'border-teal-600 text-teal-700' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
    >
      {children}
    </button>
  );
}

function KPI({ label, value, accent = 'default', mono = false }: { label: string; value: string | number; accent?: 'default' | 'warning' | 'danger'; mono?: boolean }) {
  const accentClass = accent === 'danger' ? 'text-error' : accent === 'warning' ? 'text-warning' : 'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-2xl font-bold mt-1 ${accentClass} ${mono ? 'font-mono-data' : ''}`}>{value}</div>
    </div>
  );
}
