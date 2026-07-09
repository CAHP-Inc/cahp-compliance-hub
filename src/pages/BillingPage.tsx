import { useState, useMemo, type ReactNode } from 'react';
import {
  useSharePointList,
  deleteListItem,
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
  BILLING_PHASE_STATUSES,
  upsertAbatementRecord,
  assignBillingParcel,
  computeBillingMonths,
  monthlyFee,
  buildMonthlyInvoiceDescription,
  DEFAULT_FEE_PERCENT,
} from '../lib/billing';

// Tracking-only (billing is done in QuickBooks). ONE approval per parcel, but it
// produces a billable row PER TAX YEAR (each year has its own tax bills/amount).
// Approved parcels come from submittals; you add each year you need to bill.

const REFUND_STATES: RefundStatusValue[] = ['Needed', 'Requested', 'Approved & Sent', 'No Request Needed'];
const REFUND_STYLES: Record<string, string> = {
  '': 'bg-gray-100 text-gray-500', 'Needed': 'bg-amber-100 text-amber-800', 'Requested': 'bg-blue-100 text-blue-800',
  'Approved & Sent': 'bg-green-100 text-green-800', 'No Request Needed': 'bg-gray-100 text-gray-600',
};
const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];
const NEXT_TAX_YEAR = '2027';
const usd = (n: number) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 });
const usd0 = (n: number) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });

type Rec = {
  id: string; billingId: string; submittalId: string | null; pid: string; tmid: string;
  owner: string; property: string; parcel: string; taxYear: string;
  lastFull: number; mostRecent: number; savings: number; feePercent: number; annual: number;
  previouslyAbated: boolean; billStartDate: string; months: number; monthly: number;
  refundStatus: RefundStatusValue | ''; description: string;
};
type Parcel = { key: string; pid: string; tmid: string; submittalId: string; label: string; parcel: string };
type Tab = 'abatements' | 'annual' | 'monthly' | 'ongoing';

function useBillingData() {
  const billings = useSharePointList<Billing>(LIST_NAMES.Billing, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const taxmaps = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const propById = useMemo(() => new Map((properties.data ?? []).map((p) => [String(p.id), p])), [properties.data]);
  const tmidById = useMemo(() => new Map((taxmaps.data ?? []).map((t) => [String(t.id), t])), [taxmaps.data]);
  const submById = useMemo(() => new Map((submittals.data ?? []).map((s) => [String(s.id), s])), [submittals.data]);

  // Approved parcels (for the Add-a-year picker): billing-phase submittals with a TMID.
  const approvedParcels = useMemo<Parcel[]>(() => {
    const m = new Map<string, Parcel>();
    for (const s of submittals.data ?? []) {
      if (!(s.fields.SubmittalStatus && BILLING_PHASE_STATUSES.includes(s.fields.SubmittalStatus))) continue;
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      const tmid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
      if (!pid || !tmid) continue;
      const key = `${pid}|${tmid}`;
      if (m.has(key)) continue;
      const prop = propById.get(pid);
      const owner = prop?.fields.LegalEntity || prop?.fields.Title || '';
      m.set(key, { key, pid, tmid, submittalId: String(s.id), label: `${owner ? owner + ' — ' : ''}${prop?.fields.Title ?? ''}`, parcel: tmidById.get(tmid)?.fields.Title ?? `TMID ${tmid}` });
    }
    return [...m.values()].sort((a, b) => a.label.localeCompare(b.label) || a.parcel.localeCompare(b.parcel));
  }, [submittals.data, propById, tmidById]);

  const records = useMemo<Rec[]>(() => {
    if (!billings.data) return [];
    const resolveTmid = (b: Billing) =>
      (b.fields.BillTaxMapIDLookupId ? String(b.fields.BillTaxMapIDLookupId) : '')
      || (b.fields.BillSubmittalLookupId ? String(submById.get(String(b.fields.BillSubmittalLookupId))?.fields.TaxMapIDLookupId ?? '') : '');
    // One row per parcel + year: keep the most complete (entered bills > explicit TMID > newest).
    const score = (b: Billing) => (b.fields.LastFullTaxBill != null ? 4 : 0) + (b.fields.BillTaxMapIDLookupId ? 2 : 0);
    const best = new Map<string, Billing>();
    for (const b of billings.data) {
      if (!isPercentInvoice(b) || isNAPercent(b) || isBaselineRow(b)) continue;
      if (!((b.fields.BillApprovedAbatement ?? 0) > 0 || (b.fields.AmountBilled ?? 0) > 0)) continue;
      const key = `${b.fields.PropertyLookupId ?? ''}|${resolveTmid(b)}|${b.fields.cahpTaxYear ?? ''}`;
      const cur = best.get(key);
      if (!cur || score(b) > score(cur) || (score(b) === score(cur) && Number(b.id) > Number(cur.id))) best.set(key, b);
    }
    // Drop a "whole property" (no-TMID) row when a real parceled row exists for the
    // same property+year — that's the stale duplicate.
    const kept = [...best.values()];
    const parceledPY = new Set(
      kept.filter((b) => resolveTmid(b) !== '').map((b) => `${b.fields.PropertyLookupId ?? ''}|${b.fields.cahpTaxYear ?? ''}`),
    );
    return kept
      .filter((b) => resolveTmid(b) !== '' || !parceledPY.has(`${b.fields.PropertyLookupId ?? ''}|${b.fields.cahpTaxYear ?? ''}`))
      .map((b): Rec => {
        const f = b.fields;
        const pid = f.PropertyLookupId ? String(f.PropertyLookupId) : '';
        const submittalId = f.BillSubmittalLookupId ? String(f.BillSubmittalLookupId) : null;
        const tmid = resolveTmid(b);
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
          id: String(b.id), billingId: String(b.id), submittalId, pid, tmid, owner,
          property: property?.fields.Title ?? '—',
          parcel: tmid ? (tmidById.get(tmid)?.fields.Title ?? `TMID ${tmid}`) : 'Whole property',
          taxYear: String(f.cahpTaxYear ?? ''), lastFull, mostRecent, savings, feePercent, annual,
          previouslyAbated, billStartDate, months, monthly,
          refundStatus: (f.RefundStatus ?? '') as RefundStatusValue | '', description,
        };
      })
      .sort((a, b) => a.owner.localeCompare(b.owner) || a.parcel.localeCompare(b.parcel) || b.taxYear.localeCompare(a.taxYear));
  }, [billings.data, propById, tmidById, submById]);

  return { billings, properties, taxmaps, submittals, propById, approvedParcels, records };
}
type DataCtx = ReturnType<typeof useBillingData>;

export function BillingPage() {
  const [tab, setTab] = useState<Tab>('abatements');
  const data = useBillingData();
  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">CAHP Fee Tracking</h1>
        <p className="text-sm text-gray-500 mt-1">
          One approval per parcel, billed per tax year (each year its own amount). Parcels come from approved submittals; add each year you need. Billing is done in QuickBooks — this is for tracking and invoice text only.
        </p>
      </div>
      <div className="border-b border-gray-200 mb-4 flex gap-1 overflow-x-auto">
        <TabButton active={tab === 'abatements'} onClick={() => setTab('abatements')}>Abatements (Data)</TabButton>
        <TabButton active={tab === 'annual'} onClick={() => setTab('annual')}>Annual</TabButton>
        <TabButton active={tab === 'monthly'} onClick={() => setTab('monthly')}>Current Year Monthly</TabButton>
        <TabButton active={tab === 'ongoing'} onClick={() => setTab('ongoing')}>Estimated Ongoing</TabButton>
      </div>
      {(data.billings.loading || data.properties.loading || data.submittals.loading) ? (
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

// =============================================================================
// Abatements (Data)
// =============================================================================
function AbatementsTab({ data }: { data: DataCtx }) {
  const { billings, taxmaps, approvedParcels, records } = data;
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [copied, setCopied] = useState<string | null>(null);
  const [draft, setDraft] = useState<Record<string, { lastFull: string; mostRecent: string; fee: string }>>({});
  const [addKey, setAddKey] = useState('');
  const [addYear, setAddYear] = useState<CahpTaxYear>('2025');
  const [addFull, setAddFull] = useState('');
  const [addRecent, setAddRecent] = useState('');
  const [addFee, setAddFee] = useState(String(DEFAULT_FEE_PERCENT));

  const parcelsForProp = (pid: string) => (taxmaps.data ?? []).filter((t) => String(t.fields.LinkedPropertyLookupId ?? '') === pid);
  const draftFor = (r: Rec) => draft[r.id] ?? { lastFull: String(r.lastFull || ''), mostRecent: String(r.mostRecent || ''), fee: String(r.feePercent) };
  const setD = (r: Rec, patch: Partial<{ lastFull: string; mostRecent: string; fee: string }>) => setDraft((p) => ({ ...p, [r.id]: { ...draftFor(r), ...patch } }));
  const base = (r: Rec) => ({ billingId: r.billingId, propertyId: r.pid, submittalId: r.submittalId, taxYear: r.taxYear, taxMapId: r.tmid || null });

  const saveNumbers = async (r: Rec) => {
    const d = draftFor(r);
    const lastFull = parseFloat(d.lastFull), mostRecent = parseFloat(d.mostRecent), fee = parseFloat(d.fee);
    if (isNaN(lastFull) || isNaN(mostRecent) || isNaN(fee)) { setError('Enter valid numbers.'); return; }
    if (lastFull - mostRecent < 0) { setError('Last full bill must be ≥ the most recent bill.'); return; }
    setError(null); setBusy(r.id);
    try {
      await upsertAbatementRecord({ ...base(r), lastFullTaxBill: lastFull, mostRecentTaxBill: mostRecent, feePercent: fee });
      setDraft((p) => { const n = { ...p }; delete n[r.id]; return n; });
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };
  const saveInput = async (r: Rec, patch: { previouslyAbated?: boolean; billStartDate?: string; refundStatus?: string }) => {
    setError(null); setBusy(r.id);
    try {
      await upsertAbatementRecord({ ...base(r), fallbackSavings: r.savings,
        previouslyAbated: patch.previouslyAbated,
        billStartDate: patch.billStartDate !== undefined ? (patch.billStartDate ? new Date(patch.billStartDate + 'T00:00:00Z').toISOString() : null) : undefined,
        refundStatus: patch.refundStatus });
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };
  const setParcel = async (r: Rec, tmid: string) => {
    if (!tmid) return;
    setError(null); setBusy(r.id);
    try { await assignBillingParcel(r.billingId, tmid); billings.refetch(); }
    catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };
  const removeRow = async (r: Rec) => {
    if (!window.confirm(`Delete the tracking record for ${r.owner} — ${r.parcel} (TY ${r.taxYear})? This removes it from all tabs. The submittal is not affected.`)) return;
    setError(null); setBusy(r.id);
    try { await deleteListItem(LIST_NAMES.Billing, r.billingId); billings.refetch(); }
    catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };
  const addRow = async () => {
    const parcel = approvedParcels.find((p) => p.key === addKey);
    const full = parseFloat(addFull), recent = parseFloat(addRecent), fee = parseFloat(addFee);
    if (!parcel) { setError('Pick an approved parcel.'); return; }
    if (isNaN(full) || isNaN(recent)) { setError('Enter both tax bills.'); return; }
    if (full - recent < 0) { setError('Last full bill must be ≥ the most recent bill.'); return; }
    if (records.some((r) => r.pid === parcel.pid && r.tmid === parcel.tmid && r.taxYear === addYear)) { setError('That parcel already has a row for that year.'); return; }
    setError(null); setBusy('add');
    try {
      await upsertAbatementRecord({ billingId: null, propertyId: parcel.pid, submittalId: parcel.submittalId, taxYear: addYear, taxMapId: parcel.tmid, lastFullTaxBill: full, mostRecentTaxBill: recent, feePercent: isNaN(fee) ? DEFAULT_FEE_PERCENT : fee });
      setAddFull(''); setAddRecent('');
      billings.refetch();
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); } finally { setBusy(null); }
  };
  const copy = async (id: string, text: string) => { try { await navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1500); } catch { /* ignore */ } };
  const numCls = 'w-24 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data disabled:opacity-50';
  const selCls = 'border border-gray-300 rounded px-2 py-1 text-sm bg-white';

  return (
    <div>
      {error && <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">{error}</div>}

      {/* Add a year for an approved parcel */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card p-3 mb-6">
        <div className="text-sm font-semibold text-teal-700 mb-2">Add a billable year</div>
        <div className="flex flex-wrap items-end gap-2 text-sm">
          <label className="flex flex-col gap-0.5">Approved parcel
            <select className={selCls} value={addKey} onChange={(e) => setAddKey(e.target.value)}>
              <option value="">Select…</option>
              {approvedParcels.map((p) => <option key={p.key} value={p.key}>{p.label} · {p.parcel}</option>)}
            </select>
          </label>
          <label className="flex flex-col gap-0.5">Tax Year
            <select className={selCls} value={addYear} onChange={(e) => setAddYear(e.target.value as CahpTaxYear)}>{TAX_YEARS.map((y) => <option key={y} value={y}>{y}</option>)}</select>
          </label>
          <label className="flex flex-col gap-0.5">Last Full<input type="number" className={numCls} value={addFull} onChange={(e) => setAddFull(e.target.value)} /></label>
          <label className="flex flex-col gap-0.5">Most Recent<input type="number" className={numCls} value={addRecent} onChange={(e) => setAddRecent(e.target.value)} /></label>
          <label className="flex flex-col gap-0.5">CAHP %<input type="number" className="w-16 px-2 py-1 border border-gray-300 rounded text-sm text-right font-mono-data" value={addFee} onChange={(e) => setAddFee(e.target.value)} /></label>
          <button className="bg-teal-700 text-white rounded px-4 py-1.5 text-sm font-medium disabled:opacity-50" disabled={busy !== null} onClick={addRow}>{busy === 'add' ? 'Adding…' : 'Add year'}</button>
        </div>
        {approvedParcels.length === 0 && <div className="text-xs text-gray-400 mt-1">No approved parcels found — a submittal must be in an approved status with a Tax Map ID set.</div>}
      </div>

      <p className="text-sm text-gray-500 mb-3">One row per parcel per tax year. Edit tax bills / CAHP %, set <strong>Prev.</strong> or a <strong>Start Date</strong> to prorate, and the <strong>Refund</strong> status. Copy invoice text for accounting.</p>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase"><tr>
            <th className="text-left px-3 py-2">Owner / Property</th><th className="text-left px-3 py-2">Parcel (TMID)</th><th className="text-left px-3 py-2">Yr</th>
            <th className="text-right px-3 py-2">Last Full</th><th className="text-right px-3 py-2">Most Recent</th><th className="text-right px-3 py-2">%</th>
            <th className="text-right px-3 py-2">Savings</th><th className="text-right px-3 py-2">Annual</th><th className="text-center px-3 py-2">Prev.</th>
            <th className="text-left px-3 py-2">Start</th><th className="text-left px-3 py-2">Refund</th><th className="px-3 py-2"></th>
          </tr></thead>
          <tbody className="divide-y divide-gray-100">
            {records.map((r) => {
              const d = draftFor(r);
              const dirty = draft[r.id] !== undefined;
              return (
                <tr key={r.id} className={busy === r.id ? 'opacity-60' : ''}>
                  <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}</div></td>
                  <td className="px-3 py-2 font-mono-data text-xs">
                    {r.tmid ? r.parcel : (
                      <select className="border border-amber-300 bg-amber-50 rounded px-1 py-0.5 text-xs" value="" disabled={busy === r.id} onChange={(e) => setParcel(r, e.target.value)}>
                        <option value="">Set parcel…</option>
                        {parcelsForProp(r.pid).map((t) => <option key={t.id} value={String(t.id)}>{t.fields.Title}</option>)}
                      </select>
                    )}
                  </td>
                  <td className="px-3 py-2">{r.taxYear}</td>
                  <td className="px-3 py-2 text-right"><input type="number" className={numCls} value={d.lastFull} disabled={busy === r.id} onChange={(e) => setD(r, { lastFull: e.target.value })} /></td>
                  <td className="px-3 py-2 text-right"><input type="number" className={numCls} value={d.mostRecent} disabled={busy === r.id} onChange={(e) => setD(r, { mostRecent: e.target.value })} /></td>
                  <td className="px-3 py-2 text-right"><input type="number" className="w-12 px-1 py-1 border border-gray-300 rounded text-sm text-right font-mono-data" value={d.fee} disabled={busy === r.id} onChange={(e) => setD(r, { fee: e.target.value })} /></td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd0(r.savings)}</td>
                  <td className="px-3 py-2 text-right font-mono-data">{usd0(r.annual)}</td>
                  <td className="px-3 py-2 text-center"><input type="checkbox" checked={r.previouslyAbated} disabled={busy === r.id} onChange={(e) => saveInput(r, { previouslyAbated: e.target.checked })} /></td>
                  <td className="px-3 py-2"><input type="date" className="border border-gray-300 rounded px-1 py-1 text-xs disabled:bg-gray-100" value={r.billStartDate} disabled={r.previouslyAbated || busy === r.id} onChange={(e) => saveInput(r, { billStartDate: e.target.value })} /></td>
                  <td className="px-3 py-2"><select className={`border border-gray-300 rounded px-1 py-1 text-xs ${REFUND_STYLES[r.refundStatus]}`} value={r.refundStatus} disabled={busy === r.id} onChange={(e) => saveInput(r, { refundStatus: e.target.value })}><option value="">—</option>{REFUND_STATES.map((s) => <option key={s} value={s}>{s}</option>)}</select></td>
                  <td className="px-3 py-2 text-right whitespace-nowrap sticky right-0 bg-white">{dirty && <button className="text-xs bg-teal-700 text-white rounded px-2 py-1 mr-1" disabled={busy === r.id} onClick={() => saveNumbers(r)}>Save</button>}<button className="text-xs text-teal-700 hover:underline" onClick={() => copy(r.id, r.description)}>{copied === r.id ? 'Copied!' : 'Invoice text'}</button><button className="text-xs text-red-600 hover:underline ml-2" disabled={busy === r.id} onClick={() => removeRow(r)}>Delete</button></td>
                </tr>
              );
            })}
            {records.length === 0 && <tr><td colSpan={12} className="px-3 py-6 text-center text-gray-400">No years recorded yet — add one above.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ParcelCols() {
  return (<><th className="text-left px-3 py-2">Owner / Property</th><th className="text-left px-3 py-2">Parcel (TMID)</th><th className="text-left px-3 py-2">Tax Yr</th></>);
}
function ParcelCells({ r }: { r: Rec }) {
  return (<><td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}</div></td><td className="px-3 py-2 font-mono-data text-xs">{r.parcel}</td><td className="px-3 py-2">{r.taxYear}</td></>);
}

function AnnualTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const total = records.reduce((s, r) => s + r.annual, 0);
  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        <KPI label="Parcel-Years" value={records.length} />
        <KPI label="Annual CAHP Total" value={usd(total)} mono />
        <KPI label="Refunds Outstanding" value={records.filter((r) => r.refundStatus === 'Needed' || r.refundStatus === 'Requested').length} accent="warning" />
      </div>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase"><tr><ParcelCols /><th className="text-right px-3 py-2">Total Savings</th><th className="text-right px-3 py-2">%</th><th className="text-right px-3 py-2">Annual CAHP</th><th className="text-left px-3 py-2">Refund</th></tr></thead>
          <tbody className="divide-y divide-gray-100">
            {records.map((r) => (
              <tr key={r.id}><ParcelCells r={r} />
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.savings)}</td>
                <td className="px-3 py-2 text-right">{r.feePercent}%</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.annual)}</td>
                <td className="px-3 py-2"><span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${REFUND_STYLES[r.refundStatus]}`}>{r.refundStatus || '—'}</span></td>
              </tr>
            ))}
          </tbody>
          {records.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={5}>Annual Total</td><td className="px-3 py-2 text-right font-mono-data">{usd(total)}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function MonthlyTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const [copied, setCopied] = useState<string | null>(null);
  // Current-year view: exclude parcel-years whose billing starts after 12/31 this year.
  const cutoff = `${new Date().getFullYear()}-12-31`;
  const rows = records.filter((r) => !r.billStartDate || r.billStartDate <= cutoff);
  const total = rows.reduce((s, r) => s + r.monthly, 0);
  const copy = async (id: string, text: string) => { try { await navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1500); } catch { /* ignore */ } };
  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        <KPI label="Parcel-Years (this yr)" value={rows.length} />
        <KPI label="Monthly CAHP Total" value={usd(total)} mono />
        <KPI label="Annualized" value={usd(total * 12)} mono />
      </div>
      <p className="text-sm text-gray-500 mb-3">Parcel-years currently being billed this year, annual CAHP amortized to a monthly amount. Years whose start date is after 12/31 this year are excluded until they begin. Set proration on the <strong>Abatements</strong> tab.</p>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase"><tr><ParcelCols /><th className="text-right px-3 py-2">Annual</th><th className="text-right px-3 py-2">Months</th><th className="text-right px-3 py-2">Monthly</th><th className="px-3 py-2"></th></tr></thead>
          <tbody className="divide-y divide-gray-100">
            {rows.map((r) => (
              <tr key={r.id}><ParcelCells r={r} />
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                <td className="px-3 py-2 text-right">{r.months}</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.monthly)}</td>
                <td className="px-3 py-2 text-right"><button className="text-xs text-teal-700 hover:underline" onClick={() => copy(r.id, r.description)}>{copied === r.id ? 'Copied!' : 'Invoice text'}</button></td>
              </tr>
            ))}
          </tbody>
          {rows.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={5}>Monthly Total</td><td className="px-3 py-2 text-right font-mono-data">{usd(total)}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function OngoingTab({ data }: { data: DataCtx }) {
  const { records } = data;
  const list = useMemo(() => {
    const byParcel = new Map<string, Rec>();
    for (const r of records) {
      const key = `${r.pid}|${r.tmid}`;
      const cur = byParcel.get(key);
      if (!cur || Number(r.taxYear) > Number(cur.taxYear)) byParcel.set(key, r);
    }
    return [...byParcel.values()].sort((a, b) => a.owner.localeCompare(b.owner) || a.parcel.localeCompare(b.parcel));
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
      <p className="text-sm text-gray-500 mb-3">Ongoing billing starting <strong>TY {NEXT_TAX_YEAR}</strong>, using each parcel's <strong>most recent</strong> basis year at the full 12 months (no proration).</p>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 text-gray-600 text-xs uppercase"><tr><th className="text-left px-3 py-2">Owner / Property</th><th className="text-left px-3 py-2">Parcel (TMID)</th><th className="text-left px-3 py-2">Basis Yr</th><th className="text-right px-3 py-2">Total Savings</th><th className="text-right px-3 py-2">Est. Annual CAHP</th><th className="text-right px-3 py-2">Est. Monthly</th></tr></thead>
          <tbody className="divide-y divide-gray-100">
            {list.map((r) => (
              <tr key={`ong-${r.id}`}>
                <td className="px-3 py-2"><div className="font-medium text-gray-800">{r.owner}</div><div className="text-xs text-gray-500">{r.property}</div></td>
                <td className="px-3 py-2 font-mono-data text-xs">{r.parcel}</td>
                <td className="px-3 py-2">{r.taxYear}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.savings)}</td>
                <td className="px-3 py-2 text-right font-mono-data">{usd(r.annual)}</td>
                <td className="px-3 py-2 text-right font-mono-data font-semibold">{usd(r.annual / 12)}</td>
              </tr>
            ))}
          </tbody>
          {list.length > 0 && <tfoot className="bg-gray-50 font-semibold text-gray-800"><tr><td className="px-3 py-2" colSpan={4}>Totals</td><td className="px-3 py-2 text-right font-mono-data">{usd(estAnnual)}</td><td className="px-3 py-2 text-right font-mono-data">{usd(estMonthly)}</td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: ReactNode }) {
  return <button onClick={onClick} className={`px-4 py-2 text-sm font-medium border-b-2 whitespace-nowrap -mb-px ${active ? 'border-teal-600 text-teal-700' : 'border-transparent text-gray-500 hover:text-gray-700'}`}>{children}</button>;
}
function KPI({ label, value, accent = 'default', mono = false }: { label: string; value: string | number; accent?: 'default' | 'warning' | 'danger'; mono?: boolean }) {
  const accentClass = accent === 'danger' ? 'text-error' : accent === 'warning' ? 'text-warning' : 'text-teal-700';
  return <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card"><div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div><div className={`text-2xl font-bold mt-1 ${accentClass} ${mono ? 'font-mono-data' : ''}`}>{value}</div></div>;
}
