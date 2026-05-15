import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Billing,
  type Disbursement,
  type Property,
  type BillingStatusValue,
  type QBSyncStatus,
  type DisbursementStatus,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const BILLING_STATUS_STYLES: Record<BillingStatusValue, string> = {
  'Pending Approval': 'bg-gray-100 text-gray-800',
  'Ready to Invoice': 'bg-amber-100 text-amber-800',
  'Invoiced': 'bg-blue-100 text-blue-800',
  'Paid': 'bg-green-100 text-green-800',
  'Disputed': 'bg-red-100 text-red-800',
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

type Tab = 'invoices' | 'disbursements' | 'reconciliation';

export function BillingPage() {
  const [tab, setTab] = useState<Tab>('invoices');

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Billing & Disbursements</h1>
        <p className="text-sm text-gray-500 mt-1">
          CAHP fee invoices (revenue) and DOR refund disbursements (passthrough to owners). Records are auto-created when a submittal is approved.
        </p>
      </div>

      {/* Tab nav */}
      <div className="border-b border-gray-200 mb-4 flex gap-1 overflow-x-auto">
        <TabButton active={tab === 'invoices'} onClick={() => setTab('invoices')}>
          CAHP Fee Invoices
        </TabButton>
        <TabButton active={tab === 'disbursements'} onClick={() => setTab('disbursements')}>
          Refund Disbursements
        </TabButton>
        <TabButton active={tab === 'reconciliation'} onClick={() => setTab('reconciliation')}>
          Reconciliation
        </TabButton>
      </div>

      {tab === 'invoices' && <InvoicesTab />}
      {tab === 'disbursements' && <DisbursementsTab />}
      {tab === 'reconciliation' && <ReconciliationTab />}
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
    const readyCount = billings.data.filter((b) => b.fields.BillingStatus === 'Ready to Invoice').length;
    const invoicedCount = billings.data.filter((b) => b.fields.BillingStatus === 'Invoiced').length;
    const outstanding = billings.data
      .filter((b) => b.fields.BillingStatus === 'Invoiced')
      .reduce((sum, b) => sum + (b.fields.AmountBilled ?? 0), 0);
    return { ytdBilled, readyCount, invoicedCount, outstanding };
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
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="YTD Billed" value={`$${stats.ytdBilled.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} mono />
        <KPI label="Ready to Invoice" value={stats.readyCount} accent={stats.readyCount > 0 ? 'warning' : 'default'} />
        <KPI label="Invoiced" value={stats.invoicedCount} />
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
            CAHP fee invoices are auto-created when a Submittal transitions to Approved (via the Approval Workflow modal).
          </p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
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
            Disbursement records are auto-created when a Submittal transitions to Approved (via the Approval Workflow modal).
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
