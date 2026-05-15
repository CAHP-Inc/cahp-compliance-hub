/**
 * Reports infrastructure — PR-14a
 *
 * Each report is a registered descriptor with a `run` function that returns
 * a Promise of report rows. The runner converts to CSV/JSON and triggers a
 * client-side download. No server-side report generation — the browser does
 * all the work using already-loaded SharePoint data.
 */

// =============================================================================
// Report descriptors
// =============================================================================

export type ReportCategory =
  | 'Owner Reports'
  | 'DOR Audit Pack'
  | 'Year-End / Annual'
  | 'Financial'
  | 'Operational'
  | 'Backup and Export';

export type ReportStatus =
  | 'available'           // Run Now works
  | 'pending-billing'     // Disabled — needs Billing module to be reactivated
  | 'pending-pr14b';      // Coming in PR-14b

export interface ReportDescriptor {
  id: string;
  name: string;
  description: string;
  category: ReportCategory;
  status: ReportStatus;
  /** Default filename (without extension) */
  filenameBase: string;
}

// =============================================================================
// Report Registry — the canonical list of reports per spec §3.14
// =============================================================================

export const REPORTS: ReportDescriptor[] = [
  // ─── Owner Reports ───
  {
    id: 'year-end-tax-summary',
    name: 'Year-End Tax Summary',
    description: 'CAHP fees billed + tax savings achieved per owner for the calendar year.',
    category: 'Owner Reports',
    status: 'pending-billing',
    filenameBase: 'year-end-tax-summary',
  },
  {
    id: 'refund-history',
    name: 'Refund History',
    description: 'All DOR refunds received YTD, by owner and property.',
    category: 'Owner Reports',
    status: 'pending-billing',
    filenameBase: 'refund-history',
  },
  {
    id: 'cahp-fee-statement',
    name: 'CAHP Fee Statement',
    description: 'Invoiceable + invoiced fee history with payment status, by owner.',
    category: 'Owner Reports',
    status: 'pending-billing',
    filenameBase: 'cahp-fee-statement',
  },
  {
    id: 'property-holdings',
    name: 'Property Holdings Statement',
    description: 'Each owner with their direct and beneficial property holdings.',
    category: 'Owner Reports',
    status: 'pending-pr14b',
    filenameBase: 'property-holdings',
  },

  // ─── DOR Audit Pack ───
  {
    id: 'property-audit-pack',
    name: 'Per-Property Full Record Bundle',
    description: 'Full filing history, correspondence, submittals, org chart snapshots for a single property.',
    category: 'DOR Audit Pack',
    status: 'pending-pr14b',
    filenameBase: 'property-audit-pack',
  },
  {
    id: 'portfolio-audit-pack',
    name: 'Portfolio Audit Pack',
    description: 'All properties bundled — defensive audit prep for DOR portfolio inquiries.',
    category: 'DOR Audit Pack',
    status: 'pending-pr14b',
    filenameBase: 'portfolio-audit-pack',
  },
  {
    id: 'org-chart-history',
    name: 'Org Chart History per Property',
    description: 'All frozen org chart snapshots over time for one property — proves ownership chain at each filing.',
    category: 'DOR Audit Pack',
    status: 'pending-pr14b',
    filenameBase: 'org-chart-history',
  },

  // ─── Year-End / Annual ───
  {
    id: 'annual-filing-report',
    name: 'Annual Filing Report',
    description: 'All submittals filed in the calendar year with status, dates, and outcomes.',
    category: 'Year-End / Annual',
    status: 'available',
    filenameBase: 'annual-filing-report',
  },
  {
    id: 'year-over-year',
    name: 'Year-over-Year Comparison',
    description: 'Tax savings and filing volume compared to prior years.',
    category: 'Year-End / Annual',
    status: 'pending-billing',
    filenameBase: 'year-over-year',
  },
  {
    id: 'compliance-status',
    name: 'Compliance Status Report',
    description: 'Each property with current compliance status and outstanding deadlines.',
    category: 'Year-End / Annual',
    status: 'available',
    filenameBase: 'compliance-status',
  },

  // ─── Financial — all billing-dependent ───
  {
    id: 'fee-revenue-summary',
    name: 'CAHP Fee Revenue Summary',
    description: 'Revenue by month/quarter/year with QB sync status.',
    category: 'Financial',
    status: 'pending-billing',
    filenameBase: 'fee-revenue-summary',
  },
  {
    id: 'disbursement-report',
    name: 'Disbursement Report',
    description: 'All disbursements to owners with payment method and 1099 threshold flags.',
    category: 'Financial',
    status: 'pending-billing',
    filenameBase: 'disbursement-report',
  },
  {
    id: 'outstanding-balances',
    name: 'Outstanding Balances',
    description: 'Invoiced but unpaid CAHP fees + pending disbursements.',
    category: 'Financial',
    status: 'pending-billing',
    filenameBase: 'outstanding-balances',
  },

  // ─── Operational ───
  {
    id: 'outstanding-items-by-owner',
    name: 'Outstanding Items by Owner',
    description: 'Open items grouped by assigned-to, sorted by due date.',
    category: 'Operational',
    status: 'available',
    filenameBase: 'outstanding-items-by-owner',
  },
  {
    id: 'document-expiration-calendar',
    name: 'Document Expiration Calendar',
    description: 'Documents expiring in the next 90/180/365 days.',
    category: 'Operational',
    status: 'pending-pr14b',
    filenameBase: 'document-expiration-calendar',
  },
  {
    id: 'untagged-documents-report',
    name: 'Untagged Documents Report',
    description: 'Files in SharePoint libraries that lack PropertyID metadata.',
    category: 'Operational',
    status: 'pending-pr14b',
    filenameBase: 'untagged-documents-report',
  },

  // ─── Backup and Export ───
  {
    id: 'full-database-export',
    name: 'Full Database Export (Excel)',
    description: 'All SharePoint lists exported as a multi-sheet workbook.',
    category: 'Backup and Export',
    status: 'pending-pr14b',
    filenameBase: 'full-database-export',
  },
  {
    id: 'sharepoint-library-snapshot',
    name: 'SharePoint Library Snapshot (JSON)',
    description: 'Document library metadata snapshot — filenames, paths, PropertyIDs.',
    category: 'Backup and Export',
    status: 'pending-pr14b',
    filenameBase: 'sharepoint-library-snapshot',
  },
];

// =============================================================================
// Export helpers
// =============================================================================

/**
 * Convert an array of objects to CSV and trigger a browser download.
 * Header row is keys of the first object.
 */
export function downloadCSV(rows: Record<string, unknown>[], filename: string): void {
  if (rows.length === 0) {
    // Still download with just a "(no data)" line so the user knows the report ran
    const blob = new Blob(['(no data)'], { type: 'text/csv' });
    triggerDownload(blob, filename);
    return;
  }

  const headers = Object.keys(rows[0]);
  const escape = (val: unknown): string => {
    if (val == null) return '';
    const s = String(val);
    // RFC 4180: wrap in quotes if contains comma, quote, or newline; double-quote any quotes
    if (/[",\n\r]/.test(s)) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };

  const lines = [
    headers.join(','),
    ...rows.map((row) => headers.map((h) => escape(row[h])).join(',')),
  ];
  const csv = lines.join('\n');

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  triggerDownload(blob, filename);
}

/** Trigger a JSON download. */
export function downloadJSON(data: unknown, filename: string): void {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  triggerDownload(blob, filename);
}

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Format a filename with the report base + ISO date + extension.
 * e.g., "compliance-status-2026-05-15.csv"
 */
export function timestampedFilename(base: string, ext: string): string {
  const date = new Date().toISOString().slice(0, 10);
  return `${base}-${date}.${ext}`;
}
