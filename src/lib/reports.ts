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
  | 'Operational'
  | 'Backup and Export';

export type ReportStatus =
  | 'available'           // Run Now works
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
    id: 'property-holdings',
    name: 'Property Holdings Statement',
    description: 'Each owner with their direct and beneficial property holdings.',
    category: 'Owner Reports',
    status: 'available',
    filenameBase: 'property-holdings',
  },

  // ─── DOR Audit Pack ───
  {
    id: 'property-audit-pack',
    name: 'Per-Property Full Record Bundle',
    description: 'Full filing history, correspondence, submittals, org chart snapshots for a single property.',
    category: 'DOR Audit Pack',
    status: 'available',
    filenameBase: 'property-audit-pack',
  },
  {
    id: 'portfolio-audit-pack',
    name: 'Portfolio Audit Pack',
    description: 'All properties bundled — defensive audit prep for DOR portfolio inquiries.',
    category: 'DOR Audit Pack',
    status: 'available',
    filenameBase: 'portfolio-audit-pack',
  },
  {
    id: 'org-chart-history',
    name: 'Org Chart History per Property',
    description: 'All frozen org chart snapshots over time for one property — proves ownership chain at each filing.',
    category: 'DOR Audit Pack',
    status: 'available',
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
    id: 'compliance-status',
    name: 'Compliance Status Report',
    description: 'Each property with current compliance status and outstanding deadlines.',
    category: 'Year-End / Annual',
    status: 'available',
    filenameBase: 'compliance-status',
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
    status: 'available',
    filenameBase: 'document-expiration-calendar',
  },
  {
    id: 'untagged-documents-report',
    name: 'Untagged Documents Report',
    description: 'Files in SharePoint libraries that lack PropertyID metadata.',
    category: 'Operational',
    status: 'available',
    filenameBase: 'untagged-documents-report',
  },

  // ─── Backup and Export ───
  {
    id: 'full-database-export',
    name: 'Full Database Export (Excel)',
    description: 'All SharePoint lists exported as a multi-sheet workbook.',
    category: 'Backup and Export',
    status: 'available',
    filenameBase: 'full-database-export',
  },
  {
    id: 'sharepoint-library-snapshot',
    name: 'SharePoint Library Snapshot (JSON)',
    description: 'Document library metadata snapshot — filenames, paths, PropertyIDs.',
    category: 'Backup and Export',
    status: 'available',
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

/**
 * Trigger an Excel (xlsx) download with one or more sheets.
 * Each sheet name maps to an array of row objects.
 */
export async function downloadXLSX(sheets: Record<string, Record<string, unknown>[]>, filename: string): Promise<void> {
  const XLSX = await import('xlsx');
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([sheetName, rows]) => {
    const safeName = sheetName.slice(0, 31); // Excel sheet name limit
    const ws = rows.length > 0
      ? XLSX.utils.json_to_sheet(rows)
      : XLSX.utils.aoa_to_sheet([['(no data)']]);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });
  const arrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  const blob = new Blob([arrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
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
