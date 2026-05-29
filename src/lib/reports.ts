/**
 * Reports infrastructure.
 *
 * Each report is a registered descriptor. The Reports page composes the
 * shared report context (lists already loaded via SharePoint hooks), opens
 * the Run Report config modal, and dispatches to the runner with the
 * caller-collected options. Runners are pure: they take the context + options
 * and return a built blob + metadata. The caller is responsible for
 * downloading, audit-logging, and offering the optional "Send via Email"
 * handoff.
 */

import { toDateInputValue } from './dates';
import type { CahpState, CahpTaxYear } from './sharepoint';

// =============================================================================
// Descriptor types
// =============================================================================

export type ReportCategory =
  | 'Owner Reports'
  | 'DOR Audit Pack'
  | 'Year-End / Annual'
  | 'Operational'
  | 'Backup and Export';

export type ReportAudience = 'internal' | 'owner';

export type ReportFormat = 'csv' | 'xlsx' | 'json' | 'pdf';

export type ReportScopeKind = 'portfolio' | 'property' | 'owner' | 'state';

/** Parameter capabilities a report supports. Drives which controls render in the modal. */
export interface ReportParamSpec {
  scope?: ReportScopeKind[];          // Which scope options the user can pick. Omit for portfolio-only.
  dateRange?: boolean;                // Free-form date-range pickers (from/to)
  taxYear?: boolean;                  // Tax-year dropdown (Annual Filing)
  expirationWindow?: boolean;         // 90/180/365 day bucket picker (Doc Expiration)
  quarter?: boolean;                  // Quarter + year pair (Quarterly Statement)
  internalColumnsToggle?: boolean;    // "Include internal-only notes" — defaults true for internal audience, false for owner
}

export interface ReportDescriptor {
  id: string;
  name: string;
  description: string;
  category: ReportCategory;
  audience: ReportAudience;
  params: ReportParamSpec;
  supportedFormats: ReportFormat[];
  defaultFormat: ReportFormat;
  /** Default filename (without extension) */
  filenameBase: string;
}

// =============================================================================
// Run options — collected by the modal, passed to the runner
// =============================================================================

export type ReportScope =
  | { kind: 'portfolio' }
  | { kind: 'property'; propertyId: string; propertyTitle?: string }
  | { kind: 'owner'; ownerId: string; ownerTitle?: string }
  | { kind: 'state'; state: CahpState };

export interface RunOptions {
  format: ReportFormat;
  scope: ReportScope;
  dateFrom?: string;            // ISO date (YYYY-MM-DD)
  dateTo?: string;              // ISO date (YYYY-MM-DD)
  taxYear?: CahpTaxYear;
  expirationWindow?: 90 | 180 | 365;
  quarter?: { year: number; q: 1 | 2 | 3 | 4 };
  includeInternalColumns?: boolean;
}

/** Result returned by every runner. The caller downloads the blob + logs. */
export interface RunResult {
  rowCount: number;
  filename: string;
  blob: Blob;
  contentType: string;
  /** Human-readable scope label used in the post-run toast (e.g., "Marwar Ventures LLC"). */
  scopeLabel?: string;
}

// =============================================================================
// Report Registry
// =============================================================================

export const REPORTS: ReportDescriptor[] = [
  // ─── Owner Reports ───
  {
    id: 'owner-quarterly-statement',
    name: 'Owner Quarterly Statement (PDF)',
    description:
      'Branded one-owner summary: holdings, abatement status, recent communications, open items the owner owes. Suitable to email directly.',
    category: 'Owner Reports',
    audience: 'owner',
    params: { scope: ['owner'], quarter: true, internalColumnsToggle: false },
    supportedFormats: ['pdf'],
    defaultFormat: 'pdf',
    filenameBase: 'owner-quarterly-statement',
  },
  {
    id: 'property-holdings',
    name: 'Property Holdings Statement',
    description:
      'Each owner with their direct and beneficial property holdings. Scope to one owner to produce a deliverable for that owner.',
    category: 'Owner Reports',
    audience: 'owner',
    params: { scope: ['portfolio', 'owner', 'state'], internalColumnsToggle: true },
    supportedFormats: ['csv', 'xlsx', 'pdf'],
    defaultFormat: 'csv',
    filenameBase: 'property-holdings',
  },
  {
    id: 'outstanding-items-by-owner-contact',
    name: 'Outstanding Items by Owner Contact',
    description:
      'Open items the property owner (not internal staff) must produce. Groups by owner contact; safe to email — internal notes hidden by default.',
    category: 'Owner Reports',
    audience: 'owner',
    params: { scope: ['portfolio', 'owner', 'property'], internalColumnsToggle: true },
    supportedFormats: ['csv', 'xlsx', 'pdf'],
    defaultFormat: 'pdf',
    filenameBase: 'outstanding-items-by-owner-contact',
  },

  // ─── DOR Audit Pack ───
  {
    id: 'property-audit-pack',
    name: 'Per-Property Full Record Bundle',
    description:
      'Full filing history, correspondence, submittals, org chart snapshots for a single property.',
    category: 'DOR Audit Pack',
    audience: 'internal',
    params: { scope: ['property'] },
    supportedFormats: ['xlsx'],
    defaultFormat: 'xlsx',
    filenameBase: 'property-audit-pack',
  },
  {
    id: 'portfolio-audit-pack',
    name: 'Portfolio Audit Pack',
    description:
      'All properties bundled — defensive audit prep for DOR portfolio inquiries.',
    category: 'DOR Audit Pack',
    audience: 'internal',
    params: { scope: ['portfolio', 'state'] },
    supportedFormats: ['xlsx'],
    defaultFormat: 'xlsx',
    filenameBase: 'portfolio-audit-pack',
  },
  {
    id: 'org-chart-history',
    name: 'Org Chart History per Property',
    description:
      'All frozen org chart snapshots over time for one property — renders each as a chart page in a multi-page PDF.',
    category: 'DOR Audit Pack',
    audience: 'internal',
    params: { scope: ['property'] },
    supportedFormats: ['pdf', 'json'],
    defaultFormat: 'pdf',
    filenameBase: 'org-chart-history',
  },
  {
    id: 'dor-correspondence-log',
    name: 'DOR Correspondence Log',
    description:
      'Inbound + outbound DOR letters and email threads, filterable by date range, state, or property.',
    category: 'DOR Audit Pack',
    audience: 'internal',
    params: {
      scope: ['portfolio', 'property', 'state'],
      dateRange: true,
      internalColumnsToggle: true,
    },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'dor-correspondence-log',
  },

  // ─── Year-End / Annual ───
  {
    id: 'annual-filing-report',
    name: 'Annual Filing Report',
    description: 'All submittals filed in the selected tax year with status, dates, and outcomes.',
    category: 'Year-End / Annual',
    audience: 'internal',
    params: { scope: ['portfolio', 'state'], taxYear: true },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'annual-filing-report',
  },
  {
    id: 'compliance-status',
    name: 'Compliance Status Report',
    description:
      'Each property with current status. Surfaces Overdue, At Risk (≤30 days), On Track buckets.',
    category: 'Year-End / Annual',
    audience: 'internal',
    params: { scope: ['portfolio', 'state'] },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'compliance-status',
  },

  // ─── Operational ───
  {
    id: 'outstanding-items-by-assignee',
    name: 'Outstanding Items by Assignee',
    description:
      'Open items grouped by the assigned team member (or vendor), sorted by overdue then due date. Internal triage view.',
    category: 'Operational',
    audience: 'internal',
    params: { scope: ['portfolio', 'property', 'state'], internalColumnsToggle: true },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'outstanding-items-by-assignee',
  },
  {
    id: 'communications-history',
    name: 'Communications History',
    description:
      'Owner Communications log (email, phone, meetings) filterable by date range, owner, or property. Internal QA + owner-statement support.',
    category: 'Operational',
    audience: 'internal',
    params: {
      scope: ['portfolio', 'owner', 'property'],
      dateRange: true,
      internalColumnsToggle: true,
    },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'communications-history',
  },
  {
    id: 'document-expiration-calendar',
    name: 'Document Expiration Calendar',
    description:
      'Documents and compliance deadlines expiring within the chosen window (90 / 180 / 365 days), bucketed by urgency.',
    category: 'Operational',
    audience: 'internal',
    params: { scope: ['portfolio', 'property'], expirationWindow: true },
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'document-expiration-calendar',
  },
  {
    id: 'untagged-documents-report',
    name: 'Untagged Documents Report',
    description: 'Files in SharePoint libraries that lack PropertyID metadata.',
    category: 'Operational',
    audience: 'internal',
    params: {},
    supportedFormats: ['csv', 'xlsx'],
    defaultFormat: 'csv',
    filenameBase: 'untagged-documents-report',
  },

  // ─── Backup and Export ───
  {
    id: 'full-database-export',
    name: 'Full Database Export (Excel)',
    description: 'All SharePoint lists exported as a multi-sheet workbook.',
    category: 'Backup and Export',
    audience: 'internal',
    params: {},
    supportedFormats: ['xlsx'],
    defaultFormat: 'xlsx',
    filenameBase: 'full-database-export',
  },
  {
    id: 'sharepoint-library-snapshot',
    name: 'SharePoint Library Snapshot (JSON)',
    description: 'Document library metadata snapshot — filenames, paths, PropertyIDs.',
    category: 'Backup and Export',
    audience: 'internal',
    params: {},
    supportedFormats: ['json'],
    defaultFormat: 'json',
    filenameBase: 'sharepoint-library-snapshot',
  },
];

// =============================================================================
// Blob builders — return Blob + content-type instead of triggering download
// =============================================================================

const CSV_MIME = 'text/csv;charset=utf-8';
const JSON_MIME = 'application/json';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const PDF_MIME = 'application/pdf';

export interface BuiltFile {
  blob: Blob;
  contentType: string;
}

export function buildCSV(rows: Record<string, unknown>[]): BuiltFile {
  if (rows.length === 0) {
    return { blob: new Blob(['(no data)'], { type: CSV_MIME }), contentType: CSV_MIME };
  }
  const headers = Object.keys(rows[0]);
  const escape = (val: unknown): string => {
    if (val == null) return '';
    const s = String(val);
    if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };
  const lines = [
    headers.join(','),
    ...rows.map((row) => headers.map((h) => escape(row[h])).join(',')),
  ];
  return { blob: new Blob([lines.join('\n')], { type: CSV_MIME }), contentType: CSV_MIME };
}

export function buildJSON(data: unknown): BuiltFile {
  return {
    blob: new Blob([JSON.stringify(data, null, 2)], { type: JSON_MIME }),
    contentType: JSON_MIME,
  };
}

export async function buildXLSX(
  sheets: Record<string, Record<string, unknown>[]>,
): Promise<BuiltFile> {
  const XLSX = await import('xlsx');
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([sheetName, rows]) => {
    const safeName = sheetName.slice(0, 31);
    const ws = rows.length > 0
      ? XLSX.utils.json_to_sheet(rows)
      : XLSX.utils.aoa_to_sheet([['(no data)']]);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });
  const arrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  return { blob: new Blob([arrayBuffer], { type: XLSX_MIME }), contentType: XLSX_MIME };
}

/** Wrap a jsPDF instance into a Blob. */
export function buildPDF(pdf: { output: (kind: 'blob') => Blob }): BuiltFile {
  return { blob: pdf.output('blob'), contentType: PDF_MIME };
}

// =============================================================================
// Browser download trigger
// =============================================================================

export function triggerDownload(blob: Blob, filename: string): void {
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
 * Convert a blob to base64 (no `data:` prefix) for emailing via Graph.
 * Throws if the blob can't be read.
 */
export async function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result ?? '');
      const comma = result.indexOf(',');
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = () => reject(reader.error ?? new Error('Failed to read blob.'));
    reader.readAsDataURL(blob);
  });
}

/**
 * Filename = base + ISO date + extension. e.g., "compliance-status-2026-05-29.csv"
 * If a scope label is supplied, it's slugged and folded in before the date.
 */
export function timestampedFilename(
  base: string,
  ext: string,
  scopeLabel?: string,
): string {
  const date = toDateInputValue(new Date());
  const stem = scopeLabel
    ? `${base}-${scopeLabel.replace(/[^\w-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '')}`
    : base;
  return `${stem}-${date}.${ext}`;
}
