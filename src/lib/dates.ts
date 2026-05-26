/**
 * Date utilities for SharePoint date fields.
 *
 * The app always displays dates and times in Eastern Time (America/New_York).
 * This avoids the UTC-shift bug for date-only fields and keeps timestamps
 * consistent regardless of the viewer's browser timezone.
 *
 * Date-only fields (DueDate, EffectiveDate, DateFiled, etc.) come back from
 * SharePoint as ISO strings like "2025-05-18T00:00:00Z". A naive
 * `new Date(...).toLocaleDateString()` in Eastern Time shifts that to
 * "5/17/2025". Use the helpers below to render correctly.
 */

export const EASTERN_TZ = 'America/New_York';

/**
 * Render a SharePoint date-only field as MM/DD/YYYY in Eastern Time.
 * Strips the time component before formatting so the date never shifts.
 */
export function formatDateOnly(value: string | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? fallback : formatDateET(parsed);
  }
  const [, year, month, day] = match;
  return `${month}/${day}/${year}`;
}

/**
 * Get a Date object representing a date-only field in Eastern Time, anchored
 * at noon ET (avoids any TZ rounding into adjacent days). Returns null for
 * empty/invalid inputs.
 */
export function parseDateOnly(value: string | undefined | null): Date | null {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  const [, year, month, day] = match;
  return new Date(Number(year), Number(month) - 1, Number(day), 12, 0, 0);
}

/**
 * For writing a date-only value to SharePoint. Takes a YYYY-MM-DD string
 * (from <input type="date">) and produces a noon-UTC ISO string. Storing as
 * noon UTC ensures the date is unambiguous regardless of the consuming TZ.
 */
export function toDateOnlyISO(value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return undefined;
  const [, year, month, day] = match;
  return `${year}-${month}-${day}T12:00:00.000Z`;
}

/**
 * Convert a Date or ISO string to YYYY-MM-DD in Eastern Time. Use this when
 * binding to <input type="date"> so the picker reflects the ET calendar day,
 * not the browser's local day.
 */
export function toDateInputValue(value: string | Date | undefined | null): string {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: EASTERN_TZ,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(d);
  const year = parts.find((p) => p.type === 'year')?.value ?? '';
  const month = parts.find((p) => p.type === 'month')?.value ?? '';
  const day = parts.find((p) => p.type === 'day')?.value ?? '';
  return `${year}-${month}-${day}`;
}

/**
 * Render a full datetime in Eastern Time.
 * Example: "May 18, 2025, 3:42 PM ET"
 */
export function formatDateTime(value: string | Date | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return fallback;
  return d.toLocaleString('en-US', {
    timeZone: EASTERN_TZ,
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }) + ' ET';
}

/**
 * Render just the date portion of a datetime in Eastern Time.
 * For date-only fields, prefer `formatDateOnly`.
 */
export function formatDateET(value: string | Date | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return fallback;
  return d.toLocaleDateString('en-US', { timeZone: EASTERN_TZ });
}

/**
 * Render just the time portion of a datetime in Eastern Time.
 * Example: "3:42 PM ET"
 */
export function formatTimeET(value: string | Date | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return fallback;
  return d.toLocaleTimeString('en-US', {
    timeZone: EASTERN_TZ,
    hour: 'numeric',
    minute: '2-digit',
  }) + ' ET';
}

/**
 * "5m ago", "3h ago", or a fully-formatted ET datetime once older than 7 days.
 */
export function formatRelativeOrDate(value: string | Date | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return fallback;
  const diffMs = Date.now() - d.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  const diffHr = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHr / 24);
  if (diffMin < 1) return 'just now';
  if (diffMin < 60) return `${diffMin}m ago`;
  if (diffHr < 24) return `${diffHr}h ago`;
  if (diffDay < 7) return `${diffDay}d ago`;
  return formatDateTime(d);
}
