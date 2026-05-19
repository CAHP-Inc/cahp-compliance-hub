/**
 * Date utilities for SharePoint date fields.
 *
 * SharePoint stores Date / DateTime fields as ISO 8601 in UTC
 * (e.g., "2025-05-18T00:00:00Z"). When rendered with toLocaleDateString()
 * in a non-UTC timezone (e.g., America/New_York which is UTC-4 or UTC-5),
 * a UTC-midnight date shifts to the previous calendar day locally:
 *
 *   new Date("2025-05-18T00:00:00Z").toLocaleDateString()
 *     // in EDT → "5/17/2025"  ❌ off by one
 *
 * For DATE-ONLY fields (DueDate, EffectiveDate, DateFiled, etc.), the
 * intent is a calendar date, not a precise moment in time. We need to
 * render the date as it was entered, regardless of the viewer's timezone.
 */

/**
 * Render a SharePoint date-only field as a localized date string,
 * without timezone shift.
 */
export function formatDateOnly(value: string | undefined | null, fallback = '—'): string {
  if (!value) return fallback;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? fallback : parsed.toLocaleDateString();
  }
  const [, year, month, day] = match;
  const localDate = new Date(Number(year), Number(month) - 1, Number(day), 12, 0, 0);
  return localDate.toLocaleDateString();
}

/**
 * Get a Date object representing a date-only field in local time, anchored
 * at noon. Returns null for empty/invalid inputs.
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
 * (from <input type="date">) and produces a noon-UTC ISO string.
 */
export function toDateOnlyISO(value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return undefined;
  const [, year, month, day] = match;
  return `${year}-${month}-${day}T12:00:00.000Z`;
}
