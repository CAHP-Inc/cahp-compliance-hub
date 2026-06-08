/**
 * DOR clocks — the time-sensitive deadlines on a DOR submittal:
 *
 *  - Awaiting DOR response (DOR's court): DOR's published turnaround is ~12 weeks
 *    from the date we file a submittal or send our response to a DOR letter.
 *    Filing and responding both (re)start this expectation.
 *  - RFI response (our court): when DOR sends a "Letter Received - Action Needed"
 *    request for information, we have 30 days from receipt to respond.
 *
 * Centralized here so the submittal detail transition logic and the DOR Deadlines
 * watch widgets share one definition.
 */

import { parseDateOnly, toDateInputValue, toDateOnlyISO } from './dates';
import type { Submittal, SubmittalStatusValue } from './sharepoint';

export const DOR_RESPONSE_WEEKS = 12;
export const DOR_RFI_RESPONSE_DAYS = 30;

/** Statuses where the ball is in DOR's court and we're awaiting their response. */
export const AWAITING_DOR_STATUSES: SubmittalStatusValue[] = ['Filed', 'Responded - Awaiting DOR'];

/** Status where DOR has sent an RFI and the 30-day clock to respond is running. */
export const RFI_STATUS: SubmittalStatusValue = 'Letter Received - Action Needed';

/**
 * Add `days` to a base date (a YYYY-MM-DD or ISO date-only string, or undefined
 * for "today") and return the result as a date-only ISO string (noon UTC).
 */
export function addDaysISO(baseDate: string | undefined, days: number): string | undefined {
  const base = parseDateOnly(baseDate) ?? new Date();
  base.setDate(base.getDate() + days);
  return toDateOnlyISO(toDateInputValue(base));
}

/** When a DOR response should be expected — base date + 12 weeks. */
export function dorResponseDueISO(baseDate: string | undefined): string | undefined {
  return addDaysISO(baseDate, DOR_RESPONSE_WEEKS * 7);
}

/** When our RFI response is due — receipt date + 30 days. */
export function dorRfiDueISO(baseDate: string | undefined): string | undefined {
  return addDaysISO(baseDate, DOR_RFI_RESPONSE_DAYS);
}

export type DORDeadlineKind = 'awaiting-dor' | 'rfi-response';

export type DORDeadlineEntry = {
  submittal: Submittal;
  kind: DORDeadlineKind;
  due: Date | null;
  daysOut: number | null;
  isOverdue: boolean;
};

/**
 * The deadline date for a submittal, if it has a live DOR clock running. Prefers
 * the stored NextActionDue; falls back to deriving from the relevant base date so
 * rows created before the clock existed still surface. Returns null (with a null
 * kind) when no DOR clock applies.
 */
function dorDeadlineFor(s: Submittal): { kind: DORDeadlineKind; due: Date | null } | null {
  const status = s.fields.SubmittalStatus;
  if (!status) return null;

  if (AWAITING_DOR_STATUSES.includes(status)) {
    if (s.fields.NextActionDue) return { kind: 'awaiting-dor', due: new Date(s.fields.NextActionDue) };
    const base =
      status === 'Responded - Awaiting DOR'
        ? s.fields.DateResponded ?? s.fields.DateFiled
        : s.fields.DateFiled;
    const iso = base ? dorResponseDueISO(base) : undefined;
    return { kind: 'awaiting-dor', due: iso ? new Date(iso) : null };
  }

  if (status === RFI_STATUS) {
    if (s.fields.NextActionDue) return { kind: 'rfi-response', due: new Date(s.fields.NextActionDue) };
    const iso = s.fields.DateLetterReceived ? dorRfiDueISO(s.fields.DateLetterReceived) : undefined;
    return { kind: 'rfi-response', due: iso ? new Date(iso) : null };
  }

  return null;
}

/**
 * All submittals with a live DOR clock (awaiting a DOR response, or owing DOR an
 * RFI response), each annotated with its deadline date and how many days out it
 * is. Sorted soonest-due (and overdue) first; entries with no derivable date
 * sort last.
 */
export function selectDORDeadlines(submittals: Submittal[] | undefined | null): DORDeadlineEntry[] {
  if (!submittals) return [];
  const now = Date.now();
  const entries: DORDeadlineEntry[] = [];
  for (const s of submittals) {
    const d = dorDeadlineFor(s);
    if (!d) continue;
    const daysOut = d.due ? Math.round((d.due.getTime() - now) / 86_400_000) : null;
    entries.push({ submittal: s, kind: d.kind, due: d.due, daysOut, isOverdue: daysOut !== null && daysOut < 0 });
  }
  return entries.sort((a, b) => {
    const ad = a.due ? a.due.getTime() : Infinity;
    const bd = b.due ? b.due.getTime() : Infinity;
    return ad - bd;
  });
}
