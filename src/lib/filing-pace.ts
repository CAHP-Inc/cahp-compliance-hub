import { EASTERN_TZ } from './dates';

/**
 * Filing-pace tracker for the 2026 SC initial-filing push.
 *
 * Goal: file every SC parcel by the end of Thu Jun 25, 2026 — the SC_FILING_FREEZE
 * date. We only work Mon–Fri, so pace is measured in WORKING days, not calendar
 * days.
 *
 * These two dates are the only knobs for this season. Bump them here if the
 * goal moves; everything else (targets, per-day rate) is derived and stays
 * correct as the total/filed counts change.
 */
export const FILING_PACE_START = '2026-06-03';   // campaign start (anchors the % ramp)
export const FILING_TARGET_DATE = '2026-06-25';  // finish line (100% by end of this day) = SC freeze

export type PaceStatus =
  | 'complete'       // everything filed
  | 'before-window'  // viewing before the campaign start
  | 'past-target'    // viewing after the target date, still unfiled
  | 'ahead'
  | 'on-pace'
  | 'behind';

export interface FilingPace {
  filed: number;
  total: number;
  remaining: number;
  filedPct: number;             // 0..100
  /** On-pace cumulative target by END of today (linear ramp to 100% on the last day). */
  targetCount: number;
  targetPct: number;            // 0..100
  /** filed - targetCount. >0 ahead, <0 behind. */
  delta: number;
  totalWorkingDays: number;     // working days across the whole window
  workingDayIndex: number;      // 1..totalWorkingDays for today; 0 before the window
  workingDaysRemaining: number; // today..target inclusive (working days)
  /** ceil(remaining / workingDaysRemaining) — rises if behind, falls if ahead. */
  perDayNeeded: number;
  targetDateLabel: string;      // e.g. "Fri, Jun 19"
  status: PaceStatus;
}

// 'YYYY-MM-DD' → a UTC-noon Date, so getUTCDay() reflects that exact calendar day.
function ymdToUTC(ymd: string): Date {
  const [y, m, d] = ymd.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d, 12));
}

// The Eastern-time calendar day ('YYYY-MM-DD') for a given instant.
function easternYMD(now: Date): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: EASTERN_TZ,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(now);
  const g = (t: string) => parts.find((p) => p.type === t)?.value ?? '';
  return `${g('year')}-${g('month')}-${g('day')}`;
}

function isWorkingDay(utc: Date): boolean {
  const dow = utc.getUTCDay(); // 0 Sun .. 6 Sat
  return dow >= 1 && dow <= 5;
}

/** Count Mon–Fri days in [fromYmd, toYmd] inclusive. 0 if from is after to. */
export function countWorkingDays(fromYmd: string, toYmd: string): number {
  const from = ymdToUTC(fromYmd);
  const to = ymdToUTC(toYmd);
  if (from.getTime() > to.getTime()) return 0;
  let count = 0;
  const cur = new Date(from);
  while (cur.getTime() <= to.getTime()) {
    if (isWorkingDay(cur)) count++;
    cur.setUTCDate(cur.getUTCDate() + 1);
  }
  return count;
}

/**
 * Compute filing pace from live counts. Pure — pass `now` in tests.
 *
 * The "% goal per date" is a linear ramp: by the end of working day k (of N),
 * you should be at k/N of the total. The per-day rate needed to finish is
 * computed forward from today, so it self-corrects for a changing total and
 * for being ahead/behind.
 */
export function computeFilingPace(filed: number, total: number, now: Date = new Date()): FilingPace {
  const todayYmd = easternYMD(now);
  const totalWorkingDays = countWorkingDays(FILING_PACE_START, FILING_TARGET_DATE);

  const targetDateLabel = ymdToUTC(FILING_TARGET_DATE).toLocaleDateString('en-US', {
    timeZone: 'UTC',
    weekday: 'short',
    month: 'short',
    day: 'numeric',
  });

  const remaining = Math.max(0, total - filed);
  const filedPct = total > 0 ? (filed / total) * 100 : 0;

  let workingDayIndex: number;
  let workingDaysRemaining: number;
  if (todayYmd < FILING_PACE_START) {
    workingDayIndex = 0;
    workingDaysRemaining = totalWorkingDays;
  } else if (todayYmd > FILING_TARGET_DATE) {
    workingDayIndex = totalWorkingDays;
    workingDaysRemaining = 0;
  } else {
    workingDayIndex = countWorkingDays(FILING_PACE_START, todayYmd);
    workingDaysRemaining = countWorkingDays(todayYmd, FILING_TARGET_DATE);
  }

  const ramp = totalWorkingDays > 0 ? workingDayIndex / totalWorkingDays : 1;
  const targetPct = Math.min(100, ramp * 100);
  const targetCount = Math.min(total, Math.ceil(total * ramp));
  const delta = filed - targetCount;
  const perDayNeeded = workingDaysRemaining > 0 ? Math.ceil(remaining / workingDaysRemaining) : remaining;

  let status: PaceStatus;
  if (remaining === 0) status = 'complete';
  else if (todayYmd < FILING_PACE_START) status = 'before-window';
  else if (todayYmd > FILING_TARGET_DATE) status = 'past-target';
  else if (delta > 0) status = 'ahead';
  else if (delta === 0) status = 'on-pace';
  else status = 'behind';

  return {
    filed,
    total,
    remaining,
    filedPct,
    targetCount,
    targetPct,
    delta,
    totalWorkingDays,
    workingDayIndex,
    workingDaysRemaining,
    perDayNeeded,
    targetDateLabel,
    status,
  };
}
