/**
 * One-call analysis: classify units, roll up the portfolio, detect a group
 * filing, and compute per-LLC breakdowns. Shared by the modal (live preview)
 * and the document builders so they never diverge.
 */
import {
  classify,
  evaluateScopes,
  rollup,
  type Rollup,
  type ScopeResult,
  type Unit,
} from './limits';

export interface Analysis {
  units: Unit[];
  roll: Rollup;
  scopes: ScopeResult;
  isGroup: boolean;
  sources: string[];
  perSrc: Record<string, Rollup>;
  perScope: Record<string, ScopeResult>;
  taxYear: number;
  utilityAllowance: number;
}

export function analyze(
  units: Unit[],
  opts: { taxYear: number; utilityAllowance?: number; forceGroup?: boolean; state?: string },
): Analysis {
  const utilityAllowance = opts.utilityAllowance ?? 0;
  // Fresh classification each call (units are mutated in place — reset first).
  for (const u of units) {
    u.tier = null;
    u.ceil50 = u.ceil60 = u.ceil80 = null;
    u.notes = [];
  }
  classify(units, utilityAllowance, opts.state ?? 'SC');
  const roll = rollup(units);
  const scopes = evaluateScopes(roll);

  const sources = [...new Set(units.map((u) => u.source).filter(Boolean))].sort();
  const isGroup = Boolean(opts.forceGroup) || sources.length > 1;

  const perSrc: Record<string, Rollup> = {};
  const perScope: Record<string, ScopeResult> = {};
  if (isGroup) {
    for (const s of sources) {
      const subset = units.filter((u) => u.source === s);
      perSrc[s] = rollup(subset);
      perScope[s] = evaluateScopes(perSrc[s]);
    }
  }

  return { units, roll, scopes, isGroup, sources, perSrc, perScope, taxYear: opts.taxYear, utilityAllowance };
}
