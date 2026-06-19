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

/**
 * Vacant units with no listed market rent inherit the maximum rent of
 * comparable units — same bedroom count in the same property (source), and
 * failing that, same bedroom count anywhere in the roll. A between-tenants
 * affordable unit is thus represented at the property's asking rate for that
 * unit type rather than $0/blank (which would otherwise read as either deeply
 * sub-50% or undeterminable). Units that already carry a rent are untouched.
 */
function imputeVacantRents(units: Unit[]): void {
  const hasRent = (u: Unit): boolean => typeof u.grossRent === 'number' && u.grossRent > 0;
  for (const u of units) {
    if (u.occupied || hasRent(u) || u.nonResidential) continue;
    const sameBeds = (c: Unit): boolean => c !== u && hasRent(c) && c.bedrooms === u.bedrooms;
    let pool = units.filter((c) => sameBeds(c) && c.source === u.source);
    if (pool.length === 0) pool = units.filter(sameBeds);
    if (pool.length === 0) continue;
    const max = Math.max(...pool.map((c) => c.grossRent as number));
    u.grossRent = max;
    if (u.marketRent == null || u.marketRent <= 0) u.marketRent = max;
    const bd = u.bedrooms == null ? 'comparable' : u.bedrooms === 0 ? 'studio' : `${u.bedrooms}BR`;
    u.notes.push(`Vacant — market rent imputed at $${max.toLocaleString()} (max of comparable ${bd} units).`);
  }
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
  imputeVacantRents(units);
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
