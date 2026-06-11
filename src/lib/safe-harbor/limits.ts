/**
 * Safe Harbor rent-test limits + classification logic.
 *
 * TypeScript port of scripts/safe-harbor/generate_cert.py (classify / rollup /
 * evaluate_scopes), kept deliberately faithful so the in-app generator and the
 * command-line generator produce identical determinations.
 *
 * Qualification is RENT-BASED: each unit's gross rent (contract rent, or the
 * market/asking rent for a vacant unit, plus any tenant-paid utility allowance)
 * is compared to the published Maximum Allowable Gross Rent for its county +
 * bedroom size at 50% / 60% / 80% AMI. The portfolio is then tested against both
 * Rev. Proc. 96-32 set-aside scopes.
 */

export const FY = 2026;
export const LIMITS_EFFECTIVE = '2026-05-01';
export const LIMITS_SOURCE =
  'South Carolina State Housing Finance & Development Authority (SC Housing) ' +
  'Trust Fund Rent Limits, derived from HUD FY2026 MTSP Income Limits ' +
  '(released 2026-05-01).';
export const LIMITS_SOURCE_URLS = {
  rents50:
    'https://schousing.sc.gov/sites/schousing/files/Documents/Development/Income%20and%20Rent%20Limits/2026/SCHTF_50_Rents.pdf',
  rents80:
    'https://schousing.sc.gov/sites/schousing/files/Documents/Development/Income%20and%20Rent%20Limits/2026/SCHTF_80%20Rents.pdf',
  hudMtsp: 'https://www.huduser.gov/portal/datasets/mtsp.html',
};

/** Maps the city found in a rent-roll address to its SC county. Extend as needed. */
export const CITY_TO_COUNTY: Record<string, string> = {
  Greenville: 'Greenville',
  'Travelers Rest': 'Greenville',
  Greer: 'Greenville',
  Mauldin: 'Greenville',
  Simpsonville: 'Greenville',
  Spartanburg: 'Spartanburg',
};

interface CountyLimits {
  msa: string;
  /** rentLimits[level][bedroom] = max gross rent. level ∈ {'50','80'}; bedroom 0..4. */
  rentLimits: Record<'50' | '80', Record<string, number>>;
}

export const COUNTIES: Record<string, CountyLimits> = {
  Greenville: {
    msa: 'Greenville-Anderson, SC MSA',
    rentLimits: {
      '50': { 0: 930, 1: 996, 2: 1196, 3: 1382, 4: 1542 },
      '80': { 0: 1488, 1: 1595, 2: 1913, 3: 2211, 4: 2467 },
    },
  },
  Spartanburg: {
    msa: 'Spartanburg, SC MSA',
    rentLimits: {
      '50': { 0: 790, 1: 846, 2: 1015, 3: 1173, 4: 1310 },
      '80': { 0: 1265, 1: 1355, 2: 1626, 3: 1878, 4: 2096 },
    },
  },
};

export type AmiTier = 'le50' | 'le60' | 'le80' | 'market' | 'review' | 'nonResidential';

export interface Unit {
  source: string;            // LLC / property group the unit belongs to
  prop: string;
  unit: string;
  county: string | null;
  bedrooms: number | null;
  baths: string | null;
  tenant: string;
  status: string;
  marketRent: number | null;
  contractRent: number | null;
  occupied: boolean;
  nonResidential: boolean;
  grossRent: number | null;
  tier: AmiTier | null;
  ceil50: number | null;
  ceil60: number | null;
  ceil80: number | null;
  notes: string[];
}

/** (ceil50, ceil60, ceil80) for a county+bedroom; null when unavailable. 60% = floor(1.2×50%). */
export function ceilingsFor(
  county: string,
  bedrooms: number,
): [number, number, number] | null {
  const cd = COUNTIES[county];
  if (!cd) return null;
  const bkey = String(Math.min(bedrooms, 4));
  const c50 = cd.rentLimits['50'][bkey];
  const c80 = cd.rentLimits['80'][bkey];
  if (c50 === undefined || c80 === undefined) return null;
  return [c50, Math.floor(c50 * 1.2), c80];
}

/** Classify every unit in place (sets tier + ceilings + notes). */
export function classify(units: Unit[], utilityAllowance = 0): void {
  for (const u of units) {
    if (u.nonResidential) {
      u.tier = 'nonResidential';
      u.notes.push('Lot/land only — excluded from residential unit count.');
      continue;
    }
    // Units we can't classify (missing county, bedroom, or rent) default to
    // Market — the conservative outcome (counts against qualification, never for).
    if (!u.county || !COUNTIES[u.county]) {
      u.tier = 'market';
      u.notes.push('County not determined from address — counted as Market.');
      continue;
    }
    if (u.bedrooms === null) {
      u.tier = 'market';
      u.notes.push("Bedroom count missing (BD/BA '--') — counted as Market.");
      continue;
    }
    if (u.grossRent === null) {
      u.tier = 'market';
      u.notes.push('No rent available to test — counted as Market.');
      continue;
    }
    if (u.bedrooms > 4) {
      u.notes.push(`${u.bedrooms}BR uses the published 4BR ceiling (tables cap at 4BR).`);
    }
    const cs = ceilingsFor(u.county, u.bedrooms);
    if (!cs) {
      u.tier = 'market';
      u.notes.push('No rent ceiling published for this county/bedroom — counted as Market.');
      continue;
    }
    [u.ceil50, u.ceil60, u.ceil80] = cs;
    const gr = u.grossRent + (utilityAllowance || 0);
    if (gr <= u.ceil50) u.tier = 'le50';
    else if (gr <= u.ceil60) u.tier = 'le60';
    else if (gr <= u.ceil80) u.tier = 'le80';
    else u.tier = 'market';
  }
}

export interface Rollup {
  denom: number;
  nReview: number;
  nNonRes: number;
  counts: Record<'le50' | 'le60' | 'le80' | 'market', number>;
  pct: Record<'le50' | 'le60' | 'le80' | 'market', number>;
  bounds: Record<'le50' | 'le60' | 'le80' | 'market', [number, number]>;
  review: Unit[];
}

export function rollup(units: Unit[]): Rollup {
  const residential = units.filter((u) => !u.nonResidential);
  const classified = residential.filter(
    (u) => u.tier === 'le50' || u.tier === 'le60' || u.tier === 'le80' || u.tier === 'market',
  );
  const review = residential.filter((u) => u.tier === 'review');
  const nNonRes = units.filter((u) => u.nonResidential).length;

  // Cumulative: a ≤50% unit also satisfies ≤60% and ≤80%.
  const cLe50 = classified.filter((u) => u.tier === 'le50').length;
  const cLe60 = cLe50 + classified.filter((u) => u.tier === 'le60').length;
  const cLe80 = cLe60 + classified.filter((u) => u.tier === 'le80').length;
  const cMarket = classified.filter((u) => u.tier === 'market').length;

  const denom = residential.length;
  const pct = (n: number) => (denom ? Math.round((1000 * n) / denom) / 10 : 0);
  const nReview = review.length;

  return {
    denom,
    nReview,
    nNonRes,
    counts: { le50: cLe50, le60: cLe60, le80: cLe80, market: cMarket },
    pct: { le50: pct(cLe50), le60: pct(cLe60), le80: pct(cLe80), market: pct(cMarket) },
    bounds: {
      le50: [pct(cLe50), pct(cLe50 + nReview)],
      le60: [pct(cLe60), pct(cLe60 + nReview)],
      le80: [pct(cLe80), pct(cLe80 + nReview)],
      market: [pct(cMarket), pct(cMarket + nReview)],
    },
    review,
  };
}

export type ScopeStatus = 'QUALIFIES' | 'DOES NOT QUALIFY' | 'PROVISIONAL — depends on unresolved units';

export interface ScopeVerdict {
  deepPct: number;
  deepThreshold: number;
  marketPct: number;
  status: ScopeStatus;
}

export interface ScopeResult {
  s2050: ScopeVerdict;
  s4060: ScopeVerdict;
  qualifies: string[];
  headline: string;
  chosen: string | null;
}

export function evaluateScopes(roll: Rollup): ScopeResult {
  const { pct, bounds, nReview } = roll;
  const hasReview = nReview > 0;

  const verdict = (deepKey: 'le50' | 'le60', threshold: number): ScopeVerdict => {
    const [deepBest, deepWorst] = bounds[deepKey];
    const [mktBest, mktWorst] = bounds.market;
    const worstPass = deepBest >= threshold && mktWorst <= 25;
    const bestPass = deepWorst >= threshold && mktBest <= 25;
    let status: ScopeStatus;
    if (worstPass) status = 'QUALIFIES';
    else if (!bestPass) status = 'DOES NOT QUALIFY';
    else status = 'PROVISIONAL — depends on unresolved units';
    return { deepPct: pct[deepKey], deepThreshold: threshold, marketPct: pct.market, status };
  };

  const s2050 = verdict('le50', 20);
  const s4060 = verdict('le60', 40);
  const qualifies = (
    [
      ['20/50', s2050],
      ['40/60', s4060],
    ] as const
  )
    .filter(([, s]) => s.status === 'QUALIFIES')
    .map(([name]) => name);

  let headline: string;
  if (qualifies.length) {
    const chosen = qualifies.includes('20/50') ? '20/50' : qualifies[0];
    headline =
      `QUALIFIES under the ${chosen} scope` +
      (qualifies.length === 2 ? ' (qualifies under both)' : '');
  } else if (hasReview) {
    headline = 'PROVISIONAL — resolve the units flagged for review, then re-run';
  } else {
    headline = 'DOES NOT QUALIFY under either scope as configured';
  }

  return { s2050, s4060, qualifies, headline, chosen: qualifies[0] ?? null };
}
