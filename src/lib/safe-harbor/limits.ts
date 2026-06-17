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
  'Honea Path': 'Anderson',
  Anderson: 'Anderson',
  Belton: 'Anderson',
  Williamston: 'Anderson',
  Pendleton: 'Anderson',
  Powdersville: 'Anderson',
  Easley: 'Pickens',
  Clemson: 'Pickens',
  Pickens: 'Pickens',
  Seneca: 'Oconee',
};

interface CountyLimits {
  msa: string;
  /** rentLimits[level][bedroom] = max gross rent. level ∈ {'50','80'}; bedroom 0..4. */
  rentLimits: Record<'50' | '80', Record<string, number>>;
}

// FY2026 SC Housing Trust Fund max gross rent by county [0BR,1BR,2BR,3BR,4BR].
const RAW_50: Record<string, number[]> = {
  Abbeville: [653,700,840,970,1082], Aiken: [791,848,1017,1175,1311], Allendale: [653,700,840,970,1082],
  Anderson: [832,891,1070,1236,1380], Bamberg: [653,700,840,970,1082], Barnwell: [653,700,840,970,1082],
  Beaufort: [988,1059,1271,1469,1638], Berkeley: [1028,1101,1322,1527,1703], Calhoun: [808,866,1040,1201,1340],
  Charleston: [1028,1101,1322,1527,1703], Cherokee: [653,700,840,970,1082], Chester: [653,700,840,970,1082],
  Chesterfield: [653,700,840,970,1082], Clarendon: [653,700,840,970,1082], Colleton: [653,700,840,970,1082],
  Darlington: [653,700,840,970,1082], Dillon: [653,700,840,970,1082], Dorchester: [1028,1101,1322,1527,1703],
  Edgefield: [791,848,1017,1175,1311], Fairfield: [808,866,1040,1201,1340], Florence: [695,744,892,1031,1150],
  Georgetown: [750,803,963,1113,1241], Greenville: [930,996,1196,1382,1542], Greenwood: [653,700,840,970,1082],
  Hampton: [653,700,840,970,1082], Horry: [752,806,967,1118,1247], Jasper: [731,783,940,1086,1212],
  Kershaw: [812,870,1045,1206,1346], Lancaster: [897,961,1152,1332,1486], Laurens: [700,750,900,1038,1158],
  Lee: [653,700,840,970,1082], Lexington: [808,866,1040,1201,1340], Marion: [653,700,840,970,1082],
  Marlboro: [653,700,840,970,1082], McCormick: [733,786,943,1090,1216], Newberry: [691,740,888,1026,1145],
  Oconee: [777,833,1000,1156,1290], Orangeburg: [653,700,840,970,1082], Pickens: [930,996,1196,1382,1542],
  Richland: [808,866,1040,1201,1340], Saluda: [808,866,1040,1201,1340], Spartanburg: [790,846,1015,1173,1310],
  Sumter: [653,700,840,970,1082], Union: [653,700,840,970,1082], Williamsburg: [653,700,840,970,1082],
  York: [1027,1101,1321,1526,1702],
};
const RAW_80: Record<string, number[]> = {
  Abbeville: [1045,1120,1343,1552,1732], Aiken: [1266,1356,1627,1880,2097], Allendale: [1045,1120,1343,1552,1732],
  Anderson: [1332,1427,1712,1978,2207], Bamberg: [1045,1120,1343,1552,1732], Barnwell: [1045,1120,1343,1552,1732],
  Beaufort: [1582,1695,2035,2350,2622], Berkeley: [1645,1762,2115,2444,2726], Calhoun: [1293,1386,1663,1921,2143],
  Charleston: [1645,1762,2115,2444,2726], Cherokee: [1045,1120,1343,1552,1732], Chester: [1045,1120,1343,1552,1732],
  Chesterfield: [1045,1120,1343,1552,1732], Clarendon: [1045,1120,1343,1552,1732], Colleton: [1045,1120,1343,1552,1732],
  Darlington: [1045,1120,1343,1552,1732], Dillon: [1045,1120,1343,1552,1732], Dorchester: [1645,1762,2115,2444,2726],
  Edgefield: [1266,1356,1627,1880,2097], Fairfield: [1293,1386,1663,1921,2143], Florence: [1111,1190,1428,1650,1841],
  Georgetown: [1198,1284,1541,1781,1987], Greenville: [1488,1595,1913,2211,2467], Greenwood: [1045,1120,1343,1552,1732],
  Hampton: [1045,1120,1343,1552,1732], Horry: [1205,1290,1548,1789,1996], Jasper: [1170,1253,1505,1738,1938],
  Kershaw: [1298,1391,1670,1930,2152], Lancaster: [1435,1537,1845,2132,2378], Laurens: [1118,1198,1438,1661,1853],
  Lee: [1045,1120,1343,1552,1732], Lexington: [1293,1386,1663,1921,2143], Marion: [1045,1120,1343,1552,1732],
  Marlboro: [1045,1120,1343,1552,1732], McCormick: [1173,1257,1508,1743,1945], Newberry: [1105,1183,1420,1640,1830],
  Oconee: [1245,1333,1600,1848,2062], Orangeburg: [1045,1120,1343,1552,1732], Pickens: [1488,1595,1913,2211,2467],
  Richland: [1293,1386,1663,1921,2143], Saluda: [1293,1386,1663,1921,2143], Spartanburg: [1265,1355,1626,1878,2096],
  Sumter: [1045,1120,1343,1552,1732], Union: [1045,1120,1343,1552,1732], Williamsburg: [1045,1120,1343,1552,1732],
  York: [1643,1761,2113,2441,2723],
};
const MSA_OVERRIDE: Record<string, string> = {
  Greenville: 'Greenville-Anderson, SC MSA', Anderson: 'Greenville-Anderson, SC MSA',
  Pickens: 'Greenville-Anderson, SC MSA', Spartanburg: 'Spartanburg, SC MSA',
};
const tiers = (a: number[]): Record<string, number> => ({ 0: a[0], 1: a[1], 2: a[2], 3: a[3], 4: a[4] });
export const COUNTIES: Record<string, CountyLimits> = Object.fromEntries(
  Object.keys(RAW_50).map((c) => [
    c,
    { msa: MSA_OVERRIDE[c] ?? `${c} County, SC`, rentLimits: { '50': tiers(RAW_50[c]), '80': tiers(RAW_80[c]) } },
  ]),
);

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
    headline = `QUALIFIES under the ${chosen} scope`;
  } else if (hasReview) {
    headline = 'PROVISIONAL — resolve the units flagged for review, then re-run';
  } else {
    headline = 'DOES NOT QUALIFY under either scope as configured';
  }

  return { s2050, s4060, qualifies, headline, chosen: qualifies[0] ?? null };
}
