/**
 * Parse an AppFolio Rent Roll .xlsx (in the browser) into Safe Harbor units.
 *
 * Uses the `xlsx` dependency already in the app. Mirrors the parser in
 * scripts/safe-harbor/generate_cert.py: it finds the header row (col A ==
 * "Property"), reads the "Property Groups" line to tag the source LLC, and
 * stops at the "Total" row.
 */
import * as XLSX from 'xlsx';
import { CITY_TO_COUNTY, type Unit } from './limits';

export interface ParsedRoll {
  units: Unit[];
  exported: string;
  source: string;
  filename: string;
}

function toNum(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[,$]/g, '').trim();
  const n = Number(s);
  return Number.isNaN(n) ? null : n;
}

export function detectCounty(address: string): string | null {
  const addr = address || '';
  const cities = Object.keys(CITY_TO_COUNTY).sort((a, b) => b.length - a.length);
  for (const city of cities) {
    if (new RegExp(`\\b${city.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(addr)) {
      return CITY_TO_COUNTY[city];
    }
  }
  const m = addr.match(/\b(\d{5})(?:-\d{4})?\b/);
  if (m) {
    if (m[1].startsWith('293')) return 'Spartanburg';
    if (m[1].startsWith('296')) return 'Greenville';
  }
  return null;
}

export function parseBedrooms(bdba: unknown): [number | null, string | null] {
  if (!bdba) return [null, null];
  const [left, right] = String(bdba).split('/');
  const l = (left ?? '').trim();
  const r = (right ?? '').trim();
  const bd = /^\d+$/.test(l) ? parseInt(l, 10) : null;
  return [bd, r || null];
}

export async function parseRentRoll(file: File): Promise<ParsedRoll> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, blankrows: false });

  // RealPage OneSite export ("OneSite Rents v3.0" in the first cells) — a wholly
  // different layout from AppFolio, so route to a dedicated parser.
  if (rows.slice(0, 6).some((r) => typeof r?.[0] === 'string' && /onesite/i.test(r[0] as string))) {
    return parseOneSite(rows, file.name);
  }

  let exported = '';
  let source = '';
  let headerIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const a = rows[i]?.[0];
    if (typeof a === 'string') {
      if (a.startsWith('Exported On')) exported = a.replace('Exported On:', '').trim();
      if (a.startsWith('Property Groups')) {
        source = a.split(':').slice(1).join(':').trim().replace(/^Owner-\s*/, '').trim();
      }
      if (a.trim() === 'Property') {
        headerIdx = i;
        break;
      }
    }
  }
  if (headerIdx === -1) {
    throw new Error(
      `${file.name}: could not find the rent-roll header row (column A === "Property"). ` +
        'Is this an AppFolio Rent Roll export?',
    );
  }
  if (!source) source = file.name.replace(/\.xlsx?$/i, '');

  const units: Unit[] = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i] ?? [];
    const a = r[0];
    if (a === null || a === undefined || String(a).trim() === '') continue;
    if (String(a).trim().toLowerCase() === 'total') break;

    const prop = String(a).trim();
    const unitLbl = r[1] ? String(r[1]).trim() : '';
    const [bedrooms, baths] = parseBedrooms(r[3]);
    const tenant = r[4] ? String(r[4]).trim() : '';
    const status = r[5] ? String(r[5]).trim() : '';
    const marketRent = toNum(r[7]);
    const contractRent = toNum(r[8]);

    const county = detectCounty(prop);
    const nonResidential = /lot only/i.test(prop);
    const occupied = Boolean(tenant) && !status.toLowerCase().startsWith('vacant');
    const grossRent = occupied && contractRent ? contractRent : marketRent;

    units.push({
      source,
      prop,
      unit: unitLbl,
      county,
      bedrooms,
      baths,
      tenant,
      status,
      marketRent,
      contractRent,
      occupied,
      nonResidential,
      grossRent,
      tier: null,
      ceil50: null,
      ceil60: null,
      ceil80: null,
      notes: [],
    });
  }

  return { units, exported, source, filename: file.name };
}

/**
 * RealPage OneSite "Rents" export. Wide, merged-cell layout (1 unit per 3 rows):
 *   A=Unit, C=Floorplan ("1BR"/"2BR"/"Studio"), R=Unit/Lease Status, T=Name,
 *   AG(33)=Market rent, AO–AQ(41–43)=Lease (contract) rent, AY(51)=Total billing.
 * No per-unit address/county — the modal fills county from the selected property.
 * Column indices are 0-based here (A=0): C=2, R=17, T=19, AG=32, AO=40..AQ=42.
 */
function parseOneSite(rows: unknown[][], filename: string): ParsedRoll {
  const units: Unit[] = [];
  // Header row: col A contains "Unit" (often "\nUnit").
  let headerIdx = rows.findIndex((r) => typeof r?.[0] === 'string' && /unit/i.test(r[0] as string) && /floorplan/i.test(String(r?.[2] ?? '')));
  if (headerIdx === -1) headerIdx = rows.findIndex((r) => typeof r?.[0] === 'string' && /^\s*\n?unit\s*$/i.test(r[0] as string));

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i] ?? [];
    const a = r[0];
    if (a === null || a === undefined || String(a).trim() === '') continue;
    const unitLbl = String(a).trim();
    const floorplan = r[2] != null ? String(r[2]).trim() : '';
    const status = r[17] != null ? String(r[17]).trim() : '';
    // Skip non-unit rows (summary/section headers) — a real unit has a floorplan or status.
    if (!floorplan && !status) continue;
    if (/^(summary|total|floorplan|future\b|applicant|non-?revenue)/i.test(unitLbl)) continue;

    const [bedrooms] = parseBedrooms(/\bstudio|eff/i.test(floorplan) ? '0/' : `${(floorplan.match(/(\d+)\s*BR/i)?.[1]) ?? ''}/`);
    const name = r[19] != null ? String(r[19]).trim() : '';
    const marketRent = toNum(r[32]); // AG
    const leaseRent = Math.max(toNum(r[40]) ?? 0, toNum(r[41]) ?? 0, toNum(r[42]) ?? 0) || null; // AO–AQ
    const occupied = /^occupied/i.test(status) || (Boolean(name) && name.toUpperCase() !== 'VACANT');
    const grossRent = occupied ? (leaseRent ?? marketRent) : marketRent;

    units.push({
      source: filename.replace(/\.xlsx?$/i, ''),
      prop: `Unit ${unitLbl}`,
      unit: unitLbl,
      county: null, // OneSite has no address — modal assigns the property's county
      bedrooms,
      baths: null,
      tenant: name && name.toUpperCase() !== 'VACANT' ? name : '',
      status,
      marketRent,
      contractRent: leaseRent,
      occupied,
      nonResidential: false,
      grossRent,
      tier: null,
      ceil50: null,
      ceil60: null,
      ceil80: null,
      notes: [],
    });
  }
  return { units, exported: '', source: filename.replace(/\.xlsx?$/i, ''), filename };
}
