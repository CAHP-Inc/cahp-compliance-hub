/**
 * Build the Exhibit A workbook (unit-by-unit AMI analysis + summary + per-LLC
 * breakdown for group filings) as a downloadable .xlsx Blob, using the app's
 * existing `xlsx` dependency.
 */
import * as XLSX from 'xlsx';
import type { Analysis } from './analyze';
import type { CertConfig } from './entity';
import type { AmiTier, Unit } from './limits';

const TIER_LABEL: Record<AmiTier, string> = {
  le50: '<=50% AMI (very low-income)',
  le60: '<=60% AMI',
  le80: '<=80% AMI (low-income)',
  market: 'Market (>80% AMI)',
  review: '** NEEDS REVIEW **',
  nonResidential: 'Non-residential (excluded)',
};

export function buildExhibitWorkbook(analysis: Analysis, config: CertConfig): XLSX.WorkBook {
  const { units, roll, scopes, isGroup, sources, perSrc, perScope, utilityAllowance } = analysis;
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: every unit ──
  const header = [
    'Source LLC', 'Property', 'Unit', 'County', 'BR', 'Tenant', 'Status',
    'Market Rent', 'Contract Rent', 'Gross Rent Tested',
    '50% Ceil', '60% Ceil', '80% Ceil', 'AMI Tier', 'Notes',
  ];
  const sorted = [...units].sort((a, b) =>
    (a.source || '').localeCompare(b.source || '') || a.prop.localeCompare(b.prop),
  );
  const rows = sorted.map((u: Unit) => [
    u.source, u.prop, u.unit, u.county || '?',
    u.bedrooms ?? '--', u.tenant, u.status, u.marketRent, u.contractRent,
    u.grossRent === null ? null : u.grossRent + (utilityAllowance || 0),
    u.ceil50, u.ceil60, u.ceil80,
    u.tier ? TIER_LABEL[u.tier] : '?',
    u.notes.join('; '),
  ]);
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  ws['!cols'] = [22, 34, 14, 12, 5, 22, 16, 12, 13, 16, 9, 9, 9, 26, 40].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'Exhibit A - Unit AMI');

  // ── Sheet 2: summary ──
  const p = roll.pct, c = roll.counts;
  const sum: (string | number)[][] = [
    ['Safe Harbor Summary', config.company.legalName],
    [],
    ['Total residential units (denominator)', roll.denom],
    ['Units needing review', roll.nReview],
    ['Non-residential units excluded', roll.nNonRes],
    [],
    ['AMI Tier (cumulative)', 'Units', '% of Total', 'Required', 'Result'],
    ['Low-Income (<=80% AMI)', c.le80, `${p.le80}%`, '>=75%', p.le80 >= 75 ? 'PASS' : 'FAIL'],
    ['  <=60% AMI', c.le60, `${p.le60}%`, '>=40% (40/60 scope)', p.le60 >= 40 ? 'PASS' : 'FAIL'],
    ['  Very Low-Income (<=50% AMI)', c.le50, `${p.le50}%`, '>=20% (20/50 scope)', p.le50 >= 20 ? 'PASS' : 'FAIL'],
    ['Market (>80% AMI)', c.market, `${p.market}%`, '<=25%', p.market <= 25 ? 'PASS' : 'FAIL'],
    [],
    ['20/50 scope', scopes.s2050.status],
    ['40/60 scope', scopes.s4060.status],
    ['DETERMINATION', scopes.headline],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(sum);
  ws2['!cols'] = [30, 10, 12, 22, 10].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');

  // ── Sheet 3: per-LLC (group only) ──
  if (isGroup) {
    const memberByName = new Map((config.portfolio?.members ?? []).map((m) => [m.name, m]));
    const aoa: (string | number)[][] = [
      [`Portfolio: ${config.portfolio?.groupName ?? config.company.legalName}`],
      [],
      ['Source LLC', 'Nonprofit %', 'Class', 'Units', '<=50%', '<=60%', '<=80%', 'Market', '20/50', '40/60'],
    ];
    for (const s of sources) {
      const r = perSrc[s], sc = perScope[s];
      const m = memberByName.get(s);
      aoa.push([
        s, m?.ownershipPercent == null ? '—' : `${m.ownershipPercent}%`, m?.memberClass || '—',
        r.denom, `${r.pct.le50}%`, `${r.pct.le60}%`, `${r.pct.le80}%`, `${r.pct.market}%`,
        sc.s2050.status, sc.s4060.status,
      ]);
    }
    aoa.push([]);
    aoa.push([
      'PORTFOLIO TOTAL', '', '', roll.denom, `${p.le50}%`, `${p.le60}%`, `${p.le80}%`, `${p.market}%`,
      scopes.s2050.status, scopes.s4060.status,
    ]);
    const ws3 = XLSX.utils.aoa_to_sheet(aoa);
    ws3['!cols'] = [24, 11, 8, 7, 8, 8, 8, 8, 24, 24].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws3, 'Per-LLC');
  }

  return wb;
}

export function buildExhibitBlob(analysis: Analysis, config: CertConfig): Blob {
  const wb = buildExhibitWorkbook(analysis, config);
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  return new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
