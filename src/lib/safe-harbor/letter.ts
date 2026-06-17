/**
 * Build the Safe Harbor Certification letter as an editable .docx (via `docx`)
 * and a PDF snapshot (via `jspdf`). Content mirrors the command-line generator
 * and the GSP TTAC template, with two adjustments requested for the hub:
 *   1. The certifying party is the PROPERTY MANAGEMENT COMPANY (as authorized
 *      agent for the owner), not the nonprofit.
 *   2. A simple, fully-fillable signature block: Signature / Name / Title /
 *      Company, preceded by a one-line authority statement.
 */
import {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { jsPDF } from 'jspdf';
import type { Analysis } from './analyze';
import type { CertConfig } from './entity';
import { countiesForState, FY, LIMITS_EFFECTIVE, LIMITS_SOURCE } from './limits';

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine',
  'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen',
  'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
function wordsNum(n: number): string {
  if (n < 20) return ONES[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : '');
  return String(n);
}
const blank = (v: string | undefined, w = 24) => (v && v.trim() ? v : '_'.repeat(w));

interface TierRow { label: string; units: number; pct: number; req: string; pass: boolean }

/** Shared, render-agnostic content model so docx + pdf never diverge. */
export function letterContent(analysis: Analysis, config: CertConfig) {
  const { roll, scopes, isGroup, sources, perSrc } = analysis;
  const co = config.company, prop = config.property, non = config.nonprofit;
  const p = roll.pct, cnt = roll.counts;
  const companyName = isGroup ? config.portfolio!.groupName : co.legalName;
  const companyType = isGroup ? config.portfolio!.groupStateType : co.stateType;
  const subsidiaryDesc = config.portfolio?.subsidiaryDescription ?? 'wholly-owned single-purpose subsidiary LLCs';
  const taxYear = config.filing.taxYear;

  const countiesStr =
    prop.counties.join(' and ') + ' ' + (prop.counties.length > 1 ? 'Counties' : 'County');
  const stateCounties = countiesForState(prop.state);
  const msaStr = [...new Set(prop.counties.map((c) => stateCounties[c]?.msa).filter(Boolean))].join('; ');

  const companyClause = isGroup
    ? `${companyName} (the “Company”) and its ${subsidiaryDesc} listed in Section 2`
    : `${companyName} (the “Company”)`;

  // Nonprofit ownership comes from the hub and varies by LLC. For a group it is
  // shown per-subsidiary in Section 2; for a single LLC it's stated inline.
  const members = config.portfolio?.members ?? [];
  const groupHasOwnership = isGroup && members.length > 0;
  const pctStr = non.ownershipPercent == null ? '____%' : `${non.ownershipPercent}%`;
  const clsStr = non.memberClass || '____';
  const ownershipParam = groupHasOwnership
    ? 'Varies by subsidiary — see Section 2'
    : `${pctStr} ${clsStr} Interest`;
  const ownershipClause = groupHasOwnership
    ? 'holding the membership interest in each subsidiary set forth in Section 2'
    : non.ownershipPercent == null
      ? 'holding the membership interest set forth in its operating agreement'
      : `holding a ${Number.isInteger(non.ownershipPercent) ? wordsNum(non.ownershipPercent) : non.ownershipPercent} percent (${non.ownershipPercent}%) ${clsStr} Interest`;

  const params: [string, string][] = [
    ['Property', prop.addressLine || '_'.repeat(40)],
    ['Description', `${roll.denom} ${prop.description}` +
      (isGroup ? ` held across ${sources.length} single-purpose LLCs (see Section 2)` : '')],
    ['Entity', `${companyName} (${companyType})` + (isGroup ? ' — portfolio / group filing' : '')],
    ['Nonprofit Managing Member', `${non.managingMemberName} (instrumentality of ${non.parentName})`],
    ['Nonprofit Ownership', ownershipParam],
    ['Certifying Party', `Property management company, as ${config.certification.relationshipToOwner}`],
    ['CAHP EIN', blank(non.parentEin, 14)],
    ['Entity EIN', blank(co.ein, 14)],
    ['County', countiesStr],
    ['DOR Account ID', blank(co.dorAccountId, 14)],
    ['Tax Year', String(taxYear)],
    ['Filing Type', config.filing.filingType],
    ['HUD/SC Housing Limits', `FY${FY} (effective ${LIMITS_EFFECTIVE}); ${msaStr}`],
  ];

  // Certify under a single set-aside program: prefer 20/50 (50% AMI), else
  // 40/60 (60% AMI). Show only that program's deep tier (plus the shared 75%
  // low-income / 25% market tests). When neither qualifies, show all tiers.
  const lowRow: TierRow = { label: 'Low-Income (≤80% AMI)', units: cnt.le80, pct: p.le80, req: '≥75%', pass: p.le80 >= 75 };
  const marketRow: TierRow = { label: 'Market (>80% AMI)', units: cnt.market, pct: p.market, req: '≤25%', pass: p.market <= 25 };
  const row50: TierRow = { label: 'Very Low-Income (≤50% AMI)', units: cnt.le50, pct: p.le50, req: '≥20%', pass: p.le50 >= 20 };
  const row60: TierRow = { label: 'Low-Income (≤60% AMI)', units: cnt.le60, pct: p.le60, req: '≥40%', pass: p.le60 >= 40 };
  const tierRows: TierRow[] =
    scopes.chosen === '20/50' ? [row50, lowRow, marketRow]
    : scopes.chosen === '40/60' ? [row60, lowRow, marketRow]
    : [lowRow, row60, row50, marketRow];
  const programLabel =
    scopes.chosen === '20/50' ? 'the 20%-at-50%-AMI set-aside (Rev. Proc. 96-32 §3.02)'
    : scopes.chosen === '40/60' ? 'the 40%-at-60%-AMI set-aside (Rev. Proc. 96-32 §3.02)'
    : null;

  const memberByName = new Map(members.map((m) => [m.name, m]));
  const perLlcRows = isGroup
    ? sources.map((s) => {
        const r = perSrc[s];
        const m = memberByName.get(s);
        const own = m?.ownershipPercent == null ? '—' : `${m.ownershipPercent}%`;
        return { s, denom: r.denom, le50: r.pct.le50, le60: r.pct.le60, le80: r.pct.le80, market: r.pct.market, own, cls: m?.memberClass || '—' };
      })
    : [];

  const ownerRef = isGroup
    ? `${companyName} and the subsidiary LLCs listed in Section 2`
    : companyName;

  return {
    taxYear, companyName, companyClause, countiesStr, msaStr, ownerRef,
    subsidiaryDesc, params, tierRows, perLlcRows, p, roll, non,
    determination: scopes.headline,
    programLabel,
    ownershipClause,
    groupHasOwnership,
    citation: config.jurisdiction.statuteCitation,
    recipient: config.jurisdiction.recipient,
    reviewWarning: roll.nReview
      ? `${roll.nReview} unit(s) could not be auto-classified (missing bedroom count, rent, or ` +
        `county) and are listed in Exhibit A under “NEEDS REVIEW.” The percentages above treat ` +
        `them conservatively. Resolve them and re-run before signing.`
      : null,
    authorityLine:
      `This certification is executed by the property management company that manages the ` +
      `Property, as ${config.certification.relationshipToOwner} for ${ownerRef}.`,
    isGroup, sources,
  };
}

// ───────────────────────────── DOCX ─────────────────────────────
const TEAL = '0F6E6E';
const para = (text: string, opts: { bold?: boolean; size?: number; align?: (typeof AlignmentType)[keyof typeof AlignmentType]; color?: string } = {}) =>
  new Paragraph({
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, size: opts.size ? opts.size * 2 : undefined, color: opts.color })],
  });
const heading = (text: string) =>
  new Paragraph({ spacing: { before: 200, after: 60 }, children: [new TextRun({ text, bold: true, color: TEAL })] });

function kvTable(params: [string, string][]): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: params.map(([k, v]) =>
      new TableRow({
        children: [
          new TableCell({ width: { size: 32, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: k, bold: true })] })] }),
          new TableCell({ width: { size: 68, type: WidthType.PERCENTAGE }, children: [new Paragraph(v)] }),
        ],
      }),
    ),
  });
}

function gridTable(headers: string[], rows: (string | number)[][]): Table {
  const mk = (text: string, bold = false) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold })] })] });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h) => mk(h, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((c) => mk(String(c))) })),
    ],
  });
}

export async function buildLetterDocx(analysis: Analysis, config: CertConfig): Promise<Blob> {
  const m = letterContent(analysis, config);
  const children: (Paragraph | Table)[] = [];

  children.push(para('SAFE HARBOR CERTIFICATION', { bold: true, size: 15, align: AlignmentType.CENTER, color: TEAL }));
  children.push(para(m.citation.toUpperCase(), { bold: true, align: AlignmentType.CENTER }));
  children.push(para('IRS REVENUE PROCEDURE 96-32', { bold: true, align: AlignmentType.CENTER }));
  children.push(new Paragraph(''));
  children.push(kvTable(m.params));
  children.push(new Paragraph(''));
  children.push(para(`TO ${m.recipient.toUpperCase()}:`, { bold: true }));
  children.push(para(
    `The undersigned property management company, as ${config.certification.relationshipToOwner} for ` +
    `${m.companyClause}, hereby certifies under penalty of perjury as follows:`));

  children.push(heading('1. Entity Structure and Nonprofit Status.'));
  children.push(para(
    `${m.non.parentName} (“CAHP”) is a South Carolina nonprofit corporation exempt from federal ` +
    `income tax under Section 501(c)(3) of the Internal Revenue Code. ${m.non.managingMemberName} is ` +
    `a wholly owned instrumentality of CAHP and serves as the managing member of the Company, ` +
    `${m.ownershipClause}. The Property is managed day-to-day by the undersigned property management ` +
    `company, which executes this certification as ${config.certification.relationshipToOwner} for the Company.`));

  children.push(heading('2. Property Description.'));
  if (m.isGroup) {
    children.push(para(
      `The Company holds the residential rental real property through ${wordsNum(m.sources.length)} ` +
      `(${m.sources.length}) ${m.subsidiaryDesc}, together comprising ${wordsNum(m.roll.denom)} ` +
      `(${m.roll.denom}) residential rental units located in ${m.countiesStr}, South Carolina. Unit ` +
      `detail accompanies this certification in the submitted rent roll.`));
  } else {
    children.push(para(
      `The Company owns the residential rental real property described above, consisting of ` +
      `${wordsNum(m.roll.denom)} (${m.roll.denom}) ${config.property.description} located in ` +
      `${m.countiesStr}, South Carolina.`));
  }

  children.push(heading('3. Safe Harbor Qualification — Revenue Procedure 96-32.'));
  children.push(para(
    `Qualification is established on a rent-restriction basis: a unit counts toward an AMI tier when ` +
    `its gross rent (contract rent plus any tenant-paid utility allowance) does not exceed the ` +
    `Maximum Allowable Gross Rent published for the applicable county, bedroom size, and that AMI ` +
    `tier under the HUD MTSP methodology (FY${FY} limits, effective ${LIMITS_EFFECTIVE}). The ` +
    `required percentage of residential units is so rent-restricted, satisfying the Rev. Proc. 96-32 ` +
    `set-aside as follows:`));
  children.push(new Paragraph({ children: [new TextRun({ text: 'Set-Aside Determination: ', bold: true }), new TextRun({ text: m.determination, bold: true })] }));
  if (m.programLabel) {
    children.push(para(`The Property is certified under ${m.programLabel}; the test for that program is shown below.`));
  }
  children.push(gridTable(
    ['AMI Tier', 'Units', '% Total', 'Required', 'Result'],
    m.tierRows.map((t) => [t.label, t.units, `${t.pct}%`, t.req, t.pass ? 'PASS' : 'FAIL']),
  ));
  if (m.reviewWarning) children.push(para(`⚠ ${m.reviewWarning}`, { bold: true, color: 'B00000' }));

  children.push(heading('4. Exemption Request.'));
  children.push(para(
    `Based on the foregoing, the Company respectfully requests a full exemption from ad valorem ` +
    `property taxation under ${m.citation} for tax year ${m.taxYear}.`));

  children.push(heading('5. Enclosures.'));
  [
    '(a) Current rent roll (submitted herewith);',
    `(b) FY${FY} HUD/SC Housing rent limits for ${m.msaStr};`,
    `(c) Confirmation of ${m.non.managingMemberName} ${m.non.ownershipPercent}% ${m.non.memberClass} ownership interest;`,
    `(d) IRS 501(c)(3) Determination Letter for ${m.non.parentName}.`,
  ].forEach((e) => children.push(para(e)));

  children.push(heading('6. Ongoing Compliance.'));
  children.push(para(
    `The Company and its property manager commit to maintaining Safe Harbor compliance on an ongoing ` +
    `basis and filing annual certifications by ${config.filing.annualCertificationDeadline} of each ` +
    `year, as required by law.`));

  // ── Signature block (simple, fully fillable) ──
  children.push(new Paragraph(''));
  children.push(para(`Certified under penalty of perjury this _____ day of _______________, ${m.taxYear}.`));
  children.push(para(m.authorityLine, { size: 9 }));
  children.push(new Paragraph(''));
  ['Signature: ______________________________',
   'Name:      ______________________________',
   'Title:     ______________________________',
   'Company:   ______________________________',
   'Date:      ______________________________'].forEach((l) => children.push(para(l)));

  children.push(new Paragraph(''));
  children.push(para(`Rent limits source: ${LIMITS_SOURCE} Verify against the live source before filing.`, { size: 7.5 }));

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
    sections: [{ children }],
  });
  return Packer.toBlob(doc);
}

// ───────────────────────────── PDF ─────────────────────────────
export function buildLetterPdf(analysis: Analysis, config: CertConfig): Blob {
  const m = letterContent(analysis, config);
  const doc = new jsPDF({ unit: 'pt', format: 'letter' });
  const margin = 54;
  const width = doc.internal.pageSize.getWidth() - margin * 2;
  let y = margin;
  const ensure = (h = 14) => {
    if (y + h > doc.internal.pageSize.getHeight() - margin) { doc.addPage(); y = margin; }
  };
  const line = (text: string, opts: { bold?: boolean; size?: number; center?: boolean } = {}) => {
    doc.setFont('helvetica', opts.bold ? 'bold' : 'normal');
    doc.setFontSize(opts.size ?? 10);
    const wrapped = doc.splitTextToSize(text, width);
    for (const w of wrapped) {
      ensure();
      doc.text(w, opts.center ? doc.internal.pageSize.getWidth() / 2 : margin, y, { align: opts.center ? 'center' : 'left' });
      y += (opts.size ?? 10) + 4;
    }
  };
  const gap = (h = 6) => { y += h; };

  line('SAFE HARBOR CERTIFICATION', { bold: true, size: 15, center: true });
  line(`${m.citation.toUpperCase()}  ·  IRS REVENUE PROCEDURE 96-32`, { bold: true, size: 9, center: true });
  gap();
  m.params.forEach(([k, v]) => line(`${k}:  ${v}`, { size: 9 }));
  gap();
  line(`TO ${m.recipient.toUpperCase()}:`, { bold: true, size: 9 });
  line(`The undersigned property management company, as ${config.certification.relationshipToOwner} for ${m.companyClause}, hereby certifies under penalty of perjury as follows:`);
  gap();
  line('3. Safe Harbor Qualification — Revenue Procedure 96-32.', { bold: true });
  line(`Set-Aside Determination: ${m.determination}`, { bold: true });
  if (m.programLabel) line(`The Property is certified under ${m.programLabel}; the test for that program is shown below.`);
  m.tierRows.forEach((t) =>
    line(`  ${t.label}:  ${t.units} units (${t.pct}%)  — required ${t.req} — ${t.pass ? 'PASS' : 'FAIL'}`, { size: 9 }));
  if (m.reviewWarning) { gap(); line(`⚠ ${m.reviewWarning}`, { bold: true, size: 9 }); }
  gap();
  line(`4. Exemption Request. The Company requests a full exemption under ${m.citation} for tax year ${m.taxYear}.`);
  gap(12);
  line(`Certified under penalty of perjury this _____ day of _______________, ${m.taxYear}.`);
  line(m.authorityLine, { size: 8 });
  gap();
  ['Signature: ______________________________',
   'Name:      ______________________________',
   'Title:     ______________________________',
   'Company:   ______________________________',
   'Date:      ______________________________'].forEach((l) => line(l));
  gap();
  line(`Rent limits source: ${LIMITS_SOURCE} Verify against the live source before filing.`, { size: 7 });

  return doc.output('blob');
}
