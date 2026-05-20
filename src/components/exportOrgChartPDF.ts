/**
 * Org Chart PDF Export — Native Vector
 * -------------------------------------
 * Builds a polished DOR-style org chart PDF directly with jsPDF vector
 * primitives, then uploads it to the Org Charts SharePoint library.
 *
 * Replaces an earlier html2canvas approach which left the property-entity
 * box at the bottom of the chart visually blank: html2canvas failed to
 * paint white text on the inline-styled teal card, even with all colors
 * force-inlined in an onclone pass.
 *
 * Layout matches the convention DOR prefers (and the 135 Oakwood reference
 * chart Brandy uses as the polish target): beneficial owners up top, the
 * property entity at the bottom, with labeled edges showing each
 * relationship's class and percentage.
 */

import jsPDF from 'jspdf';
import { uploadDocument } from '../lib/sharepoint';
import type { OwnershipNode, Property } from '../lib/sharepoint';

// =============================================================================
// Public API
// =============================================================================

export interface ExportOrgChartOptions {
  /** Full ownership tree rooted at the property (direct owners with ancestors via `children`). */
  tree: OwnershipNode[];
  /** Property record — title, legal entity, state, address all come from here. */
  property: Property;
  /** Name of the direct Managing Member, surfaced in the property card subtitle. */
  managerName?: string;
  /** Filenames already in the Org Charts library — used to auto-suffix _v2/_v3 on collision. */
  existingFilenames: string[];
  /** Progress callback (0-100). */
  onProgress?: (percent: number, label: string) => void;
}

export interface ExportOrgChartResult {
  webUrl: string;
  filename: string;
}

// =============================================================================
// Card classification + colors
// =============================================================================

type CardRole =
  | 'property'
  | 'nonprofit'
  | 'managing-llc'
  | 'sponsor-llc'
  | 'lp-pool'
  | 'individual'
  | 'trust'
  | 'corporation'
  | 'generic';

type RGB = [number, number, number];

const PALETTE: Record<CardRole, { fill: RGB; text: RGB; accent: RGB }> = {
  // Olive-green for the property entity at the bottom — matches 135 Oakwood
  property: { fill: [99, 131, 70], text: [255, 255, 255], accent: [222, 232, 208] },
  // Dark navy for nonprofits
  nonprofit: { fill: [30, 58, 95], text: [255, 255, 255], accent: [203, 213, 225] },
  // Teal for the managing/sole-member LLC layer (CAHP SC pattern)
  'managing-llc': { fill: [15, 118, 110], text: [255, 255, 255], accent: [204, 251, 241] },
  // Amber/orange for sponsor LLCs (VanRock pattern — Class A non-managing)
  'sponsor-llc': { fill: [184, 118, 58], text: [255, 255, 255], accent: [254, 215, 170] },
  // Dark slate for LP/GP investor pools
  'lp-pool': { fill: [71, 85, 105], text: [255, 255, 255], accent: [203, 213, 225] },
  // Darker slate for individuals
  individual: { fill: [55, 65, 81], text: [255, 255, 255], accent: [203, 213, 225] },
  trust: { fill: [124, 58, 237], text: [255, 255, 255], accent: [221, 214, 254] },
  corporation: { fill: [30, 64, 175], text: [255, 255, 255], accent: [191, 219, 254] },
  generic: { fill: [107, 114, 128], text: [255, 255, 255], accent: [229, 231, 235] },
};

function classifyCard(node: OwnershipNode): CardRole {
  const type = node.owner?.fields.OwnerType;
  const relType = node.relationship.fields.RelationshipType;
  if (type === 'Nonprofit') return 'nonprofit';
  if (type === 'Limited Partnership' || type === 'General Partnership') return 'lp-pool';
  if (type === 'LLC') {
    if (relType === 'Managing Member' || relType === 'Sole Member') return 'managing-llc';
    return 'sponsor-llc';
  }
  if (type === 'Individual') return 'individual';
  if (type === 'Trust') return 'trust';
  if (type === 'Corporation') return 'corporation';
  return 'generic';
}

// =============================================================================
// US state code → full name (for "South Carolina LLC" lines on cards)
// =============================================================================

const STATE_NAMES: Record<string, string> = {
  AL: 'Alabama', AK: 'Alaska', AZ: 'Arizona', AR: 'Arkansas', CA: 'California',
  CO: 'Colorado', CT: 'Connecticut', DE: 'Delaware', FL: 'Florida', GA: 'Georgia',
  HI: 'Hawaii', ID: 'Idaho', IL: 'Illinois', IN: 'Indiana', IA: 'Iowa',
  KS: 'Kansas', KY: 'Kentucky', LA: 'Louisiana', ME: 'Maine', MD: 'Maryland',
  MA: 'Massachusetts', MI: 'Michigan', MN: 'Minnesota', MS: 'Mississippi', MO: 'Missouri',
  MT: 'Montana', NE: 'Nebraska', NV: 'Nevada', NH: 'New Hampshire', NJ: 'New Jersey',
  NM: 'New Mexico', NY: 'New York', NC: 'North Carolina', ND: 'North Dakota', OH: 'Ohio',
  OK: 'Oklahoma', OR: 'Oregon', PA: 'Pennsylvania', RI: 'Rhode Island', SC: 'South Carolina',
  SD: 'South Dakota', TN: 'Tennessee', TX: 'Texas', UT: 'Utah', VT: 'Vermont',
  VA: 'Virginia', WA: 'Washington', WV: 'West Virginia', WI: 'Wisconsin', WY: 'Wyoming',
  DC: 'District of Columbia',
};
function spellState(code: string | undefined): string {
  if (!code) return '';
  return STATE_NAMES[code.trim().toUpperCase()] ?? code;
}

// =============================================================================
// Card / edge model build
// =============================================================================

const PROPERTY_ID = '__PROPERTY__';

interface CardModel {
  id: string;
  role: CardRole;
  title: string;
  subLines: string[];
  /** 0 = property (bottom), 1 = direct owner, 2+ = ancestors. */
  level: number;
  /** Fractional column index — 0 is leftmost leaf column. */
  col: number;
  /** ID of the downstream card this card connects DOWN to (null for property). */
  downstreamId: string | null;
  /** Label to render on the edge from this card to its downstream. */
  edgeLabel: string;
}

function buildSubLines(node: OwnershipNode): string[] {
  const lines: string[] = [];
  const type = node.owner?.fields.OwnerType;
  const state = spellState(node.owner?.fields.OwnerState);
  const description = node.owner?.fields.EntityDescription;

  // Manual EntityDescription overrides the auto-derived type line.
  if (description) {
    lines.push(description);
  } else if (type === 'Nonprofit') {
    lines.push(`${state ? state + ' ' : ''}Nonprofit Corp.`);
  } else if (type === 'LLC') {
    const isSingleMember =
      node.children.length === 1 &&
      node.children[0].relationship.fields.OwnershipPercent === 100;
    lines.push(`${state ? state + ' ' : ''}${isSingleMember ? 'Single-Member ' : ''}LLC`);
  } else if (type === 'Limited Partnership') {
    lines.push(`${state ? state + ' ' : ''}Limited Partnership`);
  } else if (type === 'General Partnership') {
    lines.push(`${state ? state + ' ' : ''}General Partnership`);
  } else if (type === 'Trust') {
    lines.push('Trust');
  } else if (type === 'Corporation') {
    lines.push(`${state ? state + ' ' : ''}Corporation`);
  } else if (type === 'Individual') {
    lines.push('Individual');
  } else if (type) {
    lines.push(type);
  }

  if (node.owner?.fields.IsTaxExempt && type === 'Nonprofit') {
    lines.push('IRC § 501(c)(3) Tax-Exempt');
  }
  if (node.owner?.fields.SponsorName) {
    lines.push(`Sponsor: ${node.owner.fields.SponsorName}`);
  }
  return lines;
}

function buildPropertySubLines(property: Property, managerName: string | undefined): string[] {
  const lines: string[] = [];
  const state = spellState(property.fields.cahpState);
  if (state) lines.push(`${state} LLC`);
  if (managerName) lines.push(`Manager-Managed by ${managerName}`);
  if (property.fields.PropertyAddress) lines.push(property.fields.PropertyAddress);
  return lines;
}

function buildEdgeLabel(node: OwnershipNode): string {
  const cls = node.relationship.fields.MemberClass;
  const type = node.relationship.fields.RelationshipType;
  const pct = node.relationship.fields.OwnershipPercent;

  // Special-case the 100% sole-member chain (CAHP Inc → CAHP SC)
  if (type === 'Sole Member' && pct === 100) return '100% Sole Member';

  const roleParts: string[] = [];
  if (type === 'Managing Member') roleParts.push('Manager');
  if (cls) roleParts.push(cls);
  if (!type && !cls) return pct != null ? `${pct}%` : '';
  if (type === 'Member' && !cls) roleParts.push('Member');
  if (type === 'Sole Member' && !cls) roleParts.push('Sole Member');
  if (type && type !== 'Managing Member' && type !== 'Member' && type !== 'Sole Member') {
    roleParts.push(type);
  }

  const role = roleParts.join(' / ');
  const pctStr = pct != null ? `${pct}%` : '';
  if (role && pctStr) return `${role} · ${pctStr}`;
  return role || pctStr;
}

/**
 * Lays out the ownership tree into positioned cards.
 *
 * Tree shape reminder: `node.children` are PARENTS (upstream owners),
 * which we want rendered ABOVE the node in the chart. The property sits
 * at level 0 (bottom-center), direct owners at level 1, ancestors above.
 *
 * Within each subtree we use a classic centered-tree layout: a node's
 * column is the centroid of its children's columns; leaves occupy 1
 * column each. Multi-parent nodes therefore widen the layout naturally.
 */
function buildModel(
  tree: OwnershipNode[],
  property: Property,
  managerName: string | undefined,
): { cards: CardModel[]; totalCols: number } {
  const cards: CardModel[] = [];

  function walk(
    node: OwnershipNode,
    level: number,
    downstreamId: string,
  ): { width: number; items: CardModel[] } {
    const cardId = `rel-${node.relationship.id}`;
    const items: CardModel[] = [];
    let widthAccum = 0;

    for (const parent of node.children) {
      const sub = walk(parent, level + 1, cardId);
      for (const item of sub.items) {
        items.push({ ...item, col: item.col + widthAccum });
      }
      widthAccum += sub.width;
    }

    const myWidth = Math.max(1, widthAccum);
    const myCol = widthAccum === 0 ? 0 : (widthAccum - 1) / 2;

    items.push({
      id: cardId,
      role: classifyCard(node),
      title: node.owner?.fields.Title ?? '(unresolved)',
      subLines: buildSubLines(node),
      level,
      col: myCol,
      downstreamId,
      edgeLabel: buildEdgeLabel(node),
    });

    return { width: myWidth, items };
  }

  let widthAccum = 0;
  for (const direct of tree) {
    const sub = walk(direct, 1, PROPERTY_ID);
    for (const item of sub.items) {
      cards.push({ ...item, col: item.col + widthAccum });
    }
    widthAccum += sub.width;
  }

  cards.push({
    id: PROPERTY_ID,
    role: 'property',
    title: property.fields.LegalEntity || property.fields.Title || '(unnamed)',
    subLines: buildPropertySubLines(property, managerName),
    level: 0,
    col: widthAccum > 0 ? (widthAccum - 1) / 2 : 0,
    downstreamId: null,
    edgeLabel: '',
  });

  return { cards, totalCols: Math.max(1, widthAccum) };
}

// =============================================================================
// Drawing primitives
// =============================================================================

const CARD_WIDTH = 200;
const CARD_HEIGHT = 84;
const COL_GAP = 28;
const ROW_GAP = 48;  // vertical space between card rows (for arrow + label)
const COL_WIDTH = CARD_WIDTH + COL_GAP;
const ROW_HEIGHT = CARD_HEIGHT + ROW_GAP;

const PAGE_MARGIN = 36;
const HEADER_HEIGHT = 56;
const FOOTER_HEIGHT = 28;

interface BoxRect {
  x: number;
  y: number;
  w: number;
  h: number;
  cx: number;
  cy: number;
}

function setFill(pdf: jsPDF, c: RGB) { pdf.setFillColor(c[0], c[1], c[2]); }
function setText(pdf: jsPDF, c: RGB) { pdf.setTextColor(c[0], c[1], c[2]); }
function setDraw(pdf: jsPDF, c: RGB) { pdf.setDrawColor(c[0], c[1], c[2]); }

function drawCard(pdf: jsPDF, card: CardModel, rect: BoxRect): void {
  const palette = PALETTE[card.role];

  // Card body — rounded rectangle filled with the role color
  setFill(pdf, palette.fill);
  setDraw(pdf, palette.fill);
  pdf.roundedRect(rect.x, rect.y, rect.w, rect.h, 6, 6, 'F');

  // Title — bold, centered
  setText(pdf, palette.text);
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(11);
  const titleLines = (pdf.splitTextToSize(card.title, rect.w - 16) as string[]).slice(0, 2);
  const titleStartY = rect.y + 18;
  titleLines.forEach((line, idx) => {
    pdf.text(line, rect.cx, titleStartY + idx * 13, { align: 'center' });
  });

  // Subtitle lines — lighter weight
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(8.5);
  setText(pdf, palette.accent);
  const subStartY = titleStartY + titleLines.length * 13 + 4;
  card.subLines.slice(0, 3).forEach((line, idx) => {
    const wrapped = (pdf.splitTextToSize(line, rect.w - 16) as string[])[0] ?? line;
    pdf.text(wrapped, rect.cx, subStartY + idx * 11, { align: 'center' });
  });
}

function drawEdge(pdf: jsPDF, label: string, fromRect: BoxRect, toRect: BoxRect): void {
  const x1 = fromRect.cx;
  const y1 = fromRect.y + fromRect.h;
  const x2 = toRect.cx;
  const y2 = toRect.y;

  setDraw(pdf, [148, 163, 184]); // slate-400
  pdf.setLineWidth(0.7);

  const dx = Math.abs(x1 - x2);
  if (dx < 0.5) {
    pdf.line(x1, y1, x2, y2 - 1);
  } else {
    // Orthogonal L-shape: down, across, down
    const midY = (y1 + y2) / 2;
    pdf.line(x1, y1, x1, midY);
    pdf.line(x1, midY, x2, midY);
    pdf.line(x2, midY, x2, y2 - 1);
  }

  // Arrowhead at the target box (small filled triangle)
  setFill(pdf, [148, 163, 184]);
  pdf.triangle(x2 - 3.5, y2 - 4.5, x2 + 3.5, y2 - 4.5, x2, y2, 'F');

  if (label) {
    const midX = (x1 + x2) / 2;
    // Position the label on the horizontal segment for L-shaped, midpoint for straight
    const labelY = dx < 0.5 ? (y1 + y2) / 2 : (y1 + y2) / 2;

    pdf.setFont('helvetica', 'italic');
    pdf.setFontSize(8);
    const textWidth = pdf.getTextWidth(label);

    // Soft white pill behind the label so it doesn't collide with the line
    setFill(pdf, [255, 255, 255]);
    pdf.roundedRect(midX - textWidth / 2 - 4, labelY - 6, textWidth + 8, 11, 2, 2, 'F');

    setText(pdf, [71, 85, 105]); // slate-600
    pdf.text(label, midX, labelY + 1, { align: 'center' });
  }
}

function drawHeader(
  pdf: jsPDF,
  title: string,
  subtitle: string,
  x: number,
  y: number,
  w: number,
): void {
  setText(pdf, [17, 24, 39]); // gray-900
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(17);
  pdf.text(title, x, y + 20);

  if (subtitle) {
    pdf.setFont('helvetica', 'italic');
    pdf.setFontSize(10);
    setText(pdf, [107, 114, 128]); // gray-500
    pdf.text(subtitle, x, y + 36);
  }

  setDraw(pdf, [209, 213, 219]); // gray-300
  pdf.setLineWidth(0.5);
  pdf.line(x, y + 46, x + w, y + 46);
}

function drawFooter(pdf: jsPDF, x: number, y: number, w: number): void {
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(8);
  setText(pdf, [148, 163, 184]); // slate-400
  pdf.text('Confidential.', x, y);
  pdf.text(
    'Prepared by Carolina Affordable Housing Project, Inc.',
    x + w,
    y,
    { align: 'right' },
  );
}

// =============================================================================
// Filename helpers
// =============================================================================

function buildBaseFilename(propertyTitle: string): string {
  const safe = propertyTitle.replace(/[\\/:*?"<>|]/g, '').trim() || 'Property';
  const date = new Date().toISOString().slice(0, 10);
  return `${safe} - Org Chart - ${date}.pdf`;
}

function resolveCollision(base: string, existing: string[]): string {
  const lowerSet = new Set(existing.map((s) => s.toLowerCase()));
  if (!lowerSet.has(base.toLowerCase())) return base;
  const dotIdx = base.lastIndexOf('.');
  const stem = dotIdx >= 0 ? base.slice(0, dotIdx) : base;
  const ext = dotIdx >= 0 ? base.slice(dotIdx) : '';
  for (let v = 2; v < 100; v++) {
    const candidate = `${stem}_v${v}${ext}`;
    if (!lowerSet.has(candidate.toLowerCase())) return candidate;
  }
  return `${stem}_${Date.now()}${ext}`;
}

// =============================================================================
// Main entry point
// =============================================================================

export async function exportOrgChartPDF(
  options: ExportOrgChartOptions,
): Promise<ExportOrgChartResult> {
  const { tree, property, managerName, existingFilenames, onProgress } = options;

  onProgress?.(10, 'Laying out chart…');

  const { cards, totalCols } = buildModel(tree, property, managerName);
  const maxLevel = cards.reduce((m, c) => Math.max(m, c.level), 0);

  // Compute the natural chart dimensions, then size the page to fit.
  const chartContentW = totalCols * COL_WIDTH - COL_GAP; // last col has no trailing gap
  const chartContentH = (maxLevel + 1) * ROW_HEIGHT - ROW_GAP;
  const minPageW = chartContentW + 2 * PAGE_MARGIN;
  const minPageH = chartContentH + 2 * PAGE_MARGIN + HEADER_HEIGHT + FOOTER_HEIGHT;

  // Prefer landscape letter, but grow either dimension as needed.
  const pageW = Math.max(792, minPageW);
  const pageH = Math.max(612, minPageH);

  const pdf = new jsPDF({
    orientation: pageW > pageH ? 'landscape' : 'portrait',
    unit: 'pt',
    format: [pageW, pageH],
  });

  // Subtle off-white page background to match the 135 Oakwood reference's lift.
  setFill(pdf, [255, 255, 255]);
  pdf.rect(0, 0, pageW, pageH, 'F');

  onProgress?.(35, 'Drawing chart…');

  // Header
  const headerX = PAGE_MARGIN;
  const headerY = PAGE_MARGIN;
  const headerW = pageW - 2 * PAGE_MARGIN;
  const headerTitle = `${property.fields.Title ?? 'Property'} — Org Chart`;
  const headerSubParts: string[] = [];
  const stateName = spellState(property.fields.cahpState);
  if (stateName) headerSubParts.push(`${stateName} LLC`);
  if (property.fields.PropertyAddress) headerSubParts.push(property.fields.PropertyAddress);
  drawHeader(pdf, headerTitle, headerSubParts.join(' · '), headerX, headerY, headerW);

  // Position each card
  const chartTop = PAGE_MARGIN + HEADER_HEIGHT;
  const chartLeft = (pageW - chartContentW) / 2;

  const rectById = new Map<string, BoxRect>();
  for (const card of cards) {
    const cx = chartLeft + card.col * COL_WIDTH + CARD_WIDTH / 2;
    const rowTop = chartTop + (maxLevel - card.level) * ROW_HEIGHT;
    const rect: BoxRect = {
      x: cx - CARD_WIDTH / 2,
      y: rowTop,
      w: CARD_WIDTH,
      h: CARD_HEIGHT,
      cx,
      cy: rowTop + CARD_HEIGHT / 2,
    };
    rectById.set(card.id, rect);
  }

  // Draw edges first so cards paint on top of any overlapping line bits.
  for (const card of cards) {
    if (!card.downstreamId) continue;
    const from = rectById.get(card.id);
    const to = rectById.get(card.downstreamId);
    if (!from || !to) continue;
    drawEdge(pdf, card.edgeLabel, from, to);
  }

  // Cards
  for (const card of cards) {
    const rect = rectById.get(card.id);
    if (!rect) continue;
    drawCard(pdf, card, rect);
  }

  // Footer
  drawFooter(pdf, PAGE_MARGIN, pageH - PAGE_MARGIN, pageW - 2 * PAGE_MARGIN);

  onProgress?.(70, 'Uploading to SharePoint…');

  const pdfBlob = pdf.output('blob');
  const baseFilename = buildBaseFilename(property.fields.Title ?? 'Property');
  const finalFilename = resolveCollision(baseFilename, existingFilenames);

  const result = await uploadDocument({
    libraryName: 'Org Charts',
    filename: finalFilename,
    file: pdfBlob,
    metadata: {
      PropertyLookupId: String(property.id),
      ChartDate: new Date().toISOString().slice(0, 10),
    },
  });

  onProgress?.(100, 'Done');

  return {
    webUrl: result.webUrl,
    filename: finalFilename,
  };
}
