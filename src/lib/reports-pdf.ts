/**
 * Branded PDF helpers for owner-facing report deliverables.
 *
 * Uses jsPDF directly (already a dep via the org-chart export). Targets
 * letter-size portrait, teal header bar + gold accent rule, monospaced
 * numeric columns to match the hub's brand. Suitable for emailing straight
 * to an owner without further formatting.
 */

import jsPDF from 'jspdf';
import { toDateInputValue } from './dates';

// =============================================================================
// Page constants
// =============================================================================

const PAGE_W = 612;
const PAGE_H = 792;
const MARGIN_X = 48;
const MARGIN_TOP = 72;
const MARGIN_BOTTOM = 56;

// Brand palette
const TEAL: [number, number, number] = [15, 61, 62];        // #0F3D3E
const GOLD: [number, number, number] = [201, 162, 39];      // #C9A227
const INK: [number, number, number] = [17, 24, 39];         // gray-900
const MUTED: [number, number, number] = [107, 114, 128];    // gray-500
const ROW_ALT: [number, number, number] = [248, 250, 252];  // slate-50
const RULE: [number, number, number] = [209, 213, 219];     // gray-300

// =============================================================================
// Public document builder
// =============================================================================

export interface PdfDoc {
  /** Push a section heading to the current cursor; advances Y. */
  heading(text: string): void;
  /** Push a paragraph of body text; advances Y. */
  paragraph(text: string): void;
  /** Render a key/value pair (small caps label + value). */
  kv(label: string, value: string): void;
  /** Render a table with header row + body rows. Column widths auto-fit. */
  table(headers: string[], rows: string[][], opts?: { widths?: number[] }): void;
  /** Force a new page (used between snapshots in Org Chart History). */
  newPage(): void;
  /** Get the underlying jsPDF for primitive ops (used by the org chart drawer). */
  raw(): jsPDF;
  /** Finish and return the blob. */
  build(): Blob;
}

export interface PdfHeader {
  title: string;
  subtitle?: string;
  rightLabel?: string;        // e.g., "Q2 2026"
}

/**
 * Create a new branded PDF document. Header runs on every page; footer too.
 * The caller pushes content via the returned helpers.
 */
export function createBrandedPDF(header: PdfHeader): PdfDoc {
  const pdf = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'letter' });
  let y = MARGIN_TOP;

  const drawHeader = () => {
    setFill(pdf, TEAL);
    pdf.rect(0, 0, PAGE_W, 44, 'F');

    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(13);
    setText(pdf, [255, 255, 255]);
    pdf.text('CAHP COMPLIANCE HUB', MARGIN_X, 28);

    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    setText(pdf, [222, 232, 240]);
    pdf.text(toDateInputValue(new Date()), PAGE_W - MARGIN_X, 28, { align: 'right' });

    // Gold accent rule
    setDraw(pdf, GOLD);
    pdf.setLineWidth(2);
    pdf.line(MARGIN_X, 50, PAGE_W - MARGIN_X, 50);

    // Title block
    setText(pdf, INK);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(17);
    pdf.text(header.title, MARGIN_X, 70);

    if (header.subtitle) {
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      setText(pdf, MUTED);
      pdf.text(header.subtitle, MARGIN_X, 86);
    }

    if (header.rightLabel) {
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      setText(pdf, MUTED);
      pdf.text(header.rightLabel, PAGE_W - MARGIN_X, 86, { align: 'right' });
    }

    setDraw(pdf, RULE);
    pdf.setLineWidth(0.5);
    pdf.line(MARGIN_X, 96, PAGE_W - MARGIN_X, 96);
  };

  const drawFooter = (pageNum: number, pageCount: number) => {
    setText(pdf, MUTED);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(8);
    pdf.text('Confidential — prepared by Carolina Affordable Housing Project, Inc.', MARGIN_X, PAGE_H - 28);
    pdf.text(`Page ${pageNum} of ${pageCount}`, PAGE_W - MARGIN_X, PAGE_H - 28, { align: 'right' });
  };

  // Initialize the first page
  drawHeader();
  y = 116;

  const ensureSpace = (h: number) => {
    if (y + h <= PAGE_H - MARGIN_BOTTOM) return;
    pdf.addPage();
    drawHeader();
    y = 116;
  };

  const doc: PdfDoc = {
    heading(text) {
      ensureSpace(28);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      setText(pdf, TEAL);
      pdf.text(text, MARGIN_X, y + 12);
      setDraw(pdf, RULE);
      pdf.setLineWidth(0.5);
      pdf.line(MARGIN_X, y + 18, PAGE_W - MARGIN_X, y + 18);
      y += 30;
    },
    paragraph(text) {
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      setText(pdf, INK);
      const width = PAGE_W - 2 * MARGIN_X;
      const lines = pdf.splitTextToSize(text, width) as string[];
      ensureSpace(lines.length * 13 + 6);
      lines.forEach((line, idx) => {
        pdf.text(line, MARGIN_X, y + (idx + 1) * 13);
      });
      y += lines.length * 13 + 8;
    },
    kv(label, value) {
      ensureSpace(16);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(8);
      setText(pdf, MUTED);
      pdf.text(label.toUpperCase(), MARGIN_X, y + 10);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      setText(pdf, INK);
      pdf.text(value || '—', MARGIN_X + 130, y + 10);
      y += 16;
    },
    table(headers, rows, opts) {
      const widths = opts?.widths ?? computeWidths(headers.length);
      const totalWidth = PAGE_W - 2 * MARGIN_X;
      const colW = widths.map((w) => (w / sum(widths)) * totalWidth);

      ensureSpace(24);
      // Header row
      setFill(pdf, TEAL);
      pdf.rect(MARGIN_X, y, totalWidth, 18, 'F');
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(9);
      setText(pdf, [255, 255, 255]);
      let cx = MARGIN_X + 6;
      headers.forEach((h, i) => {
        pdf.text(h, cx, y + 12);
        cx += colW[i];
      });
      y += 18;

      // Body rows
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      rows.forEach((row, rIdx) => {
        // Compute row height across wrapped columns
        const wrappedCells = row.map((cell, i) => {
          const text = cell ?? '';
          return pdf.splitTextToSize(text, colW[i] - 12) as string[];
        });
        const maxLines = Math.max(1, ...wrappedCells.map((c) => c.length));
        const rowH = maxLines * 12 + 4;

        ensureSpace(rowH);

        if (rIdx % 2 === 1) {
          setFill(pdf, ROW_ALT);
          pdf.rect(MARGIN_X, y, totalWidth, rowH, 'F');
        }

        setText(pdf, INK);
        let xCursor = MARGIN_X + 6;
        wrappedCells.forEach((cellLines, i) => {
          cellLines.forEach((line, lineIdx) => {
            pdf.text(line, xCursor, y + 10 + lineIdx * 12);
          });
          xCursor += colW[i];
        });

        y += rowH;
      });

      y += 6;
    },
    newPage() {
      pdf.addPage();
      drawHeader();
      y = 116;
    },
    raw() {
      return pdf;
    },
    build() {
      // Walk every page to draw the footer with the final page count
      const total = pdf.getNumberOfPages();
      for (let i = 1; i <= total; i++) {
        pdf.setPage(i);
        drawFooter(i, total);
      }
      return pdf.output('blob');
    },
  };

  return doc;
}

// =============================================================================
// Internal helpers
// =============================================================================

function setFill(pdf: jsPDF, c: [number, number, number]) {
  pdf.setFillColor(c[0], c[1], c[2]);
}
function setText(pdf: jsPDF, c: [number, number, number]) {
  pdf.setTextColor(c[0], c[1], c[2]);
}
function setDraw(pdf: jsPDF, c: [number, number, number]) {
  pdf.setDrawColor(c[0], c[1], c[2]);
}

function computeWidths(n: number): number[] {
  return Array.from({ length: n }, () => 1);
}

function sum(xs: number[]): number {
  return xs.reduce((a, b) => a + b, 0);
}
