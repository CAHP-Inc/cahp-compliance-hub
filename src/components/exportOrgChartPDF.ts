/**
 * Org Chart PDF Export
 * --------------------
 * Captures the rendered DOR-friendly org chart as a PDF and uploads it to
 * the Org Charts SharePoint library, tagged to the property.
 *
 * Uses html2canvas to rasterize the chart container, then jsPDF to wrap the
 * image in a PDF page. Auto-versions the filename if a doc with the same name
 * already exists in the library.
 *
 * USAGE:
 *   const result = await exportOrgChartPDF({
 *     containerRef,           // ref to the chart's outer div
 *     propertyId,             // for tagging the doc
 *     propertyTitle,          // for the filename
 *     existingFilenames,      // for collision detection
 *   });
 */

import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { uploadDocument } from '../lib/sharepoint/client';

export interface ExportOrgChartOptions {
  /** DOM element containing the rendered chart */
  element: HTMLElement;
  /** Property ID for tagging the uploaded doc */
  propertyId: string;
  /** Property title for filename */
  propertyTitle: string;
  /** Filenames already in the library — used to suffix _v2/_v3 on collision */
  existingFilenames: string[];
  /** Progress callback (0-100) */
  onProgress?: (percent: number, label: string) => void;
}

export interface ExportOrgChartResult {
  webUrl: string;
  filename: string;
}

/**
 * Build the base filename — sanitized property title + ISO date.
 * Example: "700 Brook St - Org Chart - 2026-05-18.pdf"
 */
function buildBaseFilename(propertyTitle: string): string {
  // Strip characters SharePoint dislikes in filenames: \ / : * ? " < > |
  const safe = propertyTitle.replace(/[\\/:*?"<>|]/g, '').trim() || 'Property';
  const date = new Date().toISOString().slice(0, 10);
  return `${safe} - Org Chart - ${date}.pdf`;
}

/**
 * Auto-version: if "Foo.pdf" exists, return "Foo_v2.pdf"; if v2 exists, v3, etc.
 */
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
  // Fallback: timestamp suffix
  return `${stem}_${Date.now()}${ext}`;
}

export async function exportOrgChartPDF(
  options: ExportOrgChartOptions
): Promise<ExportOrgChartResult> {
  const { element, propertyId, propertyTitle, existingFilenames, onProgress } = options;

  onProgress?.(10, 'Capturing chart…');

  // Wait for any in-flight fonts to load — html2canvas captures synchronously
  // and otherwise can paint text before custom fonts are ready, leading to
  // missing or fallback-rendered text in the canvas.
  if (typeof document !== 'undefined' && (document as { fonts?: FontFaceSet }).fonts?.ready) {
    try {
      await (document as { fonts: FontFaceSet }).fonts.ready;
    } catch {
      // Non-fatal — proceed with capture
    }
  }

  // Rasterize the chart at 2x for crisper PDF output
  const canvas = await html2canvas(element, {
    scale: 2,
    backgroundColor: '#ffffff',
    logging: false,
    useCORS: true,
    allowTaint: true,
    // Manually inject inline styles on every descendant before render. Fixes a
    // html2canvas quirk where text on dark backgrounds with inherited color
    // sometimes renders invisible. We force-set color/background/font on each
    // node from its computed style so html2canvas has direct values to work with.
    onclone: (clonedDoc: Document, clonedEl: HTMLElement) => {
      const all = clonedEl.querySelectorAll<HTMLElement>('*');
      all.forEach((node) => {
        const computed = clonedDoc.defaultView?.getComputedStyle(node);
        if (!computed) return;
        // Force inline the most important text/visual properties so they
        // survive html2canvas's painting pass.
        node.style.color = computed.color;
        node.style.backgroundColor = computed.backgroundColor;
        node.style.fontFamily = computed.fontFamily;
        node.style.fontSize = computed.fontSize;
        node.style.fontWeight = computed.fontWeight;
        node.style.lineHeight = computed.lineHeight;
        node.style.letterSpacing = computed.letterSpacing;
        node.style.textTransform = computed.textTransform;
      });
    },
  });

  onProgress?.(50, 'Building PDF…');

  // Decide page orientation by aspect ratio
  const isLandscape = canvas.width > canvas.height;
  const pdf = new jsPDF({
    orientation: isLandscape ? 'landscape' : 'portrait',
    unit: 'pt',
    format: 'letter', // 612 x 792 pt portrait
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 36; // 0.5 inch margins
  const usableWidth = pageWidth - 2 * margin;
  const usableHeight = pageHeight - 2 * margin;

  // Scale image to fit within margins while preserving aspect ratio
  const widthRatio = usableWidth / canvas.width;
  const heightRatio = usableHeight / canvas.height;
  const scale = Math.min(widthRatio, heightRatio);
  const imgWidth = canvas.width * scale;
  const imgHeight = canvas.height * scale;
  const x = (pageWidth - imgWidth) / 2;
  const y = (pageHeight - imgHeight) / 2;

  // Use PNG image format from canvas
  const imgData = canvas.toDataURL('image/png');
  pdf.addImage(imgData, 'PNG', x, y, imgWidth, imgHeight);

  // Convert to blob for upload
  const pdfBlob = pdf.output('blob');

  onProgress?.(70, 'Uploading to SharePoint…');

  const baseFilename = buildBaseFilename(propertyTitle);
  const finalFilename = resolveCollision(baseFilename, existingFilenames);

  const result = await uploadDocument({
    libraryName: 'Org Charts',
    filename: finalFilename,
    file: pdfBlob,
    metadata: {
      PropertyLookupId: propertyId,
      ChartDate: new Date().toISOString().slice(0, 10),
    },
  });

  onProgress?.(100, 'Done');

  return {
    webUrl: result.webUrl,
    filename: finalFilename,
  };
}
