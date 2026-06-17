/**
 * Non-AppFolio rent-roll input: extract text from a PDF and/or parse pasted
 * free-text into Safe Harbor units. Formats from outside PMs vary wildly
 * (PMI-style multi-line blocks, one-line "Unit N: $amount" emails, etc.), so
 * this is a best-effort parser — the user reviews the result and supplies the
 * pieces these documents usually omit (bedrooms, county) via bulk defaults.
 */
import { detectCounty } from './rentRoll';
import type { Unit } from './limits';

/** Extract plain text from a PDF in the browser (pdf.js, loaded on demand). */
export async function extractPdfText(file: File): Promise<string> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pdfjs: any = await import('pdfjs-dist');
  // Load the worker from CDN matching the installed version (avoids Vite worker
  // bundling issues; consistent with the app's other CDN-loaded libraries).
  pdfjs.GlobalWorkerOptions.workerSrc =
    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await pdfjs.getDocument({ data }).promise;
  let text = '';
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    text += content.items.map((it: any) => (it && typeof it.str === 'string' ? it.str : '')).join(' ') + '\n';
    page.cleanup();
  }
  return text;
}

export interface ParseTextOpts {
  source: string;               // LLC / property label applied to every parsed unit
  defaultCounty?: string | null;
  defaultBedrooms?: number | null;
}

const UNIT_RE = /\bunit\b\s*[:#]?\s*([0-9A-Za-z][\w-]*)/i;
const MONEY_RE = /\$\s?([\d,]+(?:\.\d{1,2})?)/;

/**
 * Parse pasted/extracted rent-roll text into units. Recognizes a "unit anchor"
 * line (containing "Unit <id>") and, for the block of lines up to the next
 * anchor, takes the first dollar amount as the current rent. "vacant" with no
 * positive amount leaves the rent undetermined (so it defaults to Market, the
 * conservative outcome). County is auto-detected per block when possible, else
 * the bulk default. Bedrooms aren't in these documents — the bulk default fills
 * them (units with no bedroom default to Market).
 */
export function parseTextToUnits(text: string, opts: ParseTextOpts): Unit[] {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 0);
  const anchors: number[] = [];
  lines.forEach((l, i) => {
    if (UNIT_RE.test(l)) anchors.push(i);
  });

  const units: Unit[] = [];
  for (let a = 0; a < anchors.length; a++) {
    const start = anchors[a];
    const end = a + 1 < anchors.length ? anchors[a + 1] : lines.length;
    const block = lines.slice(start, end);
    const blockText = block.join(' ');
    const m = lines[start].match(UNIT_RE);
    const unitLabel = m ? `Unit ${m[1]}` : lines[start];

    const money = blockText.match(MONEY_RE);
    const amount = money ? Number(money[1].replace(/,/g, '')) : null;
    const vacant = /\bvacant\b/i.test(blockText);
    const rent = amount && amount > 0 ? amount : null; // $0 / vacant -> undetermined
    const county = detectCounty(blockText) ?? opts.defaultCounty ?? null;
    const occupied = !vacant && rent !== null;

    units.push({
      source: opts.source,
      prop: lines[start],
      unit: unitLabel,
      county,
      bedrooms: opts.defaultBedrooms ?? null,
      baths: null,
      tenant: '',
      status: vacant ? 'Vacant' : occupied ? 'Occupied' : 'Unknown',
      marketRent: rent,
      contractRent: rent,
      occupied,
      nonResidential: false,
      grossRent: rent,
      tier: null,
      ceil50: null,
      ceil60: null,
      ceil80: null,
      notes: [],
    });
  }
  return units;
}
