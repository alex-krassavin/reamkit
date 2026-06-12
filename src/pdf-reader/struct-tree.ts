// E-PDF EP3 — the logical structure tree (ISO 32000-1 §14.7). A tagged PDF (and
// Ream writes such) carries a /StructTreeRoot pointing at a tree of /StructElem
// dictionaries that recover reading order and roles. Each element has a type
// (/S — Document, P, H1, Table, TR, TD, L, LI, …) and /K children that are either
// nested elements or marked-content references (an MCID on a page) linking it to
// the text the content-stream interpreter extracted (EP2).

import type { PdfDict, PdfValue } from '@/pdf/objects';

import type { PdfFile } from './document';
import { PDF_NULL, PdfName } from '@/pdf/objects';

export interface StructMcid {
  readonly page: number; // page index
  readonly mcid: number;
}

export interface StructNode {
  readonly type: string; // /S role name
  readonly mcids: ReadonlyArray<StructMcid>; // the element's own marked content
  readonly children: ReadonlyArray<StructNode>;
  readonly alt?: string; // /Alt — alternate text (figures)
  readonly colSpan?: number; // /A /Table /ColSpan (table cells)
  readonly rowSpan?: number; // /A /Table /RowSpan
}

const MAX_NODES = 200_000; // DoS guard on a pathological tree

export function readStructTree(file: PdfFile): StructNode | undefined {
  const stRoot = file.get(file.catalog, 'StructTreeRoot');
  if (!(stRoot instanceof Map)) return undefined;

  const pageMap = new Map<PdfDict, number>();
  file.pages().forEach((p, i) => pageMap.set(p.dict, i));
  const pageIndexOf = (pgVal: PdfValue): number | undefined => {
    const pg = file.resolve(pgVal);
    return pg instanceof Map ? pageMap.get(pg) : undefined;
  };
  const seen = new Set<PdfDict>();

  const read = (value: PdfValue, parentPage: number): StructNode | undefined => {
    const elem = file.resolve(value);
    if (!(elem instanceof Map) || seen.has(elem) || seen.size > MAX_NODES) return undefined;
    seen.add(elem);
    const ownPage = pageIndexOf(elem.get('Pg') ?? PDF_NULL) ?? parentPage;
    const mcids: Array<StructMcid> = [];
    const children: Array<StructNode> = [];
    for (const kid of kidList(file, elem.get('K'))) {
      const rk = file.resolve(kid);
      if (typeof rk === 'number') {
        // A bare integer is an MCID on this element's own page.
        if (ownPage >= 0) mcids.push({ page: ownPage, mcid: rk });
      } else if (rk instanceof Map) {
        const kind = nameOf(rk.get('Type'));
        if (kind === 'MCR') {
          const m = rk.get('MCID');
          const page = pageIndexOf(rk.get('Pg') ?? PDF_NULL) ?? ownPage;
          if (typeof m === 'number' && page >= 0) mcids.push({ page, mcid: m });
        } else if (kind === 'OBJR') {
          // an object reference (annotation) — no text
        } else {
          const child = read(rk, ownPage);
          if (child) children.push(child);
        }
      }
    }
    const alt = elem.get('Alt');
    const { colSpan, rowSpan } = readSpans(file, elem.get('A') ?? PDF_NULL);
    return {
      type: nameOf(elem.get('S')),
      mcids,
      children,
      ...(typeof alt === 'string' ? { alt } : {}),
      ...(colSpan > 1 ? { colSpan } : {}),
      ...(rowSpan > 1 ? { rowSpan } : {}),
    };
  };

  const roots = kidList(file, stRoot.get('K'))
    .map((k) => read(k, -1))
    .filter((n): n is StructNode => n !== undefined);
  if (roots.length === 1) return roots[0];
  return { type: 'Document', mcids: [], children: roots };
}

// Normalise /K (a single kid or an array) to a list of unresolved kid values.
function kidList(file: PdfFile, kVal: PdfValue | undefined): Array<PdfValue> {
  if (kVal === undefined) return [];
  const k = file.resolve(kVal);
  if (Array.isArray(k)) return k;
  if (k === PDF_NULL) return [];
  return [k];
}

function nameOf(v: PdfValue | undefined): string {
  return v instanceof PdfName ? v.value : '';
}

// §14.8.5.7 — a table cell's /A attribute dict(s) carry /ColSpan and /RowSpan.
function readSpans(file: PdfFile, aVal: PdfValue): { colSpan: number; rowSpan: number } {
  let colSpan = 1;
  let rowSpan = 1;
  const a = file.resolve(aVal);
  const attrs: ReadonlyArray<PdfValue> = Array.isArray(a) ? a : [a];
  for (const entry of attrs) {
    const d = file.resolve(entry);
    if (d instanceof Map) {
      const cs = d.get('ColSpan');
      const rs = d.get('RowSpan');
      if (typeof cs === 'number') colSpan = cs;
      if (typeof rs === 'number') rowSpan = rs;
    }
  }
  return { colSpan, rowSpan };
}
