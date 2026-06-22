// E-PDF EP3 — the logical structure tree (ISO 32000-1 §14.7). A tagged PDF (and
// Ream writes such) carries a /StructTreeRoot pointing at a tree of /StructElem
// dictionaries that recover reading order and roles. Each element has a type
// (/S — Document, P, H1, Table, TR, TD, L, LI, …) and /K children that are either
// nested elements or marked-content references (an MCID on a page) linking it to
// the text the content-stream interpreter extracted (EP2).

import type { PdfDict, PdfValue } from '@/pdf/objects';

import type { PdfFile } from './document';
import { PDF_NULL, PdfName } from '@/pdf/objects';

/** A marked-content reference: a page index plus an MCID on that page. */
export interface StructMcid {
  /** Zero-based page index the MCID lives on. */
  readonly page: number;
  readonly mcid: number;
}

/**
 * One node of the recovered logical structure tree (ISO 32000-1 §14.7): a
 * `/StructElem`'s role, its own marked-content references (the text it owns) and
 * its child elements.
 */
export interface StructNode {
  /** The `/S` role name (`Document`, `P`, `H1`, `Table`, `TR`, `TD`, `L`, `LI`, …). */
  readonly type: string;
  /** The element's own marked content, linking it to interpreter-extracted text (EP2). */
  readonly mcids: ReadonlyArray<StructMcid>;
  readonly children: ReadonlyArray<StructNode>;
  /** `/Alt` — alternate text (figures). */
  readonly alt?: string;
  /** `/A /Table /ColSpan` on a table cell. */
  readonly colSpan?: number;
  /** `/A /Table /RowSpan` on a table cell. */
  readonly rowSpan?: number;
}

const MAX_NODES = 200_000; // DoS guard on a pathological tree

/**
 * Read the `/StructTreeRoot` (ISO 32000-1 §14.7) into a {@link StructNode} tree,
 * recovering reading order and roles. Walks each `/StructElem`'s `/K` children —
 * resolving nested elements, bare-integer and `/MCR` marked-content references
 * (carrying the owning page), and skipping `/OBJR` object references (no text) —
 * guards against cycles and pathological size, and lifts `/Alt` plus table cell
 * `/ColSpan`/`/RowSpan`. Multiple top-level roots are wrapped in a synthetic
 * `Document` node.
 *
 * @returns The structure-tree root, or `undefined` when the catalog has no
 *          `/StructTreeRoot` (an untagged PDF).
 */
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
