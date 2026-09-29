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
  const typeOf = roleResolver(file, stRoot);

  const read = (value: PdfValue, parentPage: number): StructNode | undefined => {
    const elem = file.resolve(value);
    if (!(elem instanceof Map) || seen.has(elem) || seen.size > MAX_NODES) return undefined;
    seen.add(elem);
    const ownPage = pageIndexOf(elem.get('Pg') ?? PDF_NULL) ?? parentPage;
    // The element's content in the order `/K` gives it: its own marked
    // content, and its child elements between.
    const content: Array<StructMcid | StructNode> = [];
    for (const kid of kidList(file, elem.get('K'))) {
      const rk = file.resolve(kid);
      if (typeof rk === 'number') {
        // A bare integer is an MCID on this element's own page.
        if (ownPage >= 0) content.push({ page: ownPage, mcid: rk });
      } else if (rk instanceof Map) {
        const kind = nameOf(rk.get('Type'));
        if (kind === 'MCR') {
          const m = rk.get('MCID');
          const page = pageIndexOf(rk.get('Pg') ?? PDF_NULL) ?? ownPage;
          if (typeof m === 'number' && page >= 0) content.push({ page, mcid: m });
        } else if (kind === 'OBJR') {
          // an object reference (annotation) — no text
        } else {
          const child = read(rk, ownPage);
          if (!child) continue;
          // §14.8.4.4 — an INLINE element is a stretch of its parent's text,
          // not a block of its own: its words stand where it stands in the
          // parent's line (see `INLINE`).
          if (inline(file, rk, child.type)) content.push(...allMcids(child));
          else content.push(child);
        }
      }
    }
    const type = typeOf(nameOf(elem.get('S')));
    const alt = elem.get('Alt');
    const { colSpan, rowSpan } = readSpans(file, elem.get('A') ?? PDF_NULL);
    return {
      type,
      ...settle(type, content),
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

/**
 * An element's own marked content and its child elements, as the node keeps
 * them.
 *
 * An element holds text of its own AND block elements only where the tree
 * mixes them — a heading carrying its number as a child, a paragraph with a
 * formula set between two of its sentences. Kept as two lists, the order
 * between them was gone and the element's own text was never read at all:
 * bug1937438_mml_from_latex.pdf's heading came back as "1" without "A small
 * example", and its sentence around a formula without its words. Each
 * stretch of the element's own content becomes a paragraph of its own
 * standing where it stood — except in a figure, whose marked content is its
 * picture.
 */
function settle(
  type: string,
  content: ReadonlyArray<StructMcid | StructNode>,
): { mcids: Array<StructMcid>; children: Array<StructNode> } {
  const nodes = content.filter((c): c is StructNode => 'type' in c);
  const own = content.filter((c): c is StructMcid => !('type' in c));
  if (nodes.length === 0 || own.length === 0 || type === 'Figure') {
    return { mcids: own, children: nodes };
  }
  const children: Array<StructNode> = [];
  let stretch: Array<StructMcid> = [];
  const close = (): void => {
    if (stretch.length > 0) children.push({ type: 'P', mcids: stretch, children: [] });
    stretch = [];
  };
  for (const c of content) {
    if ('type' in c) {
      close();
      children.push(c);
    } else stretch.push(c);
  }
  close();
  return { mcids: [], children };
}

/** Every marked content reference under a node, in reading order. */
function allMcids(node: StructNode): Array<StructMcid> {
  return [...node.mcids, ...node.children.flatMap(allMcids)];
}

/**
 * §14.8.4.4 — the inline-level structure types (PDF 2.0 adds `Em`, `Strong`,
 * `Sub`), a formula (§14.8.4.5.5, which is set in a line as often as out of
 * one), and a list label, which is the start of its item's line.
 */
const INLINE = new Set([
  'Span',
  'Quote',
  'Reference',
  'BibEntry',
  'Code',
  'Link',
  'Annot',
  'Ruby',
  'RB',
  'RT',
  'RP',
  'Warichu',
  'WT',
  'WP',
  'Em',
  'Strong',
  'Sub',
  'Formula',
  'Lbl',
]);

/** The MathML namespace (ISO 32000-2 §14.8.6): an element in it is mathematics. */
const MATHML = 'http://www.w3.org/1998/Math/MathML';

/** Whether an element is inline — by its standard type, or as MathML. */
function inline(file: PdfFile, elem: PdfDict, type: string): boolean {
  if (INLINE.has(type)) return true;
  const ns = file.resolve(elem.get('NS') ?? PDF_NULL);
  const uri = ns instanceof Map ? file.resolve(ns.get('NS') ?? PDF_NULL) : undefined;
  return typeof uri === 'string' && uri === MATHML;
}

/**
 * §14.8.4.2 `/RoleMap` — a document's own structure types, mapped to the
 * standard ones they stand for, followed until a type maps no further. A
 * LaTeX document names its elements `section`, `text-unit` and
 * `section-number`, and maps them to `H1`, `Part` and `Span`: read by their
 * own names, none of them meant anything.
 */
function roleResolver(file: PdfFile, stRoot: PdfDict): (type: string) => string {
  const map = file.resolve(stRoot.get('RoleMap') ?? PDF_NULL);
  if (!(map instanceof Map)) return (type) => type;
  return (type) => {
    let at = type;
    for (let hop = 0; hop < MAX_ROLE_HOPS; hop++) {
      const next = file.resolve(map.get(at) ?? PDF_NULL);
      if (!(next instanceof PdfName) || next.value === at) break;
      at = next.value;
    }
    return at;
  };
}

/** How far a chain of role mappings is followed before it is taken for a cycle. */
const MAX_ROLE_HOPS = 8;

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
