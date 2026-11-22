// ISO 32000-1 §12.3.3 — the document outline a PDF carries: its bookmarks,
// read so that a PDF converted to a PDF keeps them. Each item comes back with
// what it says, how deep it sits, whether it stands open, and where it jumps:
// the page, and the height on it when its destination names one — or the URI
// it opens, for an item whose action is a URI action.

import { displayOf } from './display';
import { pdfTextString } from './encodings';
import { CARRIER_LINE_PT } from './flow-build';
import type { PdfFile, PdfPage } from './document';
import type { OutlineEntry } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import { outlineTitleOf } from '@/core/outline';
import { PdfHexString, PdfName } from '@/pdf/objects';

/** One item of a source PDF's document outline, resolved to a page. */
export interface SourceOutlineItem {
  /** §12.3.3 Table 153 `/Title`, decoded (§7.9.2.2) and put on one line. */
  readonly title: string;
  /** Its depth, 1 for a top-level item. */
  readonly level: number;
  /** Table 153 `/Count` — positive for an item that stands open. */
  readonly open: boolean;
  /** The page its destination names, 0-based; absent exactly when it opens a `uri`. */
  readonly pageIndex?: number;
  /**
   * The height its destination puts at the top of the window, y-up in the
   * page as SHOWN (crop box and turn applied, see `./display`) — where the
   * destination names one: `/XYZ`'s, `/FitH`'s, `/FitBH`'s and `/FitR`'s top.
   */
  readonly topPt?: number;
  /**
   * The URI its URI action opens (§12.6.4.7), made absolute against the
   * catalog's `/URI` `/Base` (Table 207) where it is relative. An item that
   * opens one names no page.
   */
  readonly uri?: string;
}

/** A tree past this many items is read no further: a cycle the guard missed, or a bomb. */
const MAX_ITEMS = 10_000;
/** How many name-tree nodes one lookup may visit (§7.9.6). */
const MAX_TREE_NODES = 10_000;

/**
 * Read a PDF's document outline (ISO 32000-1 §12.3.3), depth first in the
 * order a viewer shows it. An item whose action is a URI action (§12.6.4.7)
 * keeps its URI and names no page. An item whose destination cannot be
 * followed to a page takes the place its first descendant jumps to (a chapter
 * heading that only groups its sections); one with neither is left out, its
 * children kept.
 *
 * @param file The parsed PDF.
 * @returns The outline's items; empty when the file has none.
 */
export function readPdfOutline(file: PdfFile): Array<SourceOutlineItem> {
  const root = file.get(file.catalog, 'Outlines');
  if (!(root instanceof Map)) return [];
  const pages = file.pages();
  const pageIndex = new Map<PdfDict, number>(pages.map((p, i) => [p.dict, i]));
  const found: Array<SourceOutlineItem> = [];
  const seen = new Set<PdfDict>();
  const walk = (first: PdfValue | undefined, level: number): void => {
    let node = first === undefined ? undefined : file.resolve(first);
    while (node instanceof Map && !seen.has(node) && found.length < MAX_ITEMS) {
      seen.add(node);
      const count = file.get(node, 'Count');
      const target = targetOf(file, node, pages, pageIndex);
      found.push({
        title: titleOf(file.get(node, 'Title')),
        level,
        open: typeof count === 'number' && count > 0,
        ...(target ?? {}),
      });
      walk(node.get('First'), level + 1);
      node = file.get(node, 'Next');
    }
  };
  walk(root.get('First'), 1);
  // An item that jumps nowhere jumps where the next item that does jumps.
  const out: Array<SourceOutlineItem> = [];
  for (const [i, item] of found.entries()) {
    if (item.title === '') continue;
    if (item.uri !== undefined) {
      out.push({ title: item.title, level: item.level, open: item.open, uri: item.uri });
      continue;
    }
    const target =
      item.pageIndex !== undefined
        ? item
        : found.slice(i + 1).find((later) => later.pageIndex !== undefined);
    if (target?.pageIndex === undefined) continue;
    out.push({
      title: item.title,
      level: item.level,
      open: item.open,
      pageIndex: target.pageIndex,
      ...(target.topPt !== undefined ? { topPt: target.topPt } : {}),
    });
  }
  return out;
}

// §7.9.2.2 — a title is a text string. What a panel shows of it is one line:
// a language escape (ESC, a language code, ESC) goes, every other control code
// reads as a space, and white space runs together.
function titleOf(value: PdfValue): string {
  const bytes = bytesOf(value);
  if (bytes === undefined) return '';
  let text = '';
  let inEscape = false;
  for (const ch of pdfTextString(bytes)) {
    const cp = ch.codePointAt(0)!;
    if (cp === 0x1b) {
      inEscape = !inEscape;
      continue;
    }
    if (inEscape) continue;
    text += cp < 0x20 || cp === 0x7f ? ' ' : ch;
  }
  return text.replace(/\s+/gu, ' ').trim();
}

// A string object's bytes, literal or hex (§7.3.4).
function bytesOf(value: PdfValue): Uint8Array | undefined {
  if (value instanceof PdfHexString) return value.bytes;
  if (typeof value === 'string') return Uint8Array.from(value, (c) => c.charCodeAt(0) & 0xff);
  return undefined;
}

// Where an item jumps (§12.3.3 Table 153): its `/Dest`, or the destination of
// a go-to action in its `/A` (§12.6.4.2), followed to an explicit array — or,
// for a URI action in its `/A` (§12.6.4.7), the URI it opens.
function targetOf(
  file: PdfFile,
  item: PdfDict,
  pages: ReadonlyArray<PdfPage>,
  pageIndex: ReadonlyMap<PdfDict, number>,
): { pageIndex: number; topPt?: number } | { uri: string } | undefined {
  let dest = file.get(item, 'Dest');
  if (!(dest instanceof Map || Array.isArray(dest) || isNamed(dest))) {
    const action = file.get(item, 'A');
    if (!(action instanceof Map)) return undefined;
    const kind = file.get(action, 'S');
    if (!(kind instanceof PdfName)) return undefined;
    if (kind.value === 'URI') {
      const uri = uriOf(file, file.get(action, 'URI'));
      return uri === undefined ? undefined : { uri };
    }
    if (kind.value !== 'GoTo') return undefined;
    dest = file.get(action, 'D');
  }
  const explicit = explicitDestination(file, dest, 0);
  if (!explicit || explicit.length < 2) return undefined;
  const target = file.resolve(explicit[0]!);
  // A page named by number is a remote go-to's form (§12.3.2.2); a local
  // destination that uses it anyway is taken at its word.
  const index =
    target instanceof Map
      ? pageIndex.get(target)
      : typeof target === 'number' && Number.isInteger(target) && target >= 0
        ? target
        : undefined;
  if (index === undefined || index >= pages.length) return undefined;
  const top = topOf(file, explicit);
  const page = pages[index]!;
  // The height as the page is SHOWN. Only an unturned page keeps its top at
  // the top; a turned one is left to the anchor's own line.
  if (top === undefined || page.rotate === 90 || page.rotate === 270) return { pageIndex: index };
  if (page.rotate === 180) return { pageIndex: index };
  return { pageIndex: index, topPt: displayOf(page).place(0, top).y };
}

// §12.6.4.7 Table 206 — the URI a URI action opens, an ASCII string. A writer
// that put UTF-8 there anyway meant the characters those bytes spell (pdf.js
// reads them so too); bytes that are no UTF-8 are taken one character each. A
// relative reference resolves against the catalog's `/URI` `/Base` (Table 207),
// as pdf.js resolves it; with no base it stays as written — "relative to the
// location of the document itself", a place a converted copy does not share.
function uriOf(file: PdfFile, value: PdfValue): string | undefined {
  const text = uriText(value);
  if (text === undefined) return undefined;
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(text)) return text;
  const uriDict = file.get(file.catalog, 'URI');
  const base = uriDict instanceof Map ? uriText(file.get(uriDict, 'Base')) : undefined;
  if (base === undefined) return text;
  try {
    return new URL(text, base).href;
  } catch {
    return text; // a base that is no absolute URI resolves nothing
  }
}

// A URI string's characters, trimmed; undefined for no string or an empty one.
function uriText(value: PdfValue): string | undefined {
  const bytes = bytesOf(value);
  if (bytes === undefined) return undefined;
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    text = Array.from(bytes, (b) => String.fromCharCode(b)).join('');
  }
  text = text.trim();
  return text === '' ? undefined : text;
}

// §12.3.2.3 — a destination named by a name object (looked up in the
// catalog's `/Dests`) or a byte string (looked up in the `/Names` `/Dests`
// name tree); the value is the array, or a dictionary whose `/D` is.
function explicitDestination(
  file: PdfFile,
  value: PdfValue,
  depth: number,
): ReadonlyArray<PdfValue> | undefined {
  if (depth > 4) return undefined;
  const v = file.resolve(value);
  if (Array.isArray(v)) return v;
  if (v instanceof Map) return explicitDestination(file, v.get('D') ?? [], depth + 1);
  if (v instanceof PdfName) {
    const dests = file.get(file.catalog, 'Dests');
    const named = dests instanceof Map ? dests.get(v.value) : undefined;
    return named === undefined ? undefined : explicitDestination(file, named, depth + 1);
  }
  const key = bytesOf(v);
  if (key === undefined) return undefined;
  const names = file.get(file.catalog, 'Names');
  const tree = names instanceof Map ? file.get(names, 'Dests') : undefined;
  const found = tree instanceof Map ? nameTreeValue(file, tree, key) : undefined;
  return found === undefined ? undefined : explicitDestination(file, found, depth + 1);
}

// §12.3.2.2 Table 151 — the top a destination puts at the window's top edge,
// where it names one.
function topOf(file: PdfFile, dest: ReadonlyArray<PdfValue>): number | undefined {
  const kind = file.resolve(dest[1]!);
  if (!(kind instanceof PdfName)) return undefined;
  const arg = (i: number): number | undefined => {
    const n = file.resolve(dest[i] ?? []);
    return typeof n === 'number' && Number.isFinite(n) ? n : undefined;
  };
  switch (kind.value) {
    case 'XYZ':
      return arg(3);
    case 'FitH':
    case 'FitBH':
      return arg(2);
    case 'FitR':
      return arg(5);
    default:
      return undefined;
  }
}

function isNamed(value: PdfValue): boolean {
  return value instanceof PdfName || value instanceof PdfHexString || typeof value === 'string';
}

// §7.9.6 — the value a name tree keeps under `key`, the keys compared byte for
// byte. Every node is searched (bounded), not just those whose `/Limits` admit
// the key: a writer that got its limits wrong still wrote the entry.
function nameTreeValue(file: PdfFile, root: PdfDict, key: Uint8Array): PdfValue | undefined {
  const stack: Array<PdfValue> = [root];
  const seen = new Set<PdfValue>();
  let visits = 0;
  while (stack.length > 0 && visits++ < MAX_TREE_NODES) {
    const node = file.resolve(stack.pop()!);
    if (!(node instanceof Map) || seen.has(node)) continue;
    seen.add(node);
    const names = file.get(node, 'Names');
    if (Array.isArray(names)) {
      for (let i = 0; i + 1 < names.length; i += 2) {
        const k = bytesOf(file.resolve(names[i]!));
        if (k && k.length === key.length && k.every((b, j) => b === key[j])) return names[i + 1];
      }
    }
    const kids = file.get(node, 'Kids');
    if (Array.isArray(kids)) stack.push(...kids);
  }
  return undefined;
}

/**
 * Hang a source PDF's outline on the document read out of it, as
 * `Paragraph.outlineEntries`, so that the PDF written from it carries the same
 * bookmarks (ISO 32000-1 §12.3.3): the same titles, nesting and open state, in
 * the source's order, each jumping to the page it jumped to.
 *
 * Each item rides a paragraph of its page: the one whose words its title
 * begins (a bookmark is nearly always its heading's text), else the page's
 * first paragraph — the carrier that opens it, as a rule — else one of no
 * height put there for it. Where the page stands as the source drew it
 * (`placed`), or the item found no paragraph of its own words, it keeps the
 * height its destination named; otherwise it jumps to its paragraph, wherever
 * the re-set page puts that. An item that opens a URI jumps to no page: it
 * rides the first paragraph of the page the item before it jumps to (else the
 * one after it), only to be laid with the rest — its `order` alone says where
 * it stands, under which item and over which.
 *
 * @param doc        The document read out of the PDF.
 * @param items      The PDF's outline (see {@link readPdfOutline}).
 * @param pageStarts Where each source page's elements begin in `doc.body`.
 * @param placed     Whether the pages were kept as they stand rather than re-set.
 * @returns The document with the outline hung on it; `doc` itself when there is none.
 */
export function withSourceOutline(
  doc: FlowDoc,
  items: ReadonlyArray<SourceOutlineItem>,
  pageStarts: ReadonlyArray<number>,
  placed: boolean,
): FlowDoc {
  if (items.length === 0 || pageStarts.length === 0) return doc;
  const body = [...doc.body];
  const onParagraph = new Map<number, Array<OutlineEntry>>();
  const onCarrier = new Map<number, Array<OutlineEntry>>();
  for (const [order, item] of items.entries()) {
    const page = Math.min(item.pageIndex ?? neighbourPage(items, order), pageStarts.length - 1);
    const from = pageStarts[page]!;
    const to = pageStarts[page + 1] ?? body.length;
    const paragraphs: Array<number> = [];
    for (let i = from; i < to; i++) if (body[i]?.kind === 'paragraph') paragraphs.push(i);
    const wanted = folded(item.title);
    // An item that opens a URI says nothing of the words on its page.
    const own = (item.uri !== undefined ? [] : paragraphs).find((i) => {
      const el = body[i]!;
      if (el.kind !== 'paragraph') return false;
      const said = folded(outlineTitleOf(el.paragraph.runs));
      return (
        said !== '' &&
        (said.startsWith(wanted) ||
          (wanted.startsWith(said) && said.length * 5 >= wanted.length * 3))
      );
    });
    const keepsHeight = placed || own === undefined;
    const entry: OutlineEntry = {
      title: item.title,
      level: Math.min(item.level, 9),
      ...(item.open ? {} : { open: false }),
      ...(keepsHeight && item.topPt !== undefined ? { topPt: item.topPt } : {}),
      order,
      ...(item.uri !== undefined ? { uri: item.uri } : {}),
    };
    const at = own ?? paragraphs[0];
    const into = at !== undefined ? onParagraph : onCarrier;
    const key = at ?? page;
    into.set(key, [...(into.get(key) ?? []), entry]);
  }
  for (const [at, entries] of onParagraph) {
    const el = body[at]!;
    if (el.kind !== 'paragraph') continue;
    body[at] = {
      ...el,
      paragraph: {
        ...el.paragraph,
        outlineEntries: [...(el.paragraph.outlineEntries ?? []), ...entries],
      },
    };
  }
  // A page with no paragraph at all gets a line of no height to carry its
  // items, put where its elements begin — from the last page back, so that
  // each lands where it was found; the sections after it end one later.
  let sections = doc.sections;
  for (const page of [...onCarrier.keys()].sort((a, b) => b - a)) {
    const at = pageStarts[page]!;
    body.splice(at, 0, {
      kind: 'paragraph',
      paragraph: {
        properties: { spacingLine: CARRIER_LINE_PT, spacingLineRule: 'exact' },
        runs: [],
        outlineEntries: onCarrier.get(page)!,
      },
    });
    sections = sections.map((s) => (s.endIndex > at ? { ...s, endIndex: s.endIndex + 1 } : s));
  }
  return { ...doc, body, sections };
}

// The page of the nearest item before `at` that names one, else after it.
function neighbourPage(items: ReadonlyArray<SourceOutlineItem>, at: number): number {
  for (let i = at - 1; i >= 0; i--) {
    const page = items[i]!.pageIndex;
    if (page !== undefined) return page;
  }
  for (let i = at + 1; i < items.length; i++) {
    const page = items[i]!.pageIndex;
    if (page !== undefined) return page;
  }
  return 0;
}

// What two titles are compared by: case and spacing folded away.
function folded(text: string): string {
  return text.replace(/\s+/gu, ' ').trim().toLowerCase();
}
