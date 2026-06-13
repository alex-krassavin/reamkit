// E-PDF EP1 — the document layer. Takes a whole PDF byte buffer and exposes its
// objects: the cross-reference table (classic `xref` + `trailer`, following the
// /Prev chain of incremental updates), indirect-reference resolution with a
// cache, the page tree (walked with attribute inheritance), and stream decoding
// (FlateDecode via fflate). A brute-force object scan recovers files whose xref
// is broken or stored as an xref stream (which this layer does not yet parse).

import { unzlibSync } from 'fflate';

import { Lexer } from './lexer';
import { parseIndirectObject, parseObject } from './parser';
import type { PdfArray, PdfDict, PdfValue } from '@/pdf/objects';
import { PDF_NULL, PdfName, PdfRef, PdfStream } from '@/pdf/objects';

export type Rectangle = readonly [number, number, number, number];

export interface PdfPage {
  readonly dict: PdfDict;
  readonly mediaBox: Rectangle;
  readonly resources: PdfDict | undefined;
}

const DEFAULT_MEDIA_BOX: Rectangle = [0, 0, 612, 792]; // US Letter, the PDF default
const MAX_PAGES = 50_000; // DoS guard on a pathological page tree

export class PdfFile {
  private readonly cache = new Map<number, PdfValue>();

  private constructor(
    private readonly buf: Uint8Array,
    private readonly xref: ReadonlyMap<number, number>, // object number → byte offset
    readonly trailer: PdfDict,
  ) {}

  static parse(bytes: Uint8Array): PdfFile {
    let xref = new Map<number, number>();
    let trailer: PdfDict = new Map();
    try {
      const start = findStartXref(bytes);
      if (start >= 0) {
        const built = readXrefChain(bytes, start);
        xref = built.xref;
        trailer = built.trailer;
      }
    } catch {
      // fall through to the brute-force recovery below
    }
    // Recover when the xref was missing/broken or the trailer has no /Root
    // (e.g. an xref-stream file): scan the bytes for every `N G obj` definition.
    if (xref.size === 0 || !(trailer.get('Root') instanceof PdfRef)) {
      const scanned = bruteForceScan(bytes);
      // Prefer scanned offsets only where the xref lacked an entry.
      for (const [id, off] of scanned.xref) if (!xref.has(id)) xref.set(id, off);
      if (!(trailer.get('Root') instanceof PdfRef) && scanned.root) {
        trailer = new Map(trailer);
        trailer.set('Root', scanned.root);
      }
    }
    return new PdfFile(bytes, xref, trailer);
  }

  // Dereference one level: a PdfRef becomes the object it points at (parsed and
  // cached); any other value is returned unchanged.
  resolve(value: PdfValue): PdfValue {
    if (!(value instanceof PdfRef)) return value;
    const cached = this.cache.get(value.id);
    if (cached !== undefined) return cached;
    const offset = this.xref.get(value.id);
    if (offset === undefined || offset < 0 || offset >= this.buf.length) return PDF_NULL;
    this.cache.set(value.id, PDF_NULL); // guard against a self-referential cycle
    const lexer = new Lexer(this.buf, offset);
    const obj = parseIndirectObject(lexer, (r) => {
      const n = this.resolve(r);
      return typeof n === 'number' ? n : undefined;
    });
    const result = obj ? obj.value : PDF_NULL;
    this.cache.set(value.id, result);
    return result;
  }

  // Resolve dict[key] in one step.
  get(dict: PdfDict, key: string): PdfValue {
    return this.resolve(dict.get(key) ?? PDF_NULL);
  }

  get catalog(): PdfDict {
    const root = this.resolve(this.trailer.get('Root') ?? PDF_NULL);
    return root instanceof Map ? root : new Map();
  }

  // The leaf pages, in document order, each with its inherited MediaBox/Resources.
  pages(): Array<PdfPage> {
    const out: Array<PdfPage> = [];
    const root = this.get(this.catalog, 'Pages');
    if (root instanceof Map) this.walkPageTree(root, {}, out, new Set());
    return out;
  }

  private walkPageTree(
    node: PdfDict,
    inherited: { mediaBox?: Rectangle | undefined; resources?: PdfDict | undefined },
    out: Array<PdfPage>,
    seen: Set<PdfDict>,
  ): void {
    if (out.length >= MAX_PAGES || seen.has(node)) return;
    seen.add(node);
    const mediaBox = readRectangle(this.get(node, 'MediaBox')) ?? inherited.mediaBox;
    const resourcesVal = this.get(node, 'Resources');
    const resources = resourcesVal instanceof Map ? resourcesVal : inherited.resources;
    const type = node.get('Type');
    const kids = this.get(node, 'Kids');
    if (type instanceof PdfName && type.value === 'Pages' && Array.isArray(kids)) {
      for (const kid of kids) {
        const kidNode = this.resolve(kid);
        if (kidNode instanceof Map) this.walkPageTree(kidNode, { mediaBox, resources }, out, seen);
        if (out.length >= MAX_PAGES) break;
      }
      return;
    }
    // A leaf /Page (or an untyped node with no kids).
    out.push({ dict: node, mediaBox: mediaBox ?? DEFAULT_MEDIA_BOX, resources });
  }

  // A page's concatenated, decoded content stream bytes (/Contents may be a
  // single stream or an array of streams joined with a space, per §7.8.2).
  pageContent(page: PdfPage): Uint8Array {
    const contents = this.get(page.dict, 'Contents');
    const streams: Array<PdfStream> = [];
    if (contents instanceof PdfStream) streams.push(contents);
    else if (Array.isArray(contents)) {
      for (const c of contents) {
        const s = this.resolve(c);
        if (s instanceof PdfStream) streams.push(s);
      }
    }
    const parts = streams.map((s) => this.streamData(s));
    return concatWithSpaces(parts);
  }

  // Decode a stream's bytes, applying its /Filter chain (FlateDecode supported;
  // unknown filters pass through undecoded).
  streamData(stream: PdfStream): Uint8Array {
    let data = stream.data;
    const filter = this.resolve(stream.dict.get('Filter') ?? PDF_NULL);
    const filters: Array<PdfValue> = Array.isArray(filter) ? filter : [filter];
    for (const f of filters) {
      if (f instanceof PdfName && (f.value === 'FlateDecode' || f.value === 'Fl')) {
        try {
          data = unzlibSync(data);
        } catch {
          // leave undecoded on a malformed stream
        }
      }
    }
    return data;
  }
}

// --- cross-reference table -------------------------------------------------

function findStartXref(buf: Uint8Array): number {
  const tail = lastIndexOfAscii(buf, 'startxref');
  if (tail < 0) return -1;
  const lexer = new Lexer(buf, tail + 'startxref'.length);
  const tok = lexer.nextToken();
  return tok.kind === 'num' ? tok.value : -1;
}

// Read the classic `xref` table at `offset` and its `trailer`, then follow the
// /Prev chain to older sections (newer entries win).
function readXrefChain(
  buf: Uint8Array,
  offset: number,
): { xref: Map<number, number>; trailer: PdfDict } {
  const xref = new Map<number, number>();
  let trailer: PdfDict = new Map();
  const visited = new Set<number>();
  let at: number | undefined = offset;
  while (at !== undefined && at >= 0 && at < buf.length && !visited.has(at)) {
    visited.add(at);
    const section = readXrefSection(buf, at);
    if (!section) break;
    for (const [id, off] of section.xref) if (!xref.has(id)) xref.set(id, off);
    if (trailer.size === 0) trailer = section.trailer;
    const prev = section.trailer.get('Prev');
    at = typeof prev === 'number' ? prev : undefined;
  }
  return { xref, trailer };
}

function readXrefSection(
  buf: Uint8Array,
  offset: number,
): { xref: Map<number, number>; trailer: PdfDict } | undefined {
  const lexer = new Lexer(buf, offset);
  const head = lexer.nextToken();
  if (!(head.kind === 'keyword' && head.value === 'xref')) return undefined; // not a classic table
  const xref = new Map<number, number>();
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'keyword' && tok.value === 'trailer') break;
    if (tok.kind !== 'num') return undefined;
    const first = tok.value;
    const countTok = lexer.nextToken();
    if (countTok.kind !== 'num') return undefined;
    const count = countTok.value;
    for (let i = 0; i < count; i++) {
      const off = lexer.nextToken();
      const gen = lexer.nextToken();
      const type = lexer.nextToken();
      if (off.kind !== 'num' || gen.kind !== 'num' || type.kind !== 'keyword') return undefined;
      if (type.value === 'n' && !xref.has(first + i)) xref.set(first + i, off.value);
    }
  }
  const trailerVal = parseObject(lexer);
  return { xref, trailer: trailerVal instanceof Map ? trailerVal : new Map() };
}

// Linear recovery: scan for every `N G obj` definition, recording the latest
// offset per object number, and the first object whose /Type is /Catalog.
function bruteForceScan(buf: Uint8Array): { xref: Map<number, number>; root: PdfRef | undefined } {
  const xref = new Map<number, number>();
  let root: PdfRef | undefined;
  const lexer = new Lexer(buf);
  let prev2: { start: number; value: number } | undefined;
  let prev1: { start: number; value: number } | undefined;
  for (;;) {
    lexer.skipWhitespace();
    const start = lexer.pos;
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') break;
    if (tok.kind === 'keyword' && tok.value === 'obj' && prev2 && prev1) {
      xref.set(prev2.value, prev2.start);
      if (root === undefined) {
        const obj = parseIndirectObject(new Lexer(buf, prev2.start));
        const dict = obj?.value;
        if (dict instanceof Map) {
          const type = dict.get('Type');
          if (type instanceof PdfName && type.value === 'Catalog') {
            root = new PdfRef(prev2.value, prev1.value);
          }
        }
      }
    }
    prev2 = prev1;
    prev1 = tok.kind === 'num' ? { start, value: tok.value } : undefined;
  }
  return { xref, root };
}

// --- small helpers ----------------------------------------------------------

function readRectangle(value: PdfValue): Rectangle | undefined {
  if (!Array.isArray(value) || value.length < 4) return undefined;
  const nums = value.slice(0, 4).map((v) => (typeof v === 'number' ? v : NaN));
  if (nums.some((n) => !Number.isFinite(n))) return undefined;
  return [nums[0]!, nums[1]!, nums[2]!, nums[3]!];
}

function concatWithSpaces(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  if (parts.length === 1) return parts[0]!;
  const total = parts.reduce((n, p) => n + p.length + 1, 0);
  const out = new Uint8Array(Math.max(0, total - 1));
  let at = 0;
  parts.forEach((p, i) => {
    out.set(p, at);
    at += p.length;
    if (i < parts.length - 1) out[at++] = 0x0a;
  });
  return out;
}

function lastIndexOfAscii(buf: Uint8Array, needle: string): number {
  const n = needle.length;
  outer: for (let i = buf.length - n; i >= 0; i--) {
    for (let j = 0; j < n; j++) {
      if (buf[i + j] !== needle.charCodeAt(j)) continue outer;
    }
    return i;
  }
  return -1;
}

// Re-export for callers that walk a resolved dict's array values.
export type { PdfArray };
