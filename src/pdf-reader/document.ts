// E-PDF EP1/EP7 — the document layer. Takes a whole PDF byte buffer and exposes
// its objects: the cross-reference table — classic `xref` + `trailer` AND
// cross-reference STREAMS (/Type /XRef, PDF 1.5+), following the /Prev chain and
// hybrid /XRefStm pointers — indirect-reference resolution with a cache
// (including objects packed inside OBJECT STREAMS, /Type /ObjStm), the page tree
// (walked with attribute inheritance), and stream decoding (FlateDecode +
// predictors via fflate). A brute-force object scan recovers files whose xref is
// broken, and itself indexes any object streams it finds.

import { unzlibSync } from 'fflate';

import { buildDecryptor } from './decrypt';
import { Lexer } from './lexer';
import { parseIndirectObject, parseObject } from './parser';
import { reversePredictor } from './predictor';
import type { Decryptor } from './decrypt';
import type { PdfArray, PdfDict, PdfValue } from '@/pdf/objects';
import { PDF_NULL, PdfName, PdfRef, PdfStream } from '@/pdf/objects';

export type Rectangle = readonly [number, number, number, number];

export interface PdfPage {
  readonly dict: PdfDict;
  readonly mediaBox: Rectangle;
  readonly resources: PdfDict | undefined;
}

// A cross-reference entry: a byte offset for a normal object, or a (stream, index)
// pair for an object packed inside an object stream (§7.5.7).
type XrefEntry =
  | { readonly kind: 'uncompressed'; readonly offset: number }
  | { readonly kind: 'compressed'; readonly streamObj: number; readonly index: number };

const DEFAULT_MEDIA_BOX: Rectangle = [0, 0, 612, 792]; // US Letter, the PDF default
const MAX_PAGES = 50_000; // DoS guard on a pathological page tree
const MAX_OBJSTM_N = 200_000; // DoS guard on an object stream's object count

export class PdfFile {
  private readonly cache = new Map<number, PdfValue>();
  // objStm object number → its decoded members (objNum → value), decoded once.
  private readonly objStmCache = new Map<number, Map<number, PdfValue>>();
  // The standard security handler, when the document is encrypted (EP9).
  private decryptor: Decryptor | undefined;
  private encryptObjNum = -1;

  private constructor(
    private readonly buf: Uint8Array,
    private readonly xref: ReadonlyMap<number, XrefEntry>,
    readonly trailer: PdfDict,
  ) {}

  static parse(bytes: Uint8Array): PdfFile {
    let xref = new Map<number, XrefEntry>();
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
    // Recover when the xref was missing/broken or the trailer has no /Root: scan
    // for every `N G obj` and index any object streams found.
    if (xref.size === 0 || !(trailer.get('Root') instanceof PdfRef)) {
      const scanned = bruteForceScan(bytes);
      for (const [id, entry] of scanned.xref) if (!xref.has(id)) xref.set(id, entry);
      if (!(trailer.get('Root') instanceof PdfRef) && scanned.root) {
        trailer = new Map(trailer);
        trailer.set('Root', scanned.root);
      }
    }
    const file = new PdfFile(bytes, xref, trailer);
    file.initEncryption();
    return file;
  }

  // Build the decryptor from /Encrypt (§7.6). Runs before any other object is
  // resolved, so the /Encrypt dictionary itself is read in the clear; its object
  // is then never decrypted.
  private initEncryption(): void {
    const encVal = this.trailer.get('Encrypt');
    if (encVal === undefined) return;
    if (encVal instanceof PdfRef) this.encryptObjNum = encVal.id;
    const enc = this.resolve(encVal);
    if (!(enc instanceof Map)) return;
    const id = this.trailer.get('ID');
    this.decryptor = buildDecryptor(enc, Array.isArray(id) ? id : undefined);
  }

  // Resolve a stream's /Length when it is an indirect reference.
  private readonly lengthResolver = (r: PdfRef): number | undefined => {
    const n = this.resolve(r);
    return typeof n === 'number' ? n : undefined;
  };

  // Dereference one level: a PdfRef becomes the object it points at (parsed and
  // cached); any other value is returned unchanged.
  resolve(value: PdfValue): PdfValue {
    if (!(value instanceof PdfRef)) return value;
    const cached = this.cache.get(value.id);
    if (cached !== undefined) return cached;
    const entry = this.xref.get(value.id);
    if (entry === undefined) return PDF_NULL;
    this.cache.set(value.id, PDF_NULL); // guard against a self-referential cycle
    let result: PdfValue = PDF_NULL;
    if (entry.kind === 'uncompressed') {
      if (entry.offset >= 0 && entry.offset < this.buf.length) {
        const obj = parseIndirectObject(new Lexer(this.buf, entry.offset), this.lengthResolver);
        result = obj ? obj.value : PDF_NULL;
        // Decrypt every string/stream (§7.6) — except the /Encrypt dict itself,
        // whose strings are the keys. Compressed objects are already plaintext
        // (their object stream was decrypted as a whole).
        if (this.decryptor && value.id !== this.encryptObjNum) {
          result = this.decryptor.decrypt(result, value.id, obj?.generation ?? 0);
        }
      }
    } else {
      result = this.objectFromStream(entry.streamObj).get(value.id) ?? PDF_NULL;
    }
    this.cache.set(value.id, result);
    return result;
  }

  // Decode (once) the members of an object stream and return objNum → value.
  private objectFromStream(streamObj: number): Map<number, PdfValue> {
    const cached = this.objStmCache.get(streamObj);
    if (cached) return cached;
    const out = new Map<number, PdfValue>();
    this.objStmCache.set(streamObj, out); // cycle guard
    const entry = this.xref.get(streamObj);
    if (!entry || entry.kind !== 'uncompressed') return out;
    const obj = parseIndirectObject(new Lexer(this.buf, entry.offset), this.lengthResolver);
    if (!obj || !(obj.value instanceof PdfStream)) return out;
    let stream = obj.value;
    if (this.decryptor) {
      const dec = this.decryptor.decrypt(stream, streamObj, obj.generation);
      if (dec instanceof PdfStream) stream = dec;
    }
    const data = inflateStream(stream);
    const first = numOf(this.resolve(stream.dict.get('First') ?? PDF_NULL));
    for (const member of objStmHeader(
      data,
      numOf(this.resolve(stream.dict.get('N') ?? PDF_NULL)),
    )) {
      out.set(member.id, parseObject(new Lexer(data, first + member.off), this.lengthResolver));
    }
    return out;
  }

  // Resolve dict[key] in one step.
  get(dict: PdfDict, key: string): PdfValue {
    return this.resolve(dict.get(key) ?? PDF_NULL);
  }

  get catalog(): PdfDict {
    const root = this.resolve(this.trailer.get('Root') ?? PDF_NULL);
    return root instanceof Map ? root : new Map();
  }

  // The document is encrypted but no decryptor could be built (an unsupported
  // handler, or a non-empty user password) — its content is unreadable (EP9).
  get encryptionUnsupported(): boolean {
    return this.trailer.get('Encrypt') !== undefined && this.decryptor === undefined;
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

  // Decode a stream's bytes, applying its /Filter chain (FlateDecode + any
  // /Predictor supported; unknown filters pass through undecoded).
  streamData(stream: PdfStream): Uint8Array {
    let data = stream.data;
    const filter = this.resolve(stream.dict.get('Filter') ?? PDF_NULL);
    const filters: Array<PdfValue> = Array.isArray(filter) ? filter : [filter];
    let flate = false;
    for (const f of filters) {
      if (f instanceof PdfName && (f.value === 'FlateDecode' || f.value === 'Fl')) {
        try {
          data = unzlibSync(data);
          flate = true;
        } catch {
          // leave undecoded on a malformed stream
        }
      }
    }
    return flate ? applyStreamPredictor(this, stream.dict, data) : data;
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

interface XrefSection {
  readonly xref: Map<number, XrefEntry>;
  readonly trailer: PdfDict;
}

// Read the section at `offset` (classic table or xref stream), then follow the
// /Prev chain and any hybrid /XRefStm pointers (newest entries win).
function readXrefChain(buf: Uint8Array, offset: number): XrefSection {
  const xref = new Map<number, XrefEntry>();
  let trailer: PdfDict = new Map();
  const visited = new Set<number>();
  const queue: Array<number> = [offset];
  while (queue.length > 0) {
    const at = queue.shift()!;
    if (at < 0 || at >= buf.length || visited.has(at)) continue;
    visited.add(at);
    const section = readXrefAt(buf, at);
    if (!section) continue;
    for (const [id, entry] of section.xref) if (!xref.has(id)) xref.set(id, entry);
    if (trailer.size === 0) trailer = section.trailer;
    const xrefStm = section.trailer.get('XRefStm'); // hybrid-reference file
    if (typeof xrefStm === 'number') queue.push(xrefStm);
    const prev = section.trailer.get('Prev');
    if (typeof prev === 'number') queue.push(prev);
  }
  return { xref, trailer };
}

// Dispatch by what sits at the offset: the `xref` keyword (a classic table) or an
// `N G obj` definition (a cross-reference stream).
function readXrefAt(buf: Uint8Array, offset: number): XrefSection | undefined {
  const lexer = new Lexer(buf, offset);
  const head = lexer.nextToken();
  if (head.kind === 'keyword' && head.value === 'xref') return readClassicXref(lexer);
  const obj = parseIndirectObject(new Lexer(buf, offset));
  if (obj && obj.value instanceof PdfStream) return readXrefStream(obj.value);
  return undefined;
}

// A classic `xref` subsection list followed by `trailer << … >>`.
function readClassicXref(lexer: Lexer): XrefSection | undefined {
  const xref = new Map<number, XrefEntry>();
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
      if (type.value === 'n' && !xref.has(first + i)) {
        xref.set(first + i, { kind: 'uncompressed', offset: off.value });
      }
    }
  }
  const trailerVal = parseObject(lexer);
  return { xref, trailer: trailerVal instanceof Map ? trailerVal : new Map() };
}

// A cross-reference stream (§7.5.8): /W field widths, /Index subsections, and
// fixed-width binary rows of (type, field2, field3).
function readXrefStream(stream: PdfStream): XrefSection | undefined {
  const dict = stream.dict;
  const wv = dict.get('W');
  if (!Array.isArray(wv) || wv.length < 3) return undefined;
  const w0 = numOf(wv[0]);
  const w1 = numOf(wv[1]);
  const w2 = numOf(wv[2]);
  const rowLen = w0 + w1 + w2;
  if (rowLen <= 0) return undefined;
  const data = inflateStream(stream);
  const size = numOf(dict.get('Size'));
  const indexV = dict.get('Index');
  const index = Array.isArray(indexV) ? indexV.map(numOf) : [0, size];

  const xref = new Map<number, XrefEntry>();
  let pos = 0;
  for (let s = 0; s + 1 < index.length; s += 2) {
    const start = index[s]!;
    const count = index[s + 1]!;
    for (let i = 0; i < count && pos + rowLen <= data.length; i++) {
      const type = w0 === 0 ? 1 : readBE(data, pos, w0); // default type 1 (§7.5.8.2)
      const f2 = readBE(data, pos + w0, w1);
      const f3 = readBE(data, pos + w0 + w1, w2);
      pos += rowLen;
      const id = start + i;
      if (xref.has(id)) continue;
      if (type === 1) xref.set(id, { kind: 'uncompressed', offset: f2 });
      else if (type === 2) xref.set(id, { kind: 'compressed', streamObj: f2, index: f3 });
    }
  }
  return { xref, trailer: dict };
}

// --- object streams ---------------------------------------------------------

// The N leading `objNum offset` pairs of an object stream (§7.5.7).
function objStmHeader(data: Uint8Array, n: number): Array<{ id: number; off: number }> {
  const out: Array<{ id: number; off: number }> = [];
  const lexer = new Lexer(data, 0);
  for (let i = 0; i < Math.min(n, MAX_OBJSTM_N); i++) {
    const idTok = lexer.nextToken();
    const offTok = lexer.nextToken();
    if (idTok.kind !== 'num' || offTok.kind !== 'num') break;
    out.push({ id: idTok.value, off: offTok.value });
  }
  return out;
}

// FlateDecode + /Predictor, reading the parameters directly from the (directly
// stored) dict — used while the cross-reference is still being built, so it
// cannot rely on indirect-reference resolution.
function inflateStream(stream: PdfStream): Uint8Array {
  let data = stream.data;
  const filter = stream.dict.get('Filter') ?? PDF_NULL;
  const filters: Array<PdfValue> = Array.isArray(filter) ? filter : [filter];
  for (const f of filters) {
    if (f instanceof PdfName && (f.value === 'FlateDecode' || f.value === 'Fl')) {
      try {
        data = unzlibSync(data);
      } catch {
        return new Uint8Array(0);
      }
    }
  }
  const parmsVal = stream.dict.get('DecodeParms') ?? stream.dict.get('DP');
  const parms =
    parmsVal instanceof Map
      ? parmsVal
      : Array.isArray(parmsVal)
        ? parmsVal.find((p): p is PdfDict => p instanceof Map)
        : undefined;
  if (parms) {
    const predictor = numOf(parms.get('Predictor'));
    if (predictor >= 2) {
      data = reversePredictor(data, {
        predictor,
        colors: numOf(parms.get('Colors')) || 1,
        bitsPerComponent: numOf(parms.get('BitsPerComponent')) || 8,
        columns: numOf(parms.get('Columns')) || 1,
      });
    }
  }
  return data;
}

// streamData's predictor step (parameters may be indirect here, so resolve them).
function applyStreamPredictor(file: PdfFile, dict: PdfDict, data: Uint8Array): Uint8Array {
  const parmsVal = file.get(dict, 'DecodeParms');
  const parms = parmsVal instanceof Map ? parmsVal : undefined;
  if (!parms) return data;
  const predictor = numOf(file.get(parms, 'Predictor'));
  if (predictor < 2) return data;
  return reversePredictor(data, {
    predictor,
    colors: numOf(file.get(parms, 'Colors')) || 1,
    bitsPerComponent: numOf(file.get(parms, 'BitsPerComponent')) || 8,
    columns: numOf(file.get(parms, 'Columns')) || 1,
  });
}

// --- brute-force recovery ---------------------------------------------------

// Linear recovery: scan for every `N G obj`, recording the latest offset per
// object number and the first /Catalog; then index any object streams found, so
// the objects they pack are reachable too.
function bruteForceScan(buf: Uint8Array): {
  xref: Map<number, XrefEntry>;
  root: PdfRef | undefined;
} {
  const xref = new Map<number, XrefEntry>();
  const objStmObjs: Array<number> = [];
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
      xref.set(prev2.value, { kind: 'uncompressed', offset: prev2.start });
      const obj = parseIndirectObject(new Lexer(buf, prev2.start));
      const value = obj?.value;
      const dict = value instanceof PdfStream ? value.dict : value;
      if (dict instanceof Map) {
        const type = dict.get('Type');
        if (type instanceof PdfName && type.value === 'Catalog' && root === undefined) {
          root = new PdfRef(prev2.value, prev1.value);
        } else if (type instanceof PdfName && type.value === 'ObjStm') {
          objStmObjs.push(prev2.value);
        }
      }
    }
    prev2 = prev1;
    prev1 = tok.kind === 'num' ? { start, value: tok.value } : undefined;
  }
  // Index each object stream's members as compressed entries (offsets win), and
  // recover the /Catalog from inside one if no top-level catalog was found.
  for (const streamObj of objStmObjs) {
    const entry = xref.get(streamObj);
    if (!entry || entry.kind !== 'uncompressed') continue;
    const obj = parseIndirectObject(new Lexer(buf, entry.offset));
    if (!obj || !(obj.value instanceof PdfStream)) continue;
    const data = inflateStream(obj.value);
    const first = numOf(obj.value.dict.get('First') ?? PDF_NULL);
    const n = numOf(obj.value.dict.get('N') ?? PDF_NULL);
    objStmHeader(data, n).forEach((member, index) => {
      if (!xref.has(member.id)) xref.set(member.id, { kind: 'compressed', streamObj, index });
      if (root === undefined) {
        const v = parseObject(new Lexer(data, first + member.off));
        if (v instanceof Map) {
          const t = v.get('Type');
          if (t instanceof PdfName && t.value === 'Catalog') root = new PdfRef(member.id, 0);
        }
      }
    });
  }
  return { xref, root };
}

// --- small helpers ----------------------------------------------------------

function readBE(data: Uint8Array, offset: number, width: number): number {
  let v = 0;
  for (let i = 0; i < width; i++) v = v * 256 + (data[offset + i] ?? 0);
  return v;
}

function numOf(v: PdfValue | undefined): number {
  return typeof v === 'number' ? v : 0;
}

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
