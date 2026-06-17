// Legacy `.ppt` (PowerPoint 97–2003, [MS-PPT]) record reader + per-slide text
// extraction (PPT-1). A `.ppt` is an OLE2/CFB container whose `PowerPoint
// Document` stream is a tree of records sharing the OfficeArt header layout
// (`[recVerInstance:u16][recType:u16][recLen:u32]`, a container when the low
// nibble of recVerInstance is 0xF — the same convention `escher.ts` walks for
// `.xls`). Slides are reached indirectly: the `Current User` stream points at the
// most recent UserEditAtom, whose chain of edits names a PersistDirectoryAtom that
// maps persist-object ids to absolute byte offsets; the DocumentContainer's slide
// list (a SlideListWithText) gives the slide order via per-slide SlidePersistAtoms.
//
// PPT-1 pulls each slide's text: the TextCharsAtom (UTF-16) / TextBytesAtom
// (cp1252) atoms inside the slide's drawing, falling back to the slide list's
// outline text when a slide stores none inline. Everything is best-effort and
// graceful-on-failure — structural doubt yields missing text, never wrong text —
// mirroring the `.doc`/`.xls` readers built on the same CFB keystone.

import { isCfb, openCfb } from '@/core/ole/cfb';

// --- [MS-PPT] §2.13.24 record types (the ones PPT-1 needs) -------------------
const RT_DOCUMENT = 0x03e8;
const RT_DOCUMENT_ATOM = 0x03e9;
const RT_SLIDE_LIST_WITH_TEXT = 0x0ff0;
const RT_SLIDE_PERSIST_ATOM = 0x03f3;
const RT_SLIDE = 0x03ee;
const RT_USER_EDIT_ATOM = 0x0ff5;
const RT_PERSIST_DIRECTORY_ATOM = 0x1772;
const RT_TEXT_CHARS_ATOM = 0x0fa0;
const RT_TEXT_BYTES_ATOM = 0x0fa8;

// CurrentUserAtom.headerToken (§2.3.2) — the encrypted variant means the document
// streams are obfuscated and the text cannot be read.
const TOKEN_ENCRYPTED = 0xf3d1c4df;

// A SlideListWithText's recInstance distinguishes the three lists it can be
// (§2.4.14.3): 0 = slides, 1 = masters, 2 = notes. PPT-1 reads the slide list.
const SLWT_INSTANCE_SLIDES = 0;

// Guard against a crafted edit chain or persist directory looping forever.
const MAX_EDITS = 4096;
const MAX_DEPTH = 24;

// One slide's extracted text, as paragraphs (split at the CR paragraph mark).
export interface PptParagraph {
  readonly text: string;
}
export interface PptSlide {
  readonly paragraphs: ReadonlyArray<PptParagraph>;
}
export interface PptContent {
  readonly slides: ReadonlyArray<PptSlide>;
  // The deck page size in points, when the DocumentAtom gives a sane slide size.
  readonly slideWidthPt?: number;
  readonly slideHeightPt?: number;
  readonly encrypted: boolean;
}

// A record header (§2.3.1) read at an absolute offset in the stream — `data` is
// the bytes after the 8-byte header, bounded to the stream.
interface PptRecord {
  readonly type: number;
  readonly instance: number;
  readonly isContainer: boolean;
  readonly data: Uint8Array;
}

function recordAt(stream: Uint8Array, off: number): PptRecord | undefined {
  if (off < 0 || off + 8 > stream.length) return undefined;
  const verInstance = u16(stream, off);
  const type = u16(stream, off + 2);
  const recLen = u32(stream, off + 4);
  const start = off + 8;
  const end = Math.min(stream.length, start + recLen);
  return {
    type,
    instance: verInstance >> 4,
    isContainer: (verInstance & 0x0f) === 0x0f,
    data: stream.subarray(start, end),
  };
}

// Iterate the records laid out consecutively in one buffer (a container's data,
// or the stream from a given base).
function* records(d: Uint8Array): Generator<PptRecord> {
  let off = 0;
  while (off + 8 <= d.length) {
    const rec = recordAt(d, off);
    if (!rec) break;
    yield rec;
    const consumed = 8 + rec.data.length;
    if (consumed <= 0) break; // zero-length guard
    off += consumed;
  }
}

export function extractPptContent(bytes: Uint8Array): PptContent {
  const empty: PptContent = { slides: [], encrypted: false };
  if (!isCfb(bytes)) return empty;
  let cfb;
  try {
    cfb = openCfb(bytes);
  } catch {
    return empty;
  }
  const stream = cfb.readStream('PowerPoint Document');
  if (!stream) return empty;

  // The Current User stream tells us where the most recent edit is and whether
  // the document is encrypted.
  const currentUser = cfb.readStream('Current User');
  if (currentUser && currentUser.length >= 16) {
    const headerToken = u32(currentUser, 12);
    if (headerToken === TOKEN_ENCRYPTED) return { slides: [], encrypted: true };
  }
  // The current edit is named by the Current User stream; if that is missing or
  // stale, fall back to the last UserEditAtom laid out in the stream.
  let currentEditOffset = currentUser && currentUser.length >= 20 ? u32(currentUser, 16) : 0;
  if (currentEditOffset === 0 || currentEditOffset + 8 > stream.length) {
    currentEditOffset = findLastUserEdit(stream);
  }

  // Resolve persist ids → stream offsets by walking the edit chain (newest first).
  const { persist, docPersistId } = buildPersistDirectory(stream, currentEditOffset);
  const docOffset = docPersistId !== undefined ? persist.get(docPersistId) : undefined;
  const docRec = docOffset !== undefined ? recordAt(stream, docOffset) : findDocument(stream);
  if (!docRec || docRec.type !== RT_DOCUMENT) {
    // No resolvable document container — last resort: every RT_Slide in the stream.
    return { slides: slidesByScan(stream), encrypted: false };
  }

  // If the slide list resolved no text (an unresolvable persist directory, say),
  // fall back to scanning the stream for slide containers directly.
  const slides = readSlideList(stream, docRec.data, persist);
  const hasText = slides.some((s) => s.paragraphs.some((p) => p.text.length > 0));
  const size = readSlideSize(docRec.data);
  return {
    slides: hasText ? slides : slidesByScan(stream),
    ...(size ?? {}),
    encrypted: false,
  };
}

// Build the persist-id → offset map and the document's persist id. The edits form
// a singly linked list newest→oldest (offsetLastEdit); a persist id set by a newer
// edit wins, so we only record an id the first time we see it (§2.3.3, §2.3.6).
function buildPersistDirectory(
  stream: Uint8Array,
  currentEditOffset: number,
): { persist: Map<number, number>; docPersistId: number | undefined } {
  const persist = new Map<number, number>();
  let docPersistId: number | undefined;
  let editOff = currentEditOffset;
  for (let guard = 0; editOff !== 0 && guard < MAX_EDITS; guard++) {
    const edit = recordAt(stream, editOff);
    if (!edit || edit.type !== RT_USER_EDIT_ATOM || edit.data.length < 20) break;
    const offsetLastEdit = u32(edit.data, 8);
    const offsetPersistDir = u32(edit.data, 12);
    if (docPersistId === undefined) docPersistId = u32(edit.data, 16); // the current edit's
    mergePersistDirectory(stream, offsetPersistDir, persist);
    if (offsetLastEdit >= editOff) break; // must point strictly backward
    editOff = offsetLastEdit;
  }
  return { persist, docPersistId };
}

// One PersistDirectoryAtom (§2.3.4) → entries merged into `persist` (first wins).
// Each PersistDirectoryEntry packs a starting persist id (low 20 bits) and a count
// (high 12 bits), then that many absolute stream offsets.
function mergePersistDirectory(
  stream: Uint8Array,
  offset: number,
  persist: Map<number, number>,
): void {
  const dir = recordAt(stream, offset);
  if (!dir || dir.type !== RT_PERSIST_DIRECTORY_ATOM) return;
  const d = dir.data;
  let p = 0;
  while (p + 4 <= d.length) {
    const persistIdAndCnt = u32(d, p);
    p += 4;
    const startId = persistIdAndCnt & 0xfffff;
    const count = (persistIdAndCnt >>> 20) & 0xfff;
    for (let i = 0; i < count && p + 4 <= d.length; i++) {
      const off = u32(d, p);
      p += 4;
      if (!persist.has(startId + i)) persist.set(startId + i, off);
    }
  }
}

// The document container's slide list (SlideListWithText, recInstance 0) → one
// PptSlide per SlidePersistAtom, in document order. Each SlidePersistAtom's
// persistIdRef (its first u32) resolves through the persist directory to a
// SlideContainer; the slide's text is read from there, falling back to the outline
// text that follows the SlidePersistAtom in the list when the slide stores none.
function readSlideList(
  stream: Uint8Array,
  docData: Uint8Array,
  persist: Map<number, number>,
): Array<PptSlide> {
  const slwt = findChild(
    docData,
    (r) => r.type === RT_SLIDE_LIST_WITH_TEXT && r.instance === SLWT_INSTANCE_SLIDES,
  );
  if (!slwt) return [];

  const slides: Array<PptSlide> = [];
  let outline: Array<PptParagraph> = [];
  let pendingRef: number | undefined;

  const flush = (): void => {
    if (pendingRef === undefined) return;
    const slideOff = persist.get(pendingRef);
    const slideRec = slideOff !== undefined ? recordAt(stream, slideOff) : undefined;
    const inline = slideRec && slideRec.type === RT_SLIDE ? collectText(slideRec.data, 0) : [];
    slides.push({ paragraphs: inline.length > 0 ? inline : outline });
    outline = [];
    pendingRef = undefined;
  };

  for (const rec of records(slwt)) {
    if (rec.type === RT_SLIDE_PERSIST_ATOM) {
      flush(); // close the previous slide before starting the next
      pendingRef = rec.data.length >= 4 ? u32(rec.data, 0) : undefined;
    } else if (rec.type === RT_TEXT_CHARS_ATOM) {
      outline.push(...splitParagraphs(decodeUtf16(rec.data)));
    } else if (rec.type === RT_TEXT_BYTES_ATOM) {
      outline.push(...splitParagraphs(decodeCp1252(rec.data)));
    }
  }
  flush();
  return slides;
}

// Recursively gather the text atoms inside a slide container (its drawing's client
// text boxes), in record order, as paragraphs.
function collectText(d: Uint8Array, depth: number): Array<PptParagraph> {
  if (depth > MAX_DEPTH) return [];
  const out: Array<PptParagraph> = [];
  for (const rec of records(d)) {
    if (rec.type === RT_TEXT_CHARS_ATOM) out.push(...splitParagraphs(decodeUtf16(rec.data)));
    else if (rec.type === RT_TEXT_BYTES_ATOM) out.push(...splitParagraphs(decodeCp1252(rec.data)));
    else if (rec.isContainer) out.push(...collectText(rec.data, depth + 1));
  }
  return out;
}

// Last-resort slide discovery: every RT_Slide container in the stream, in stream
// order (usually slide order), with its inline text. Used when the persist /
// document structure cannot be resolved.
function slidesByScan(stream: Uint8Array): Array<PptSlide> {
  const slides: Array<PptSlide> = [];
  collectSlides(stream, slides, 0);
  return slides;
}

function collectSlides(d: Uint8Array, out: Array<PptSlide>, depth: number): void {
  if (depth > MAX_DEPTH) return;
  for (const rec of records(d)) {
    if (rec.type === RT_SLIDE) out.push({ paragraphs: collectText(rec.data, 0) });
    else if (rec.isContainer) collectSlides(rec.data, out, depth + 1);
  }
}

// The first matching top-level child of a container's data.
function findChild(d: Uint8Array, pred: (r: PptRecord) => boolean): Uint8Array | undefined {
  for (const rec of records(d)) if (pred(rec)) return rec.data;
  return undefined;
}

// Find the DocumentContainer by scanning top-level records — the fallback when the
// Current User → persist path does not resolve it.
function findDocument(stream: Uint8Array): PptRecord | undefined {
  for (const rec of records(stream)) if (rec.type === RT_DOCUMENT) return rec;
  return undefined;
}

// The offset of the last top-level UserEditAtom — the most recent edit, used when
// the Current User stream is missing or points out of range.
function findLastUserEdit(stream: Uint8Array): number {
  let off = 0;
  let last = 0;
  while (off + 8 <= stream.length) {
    const rec = recordAt(stream, off);
    if (!rec) break;
    if (rec.type === RT_USER_EDIT_ATOM) last = off;
    const consumed = 8 + rec.data.length;
    if (consumed <= 0) break;
    off += consumed;
  }
  return last;
}

// DocumentAtom.slideSize (§2.4.2) → the page size in points. slideSize is a
// PointStruct (x, y as i32) in master units (576 per inch); 1 inch = 72 points, so
// points = units / 8. Out-of-range values fall back to the caller's default.
const MASTER_UNITS_PER_INCH = 576;
const POINTS_PER_INCH = 72;
function readSlideSize(
  docData: Uint8Array,
): { slideWidthPt: number; slideHeightPt: number } | undefined {
  const atom = findChild(docData, (r) => r.type === RT_DOCUMENT_ATOM);
  if (!atom || atom.length < 8) return undefined;
  const cx = i32(atom, 0);
  const cy = i32(atom, 4);
  const toPt = (u: number): number => (u / MASTER_UNITS_PER_INCH) * POINTS_PER_INCH;
  const w = toPt(cx);
  const h = toPt(cy);
  // Sanity-bound to a plausible slide (3"–60") so a misread size can't distort the
  // page; the reader defaults to a standard 10"×7.5" deck otherwise.
  if (w < 216 || w > 4320 || h < 216 || h > 4320) return undefined;
  return { slideWidthPt: w, slideHeightPt: h };
}

// Split a decoded text run into paragraphs at the CR (0x0D) paragraph mark; the
// vertical tab (0x0B) is a soft line break, normalized to a space. Control
// characters and the UTF-16 BOM are stripped. A trailing empty paragraph (from a
// terminating CR) is dropped.
function splitParagraphs(text: string): Array<PptParagraph> {
  const paras: Array<PptParagraph> = [];
  for (const raw of text.split('\r')) {
    let s = '';
    for (let i = 0; i < raw.length; i++) {
      const c = raw.charCodeAt(i);
      if (c === 0x0b || c === 0x0a || c === 0x09) s += ' ';
      else if (c === 0xfeff || c < 0x20) continue;
      else s += raw[i];
    }
    paras.push({ text: s });
  }
  // Drop a single trailing empty paragraph left by the terminating CR.
  if (paras.length > 1 && paras[paras.length - 1]!.text.length === 0) paras.pop();
  return paras;
}

function decodeUtf16(d: Uint8Array): string {
  let out = '';
  for (let i = 0; i + 1 < d.length; i += 2) out += String.fromCharCode(d[i]! | (d[i + 1]! << 8));
  return out;
}

// Windows-1252 high range (0x80–0x9F); the rest is Latin-1 (== Unicode 0–0xFF).
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
  0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
  0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

function decodeCp1252(d: Uint8Array): string {
  let out = '';
  for (let i = 0; i < d.length; i++) {
    const c = d[i]!;
    out += String.fromCharCode(c >= 0x80 && c <= 0x9f ? CP1252_HIGH[c - 0x80]! : c);
  }
  return out;
}

function u16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
}
function u32(d: Uint8Array, off: number): number {
  return (
    ((d[off] ?? 0) |
      ((d[off + 1] ?? 0) << 8) |
      ((d[off + 2] ?? 0) << 16) |
      ((d[off + 3] ?? 0) << 24)) >>>
    0
  );
}
function i32(d: Uint8Array, off: number): number {
  return u32(d, off) | 0;
}
