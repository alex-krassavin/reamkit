// Legacy `.ppt` (PowerPoint 97–2003, [MS-PPT]) record reader + per-slide content
// extraction (PPT-1..7). A `.ppt` is an OLE2/CFB container whose `PowerPoint
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
// outline text when a slide stores none inline. PPT-2 pairs each text atom with the
// StyleTextPropAtom that follows it, so runs carry bold / italic / underline / size
// / colour and paragraphs carry alignment / indent level. PPT-3 reads embedded
// pictures (Pictures stream). PPT-4 reads each shape's OfficeArtClientAnchor so text
// boxes and pictures carry their slide rectangle; an un-anchored shape (e.g. a
// placeholder inheriting master geometry) falls back to reading-order flow. PPT-5
// reads decorative autoshapes (the FSP shape type + the OPT fill / line colour).
// PPT-6 resolves a shape's scheme-relative fill / line colour through the slide's
// colour scheme (its own SlideSchemeColorSchemeAtom, or the master's when the slide
// follows it), instead of dropping it as it did when only a literal sRGB value read.
// PPT-7 reads a freeform shape's exact custom geometry — the OPT's pVertices and
// pSegmentInfo complex arrays walked into a path (moveTo / lineTo / cubic curveTo /
// close) in its geometry-bounds space. Everything is best-effort and graceful-on-
// failure — structural doubt yields missing content, never wrong content —
// mirroring the `.doc`/`.xls` readers.

import { isCfb, openCfb } from '@/core/ole/cfb';

// --- [MS-PPT] §2.13.24 record types (the ones the `.ppt` reader needs) --------
const RT_DOCUMENT = 0x03e8;
const RT_DOCUMENT_ATOM = 0x03e9;
const RT_SLIDE_LIST_WITH_TEXT = 0x0ff0;
const RT_SLIDE_PERSIST_ATOM = 0x03f3;
const RT_SLIDE = 0x03ee;
const RT_USER_EDIT_ATOM = 0x0ff5;
const RT_PERSIST_DIRECTORY_ATOM = 0x1772;
const RT_DRAWING_GROUP = 0x040b;
const RT_TEXT_HEADER_ATOM = 0x0f9f;
const RT_TEXT_CHARS_ATOM = 0x0fa0;
const RT_STYLE_TEXT_PROP_ATOM = 0x0fa1;
const RT_TEXT_BYTES_ATOM = 0x0fa8;
const RT_SLIDE_ATOM = 0x03ef; // SlideAtom — carries masterIdRef + slideFlags (PPT-6)
const RT_MAIN_MASTER = 0x03f8; // MainMasterContainer — its colour scheme is inherited
const RT_COLOR_SCHEME_ATOM = 0x07f0; // SlideSchemeColorSchemeAtom — the 8-colour scheme

// OfficeArt (Escher) record types for the drawing layer (PPT-3), sharing the same
// header as the PPT records. The DrawingGroupContainer holds the picture store
// (OfficeArtBStoreContainer of OfficeArtFBSE entries); each slide's shapes
// reference a store entry by a 1-based `pib` index in their OPT property table.
const FBT_BSTORE_CONTAINER = 0xf001;
const FBT_BSE = 0xf007;
const FBT_SP_CONTAINER = 0xf004;
const FBT_FSP = 0xf00a; // OfficeArtFSP — recInstance is the shape type, body has the flags
const FBT_OPT = 0xf00b;
const FBT_CLIENT_TEXTBOX = 0xf00d; // OfficeArtClientTextbox — holds the PPT text atoms
const FBT_CLIENT_ANCHOR = 0xf010; // OfficeArtClientAnchor — the slide rectangle (PPT-4)
const PROP_PIB = 0x0104; // OPT property (low 14 bits): 1-based index into the FBSE store
const PROP_FILL_COLOR = 0x0181; // OPT fillColor (PPT-5)
const PROP_LINE_COLOR = 0x01c0; // OPT lineColor (PPT-5)
// Freeform geometry OPT properties (§2.3.6 / [MS-ODRAW]); the bounds are simple
// LONGs, the vertices / segment info complex array properties (PPT-7).
const PROP_GEO_LEFT = 0x0140;
const PROP_GEO_TOP = 0x0141;
const PROP_GEO_RIGHT = 0x0142;
const PROP_GEO_BOTTOM = 0x0143;
const PROP_VERTICES = 0x0145;
const PROP_SEGMENT_INFO = 0x0146;
// The colour scheme is the slideSchemeColorSchemeAtom — recInstance 1 (recInstance
// 6 is the alternative SchemeListElement palette, which is not used for resolution).
const COLOR_SCHEME_INSTANCE = 1;
// SlideAtom.slideFlags.fMasterScheme (§2.4.24): set ⇒ follow the master's scheme.
const SLIDE_FLAG_MASTER_SCHEME = 0x0002;
// OfficeArtCOLORREF flags byte (§2.2.2): fSchemeIndex ⇒ the red byte is an index
// (0–7) into the colour scheme; flags 0 ⇒ a literal sRGB value.
const COLORREF_SCHEME_INDEX = 0x08;
const FBSE_FODELAY_OFFSET = 28; // foDelay in the FBSE data (record offset 36 − 8-byte header)
// OfficeArtFSP flags (§2.2.40): skip groups, the patriarch and the background shape.
const FSP_FLAG_GROUP = 0x0001;
const FSP_FLAG_PATRIARCH = 0x0004;
const FSP_FLAG_BACKGROUND = 0x0400;

// CurrentUserAtom.headerToken (§2.3.2) — the encrypted variant means the document
// streams are obfuscated and the text cannot be read.
const TOKEN_ENCRYPTED = 0xf3d1c4df;

// A SlideListWithText's recInstance distinguishes the three lists it can be
// (§2.4.14.3): 0 = slides, 1 = masters, 2 = notes. PPT-1 reads the slide list.
const SLWT_INSTANCE_SLIDES = 0;

// Guard against a crafted edit chain or persist directory looping forever.
const MAX_EDITS = 4096;
const MAX_DEPTH = 24;

// A run of uniformly-formatted text within a paragraph (PPT-2).
export interface PptRun {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly sizePt?: number;
  readonly colorHex?: string;
}

// One slide paragraph: its runs plus the raw PPT alignment enum (0 left, 1 center,
// 2 right, 3 justify, 4 distribute) and indent level (0–4), mapped by the reader.
export interface PptParagraph {
  readonly runs: ReadonlyArray<PptRun>;
  readonly align?: number;
  readonly level?: number;
}
// An embedded picture referenced by a slide shape — the raw image bytes pulled
// from the Pictures stream (PPT-3).
export interface PptImage {
  readonly bytes: Uint8Array;
}
// A shape's rectangle on the slide, in points (from the OfficeArtClientAnchor).
export interface PptRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}
// One command of a freeform path (PPT-7), in geometry-bounds space (0..pathWidth,
// 0..pathHeight, y down) — the same path-space the DrawingML custom geometry uses.
export type PptPathCmd =
  | { readonly kind: 'move'; readonly x: number; readonly y: number }
  | { readonly kind: 'line'; readonly x: number; readonly y: number }
  | {
      readonly kind: 'cubic';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly kind: 'close' };
// A shape's exact custom geometry (PPT-7): the freeform path from the OPT's
// pVertices + pSegmentInfo arrays, in its geometry-bounds coordinate space.
export interface PptCustomGeometry {
  readonly pathWidth: number;
  readonly pathHeight: number;
  readonly commands: ReadonlyArray<PptPathCmd>;
}
// A decorative autoshape (PPT-5): its preset type plus any literal fill / line
// colour. Carried only by an anchored shape that has no text and no picture. A
// freeform additionally carries its exact custom geometry (PPT-7).
export interface PptAutoShape {
  readonly shapeType: number; // MSOSPT (the FSP recInstance)
  readonly fillColorHex?: string;
  readonly lineColorHex?: string;
  readonly geometry?: PptCustomGeometry;
}
// One slide shape (PPT-4..5): its text, picture and/or autoshape geometry, plus its
// slide rectangle when the shape carries an explicit anchor (else it is laid out in
// reading order).
export interface PptShape {
  readonly rectPt?: PptRect;
  readonly paragraphs?: ReadonlyArray<PptParagraph>;
  readonly image?: PptImage;
  readonly autoShape?: PptAutoShape;
}
export interface PptSlide {
  readonly shapes: ReadonlyArray<PptShape>;
}
// The document-level picture store (FBSE offsets into the Pictures stream) plus the
// Pictures stream itself — threaded into the slide walks to resolve shape blips.
interface ImageContext {
  readonly foDelays: ReadonlyArray<number>;
  readonly pictures: Uint8Array;
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
  const noImages: ImageContext = { foDelays: [], pictures: new Uint8Array(0) };
  if (!docRec || docRec.type !== RT_DOCUMENT) {
    // No resolvable document container — last resort: every RT_Slide in the stream.
    return { slides: slidesByScan(stream, noImages), encrypted: false };
  }

  // The picture store: the FBSE offsets into the Pictures stream, resolved by a
  // shape's 1-based `pib` index (PPT-3).
  const img: ImageContext = {
    foDelays: parseFbseStore(docRec.data),
    pictures: cfb.readStream('Pictures') ?? new Uint8Array(0),
  };

  // If the slide list resolved no content (an unresolvable persist directory, say),
  // fall back to scanning the stream for slide containers directly.
  const slides = readSlideList(stream, docRec.data, persist, img);
  const hasContent = slides.some(slideHasContent);
  const size = readSlideSize(docRec.data);
  return {
    slides: hasContent ? slides : slidesByScan(stream, img),
    ...(size ?? {}),
    encrypted: false,
  };
}

// Whether a slide carries any readable content (text, a picture or an autoshape).
function slideHasContent(s: PptSlide): boolean {
  return s.shapes.some(
    (sh) => shapeHasText(sh) || sh.image !== undefined || sh.autoShape !== undefined,
  );
}
function shapeHasText(sh: PptShape): boolean {
  return sh.paragraphs?.some((p) => paragraphText(p).length > 0) ?? false;
}

// A slide's shapes: the positioned text boxes and pictures from its drawing, with
// a reading-order text fallback (whole-slide inline text, else the outline) added
// as an un-anchored shape when no shape carries text.
function slideShapes(
  slideData: Uint8Array,
  img: ImageContext,
  outline: ReadonlyArray<PptParagraph>,
  scheme: SchemeColors | undefined,
): Array<PptShape> {
  const shapes = collectShapes(slideData, img, scheme);
  if (!shapes.some(shapeHasText)) {
    const inline = collectParagraphs(slideData, 0);
    const text = inline.some((p) => paragraphText(p).length > 0) ? inline : outline;
    if (text.some((p) => paragraphText(p).length > 0)) shapes.push({ paragraphs: text });
  }
  return shapes;
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
  img: ImageContext,
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
    const isSlide = slideRec !== undefined && slideRec.type === RT_SLIDE;
    const shapes = isSlide
      ? slideShapes(slideRec.data, img, outline, resolveScheme(slideRec.data, persist, stream))
      : outline.some((p) => paragraphText(p).length > 0)
        ? [{ paragraphs: outline }]
        : [];
    slides.push({ shapes });
    outline = [];
    pendingRef = undefined;
  };

  const recs = [...records(slwt)];
  for (let i = 0; i < recs.length; i++) {
    const rec = recs[i]!;
    if (rec.type === RT_SLIDE_PERSIST_ATOM) {
      flush(); // close the previous slide before starting the next
      pendingRef = rec.data.length >= 4 ? u32(rec.data, 0) : undefined;
    } else if (rec.type === RT_TEXT_CHARS_ATOM || rec.type === RT_TEXT_BYTES_ATOM) {
      outline.push(...styledParagraphs(rec, findFollowingStyle(recs, i)));
    }
  }
  flush();
  return slides;
}

// Recursively gather the text atoms inside a slide container (its drawing's client
// text boxes), in record order, pairing each with the StyleTextPropAtom that
// follows it so runs carry their formatting.
function collectParagraphs(d: Uint8Array, depth: number): Array<PptParagraph> {
  if (depth > MAX_DEPTH) return [];
  const out: Array<PptParagraph> = [];
  const recs = [...records(d)];
  for (let i = 0; i < recs.length; i++) {
    const rec = recs[i]!;
    if (rec.type === RT_TEXT_CHARS_ATOM || rec.type === RT_TEXT_BYTES_ATOM) {
      out.push(...styledParagraphs(rec, findFollowingStyle(recs, i)));
    } else if (rec.isContainer) {
      out.push(...collectParagraphs(rec.data, depth + 1));
    }
  }
  return out;
}

// The StyleTextPropAtom that applies to the text atom at index `i`: the first one
// following it before the next text atom / text header begins.
function findFollowingStyle(recs: ReadonlyArray<PptRecord>, i: number): Uint8Array | undefined {
  for (let j = i + 1; j < recs.length; j++) {
    const t = recs[j]!.type;
    if (t === RT_STYLE_TEXT_PROP_ATOM) return recs[j]!.data;
    if (t === RT_TEXT_CHARS_ATOM || t === RT_TEXT_BYTES_ATOM || t === RT_TEXT_HEADER_ATOM) break;
  }
  return undefined;
}

// A text atom (+ its optional StyleTextPropAtom) → formatted paragraphs.
function styledParagraphs(
  textRec: PptRecord,
  styleData: Uint8Array | undefined,
): Array<PptParagraph> {
  const text =
    textRec.type === RT_TEXT_CHARS_ATOM ? decodeUtf16(textRec.data) : decodeCp1252(textRec.data);
  return buildStyledParagraphs(text, styleData);
}

// Last-resort slide discovery: every RT_Slide container in the stream, in stream
// order (usually slide order), with its inline text and images. Used when the
// persist / document structure cannot be resolved.
function slidesByScan(stream: Uint8Array, img: ImageContext): Array<PptSlide> {
  const slides: Array<PptSlide> = [];
  collectSlides(stream, slides, img, 0);
  return slides;
}

function collectSlides(
  d: Uint8Array,
  out: Array<PptSlide>,
  img: ImageContext,
  depth: number,
): void {
  if (depth > MAX_DEPTH) return;
  for (const rec of records(d)) {
    if (rec.type === RT_SLIDE) {
      out.push({ shapes: slideShapes(rec.data, img, [], resolveScheme(rec.data)) });
    } else if (rec.isContainer) collectSlides(rec.data, out, img, depth + 1);
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

// === Shapes: placement + pictures (PPT-3..4) =================================

// The picture store: the DrawingGroupContainer's OfficeArtBStoreContainer holds
// one OfficeArtFBSE per stored image; each FBSE's foDelay is the byte offset into
// the Pictures stream of that image's OfficeArtBlip. A slide shape's 1-based `pib`
// indexes this array.
function parseFbseStore(docData: Uint8Array): Array<number> {
  const group = findChild(docData, (r) => r.type === RT_DRAWING_GROUP);
  if (!group) return [];
  const store = findDescendantContainer(group, FBT_BSTORE_CONTAINER, 0);
  if (!store) return [];
  const foDelays: Array<number> = [];
  for (const r of records(store)) {
    if (r.type === FBT_BSE) {
      foDelays.push(
        r.data.length >= FBSE_FODELAY_OFFSET + 4 ? u32(r.data, FBSE_FODELAY_OFFSET) : 0,
      );
    }
  }
  return foDelays;
}

// Depth-first search for the first container of `type` (the Escher tree nests an
// OfficeArtDggContainer inside the DrawingGroup; the BStoreContainer is within it).
function findDescendantContainer(
  d: Uint8Array,
  type: number,
  depth: number,
): Uint8Array | undefined {
  if (depth > MAX_DEPTH) return undefined;
  for (const r of records(d)) {
    if (r.type === type) return r.data;
    if (r.isContainer) {
      const found = findDescendantContainer(r.data, type, depth + 1);
      if (found) return found;
    }
  }
  return undefined;
}

// Every drawing shape in a slide → a PptShape: its client-anchor rectangle (when
// present), the text in its client text box, and the picture its `pib` references.
// Top-level shapes carry an OfficeArtClientAnchor (0xF010); the recursion descends
// the group/drawing containers but a grouped shape's group-relative ChildAnchor is
// left unpositioned (it falls back to reading-order flow).
function collectShapes(
  slideData: Uint8Array,
  img: ImageContext,
  scheme: SchemeColors | undefined,
): Array<PptShape> {
  const out: Array<PptShape> = [];
  collectShapeContainers(slideData, img, out, 0, scheme);
  return out;
}

function collectShapeContainers(
  d: Uint8Array,
  img: ImageContext,
  out: Array<PptShape>,
  depth: number,
  scheme: SchemeColors | undefined,
): void {
  if (depth > MAX_DEPTH) return;
  for (const r of records(d)) {
    if (r.type === FBT_SP_CONTAINER) {
      let rectPt: PptRect | undefined;
      let paragraphs: Array<PptParagraph> | undefined;
      let pib: number | undefined;
      let shapeType = 0;
      let fspFlags = 0;
      let fillColorHex: string | undefined;
      let lineColorHex: string | undefined;
      let geometry: PptCustomGeometry | undefined;
      for (const child of records(r.data)) {
        if (child.type === FBT_FSP) {
          shapeType = child.instance;
          fspFlags = child.data.length >= 8 ? u32(child.data, 4) : 0;
        } else if (child.type === FBT_CLIENT_ANCHOR) rectPt = parseAnchor(child.data);
        else if (child.type === FBT_CLIENT_TEXTBOX) paragraphs = collectParagraphs(child.data, 0);
        else if (child.type === FBT_OPT) {
          pib = optProperty(child.data, child.instance, PROP_PIB);
          fillColorHex = optColor(child.data, child.instance, PROP_FILL_COLOR, scheme);
          lineColorHex = optColor(child.data, child.instance, PROP_LINE_COLOR, scheme);
          geometry = parseFreeformGeometry(child.data, child.instance);
        }
      }
      let image: PptImage | undefined;
      if (pib !== undefined && pib >= 1 && pib <= img.foDelays.length) {
        const bytes = readBlipBytes(img.pictures, img.foDelays[pib - 1]!);
        if (bytes) image = { bytes };
      }
      const textParas =
        paragraphs && paragraphs.some((p) => paragraphText(p).length > 0) ? paragraphs : undefined;
      // A decorative autoshape: an anchored shape with a preset type (or its own
      // freeform geometry — PPT-7), a literal fill or line colour, and no
      // text/picture — and not a group / patriarch / background shape.
      const decorative =
        (fspFlags & (FSP_FLAG_GROUP | FSP_FLAG_PATRIARCH | FSP_FLAG_BACKGROUND)) === 0;
      const autoShape =
        !textParas &&
        !image &&
        rectPt &&
        (shapeType > 0 || geometry) &&
        decorative &&
        (fillColorHex || lineColorHex)
          ? {
              shapeType,
              ...(fillColorHex ? { fillColorHex } : {}),
              ...(lineColorHex ? { lineColorHex } : {}),
              ...(geometry ? { geometry } : {}),
            }
          : undefined;
      if (textParas || image || autoShape) {
        out.push({
          ...(rectPt ? { rectPt } : {}),
          ...(textParas ? { paragraphs: textParas } : {}),
          ...(image ? { image } : {}),
          ...(autoShape ? { autoShape } : {}),
        });
      }
    } else if (r.isContainer) {
      collectShapeContainers(r.data, img, out, depth + 1, scheme);
    }
  }
}

// OfficeArtClientAnchor (§2.7.1/§2.7.2) → the shape's rectangle in points. The
// data is four coordinates in the order top, left, right, bottom — 2-byte i16 in
// the common 8-byte SmallRectStruct form, 4-byte i32 in the 16-byte RectStruct
// form — in master units (576 per inch ⇒ points = units / 8). An implausible
// rectangle yields undefined, so the shape falls back to reading-order flow.
function parseAnchor(data: Uint8Array): PptRect | undefined {
  let top: number;
  let left: number;
  let right: number;
  let bottom: number;
  if (data.length >= 16) {
    top = i32(data, 0);
    left = i32(data, 4);
    right = i32(data, 8);
    bottom = i32(data, 12);
  } else if (data.length >= 8) {
    top = i16(data, 0);
    left = i16(data, 2);
    right = i16(data, 4);
    bottom = i16(data, 6);
  } else {
    return undefined;
  }
  const x = left / 8;
  const y = top / 8;
  const w = (right - left) / 8;
  const h = (bottom - top) / 8;
  // Bound to a plausible on-slide rectangle; a misread can't fling content off-page.
  if (!(w > 0 && h > 0) || w > 5000 || h > 5000 || Math.abs(x) > 5000 || Math.abs(y) > 5000) {
    return undefined;
  }
  return { x, y, w, h };
}

// §2.3.7.2 OfficeArtFOPT — `count` properties, each a 2-byte id (low 14 bits) + a
// 4-byte value. Returns the simple (non-complex) value of `wantId`.
function optProperty(d: Uint8Array, count: number, wantId: number): number | undefined {
  for (let i = 0; i < count && i * 6 + 6 <= d.length; i++) {
    const id = u16(d, i * 6);
    if ((id & 0x3fff) === wantId && (id & 0x8000) === 0) return u32(d, i * 6 + 2);
  }
  return undefined;
}

// A complex OFOPT property's blob: the `count` fixed 6-byte entries come first,
// then each complex entry's data follows in entry order, its byte length given by
// that entry's 4-byte value. Returns the bytes of `wantId`'s blob (bounds-checked).
function optComplex(d: Uint8Array, count: number, wantId: number): Uint8Array | undefined {
  let offset = count * 6; // the complex region starts after the fixed entries
  for (let i = 0; i < count && i * 6 + 6 <= d.length; i++) {
    const id = u16(d, i * 6);
    if ((id & 0x8000) === 0) continue; // fComplex clear ⇒ no trailing data
    const len = u32(d, i * 6 + 2);
    if ((id & 0x3fff) === wantId) {
      return offset + len <= d.length && len > 0 ? d.subarray(offset, offset + len) : undefined;
    }
    offset += len;
  }
  return undefined;
}

// An IMsoArray complex property: a 6-byte header (nElems u16, nElemsAlloc u16,
// cbElem i16) then nElems × cbElem bytes. A negative cbElem (e.g. 0xFFF0 = −16)
// encodes the real element size as (−cbElem) >> 2 (Apache POI EscherArrayProperty).
function arrayHeader(blob: Uint8Array): { count: number; size: number; base: number } | undefined {
  if (blob.length < 6) return undefined;
  const raw = u16(blob, 4);
  const signed = raw >= 0x8000 ? raw - 0x10000 : raw;
  const size = signed < 0 ? -signed >> 2 : signed;
  if (size <= 0) return undefined;
  return { count: u16(blob, 0), size, base: 6 };
}

// A 4-byte OPT value as a signed LONG (the geometry bounds may be negative).
function signedLong(v: number | undefined): number | undefined {
  return v === undefined ? undefined : v >= 0x80000000 ? v - 0x100000000 : v;
}

const MAX_PATH_ELEMS = 20000; // a sane cap on a freeform's vertex / segment count

// A freeform shape's exact custom geometry (PPT-7): walk the pSegmentInfo opcodes,
// pulling points from pVertices, into path commands in the geometry-bounds space.
// The opcode's top 3 bits are the MSOPATHTYPE (lineTo 0, curveTo 1, moveTo 2,
// close 3, end 4, escape 5); a lineTo/moveTo consumes one point, a curveTo three
// (a cubic bezier), close/end none. An arc / ellipse escape (which would consume
// points for a curve we do not synthesize) makes the whole path bail to its preset
// — a missing custom geometry, never a mis-aligned one. Mirrors POI HSLFAutoShape.
function parseFreeformGeometry(d: Uint8Array, count: number): PptCustomGeometry | undefined {
  const vBlob = optComplex(d, count, PROP_VERTICES);
  const sBlob = optComplex(d, count, PROP_SEGMENT_INFO);
  if (!vBlob || !sBlob) return undefined;

  // The vertices: each a POINT — an i16 pair (4-byte element) or i32 pair (8-byte).
  const vh = arrayHeader(vBlob);
  if (!vh || vh.count > MAX_PATH_ELEMS) return undefined;
  const verts: Array<{ x: number; y: number }> = [];
  for (let i = 0; i < vh.count; i++) {
    const off = vh.base + i * vh.size;
    if (off + vh.size > vBlob.length) break;
    if (vh.size >= 8) verts.push({ x: i32(vBlob, off), y: i32(vBlob, off + 4) });
    else if (vh.size >= 4) verts.push({ x: i16(vBlob, off), y: i16(vBlob, off + 2) });
    else return undefined;
  }
  if (verts.length === 0) return undefined;

  const sh = arrayHeader(sBlob);
  if (!sh || sh.size < 2 || sh.count > MAX_PATH_ELEMS) return undefined;

  // The geometry-bounds box; absent extents fall back to the vertices' own range.
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const v of verts) {
    if (v.x < minX) minX = v.x;
    if (v.x > maxX) maxX = v.x;
    if (v.y < minY) minY = v.y;
    if (v.y > maxY) maxY = v.y;
  }
  const left = signedLong(optProperty(d, count, PROP_GEO_LEFT)) ?? minX;
  const top = signedLong(optProperty(d, count, PROP_GEO_TOP)) ?? minY;
  const right = signedLong(optProperty(d, count, PROP_GEO_RIGHT)) ?? maxX;
  const bottom = signedLong(optProperty(d, count, PROP_GEO_BOTTOM)) ?? maxY;
  const pathWidth = right - left;
  const pathHeight = bottom - top;
  if (!(pathWidth > 0) || !(pathHeight > 0) || pathWidth > 1e7 || pathHeight > 1e7)
    return undefined;

  const commands: Array<PptPathCmd> = [];
  let vi = 0;
  let drew = false; // a path needs at least one drawing (line / curve) segment
  for (let i = 0; i < sh.count; i++) {
    const off = sh.base + i * sh.size;
    if (off + 2 > sBlob.length) break;
    const type = (u16(sBlob, off) >> 13) & 0x7;
    if (type === 2) {
      const v = verts[vi++];
      if (!v) break;
      commands.push({ kind: 'move', x: v.x - left, y: v.y - top });
    } else if (type === 0) {
      const v = verts[vi++];
      if (!v) break;
      commands.push({ kind: 'line', x: v.x - left, y: v.y - top });
      drew = true;
    } else if (type === 1) {
      const c1 = verts[vi++];
      const c2 = verts[vi++];
      const e = verts[vi++];
      if (!c1 || !c2 || !e) break;
      commands.push({
        kind: 'cubic',
        x1: c1.x - left,
        y1: c1.y - top,
        x2: c2.x - left,
        y2: c2.y - top,
        x: e.x - left,
        y: e.y - top,
      });
      drew = true;
    } else if (type === 3) {
      commands.push({ kind: 'close' });
    } else if (type === 4) {
      break; // msopathEnd
    } else if (type === 5) {
      // An escape: low-range codes (arc / ellipse / extension) consume points for
      // geometry we do not model — bail; high-range codes are render hints, skip.
      if (((u16(sBlob, off) >> 8) & 0x1f) < 0x0a) return undefined;
    }
  }
  return drew ? { pathWidth, pathHeight, commands } : undefined;
}

// An OPT colour property → 6-hex RGB. A literal sRGB value (flags byte 0) is taken
// directly; a scheme-relative value (fSchemeIndex) resolves through the slide's
// colour scheme (PPT-6). Palette / system colours, and any scheme index without a
// resolved scheme, are skipped (missing colour, never a wrong one).
function optColor(
  d: Uint8Array,
  count: number,
  wantId: number,
  scheme: SchemeColors | undefined,
): string | undefined {
  const v = optProperty(d, count, wantId);
  if (v === undefined) return undefined;
  const flags = v >>> 24;
  if (flags === 0) return `${hex2(v & 0xff)}${hex2((v >> 8) & 0xff)}${hex2((v >> 16) & 0xff)}`;
  if (flags === COLORREF_SCHEME_INDEX && scheme) {
    const idx = v & 0xff;
    if (idx < scheme.length) return scheme[idx];
  }
  return undefined;
}

// The eight scheme colours as 6-hex RGB, indexed by COLORREF scheme slot (0 =
// background, 1 = text&lines, 2 = shadows, 3 = title text, 4 = fills, 5 = accent,
// 6 = accent&hyperlink, 7 = accent&followed-hyperlink).
type SchemeColors = ReadonlyArray<string>;

// The colour scheme that applies to a slide's shapes: the slide's own
// SlideSchemeColorSchemeAtom, unless slideAtom.slideFlags.fMasterScheme is set, in
// which case the master's scheme is used (resolved through the persist directory).
// Returns undefined when no scheme can be located, so a scheme colour stays dropped.
function resolveScheme(
  slideData: Uint8Array,
  persist?: Map<number, number>,
  stream?: Uint8Array,
): SchemeColors | undefined {
  let schemeData = findColorScheme(slideData);
  const slideAtom = findChild(slideData, (r) => r.type === RT_SLIDE_ATOM);
  if (slideAtom && slideAtom.length >= 22 && persist && stream) {
    const followMaster = (u16(slideAtom, 20) & SLIDE_FLAG_MASTER_SCHEME) !== 0;
    if (followMaster) {
      const masterOff = persist.get(u32(slideAtom, 12)); // masterIdRef
      const masterRec = masterOff !== undefined ? recordAt(stream, masterOff) : undefined;
      if (masterRec && masterRec.type === RT_MAIN_MASTER) {
        const masterScheme = findColorScheme(masterRec.data);
        if (masterScheme) schemeData = masterScheme;
      }
    }
  }
  return schemeData ? parseColorScheme(schemeData) : undefined;
}

// The slide/master colour scheme atom (recInstance 1), skipping the alternative
// SchemeListElement palettes (recInstance 6) that share the record type.
function findColorScheme(containerData: Uint8Array): Uint8Array | undefined {
  return findChild(
    containerData,
    (r) => r.type === RT_COLOR_SCHEME_ATOM && r.instance === COLOR_SCHEME_INSTANCE,
  );
}

// A SlideSchemeColorSchemeAtom body → its eight ColorStruct entries (red, green,
// blue, unused) as 6-hex RGB.
function parseColorScheme(d: Uint8Array): SchemeColors | undefined {
  if (d.length < 32) return undefined;
  const colors: Array<string> = [];
  for (let i = 0; i < 8; i++) {
    const o = i * 4;
    colors.push(`${hex2(d[o]!)}${hex2(d[o + 1]!)}${hex2(d[o + 2]!)}`);
  }
  return colors;
}

// The OfficeArtBlip at `foDelay` in the Pictures stream → its raw image bytes. The
// blip header carries one or two UIDs before the payload; a bounded magic scan
// finds the PNG/JPEG start regardless (mirrors the `.xls` Escher reader).
function readBlipBytes(pictures: Uint8Array, foDelay: number): Uint8Array | undefined {
  if (foDelay < 0 || foDelay + 8 > pictures.length) return undefined;
  const recLen = u32(pictures, foDelay + 4);
  const start = foDelay + 8;
  const blip = pictures.subarray(start, Math.min(pictures.length, start + recLen));
  const limit = Math.min(blip.length, 80);
  for (let off = 0; off < limit; off++) {
    for (const magic of IMAGE_MAGICS) {
      if (matchesMagic(blip, off, magic)) return blip.subarray(off);
    }
  }
  return undefined;
}

// PNG and JPEG signatures — the raster formats the renderer can embed.
const IMAGE_MAGICS: ReadonlyArray<ReadonlyArray<number>> = [
  [0x89, 0x50, 0x4e, 0x47], // PNG
  [0xff, 0xd8, 0xff], // JPEG
];

function matchesMagic(d: Uint8Array, off: number, magic: ReadonlyArray<number>): boolean {
  if (off + magic.length > d.length) return false;
  for (let i = 0; i < magic.length; i++) if (d[off + i] !== magic[i]) return false;
  return true;
}

// === Text + formatting (PPT-2) ===============================================

// One character run / paragraph run parsed out of a StyleTextPropAtom.
interface CharRun {
  count: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  sizePt?: number;
  colorHex?: string;
}
interface ParaRun {
  count: number;
  align?: number;
  level?: number;
}

export function paragraphText(p: PptParagraph): string {
  return p.runs.map((r) => r.text).join('');
}

// A decoded text run (+ its StyleTextPropAtom) → paragraphs of formatted runs. The
// text is split at the CR (0x0D) paragraph mark; the StyleTextPropAtom's character
// runs colour each slice. Without a style atom, each paragraph is a single plain
// run.
function buildStyledParagraphs(
  text: string,
  styleData: Uint8Array | undefined,
): Array<PptParagraph> {
  const style = styleData ? parseStyleTextProp(styleData, text.length) : undefined;
  const charProps = style ? expandCharProps(style.charRuns, text.length) : undefined;
  const paraRuns = style?.paraRuns ?? [];

  const paras: Array<PptParagraph> = [];
  let runs: Array<PptRun> = [];
  let cur = '';
  let curProps: CharRun | undefined;
  let paraIndex = 0;

  const pushRun = (): void => {
    if (cur.length > 0) runs.push(toRun(cur, curProps));
    cur = '';
  };
  const endParagraph = (): void => {
    pushRun();
    const meta = paraRuns[paraIndex];
    paras.push({
      runs,
      ...(meta?.align !== undefined ? { align: meta.align } : {}),
      ...(meta?.level !== undefined ? { level: meta.level } : {}),
    });
    runs = [];
    paraIndex++;
  };

  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 0x0d) {
      endParagraph();
      continue;
    }
    const props = charProps?.[i];
    if (cur.length > 0 && !sameProps(props, curProps)) pushRun();
    if (cur.length === 0) curProps = props;
    if (c === 0x0b || c === 0x0a || c === 0x09) cur += ' ';
    else if (c !== 0xfeff && c >= 0x20) cur += text[i];
  }
  endParagraph();

  // Drop a single trailing empty paragraph left by a terminating CR.
  if (paras.length > 1 && paragraphText(paras[paras.length - 1]!).length === 0) paras.pop();
  return paras;
}

function toRun(text: string, props: CharRun | undefined): PptRun {
  if (!props) return { text };
  return {
    text,
    ...(props.bold ? { bold: true } : {}),
    ...(props.italic ? { italic: true } : {}),
    ...(props.underline ? { underline: true } : {}),
    ...(props.sizePt ? { sizePt: props.sizePt } : {}),
    ...(props.colorHex ? { colorHex: props.colorHex } : {}),
  };
}

function sameProps(a: CharRun | undefined, b: CharRun | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    a.sizePt === b.sizePt &&
    a.colorHex === b.colorHex
  );
}

// Expand the character runs into a per-character props array, so the paragraph
// splitter can group consecutive equal-property characters into runs.
function expandCharProps(charRuns: ReadonlyArray<CharRun>, textLen: number): Array<CharRun> {
  const out: Array<CharRun> = [];
  for (const run of charRuns) {
    for (let i = 0; i < run.count && out.length < textLen; i++) out.push(run);
  }
  while (out.length < textLen) out.push(charRuns[charRuns.length - 1] ?? { count: 0 });
  return out;
}

// StyleTextPropAtom (§2.9.1): a paragraph-run section then a character-run
// section, each summing to the text length + 1 (the phantom paragraph
// terminator). Within a run, optional fields follow the masks in the spec's byte
// order (NOT bit-ascending); every present field must be read to stay aligned,
// even the ones we drop. We keep the character bold/italic/underline/size/colour
// and the paragraph alignment/indent level.
function parseStyleTextProp(
  data: Uint8Array,
  textLen: number,
): { paraRuns: Array<ParaRun>; charRuns: Array<CharRun> } | undefined {
  const target = textLen + 1;
  let off = 0;

  // --- paragraph runs (§2.9.20 TextPFException) ---
  const paraRuns: Array<ParaRun> = [];
  let consumed = 0;
  while (consumed < target && off + 10 <= data.length) {
    const count = u32(data, off);
    const level = u16(data, off + 4);
    const mask = u32(data, off + 6);
    off += 10;
    let align: number | undefined;
    if ((mask & 0x0000000f) !== 0) off += 2; // bulletFlags
    if ((mask & 0x00000080) !== 0) off += 2; // bulletChar
    if ((mask & 0x00000010) !== 0) off += 2; // bulletFontRef
    if ((mask & 0x00000040) !== 0) off += 2; // bulletSize
    if ((mask & 0x00000020) !== 0) off += 4; // bulletColor
    if ((mask & 0x00000800) !== 0) {
      align = u16(data, off); // textAlignment
      off += 2;
    }
    if ((mask & 0x00001000) !== 0) off += 2; // lineSpacing
    if ((mask & 0x00002000) !== 0) off += 2; // spaceBefore
    if ((mask & 0x00004000) !== 0) off += 2; // spaceAfter
    if ((mask & 0x00000100) !== 0) off += 2; // leftMargin
    if ((mask & 0x00000400) !== 0) off += 2; // indent
    if ((mask & 0x00008000) !== 0) off += 2; // defaultTabSize
    if ((mask & 0x00100000) !== 0) off += 2 + (off + 2 <= data.length ? u16(data, off) : 0) * 4; // tabStops
    if ((mask & 0x00010000) !== 0) off += 2; // fontAlign
    if ((mask & 0x000e0000) !== 0) off += 2; // wrapFlags (charWrap/wordWrap/overflow)
    if ((mask & 0x00200000) !== 0) off += 2; // textDirection
    paraRuns.push({ count, level, ...(align !== undefined ? { align } : {}) });
    consumed += count;
    if (count <= 0) break;
  }

  // --- character runs (§2.9.11 TextCFException) ---
  const charRuns: Array<CharRun> = [];
  consumed = 0;
  while (consumed < target && off + 8 <= data.length) {
    const count = u32(data, off);
    const mask = u32(data, off + 4);
    off += 8;
    const run: CharRun = { count };
    // fontStyle (CFStyle): present if any style bit (bold/italic/underline/
    // shadow/fehint/kumi/emboss) or the fHasStyle group (bits 10–13) is set.
    if ((mask & 0x00003eb7) !== 0) {
      const style = u16(data, off);
      off += 2;
      if ((mask & 0x1) !== 0 && (style & 0x1) !== 0) run.bold = true;
      if ((mask & 0x2) !== 0 && (style & 0x2) !== 0) run.italic = true;
      if ((mask & 0x4) !== 0 && (style & 0x4) !== 0) run.underline = true;
    }
    if ((mask & 0x00010000) !== 0) off += 2; // fontRef
    if ((mask & 0x00200000) !== 0) off += 2; // oldEAFontRef
    if ((mask & 0x00400000) !== 0) off += 2; // ansiFontRef
    if ((mask & 0x00800000) !== 0) off += 2; // symbolFontRef
    if ((mask & 0x00020000) !== 0) {
      const sz = u16(data, off); // fontSize (points)
      off += 2;
      if (sz > 0) run.sizePt = sz;
    }
    if ((mask & 0x00040000) !== 0) {
      // color (ColorIndexStruct: red, green, blue, index); 0xFE = explicit sRGB.
      if (off + 4 <= data.length && data[off + 3] === 0xfe) {
        run.colorHex = hex2(data[off]!) + hex2(data[off + 1]!) + hex2(data[off + 2]!);
      }
      off += 4;
    }
    if ((mask & 0x00080000) !== 0) off += 2; // position
    charRuns.push(run);
    consumed += count;
    if (count <= 0) break;
  }

  return paraRuns.length > 0 || charRuns.length > 0 ? { paraRuns, charRuns } : undefined;
}

function hex2(n: number): string {
  return n.toString(16).padStart(2, '0').toUpperCase();
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
function i16(d: Uint8Array, off: number): number {
  const v = u16(d, off);
  return v >= 0x8000 ? v - 0x10000 : v;
}
