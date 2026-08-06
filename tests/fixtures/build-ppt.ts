// Minimal PowerPoint 97–2003 `.ppt` builder for tests — packs a `PowerPoint
// Document` stream (a DocumentContainer with a slide list, plus a SlideContainer
// per slide) and a `Current User` stream into a CFB container, so the `.ppt`
// reader has a deterministic input without a checked-in binary. Mirrors
// build-doc.ts: it writes exactly the records ppt-text.ts reads back — the record
// headers, the Current User → UserEditAtom → PersistDirectoryAtom indirection
// (with absolute stream offsets), the SlideListWithText with its per-slide
// SlidePersistAtoms, and the TextCharsAtom / TextBytesAtom text atoms — so the
// round trip validates the offset / persist / record logic rather than restating it.

import { buildCfb } from './build-cfb';

// [MS-PPT] record types the builder emits.
const RT_DOCUMENT = 0x03e8;
const RT_DOCUMENT_ATOM = 0x03e9;
const RT_SLIDE_LIST_WITH_TEXT = 0x0ff0;
const RT_SLIDE_PERSIST_ATOM = 0x03f3;
const RT_SLIDE = 0x03ee;
const RT_USER_EDIT_ATOM = 0x0ff5;
const RT_PERSIST_DIRECTORY_ATOM = 0x1772;
const RT_CURRENT_USER_ATOM = 0x0ff6;
const RT_TEXT_CHARS_ATOM = 0x0fa0;
const RT_STYLE_TEXT_PROP_ATOM = 0x0fa1;
const RT_TEXT_BYTES_ATOM = 0x0fa8;
const RT_TEXT_HEADER_ATOM = 0x0f9f;
const RT_OUTLINE_TEXT_REF_ATOM = 0x0f9e;
const RT_TEXT_MASTER_STYLE_ATOM = 0x0fa3;
const RT_PP_DRAWING = 0x040c;
const RT_DRAWING_GROUP = 0x040b;
const RT_SLIDE_ATOM = 0x03ef;
const RT_MAIN_MASTER = 0x03f8;
const RT_ENVIRONMENT = 0x03f2;
const RT_FONT_COLLECTION = 0x07d5;
const RT_FONT_ENTITY_ATOM = 0x0fb7;
const RT_COLOR_SCHEME_ATOM = 0x07f0;
const COLOR_SCHEME_INSTANCE = 1; // slideSchemeColorSchemeAtom (vs scheme-list 6)
const SLIDE_FLAG_MASTER_SCHEME = 0x0002; // slideAtom.slideFlags.fMasterScheme
const SLIDE_FLAG_MASTER_OBJECTS = 0x0001; // slideAtom.slideFlags.fMasterObjects
const SLIDE_FLAG_MASTER_BACKGROUND = 0x0004; // slideAtom.slideFlags.fMasterBackground
const COLORREF_FLAG_SCHEME = 0x08; // OfficeArtCOLORREF fSchemeIndex flags byte

// OfficeArt (Escher) record types for the picture store + picture shapes (PPT-3).
const FBT_DGG_CONTAINER = 0xf000;
const FBT_BSTORE_CONTAINER = 0xf001;
const FBT_BSE = 0xf007;
const FBT_SP_CONTAINER = 0xf004;
const FBT_FSP = 0xf00a;
const FBT_OPT = 0xf00b;
const FBT_CLIENT_TEXTBOX = 0xf00d;
const FBT_CLIENT_ANCHOR = 0xf010;
const FBT_SPGR_CONTAINER = 0xf003;
const FBT_SPGR = 0xf009;
const FBT_CHILD_ANCHOR = 0xf00f;
const PROP_FILL_TYPE = 0x0180;
const PROP_GTEXT_UNICODE_COMPLEX = 0x80c0; // gtextUNICODE with fComplex
const PROP_GTEXT_SIZE = 0x00c3;
const PROP_GTEXT_FONT_COMPLEX = 0x80c5; // gtextFont with fComplex
const PROP_FILL_WIDTH = 0x0189;
const PROP_FILL_HEIGHT = 0x018a;
const PROP_FILL_BLIP_COMPLEX = 0x8186; // fillBlip with fComplex: the blip follows
const FBT_BLIP_PNG_INLINE = 0xf01e;
const FBT_CLIENT_DATA = 0xf011;
const RT_PLACEHOLDER_ATOM = 0x0bc3;
const FBT_BLIP_PNG = 0xf01e;
const PROP_PIB_ID = 0x4104; // OPT property id: pib (0x0104) with the fBid flag (0x4000)
const PROP_FILL_COLOR = 0x0181;
const PROP_LINE_COLOR = 0x01c0;
const PROP_FILL_BOOLS = 0x01bf; // fill style booleans (fFilled + its usage bit)
const PROP_LINE_BOOLS = 0x01ff; // line style booleans (fLine + its usage bit)
const PROP_GEO_LEFT = 0x0140; // geometry bounds (simple LONGs), PPT-7
const PROP_GEO_TOP = 0x0141;
const PROP_GEO_RIGHT = 0x0142;
const PROP_GEO_BOTTOM = 0x0143;
const PROP_VERTICES_COMPLEX = 0x8145; // pVertices (0x0145) with the fComplex flag (0x8000)
const PROP_SEGMENT_INFO_COMPLEX = 0x8146; // pSegmentInfo (0x0146) with fComplex
const FSP_FLAG_BACKGROUND = 0x0400; // OfficeArtFSP.fBackground
const ARRAY_HEADER_BYTES = 6; // IMsoArray: nElems, nElemsAlloc, cbElem
const MASTER_PER_POINT = 8; // 576 master units / inch ÷ 72 points / inch

const TOKEN_UNENCRYPTED = 0xe391c05f;
const TOKEN_ENCRYPTED = 0xf3d1c4df;

// DocumentAtom.slideSize unit: master units, 576 per inch (must match ppt-text.ts).
const MASTER_UNITS_PER_INCH = 576;

// A run of character formatting in a StyleTextPropAtom, over `length` characters.
export interface PptStyleRun {
  readonly length: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly sizePt?: number;
  readonly colorHex?: string; // 6-hex → an explicit-RGB ColorIndexStruct (index 0xFE)
  // A ColorIndexStruct naming a SCHEME slot (index 0–7) instead of an sRGB.
  readonly colorSchemeIndex?: number;
  // An index into the deck's font collection (see BuildPptOptions.fonts).
  readonly fontRef?: number;
}

// A paragraph run in a StyleTextPropAtom, over `length` characters.
export interface PptParaStyleRun {
  readonly length: number;
  readonly align?: number; // TextAlignmentEnum: 0 left, 1 center, 2 right, 3 justify, 4 distribute
  readonly level?: number; // indent level 0–4
  // §2.9.20 bulletFlags.fHasBullet + bulletChar (the character, often a symbol
  // font's — Wingdings 0x6C is the filled circle).
  readonly hasBullet?: boolean;
  readonly bulletChar?: number;
}

export interface PptSlideInput {
  // Inline drawing text (a TextCharsAtom, UTF-16), paragraphs split by '\r'.
  readonly text?: string;
  // Inline drawing text as a TextBytesAtom (cp1252) instead of UTF-16.
  readonly textBytes?: string;
  // Outline text in the slide list (the fallback when a slide has no inline text).
  readonly outline?: string;
  // Several outline texts, each with its TextHeaderAtom text type — what a shape's
  // OutlineTextRefAtom indexes into.
  readonly outlineTexts?: ReadonlyArray<{ readonly textType: number; readonly text: string }>;
  // Wrap the inline text in a PPDrawing container, to exercise recursive descent.
  readonly nested?: boolean;
  // A StyleTextPropAtom for the inline text: character runs and/or paragraph runs.
  // Lengths sum to the raw text length; the builder adds the phantom terminator.
  readonly charRuns?: ReadonlyArray<PptStyleRun>;
  readonly paraRuns?: ReadonlyArray<PptParaStyleRun>;
  // A picture shape referencing the deck image at this 1-based index (pib).
  readonly imageRef?: number;
  // Positioned shapes: each an SpContainer with an OfficeArtClientAnchor (the
  // rectangle, given in points) and a client text box and/or a picture (PPT-4).
  readonly boxes?: ReadonlyArray<PptBoxInput>;
  // Grouped shapes: an SpgrContainer whose first child is the group shape (PPT-15).
  readonly groups?: ReadonlyArray<PptGroupInput>;
  // The slide's own colour scheme (8 × 6-hex RGB) — a SlideSchemeColorSchemeAtom
  // in the SlideContainer, so a shape's scheme-relative colour resolves (PPT-6).
  readonly colorScheme?: ReadonlyArray<string>;
  // Set slideAtom.slideFlags.fMasterScheme so the slide follows the master at
  // masterIndex (default 0) for its scheme instead of its own (PPT-6).
  readonly followMasterScheme?: boolean;
  // Set slideAtom.slideFlags.fMasterObjects, so the master's shapes are drawn
  // under the slide's own (PPT-14).
  readonly followMasterObjects?: boolean;
  // Set slideAtom.slideFlags.fMasterBackground, so the slide shows the master's
  // background rather than the one it holds of its own (PPT-12).
  readonly followMasterBackground?: boolean;
  readonly masterIndex?: number;
}

// A positioned drawing shape for the fixture.
export interface PptBoxInput {
  readonly anchor?: {
    readonly x: number;
    readonly y: number;
    readonly w: number;
    readonly h: number;
  };
  readonly text?: string; // a client text box (UTF-16), paragraphs split by '\r'
  // A client text box holding only an OutlineTextRefAtom at this 0-based index.
  readonly outlineRef?: number;
  readonly imageRef?: number; // a picture shape (1-based pib index)
  readonly shapeType?: number; // an autoshape: the FSP recInstance (MSOSPT)
  readonly fillColorHex?: string; // OPT fillColor (6-hex literal RGB)
  // §2.3.7.43 / §2.3.8.44 — state fFilled / fLine as OFF, so the colours above
  // are stated but not used.
  readonly noFill?: boolean;
  readonly noLine?: boolean;
  // An OfficeArtClientData holding a PlaceholderAtom: on a master this shape is
  // a prototype, not decoration (PPT-14).
  readonly placeholder?: boolean;
  // MS-ODRAW fBackground: the shape states the SLIDE's background, not content.
  readonly background?: boolean;
  readonly lineColorHex?: string; // OPT lineColor (6-hex literal RGB)
  readonly fillSchemeIndex?: number; // OPT fillColor as a scheme index (0–7), PPT-6
  readonly lineSchemeIndex?: number; // OPT lineColor as a scheme index (0–7), PPT-6
  readonly fillSysColor?: number; // OPT fillColor as a Windows system-colour index, PPT-8
  readonly freeform?: PptFreeformInput; // exact custom geometry (PPT-7)
  // A picture fill: MSOFILLTYPE (2 texture / 3 picture) plus the blip carried
  // INLINE as the fillBlip property's complex data (PPT-15).
  readonly pictureFill?: {
    readonly fillType: number;
    readonly png: Uint8Array;
    // MS-ODRAW §2.3.7.11/.12 `fillWidth` / `fillHeight`, in EMU.
    readonly tileEmu?: readonly [number, number];
  };
  // A grouped shape's rectangle, in the enclosing group's coordinate space.
  readonly childAnchor?: readonly [number, number, number, number];
  // MS-ODRAW §2.3.22 WordArt: gtextUNICODE / gtextSize / gtextFont.
  readonly wordArt?: { readonly text: string; readonly sizePt?: number; readonly font?: string };
}

/** A group of shapes: where it sits on the slide, and the space its children use. */
export interface PptGroupInput {
  readonly anchor: {
    readonly x: number;
    readonly y: number;
    readonly w: number;
    readonly h: number;
  };
  readonly box: readonly [number, number, number, number]; // x, y, right, bottom
  readonly boxes: ReadonlyArray<PptBoxInput>;
}

// A freeform shape's geometry: the bounds extent (left/top are 0) plus the raw
// pVertices (i16 POINT pairs) and pSegmentInfo (u16 opcodes) arrays (PPT-7).
export interface PptFreeformInput {
  readonly geoRight: number;
  readonly geoBottom: number;
  readonly vertices: ReadonlyArray<readonly [number, number]>;
  readonly segments: ReadonlyArray<number>;
  // State each array's length without its 6-byte header, as some producers do.
  readonly arrayLenExcludesHeader?: boolean;
}

export interface BuildPptOptions {
  readonly encrypted?: boolean;
  readonly slideSizeInches?: { readonly w: number; readonly h: number };
  // §2.9.30 — the deck's typefaces, indexed by a run's `fontRef` (PPT-19).
  readonly fonts?: ReadonlyArray<string>;
  // Drop the Current User stream, to exercise the document-scan fallback.
  readonly omitCurrentUser?: boolean;
  // Deck images, stored in the Pictures stream and referenced by slide imageRef.
  readonly images?: ReadonlyArray<Uint8Array>;
  // Slide masters, each persisted as a MainMasterContainer with its own colour
  // scheme — referenced by a slide's followMasterScheme + masterIndex (PPT-6).
  readonly masters?: ReadonlyArray<{
    readonly colorScheme: ReadonlyArray<string>;
    // Decoration the master draws on every slide that follows it (PPT-14).
    readonly boxes?: ReadonlyArray<PptBoxInput>;
    // TextMasterStyleAtoms: per text type, the font size of each indent level.
    readonly textStyles?: ReadonlyArray<{
      readonly textType: number;
      readonly sizesPt: ReadonlyArray<number>;
      // The font-collection index every level of this style names (PPT-19).
      readonly fontRef?: number;
      // A bullet CHARACTER per level, with no fHasBullet flag beside it — which
      // is how a master states its outline bullets (PPT-18).
      readonly bulletChars?: ReadonlyArray<number>;
    }>;
  }>;
}

// Build an 8-byte record header + data. A container uses recVer 0xF (low nibble);
// atoms use recVer 0. recInstance occupies the high 12 bits of the first u16.
function rec(type: number, instance: number, isContainer: boolean, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(8 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint16(0, ((instance & 0xfff) << 4) | (isContainer ? 0x0f : 0x00), true);
  dv.setUint16(2, type, true);
  dv.setUint32(4, data.length, true);
  out.set(data, 8);
  return out;
}

function concat(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

// A slide's inline text record (TextChars/TextBytes) followed by its optional
// StyleTextPropAtom, or undefined when the slide carries no inline text.
function slideTextBlock(slide: PptSlideInput): Uint8Array | undefined {
  let textRec: Uint8Array;
  let textLen: number;
  if (slide.textBytes !== undefined) {
    textRec = rec(RT_TEXT_BYTES_ATOM, 0, false, encodeCp1252(slide.textBytes));
    textLen = slide.textBytes.length;
  } else if (slide.text !== undefined) {
    textRec = rec(RT_TEXT_CHARS_ATOM, 0, false, encodeUtf16(slide.text));
    textLen = slide.text.length;
  } else {
    return undefined;
  }
  if (!slide.charRuns && !slide.paraRuns) return textRec;
  const style = rec(
    RT_STYLE_TEXT_PROP_ATOM,
    0,
    false,
    buildStyleTextProp(
      slide.charRuns ?? [{ length: textLen }],
      slide.paraRuns ?? [{ length: textLen }],
    ),
  );
  return concat([textRec, style]);
}

// A StyleTextPropAtom body: the paragraph-run section then the character-run
// section, in the on-disk field order ppt-text.ts reads back. The last run of each
// section gets +1 for the phantom paragraph terminator (matching PowerPoint).
function buildStyleTextProp(
  charRuns: ReadonlyArray<PptStyleRun>,
  paraRuns: ReadonlyArray<PptParaStyleRun>,
): Uint8Array {
  const out: Array<number> = [];
  const u16 = (v: number): void => void out.push(v & 0xff, (v >> 8) & 0xff);
  const u32 = (v: number): void =>
    void out.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >> 24) & 0xff);

  paraRuns.forEach((r, i) => {
    u32(r.length + (i === paraRuns.length - 1 ? 1 : 0)); // count (+ phantom terminator)
    u16(r.level ?? 0); // indentLevel
    const mask =
      (r.align !== undefined ? 0x00000800 : 0) | // textAlignment bit
      (r.hasBullet !== undefined ? 0x00000001 : 0) | // bulletFlags
      (r.bulletChar !== undefined ? 0x00000080 : 0); // bulletChar
    u32(mask);
    // The fields follow the mask in the SPEC's byte order, not bit order.
    if (r.hasBullet !== undefined) u16(r.hasBullet ? 0x0001 : 0);
    if (r.bulletChar !== undefined) u16(r.bulletChar);
    if (r.align !== undefined) u16(r.align);
  });

  charRuns.forEach((r, i) => {
    u32(r.length + (i === charRuns.length - 1 ? 1 : 0)); // count (+ phantom terminator)
    let mask = 0;
    let style = 0;
    if (r.bold) {
      mask |= 0x1;
      style |= 0x1;
    }
    if (r.italic) {
      mask |= 0x2;
      style |= 0x2;
    }
    if (r.underline) {
      mask |= 0x4;
      style |= 0x4;
    }
    const hasStyle = (mask & 0x3eb7) !== 0;
    if (r.sizePt) mask |= 0x00020000; // size bit
    if (r.colorHex || r.colorSchemeIndex !== undefined) mask |= 0x00040000; // color bit
    if (r.fontRef !== undefined) mask |= 0x00010000; // typeface bit
    u32(mask);
    if (hasStyle) u16(style); // fontStyle (CFStyle)
    if (r.fontRef !== undefined) u16(r.fontRef); // fontRef, before the size
    if (r.sizePt) u16(r.sizePt); // fontSize (points)
    if (r.colorHex) {
      const hex = r.colorHex;
      const rr = parseInt(hex.slice(0, 2), 16);
      const gg = parseInt(hex.slice(2, 4), 16);
      const bb = parseInt(hex.slice(4, 6), 16);
      out.push(rr, gg, bb, 0xfe); // ColorIndexStruct: red, green, blue, index 0xFE = explicit
    } else if (r.colorSchemeIndex !== undefined) {
      out.push(0, 0, 0, r.colorSchemeIndex); // index 0–7 = a scheme slot
    }
  });

  return Uint8Array.from(out);
}

export function buildPpt(
  slides: ReadonlyArray<PptSlideInput>,
  opts: BuildPptOptions = {},
): Uint8Array {
  // --- DocumentAtom: slideSize (x, y) in master units, then zero padding -----
  const docAtomData = new Uint8Array(40);
  const dav = new DataView(docAtomData.buffer);
  const inches = opts.slideSizeInches ?? { w: 10, h: 7.5 };
  dav.setInt32(0, Math.round(inches.w * MASTER_UNITS_PER_INCH), true);
  dav.setInt32(4, Math.round(inches.h * MASTER_UNITS_PER_INCH), true);
  const docAtom = rec(RT_DOCUMENT_ATOM, 0, false, docAtomData);

  // --- SlideListWithText (instance 0): per slide, a SlidePersistAtom (its
  //     persist id) followed by the slide's outline text atom (if any) ---------
  const docPersistId = 1;
  const slidePersistIds = slides.map((_, i) => 2 + i);
  const masters = opts.masters ?? [];
  const masterPersistIds = masters.map((_, i) => 2 + slides.length + i);
  const slwtParts: Array<Uint8Array> = [];
  slides.forEach((slide, i) => {
    const spa = new Uint8Array(20);
    new DataView(spa.buffer).setUint32(0, slidePersistIds[i]!, true); // persistIdRef
    slwtParts.push(rec(RT_SLIDE_PERSIST_ATOM, 0, false, spa));
    if (slide.outline !== undefined) {
      slwtParts.push(rec(RT_TEXT_CHARS_ATOM, 0, false, encodeUtf16(slide.outline)));
    }
    for (const entry of slide.outlineTexts ?? []) {
      const header = new Uint8Array(4);
      new DataView(header.buffer).setUint32(0, entry.textType, true);
      slwtParts.push(rec(RT_TEXT_HEADER_ATOM, 0, false, header));
      slwtParts.push(rec(RT_TEXT_CHARS_ATOM, 0, false, encodeUtf16(entry.text)));
    }
  });
  const slwt = rec(RT_SLIDE_LIST_WITH_TEXT, 0, true, concat(slwtParts));

  // --- Pictures stream + DrawingGroup picture store (PPT-3) -------------------
  const images = opts.images ?? [];
  const blips: Array<Uint8Array> = [];
  const foDelays: Array<number> = [];
  let picOff = 0;
  for (const image of images) {
    foDelays.push(picOff);
    // OfficeArtBlipPNG: 16-byte UID + 1 tag byte, then the raw image bytes.
    const blip = rec(
      FBT_BLIP_PNG,
      0x6e0,
      false,
      concat([new Uint8Array(16), Uint8Array.of(0xff), image]),
    );
    blips.push(blip);
    picOff += blip.length;
  }
  const picturesStream = concat(blips);
  const drawingGroup =
    images.length > 0
      ? rec(RT_DRAWING_GROUP, 0, true, rec(FBT_DGG_CONTAINER, 0, true, buildBStore(foDelays)))
      : new Uint8Array(0);

  const docData = concat([docAtom, fontCollection(opts.fonts ?? []), drawingGroup, slwt]);
  const docRec = rec(RT_DOCUMENT, 0, true, docData);

  // --- one SlideContainer per slide, carrying its inline drawing text and any
  //     picture shape (an SpContainer with the pib blip reference) -------------
  const slideRecs = slides.map((slide) => {
    const parts: Array<Uint8Array> = [];
    if (
      slide.followMasterScheme ||
      slide.followMasterObjects ||
      slide.followMasterBackground ||
      slide.masterIndex !== undefined
    ) {
      const flags =
        (slide.followMasterScheme ? SLIDE_FLAG_MASTER_SCHEME : 0) |
        (slide.followMasterObjects ? SLIDE_FLAG_MASTER_OBJECTS : 0) |
        (slide.followMasterBackground ? SLIDE_FLAG_MASTER_BACKGROUND : 0);
      parts.push(slideAtom(masterPersistIds[slide.masterIndex ?? 0] ?? 0, flags));
    }
    const block = slideTextBlock(slide);
    if (block) parts.push(slide.nested ? rec(RT_PP_DRAWING, 0, true, block) : block);
    if (slide.imageRef !== undefined) parts.push(imageShapeContainer(slide.imageRef));
    for (const box of slide.boxes ?? []) parts.push(buildShapeContainer(box));
    for (const group of slide.groups ?? []) parts.push(buildGroupContainer(group));
    if (slide.colorScheme) parts.push(colorSchemeAtom(slide.colorScheme));
    return rec(RT_SLIDE, 0, true, concat(parts));
  });

  // One MainMasterContainer per master, carrying just its colour scheme (the
  // SlideSchemeColorSchemeAtom) — enough for a slide that follows the master.
  const masterRecs = masters.map((m) =>
    rec(
      RT_MAIN_MASTER,
      0,
      true,
      concat([
        colorSchemeAtom(m.colorScheme),
        ...(m.textStyles ?? []).map((st) =>
          textMasterStyleAtom(st.textType, st.sizesPt, st.fontRef, st.bulletChars),
        ),
        ...(m.boxes ?? []).map(buildShapeContainer),
      ]),
    ),
  );

  // --- assign absolute offsets: [doc][slides...][masters...][persistDir][userEdit]
  let cursor = 0;
  const docOffset = cursor;
  cursor += docRec.length;
  const slideOffsets = slideRecs.map((r) => {
    const off = cursor;
    cursor += r.length;
    return off;
  });
  const masterOffsets = masterRecs.map((r) => {
    const off = cursor;
    cursor += r.length;
    return off;
  });

  // PersistDirectoryAtom: one entry (count 1) per persist id → its offset.
  const dirEntries: Array<Uint8Array> = [];
  const addEntry = (persistId: number, offset: number): void => {
    const e = new Uint8Array(8);
    const ev = new DataView(e.buffer);
    ev.setUint32(0, (persistId & 0xfffff) | (1 << 20), true); // persistId + count 1
    ev.setUint32(4, offset, true);
    dirEntries.push(e);
  };
  addEntry(docPersistId, docOffset);
  slidePersistIds.forEach((id, i) => addEntry(id, slideOffsets[i]!));
  masterPersistIds.forEach((id, i) => addEntry(id, masterOffsets[i]!));
  const persistRec = rec(RT_PERSIST_DIRECTORY_ATOM, 0, false, concat(dirEntries));
  const persistOffset = cursor;
  cursor += persistRec.length;

  // UserEditAtom (28-byte data): offsetLastEdit @8, offsetPersistDirectory @12,
  // docPersistIdRef @16, then persistIdSeed / lastView / unused.
  const editData = new Uint8Array(28);
  const ev = new DataView(editData.buffer);
  ev.setUint8(15, 0x03); // majorVersion
  ev.setUint32(8, 0, true); // offsetLastEdit (no prior edit)
  ev.setUint32(12, persistOffset, true); // offsetPersistDirectory
  ev.setUint32(16, docPersistId, true); // docPersistIdRef
  ev.setUint32(20, 2 + slides.length + masters.length, true); // persistIdSeed (> all ids)
  const editRec = rec(RT_USER_EDIT_ATOM, 0, false, editData);
  const editOffset = cursor;
  cursor += editRec.length;

  const powerpointDocument = concat([docRec, ...slideRecs, ...masterRecs, persistRec, editRec]);

  // --- Current User stream: CurrentUserAtom → offsetToCurrentEdit ------------
  const cuData = new Uint8Array(20);
  const cv = new DataView(cuData.buffer);
  cv.setUint32(0, 0x14, true); // size
  cv.setUint32(4, opts.encrypted ? TOKEN_ENCRYPTED : TOKEN_UNENCRYPTED, true); // headerToken
  cv.setUint32(8, editOffset, true); // offsetToCurrentEdit
  cv.setUint16(16, 0x03, true); // docFileVersion fields / major
  const currentUser = rec(RT_CURRENT_USER_ATOM, 0, false, cuData);

  const streams = [{ name: 'PowerPoint Document', data: powerpointDocument }];
  if (!opts.omitCurrentUser) streams.push({ name: 'Current User', data: currentUser });
  if (picturesStream.length > 0) streams.push({ name: 'Pictures', data: picturesStream });
  return buildCfb(streams);
}

// The OfficeArtBStoreContainer: one OfficeArtFBSE per stored image, each carrying
// its foDelay (the blip's byte offset in the Pictures stream) at data offset 28.
function buildBStore(foDelays: ReadonlyArray<number>): Uint8Array {
  const fbses = foDelays.map((foDelay) => {
    const d = new Uint8Array(36);
    new DataView(d.buffer).setUint32(28, foDelay, true); // foDelay
    return rec(FBT_BSE, 2, false, d);
  });
  return rec(FBT_BSTORE_CONTAINER, foDelays.length, true, concat(fbses));
}

// A picture shape: an SpContainer whose OPT table carries the pib blip index.
function imageShapeContainer(pib: number): Uint8Array {
  const opt = new Uint8Array(6);
  const ov = new DataView(opt.buffer);
  ov.setUint16(0, PROP_PIB_ID, true); // property id (pib | fBid)
  ov.setUint32(2, pib, true); // 1-based index into the FBSE store
  return rec(FBT_SP_CONTAINER, 0, true, rec(FBT_OPT, 1, false, opt));
}

// A positioned shape: an SpContainer carrying (in order) an FSP (shape type, for
// an autoshape), an FOPT (pib / fill / line properties), an OfficeArtClientAnchor
// (the rectangle, point coords → master units) and a client text box (PPT-4..5).
// An SpgrContainer: the group's own SpContainer (FSPGR + client anchor) then its
// children, each carrying a ChildAnchor instead of a client anchor (PPT-15).
function buildGroupContainer(group: PptGroupInput): Uint8Array {
  const box = new Uint8Array(16);
  const bv = new DataView(box.buffer);
  group.box.forEach((v, i) => bv.setInt32(i * 4, v, true));
  const head = rec(
    FBT_SP_CONTAINER,
    0,
    true,
    concat([
      rec(FBT_SPGR, 1, false, box),
      rec(FBT_FSP, 0, false, new Uint8Array(8)),
      clientAnchorRec(group.anchor),
    ]),
  );
  return rec(FBT_SPGR_CONTAINER, 0, true, concat([head, ...group.boxes.map(buildShapeContainer)]));
}

// An OfficeArtClientAnchor for a rectangle given in points.
function clientAnchorRec(a: {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}): Uint8Array {
  const d = new Uint8Array(8);
  const v = new DataView(d.buffer);
  v.setInt16(0, Math.round(a.y * MASTER_PER_POINT), true); // top
  v.setInt16(2, Math.round(a.x * MASTER_PER_POINT), true); // left
  v.setInt16(4, Math.round((a.x + a.w) * MASTER_PER_POINT), true); // right
  v.setInt16(6, Math.round((a.y + a.h) * MASTER_PER_POINT), true); // bottom
  return rec(FBT_CLIENT_ANCHOR, 0, false, d);
}

function buildShapeContainer(box: PptBoxInput): Uint8Array {
  const parts: Array<Uint8Array> = [];
  if (box.shapeType !== undefined) {
    const fsp = new Uint8Array(8); // spid + flags
    if (box.background) new DataView(fsp.buffer).setUint32(4, FSP_FLAG_BACKGROUND, true);
    parts.push(rec(FBT_FSP, box.shapeType, false, fsp));
  }
  const props: Array<{ id: number; value: number; blob?: Uint8Array }> = [];
  if (box.imageRef !== undefined) props.push({ id: PROP_PIB_ID, value: box.imageRef });
  if (box.fillColorHex) props.push({ id: PROP_FILL_COLOR, value: rgbColorRef(box.fillColorHex) });
  else if (box.fillSchemeIndex !== undefined)
    props.push({ id: PROP_FILL_COLOR, value: schemeColorRef(box.fillSchemeIndex) });
  else if (box.fillSysColor !== undefined)
    props.push({ id: PROP_FILL_COLOR, value: ((box.fillSysColor & 0xffff) | (0x10 << 24)) >>> 0 });
  if (box.lineColorHex) props.push({ id: PROP_LINE_COLOR, value: rgbColorRef(box.lineColorHex) });
  // The value bit clear, its usage bit set: "this shape states it is not filled".
  if (box.noFill) props.push({ id: PROP_FILL_BOOLS, value: 0x0010 << 16 });
  if (box.noLine) props.push({ id: PROP_LINE_BOOLS, value: 0x0008 << 16 });
  else if (box.lineSchemeIndex !== undefined)
    props.push({ id: PROP_LINE_COLOR, value: schemeColorRef(box.lineSchemeIndex) });
  if (box.wordArt) {
    const str = (t: string): Uint8Array => {
      const d = new Uint8Array((t.length + 1) * 2);
      const v = new DataView(d.buffer);
      for (let i = 0; i < t.length; i++) v.setUint16(i * 2, t.charCodeAt(i), true);
      return d;
    };
    const textBlob = str(box.wordArt.text);
    props.push({ id: PROP_GTEXT_UNICODE_COMPLEX, value: textBlob.length, blob: textBlob });
    if (box.wordArt.sizePt !== undefined) {
      props.push({ id: PROP_GTEXT_SIZE, value: Math.round(box.wordArt.sizePt * 65536) });
    }
    if (box.wordArt.font !== undefined) {
      const fontBlob = str(box.wordArt.font);
      props.push({ id: PROP_GTEXT_FONT_COMPLEX, value: fontBlob.length, blob: fontBlob });
    }
  }
  if (box.pictureFill) {
    props.push({ id: PROP_FILL_TYPE, value: box.pictureFill.fillType });
    if (box.pictureFill.tileEmu) {
      props.push({ id: PROP_FILL_WIDTH, value: box.pictureFill.tileEmu[0] });
      props.push({ id: PROP_FILL_HEIGHT, value: box.pictureFill.tileEmu[1] });
    }
    // The complex data is the OfficeArtBlip record itself: header, UID, tag, PNG.
    const blob = rec(
      FBT_BLIP_PNG_INLINE,
      0x6e0,
      false,
      concat([new Uint8Array(16), Uint8Array.of(0xff), box.pictureFill.png]),
    );
    props.push({ id: PROP_FILL_BLIP_COMPLEX, value: blob.length, blob });
  }
  if (box.freeform) {
    const f = box.freeform;
    // The geometry bounds (simple LONGs; left/top default to 0) then the two
    // complex array properties — their 4-byte value is the trailing blob length.
    props.push({ id: PROP_GEO_LEFT, value: 0 });
    props.push({ id: PROP_GEO_TOP, value: 0 });
    props.push({ id: PROP_GEO_RIGHT, value: f.geoRight });
    props.push({ id: PROP_GEO_BOTTOM, value: f.geoBottom });
    const vBlob = buildVerticesBlob(f.vertices);
    const sBlob = buildSegmentsBlob(f.segments);
    // Some producers state an array property's length WITHOUT its 6-byte header,
    // which shifts every complex blob after it (Apache POI meets the same files).
    const short = f.arrayLenExcludesHeader === true ? ARRAY_HEADER_BYTES : 0;
    props.push({ id: PROP_VERTICES_COMPLEX, value: vBlob.length - short, blob: vBlob });
    props.push({ id: PROP_SEGMENT_INFO_COMPLEX, value: sBlob.length - short, blob: sBlob });
  }
  if (props.length > 0) {
    // The fixed 6-byte entries first, then any complex blobs in entry order.
    const fixed = new Uint8Array(props.length * 6);
    const ov = new DataView(fixed.buffer);
    props.forEach((p, i) => {
      ov.setUint16(i * 6, p.id, true);
      ov.setUint32(i * 6 + 2, p.value, true);
    });
    const blobs = props.filter((p) => p.blob).map((p) => p.blob as Uint8Array);
    parts.push(rec(FBT_OPT, props.length, false, concat([fixed, ...blobs])));
  }
  if (box.childAnchor) {
    const ca = new Uint8Array(16);
    const cv = new DataView(ca.buffer);
    box.childAnchor.forEach((v, i) => cv.setInt32(i * 4, v, true));
    parts.push(rec(FBT_CHILD_ANCHOR, 0, false, ca));
  } else if (box.anchor) {
    parts.push(clientAnchorRec(box.anchor));
  }
  if (box.text !== undefined) {
    parts.push(
      rec(FBT_CLIENT_TEXTBOX, 0, true, rec(RT_TEXT_CHARS_ATOM, 0, false, encodeUtf16(box.text))),
    );
  }
  if (box.placeholder) {
    const ph = new Uint8Array(8); // position (4), placementId, size, unused
    parts.push(rec(FBT_CLIENT_DATA, 0, true, rec(RT_PLACEHOLDER_ATOM, 0, false, ph)));
  }
  if (box.text === undefined && box.outlineRef !== undefined) {
    const idx = new Uint8Array(4);
    new DataView(idx.buffer).setUint32(0, box.outlineRef, true);
    parts.push(rec(FBT_CLIENT_TEXTBOX, 0, true, rec(RT_OUTLINE_TEXT_REF_ATOM, 0, false, idx)));
  }
  return rec(FBT_SP_CONTAINER, 0, true, concat(parts));
}

// An IMsoArray complex property: a 6-byte header (nElems, nElemsAlloc, cbElem)
// then the element bytes (PPT-7).
function buildArrayProp(cbElem: number, count: number, body: Uint8Array): Uint8Array {
  const out = new Uint8Array(6 + body.length);
  const dv = new DataView(out.buffer);
  dv.setUint16(0, count, true);
  dv.setUint16(2, count, true); // nElemsAlloc
  dv.setUint16(4, cbElem, true);
  out.set(body, 6);
  return out;
}

// pVertices: each POINT a pair of i16 (a 4-byte element).
function buildVerticesBlob(verts: ReadonlyArray<readonly [number, number]>): Uint8Array {
  const body = new Uint8Array(verts.length * 4);
  const dv = new DataView(body.buffer);
  verts.forEach(([x, y], i) => {
    dv.setInt16(i * 4, x, true);
    dv.setInt16(i * 4 + 2, y, true);
  });
  return buildArrayProp(4, verts.length, body);
}

// pSegmentInfo: each opcode a u16 (a 2-byte element).
function buildSegmentsBlob(segs: ReadonlyArray<number>): Uint8Array {
  const body = new Uint8Array(segs.length * 2);
  const dv = new DataView(body.buffer);
  segs.forEach((s, i) => dv.setUint16(i * 2, s, true));
  return buildArrayProp(2, segs.length, body);
}

// A 6-hex RGB → an OfficeArtCOLORREF value (red | green<<8 | blue<<16, flags 0 =
// literal sRGB).
function rgbColorRef(hex: string): number {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return (r | (g << 8) | (b << 16)) >>> 0;
}

// A scheme-relative OfficeArtCOLORREF: the red byte is the scheme index (0–7); the
// flags byte sets fSchemeIndex; green/blue are 0 (PPT-6).
function schemeColorRef(index: number): number {
  return ((index & 0xff) | (COLORREF_FLAG_SCHEME << 24)) >>> 0;
}

// The DocumentTextInfo (RT_Environment) holding a FontCollectionContainer: one
// FontEntityAtom per typeface, its lfFaceName 32 UTF-16 units NUL-terminated,
// its recInstance the index a run's `fontRef` names (PPT-19).
function fontCollection(fonts: ReadonlyArray<string>): Uint8Array {
  if (fonts.length === 0) return new Uint8Array(0);
  const atoms = fonts.map((name, i) => {
    const d = new Uint8Array(68);
    const v = new DataView(d.buffer);
    for (let c = 0; c < Math.min(name.length, 31); c++)
      v.setUint16(c * 2, name.charCodeAt(c), true);
    return rec(RT_FONT_ENTITY_ATOM, i, false, d);
  });
  return rec(RT_ENVIRONMENT, 0, true, rec(RT_FONT_COLLECTION, 0, true, concat(atoms)));
}

// A SlideAtom (§2.4.24): a 24-byte body with masterIdRef at offset 12 and
// slideFlags (fMasterScheme) at offset 20, so the slide follows its master's scheme.
function textMasterStyleAtom(
  textType: number,
  sizesPt: ReadonlyArray<number>,
  fontRef?: number,
  bulletChars?: ReadonlyArray<number>,
): Uint8Array {
  const parts: Array<Uint8Array> = [];
  const head = new Uint8Array(2);
  new DataView(head.buffer).setUint16(0, sizesPt.length, true); // cLevels
  parts.push(head);
  for (const size of sizesPt) {
    // A level: its own index (only on the types past `other`), then a
    // TextPFException stating nothing and a TextCFException stating the size.
    const level = new Uint8Array(textType >= 4 ? 2 : 0);
    const bulletChar = bulletChars?.[parts.length === 1 ? 0 : (parts.length - 1) / 3];
    // A TextPFException stating the bullet CHARACTER (mask bit 0x80) and nothing
    // else — no fHasBullet flag, which is what a real master writes.
    const pf = new Uint8Array(bulletChar === undefined ? 4 : 6);
    if (bulletChar !== undefined) {
      const pv = new DataView(pf.buffer);
      pv.setUint32(0, 0x00000080, true);
      pv.setUint16(4, bulletChar, true);
    }
    const cf = new Uint8Array(fontRef === undefined ? 6 : 8);
    const cv = new DataView(cf.buffer);
    // TextCFExceptionMask: size, and the typeface when one is named.
    cv.setUint32(0, fontRef === undefined ? 0x00020000 : 0x00030000, true);
    if (fontRef === undefined) cv.setUint16(4, size, true);
    else {
      cv.setUint16(4, fontRef, true); // fontRef comes before the size
      cv.setUint16(6, size, true);
    }
    parts.push(level, pf, cf);
  }
  return rec(RT_TEXT_MASTER_STYLE_ATOM, textType, false, concat(parts));
}

function slideAtom(masterIdRef: number, flags = SLIDE_FLAG_MASTER_SCHEME): Uint8Array {
  const d = new Uint8Array(24);
  const dv = new DataView(d.buffer);
  dv.setUint32(12, masterIdRef, true); // masterIdRef
  dv.setUint16(20, flags, true); // slideFlags
  return rec(RT_SLIDE_ATOM, 0, false, d);
}

// A SlideSchemeColorSchemeAtom (recInstance 1): 8 ColorStruct entries (red, green,
// blue, unused) parsed from 6-hex RGB.
function colorSchemeAtom(scheme: ReadonlyArray<string>): Uint8Array {
  const body = new Uint8Array(32);
  for (let i = 0; i < 8; i++) {
    const hex = scheme[i] ?? '000000';
    body[i * 4] = parseInt(hex.slice(0, 2), 16);
    body[i * 4 + 1] = parseInt(hex.slice(2, 4), 16);
    body[i * 4 + 2] = parseInt(hex.slice(4, 6), 16);
  }
  return rec(RT_COLOR_SCHEME_ATOM, COLOR_SCHEME_INSTANCE, false, body);
}

function encodeUtf16(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = (c >> 8) & 0xff;
  }
  return out;
}

const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
  0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
  0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

function encodeCp1252(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) out[i] = c;
    else {
      const idx = CP1252_HIGH.indexOf(c);
      out[i] = idx >= 0 ? 0x80 + idx : 0x3f;
    }
  }
  return out;
}
