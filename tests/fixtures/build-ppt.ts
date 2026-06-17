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
const RT_PP_DRAWING = 0x040c;

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
}

// A paragraph run in a StyleTextPropAtom, over `length` characters.
export interface PptParaStyleRun {
  readonly length: number;
  readonly align?: number; // TextAlignmentEnum: 0 left, 1 center, 2 right, 3 justify, 4 distribute
  readonly level?: number; // indent level 0–4
}

export interface PptSlideInput {
  // Inline drawing text (a TextCharsAtom, UTF-16), paragraphs split by '\r'.
  readonly text?: string;
  // Inline drawing text as a TextBytesAtom (cp1252) instead of UTF-16.
  readonly textBytes?: string;
  // Outline text in the slide list (the fallback when a slide has no inline text).
  readonly outline?: string;
  // Wrap the inline text in a PPDrawing container, to exercise recursive descent.
  readonly nested?: boolean;
  // A StyleTextPropAtom for the inline text: character runs and/or paragraph runs.
  // Lengths sum to the raw text length; the builder adds the phantom terminator.
  readonly charRuns?: ReadonlyArray<PptStyleRun>;
  readonly paraRuns?: ReadonlyArray<PptParaStyleRun>;
}

export interface BuildPptOptions {
  readonly encrypted?: boolean;
  readonly slideSizeInches?: { readonly w: number; readonly h: number };
  // Drop the Current User stream, to exercise the document-scan fallback.
  readonly omitCurrentUser?: boolean;
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
    const mask = r.align !== undefined ? 0x00000800 : 0; // textAlignment bit
    u32(mask);
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
    if (r.colorHex) mask |= 0x00040000; // color bit
    u32(mask);
    if (hasStyle) u16(style); // fontStyle (CFStyle)
    if (r.sizePt) u16(r.sizePt); // fontSize (points)
    if (r.colorHex) {
      const hex = r.colorHex;
      const rr = parseInt(hex.slice(0, 2), 16);
      const gg = parseInt(hex.slice(2, 4), 16);
      const bb = parseInt(hex.slice(4, 6), 16);
      out.push(rr, gg, bb, 0xfe); // ColorIndexStruct: red, green, blue, index 0xFE = explicit
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
  const slwtParts: Array<Uint8Array> = [];
  slides.forEach((slide, i) => {
    const spa = new Uint8Array(20);
    new DataView(spa.buffer).setUint32(0, slidePersistIds[i]!, true); // persistIdRef
    slwtParts.push(rec(RT_SLIDE_PERSIST_ATOM, 0, false, spa));
    if (slide.outline !== undefined) {
      slwtParts.push(rec(RT_TEXT_CHARS_ATOM, 0, false, encodeUtf16(slide.outline)));
    }
  });
  const slwt = rec(RT_SLIDE_LIST_WITH_TEXT, 0, true, concat(slwtParts));

  const docData = concat([docAtom, slwt]);
  const docRec = rec(RT_DOCUMENT, 0, true, docData);

  // --- one SlideContainer per slide, carrying its inline drawing text --------
  const slideRecs = slides.map((slide) => {
    const block = slideTextBlock(slide);
    if (!block) return rec(RT_SLIDE, 0, true, new Uint8Array(0));
    const inner = slide.nested ? rec(RT_PP_DRAWING, 0, true, block) : block;
    return rec(RT_SLIDE, 0, true, inner);
  });

  // --- assign absolute offsets: [doc][slides...][persistDir][userEdit] -------
  let cursor = 0;
  const docOffset = cursor;
  cursor += docRec.length;
  const slideOffsets = slideRecs.map((r) => {
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
  ev.setUint32(20, 2 + slides.length, true); // persistIdSeed (> all persist ids)
  const editRec = rec(RT_USER_EDIT_ATOM, 0, false, editData);
  const editOffset = cursor;
  cursor += editRec.length;

  const powerpointDocument = concat([docRec, ...slideRecs, persistRec, editRec]);

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
  return buildCfb(streams);
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
