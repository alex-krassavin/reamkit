// Minimal BIFF8 (.xls) builder for tests — assembles a Workbook record stream
// (globals substream + one substream per sheet, with the BoundSheet8 lbPlyPos
// offsets patched in) and wraps it in a CFB container. Plus low-level record
// builders so a test can hand-craft RK / MULRK / BOOLERR / FORMULA streams.

import { buildCfb } from './build-cfb';

const ascii = (s: string): Uint8Array => Uint8Array.from(s, (c) => c.charCodeAt(0) & 0xff);

// [type:u16][len:u16][data]
export function rec(type: number, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(4 + data.length);
  const v = new DataView(out.buffer);
  v.setUint16(0, type, true);
  v.setUint16(2, data.length, true);
  out.set(data, 4);
  return out;
}

function cellHead(row: number, col: number, ixfe = 0): Uint8Array {
  const d = new Uint8Array(6);
  const v = new DataView(d.buffer);
  v.setUint16(0, row, true);
  v.setUint16(2, col, true);
  v.setUint16(4, ixfe, true); // ixfe (XF / style index)
  return d;
}

export function numberRec(row: number, col: number, value: number, ixfe = 0): Uint8Array {
  const d = new Uint8Array(14);
  d.set(cellHead(row, col, ixfe), 0);
  new DataView(d.buffer).setFloat64(6, value, true);
  return rec(0x0203, d);
}

export function rkRec(row: number, col: number, intValue: number): Uint8Array {
  const d = new Uint8Array(10);
  d.set(cellHead(row, col), 0);
  new DataView(d.buffer).setUint32(6, ((intValue << 2) | 0x02) >>> 0, true);
  return rec(0x027e, d);
}

export function mulRkRec(row: number, colFirst: number, ints: ReadonlyArray<number>): Uint8Array {
  const d = new Uint8Array(4 + ints.length * 6 + 2);
  const v = new DataView(d.buffer);
  v.setUint16(0, row, true);
  v.setUint16(2, colFirst, true);
  for (let k = 0; k < ints.length; k++) {
    v.setUint16(4 + k * 6, 0, true); // ixfe
    v.setUint32(4 + k * 6 + 2, ((ints[k]! << 2) | 0x02) >>> 0, true);
  }
  v.setUint16(4 + ints.length * 6, colFirst + ints.length - 1, true);
  return rec(0x00bd, d);
}

export function labelSstRec(row: number, col: number, isst: number, ixfe = 0): Uint8Array {
  const d = new Uint8Array(10);
  d.set(cellHead(row, col, ixfe), 0);
  new DataView(d.buffer).setUint32(6, isst, true);
  return rec(0x00fd, d);
}

// FONT record (0x0031). bold via weight (700); italic/underline via flags.
export function fontRec(opts: {
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  colorIndex?: number;
  name?: string;
}): Uint8Array {
  const nm = ascii(opts.name ?? 'Arial');
  const d = new Uint8Array(14 + 2 + nm.length);
  const v = new DataView(d.buffer);
  v.setUint16(0, Math.round((opts.sizePt ?? 10) * 20), true); // height in twips
  v.setUint16(2, opts.italic ? 0x0002 : 0, true); // grbit
  v.setUint16(4, opts.colorIndex ?? 0x7fff, true); // icv (0x7FFF = auto)
  v.setUint16(6, opts.bold ? 700 : 400, true); // bls
  d[10] = opts.underline ? 1 : 0; // uls
  d[14] = opts.name ? opts.name.length : 4; // cch
  d[15] = 0; // compressed
  d.set(nm, 16);
  return rec(0x0031, d);
}

// FORMAT record (0x041E) — a custom number-format code at index `ifmt`.
export function formatRec(ifmt: number, code: string): Uint8Array {
  const c = ascii(code);
  const d = new Uint8Array(2 + 3 + c.length);
  const v = new DataView(d.buffer);
  v.setUint16(0, ifmt, true);
  v.setUint16(2, code.length, true);
  d[4] = 0; // compressed
  d.set(c, 5);
  return rec(0x041e, d);
}

export interface XfBorder {
  readonly style: number; // dg line style 0..13
  readonly colorIndex?: number;
}

// XF record (0x00E0, 20-byte BIFF8 cell XF) — references a font + number format,
// with an inline fill (pattern fls + fg/bg palette indices) and per-edge borders.
export function xfRec(opts: {
  fontId?: number;
  numFmtId?: number;
  halign?: number;
  valign?: number;
  wrap?: boolean;
  rotation?: number;
  indent?: number;
  fill?: { pattern: number; fg?: number; bg?: number };
  left?: XfBorder;
  right?: XfBorder;
  top?: XfBorder;
  bottom?: XfBorder;
}): Uint8Array {
  const d = new Uint8Array(20);
  const v = new DataView(d.buffer);
  v.setUint16(0, opts.fontId ?? 0, true);
  v.setUint16(2, opts.numFmtId ?? 0, true);
  v.setUint16(4, 0, true); // attrs (cell XF)
  d[6] = (opts.halign ?? 0) | (opts.wrap ? 0x08 : 0) | ((opts.valign ?? 2) << 4);
  d[7] = opts.rotation ?? 0;
  d[8] = opts.indent ?? 0;
  const brd1 =
    (opts.left?.style ?? 0) |
    ((opts.right?.style ?? 0) << 4) |
    ((opts.top?.style ?? 0) << 8) |
    ((opts.bottom?.style ?? 0) << 12);
  v.setUint16(10, brd1, true);
  const brd2 =
    ((opts.left?.colorIndex ?? 0) & 0x7f) |
    (((opts.right?.colorIndex ?? 0) & 0x7f) << 7) |
    (((opts.top?.colorIndex ?? 0) & 0x7f) << 16) |
    (((opts.bottom?.colorIndex ?? 0) & 0x7f) << 23);
  v.setUint32(12, brd2 >>> 0, true);
  const f = opts.fill;
  const brd3 = f
    ? ((f.fg ?? 0) & 0x7f) | (((f.bg ?? 0) & 0x7f) << 7) | ((f.pattern & 0x3f) << 26)
    : 0;
  v.setUint32(16, brd3 >>> 0, true);
  return rec(0x00e0, d);
}

// PALETTE record (0x0092) — override colours from index 8; each is R,G,B,0.
export function paletteRec(colors: ReadonlyArray<[number, number, number]>): Uint8Array {
  const d = new Uint8Array(2 + colors.length * 4);
  const v = new DataView(d.buffer);
  v.setUint16(0, colors.length, true);
  colors.forEach(([r, g, b], i) => {
    const o = 2 + i * 4;
    d[o] = r;
    d[o + 1] = g;
    d[o + 2] = b;
  });
  return rec(0x0092, d);
}

export function boolRec(row: number, col: number, value: boolean): Uint8Array {
  const d = new Uint8Array(8);
  d.set(cellHead(row, col), 0);
  d[6] = value ? 1 : 0;
  d[7] = 0; // fError = 0
  return rec(0x0205, d);
}

export function errRec(row: number, col: number, code: number): Uint8Array {
  const d = new Uint8Array(8);
  d.set(cellHead(row, col), 0);
  d[6] = code;
  d[7] = 1; // fError = 1
  return rec(0x0205, d);
}

export function formulaNumberRec(row: number, col: number, value: number): Uint8Array {
  const d = new Uint8Array(22);
  d.set(cellHead(row, col), 0);
  new DataView(d.buffer).setFloat64(6, value, true); // cached double (bytes 12-13 ≠ 0xFFFF)
  return rec(0x0006, d);
}

export function formulaStringRecs(row: number, col: number, text: string): Array<Uint8Array> {
  const d = new Uint8Array(22);
  d.set(cellHead(row, col), 0);
  d[6] = 0; // result type: string
  d[12] = 0xff;
  d[13] = 0xff; // the 0xFFFF marker
  const s = ascii(text);
  const str = new Uint8Array(3 + s.length);
  const sv = new DataView(str.buffer);
  sv.setUint16(0, text.length, true);
  str[2] = 0; // flags: compressed
  str.set(s, 3);
  return [rec(0x0006, d), rec(0x0207, str)];
}

export function mergeCellsRec(
  ranges: ReadonlyArray<{ r0: number; r1: number; c0: number; c1: number }>,
): Uint8Array {
  const d = new Uint8Array(2 + ranges.length * 8);
  const v = new DataView(d.buffer);
  v.setUint16(0, ranges.length, true);
  ranges.forEach((r, k) => {
    const o = 2 + k * 8;
    v.setUint16(o, r.r0, true);
    v.setUint16(o + 2, r.r1, true);
    v.setUint16(o + 4, r.c0, true);
    v.setUint16(o + 6, r.c1, true);
  });
  return rec(0x00e5, d);
}

// The StdHlink and URLMoniker CLSIDs (on-disk LE bytes).
const STD_HLINK_CLSID = [
  0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const URL_MONIKER_CLSID = [
  0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];

// An HLink record (0x01B8): an external URL hyperlink over a cell range — a Ref8U
// range, the StdHlink CLSID, then a Hyperlink Object (streamVersion 2, flags
// HasMoniker|IsAbsolute) carrying a URLMoniker with a NUL-terminated UTF-16LE URL.
export function hlinkRec(opts: {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstCol: number;
  readonly lastCol: number;
  readonly url: string;
}): Uint8Array {
  const ref8 = new Uint8Array(8);
  const rv = new DataView(ref8.buffer);
  rv.setUint16(0, opts.firstRow, true);
  rv.setUint16(2, opts.lastRow, true);
  rv.setUint16(4, opts.firstCol, true);
  rv.setUint16(6, opts.lastCol, true);

  const head = new Uint8Array(8);
  const hv = new DataView(head.buffer);
  hv.setUint32(0, 2, true); // streamVersion
  hv.setUint32(4, 0x03, true); // hlinkFlags: HasMoniker | IsAbsolute

  // URLMoniker data: a u32 byte length then a NUL-terminated UTF-16LE URL.
  const urlBytes = new Uint8Array((opts.url.length + 1) * 2); // + the NUL terminator
  for (let i = 0; i < opts.url.length; i++) {
    const c = opts.url.charCodeAt(i);
    urlBytes[i * 2] = c & 0xff;
    urlBytes[i * 2 + 1] = (c >> 8) & 0xff;
  }
  const urlLen = new Uint8Array(4);
  new DataView(urlLen.buffer).setUint32(0, urlBytes.length, true);

  return rec(
    0x01b8,
    concat([
      ref8,
      Uint8Array.from(STD_HLINK_CLSID),
      head,
      Uint8Array.from(URL_MONIKER_CLSID),
      urlLen,
      urlBytes,
    ]),
  );
}

// === Print model records (XLS-9) ============================================

// Setup (0x00A1): paper size, scale, fit-to-N, orientation (fPortrait 0x02 set
// unless landscape; fNoPls/fNoOrient clear so it is honoured) and header/footer
// margins (numHdr @16, numFtr @24 as doubles).
export function setupRec(opts: {
  readonly paperSize?: number;
  readonly scale?: number;
  readonly fitWidth?: number;
  readonly fitHeight?: number;
  readonly landscape?: boolean;
  readonly headerMarginInches?: number;
  readonly footerMarginInches?: number;
}): Uint8Array {
  const d = new Uint8Array(34);
  const v = new DataView(d.buffer);
  v.setUint16(0, opts.paperSize ?? 1, true);
  v.setUint16(2, opts.scale ?? 100, true);
  v.setUint16(6, opts.fitWidth ?? 1, true);
  v.setUint16(8, opts.fitHeight ?? 1, true);
  v.setUint16(10, opts.landscape ? 0x0000 : 0x0002, true); // grbit
  v.setFloat64(16, opts.headerMarginInches ?? 0.3, true);
  v.setFloat64(24, opts.footerMarginInches ?? 0.3, true);
  return rec(0x00a1, d);
}

// WsBool (0x0081): fFitToPage at 0x0100.
export function wsBoolRec(fitToPage: boolean): Uint8Array {
  const d = new Uint8Array(2);
  new DataView(d.buffer).setUint16(0, fitToPage ? 0x0100 : 0x0000, true);
  return rec(0x0081, d);
}

function boolU16Rec(type: number, on: boolean): Uint8Array {
  const d = new Uint8Array(2);
  new DataView(d.buffer).setUint16(0, on ? 1 : 0, true);
  return rec(type, d);
}
export const printGridlinesRec = (on: boolean): Uint8Array => boolU16Rec(0x002b, on);
export const hCenterRec = (on: boolean): Uint8Array => boolU16Rec(0x0083, on);
export const vCenterRec = (on: boolean): Uint8Array => boolU16Rec(0x0084, on);

function marginRec(type: number, inches: number): Uint8Array {
  const d = new Uint8Array(8);
  new DataView(d.buffer).setFloat64(0, inches, true);
  return rec(type, d);
}
export const leftMarginRec = (inches: number): Uint8Array => marginRec(0x0026, inches);
export const rightMarginRec = (inches: number): Uint8Array => marginRec(0x0027, inches);
export const topMarginRec = (inches: number): Uint8Array => marginRec(0x0028, inches);
export const bottomMarginRec = (inches: number): Uint8Array => marginRec(0x0029, inches);

// Header (0x0014) / Footer (0x0015): an XLUnicodeString (cch u16, flags u8, chars);
// an empty string emits a zero-length record (= no header/footer).
function hfRec(type: number, text: string): Uint8Array {
  if (text.length === 0) return rec(type, new Uint8Array(0));
  const d = new Uint8Array(3 + text.length);
  new DataView(d.buffer).setUint16(0, text.length, true);
  d[2] = 0; // flags: 8-bit chars
  for (let i = 0; i < text.length; i++) d[3 + i] = text.charCodeAt(i) & 0xff;
  return rec(type, d);
}
export const headerRec = (text: string): Uint8Array => hfRec(0x0014, text);
export const footerRec = (text: string): Uint8Array => hfRec(0x0015, text);

// HPageBreak (0x001B) / VPageBreak (0x001A): a u16 count then count×6-byte entries
// whose first u16 is the row/column that begins the next page.
function pageBreakRec(type: number, indices: ReadonlyArray<number>): Uint8Array {
  const d = new Uint8Array(2 + indices.length * 6);
  const v = new DataView(d.buffer);
  v.setUint16(0, indices.length, true);
  indices.forEach((idx, k) => v.setUint16(2 + k * 6, idx, true));
  return rec(type, d);
}
export const hPageBreakRec = (rows: ReadonlyArray<number>): Uint8Array =>
  pageBreakRec(0x001b, rows);
export const vPageBreakRec = (cols: ReadonlyArray<number>): Uint8Array =>
  pageBreakRec(0x001a, cols);

// NAME / Lbl (0x0018): a workbook defined name (XLS-10). A built-in (Print_Area
// 0x06 / Print_Titles 0x07) stores its id as the single name char; a regular name
// stores the ascii string. The rgce is one or more PtgArea3d (0x3B) tokens,
// optionally joined by tUnion (0x10) for a Print_Titles rows+cols pair.
export function nameRec(opts: {
  readonly builtinId?: number;
  readonly name?: string;
  readonly itab?: number; // 1-based sheet scope (0 = global)
  readonly areas: ReadonlyArray<{
    readonly r0: number;
    readonly r1: number;
    readonly c0: number;
    readonly c1: number;
    readonly ixti?: number;
  }>;
  readonly union?: boolean;
}): Uint8Array {
  const ptgs: Array<Uint8Array> = [];
  opts.areas.forEach((a, i) => {
    const p = new Uint8Array(11);
    const v = new DataView(p.buffer);
    p[0] = 0x3b; // PtgArea3d (ref class)
    v.setUint16(1, a.ixti ?? 0, true); // ixti
    v.setUint16(3, a.r0, true);
    v.setUint16(5, a.r1, true);
    v.setUint16(7, a.c0, true);
    v.setUint16(9, a.c1, true);
    ptgs.push(p);
    if (opts.union && i > 0) ptgs.push(Uint8Array.of(0x10)); // tUnion
  });
  const rgce = concat(ptgs);

  const nameChars =
    opts.builtinId !== undefined ? Uint8Array.of(opts.builtinId) : ascii(opts.name ?? 'Name');
  const cch = opts.builtinId !== undefined ? 1 : (opts.name ?? 'Name').length;

  const head = new Uint8Array(14);
  const hv = new DataView(head.buffer);
  hv.setUint16(0, opts.builtinId !== undefined ? 0x0020 : 0x0000, true); // grbit (fBuiltin)
  head[3] = cch;
  hv.setUint16(4, rgce.length, true); // cce
  hv.setUint16(8, opts.itab ?? 0, true); // itab (sheet scope)
  // Name = XLUnicodeStringNoCch: flags(1) = 0 (8-bit) then the chars.
  return rec(0x0018, concat([head, Uint8Array.of(0x00), nameChars, rgce]));
}

function sstData(strings: ReadonlyArray<string>): Uint8Array {
  const parts: Array<Uint8Array> = [];
  const head = new Uint8Array(8);
  const hv = new DataView(head.buffer);
  hv.setUint32(0, strings.length, true); // cstTotal
  hv.setUint32(4, strings.length, true); // cstUnique
  parts.push(head);
  for (const s of strings) {
    const bytes = ascii(s);
    const sd = new Uint8Array(3 + bytes.length);
    new DataView(sd.buffer).setUint16(0, s.length, true);
    sd[2] = 0; // compressed
    sd.set(bytes, 3);
    parts.push(sd);
  }
  return concat(parts);
}

function bofData(dt: number): Uint8Array {
  const d = new Uint8Array(16);
  const v = new DataView(d.buffer);
  v.setUint16(0, 0x0600, true); // BIFF8
  v.setUint16(2, dt, true); // 0x0005 globals, 0x0010 worksheet
  return d;
}

function dateModeData(date1904: boolean): Uint8Array {
  const d = new Uint8Array(2);
  new DataView(d.buffer).setUint16(0, date1904 ? 1 : 0, true);
  return d;
}

function boundSheetData(name: string): Uint8Array {
  const nm = ascii(name);
  const d = new Uint8Array(6 + 2 + nm.length);
  const v = new DataView(d.buffer);
  v.setUint32(0, 0, true); // lbPlyPos — patched later
  d[4] = 0; // visibility
  d[5] = 0; // dt = worksheet
  d[6] = name.length; // cch
  d[7] = 0; // flags: compressed
  d.set(nm, 8);
  return d; // raw record DATA wrapped below so we can patch in place
}

// --- Office Drawing (Escher) builders for embedded images (XLS-5) ------------

// `[verInstance:u16][type:u16][len:u32][data]`. A container sets ver = 0xF.
function escher(ver: number, instance: number, type: number, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(8 + data.length);
  const v = new DataView(out.buffer);
  v.setUint16(0, ((instance << 4) | (ver & 0x0f)) & 0xffff, true);
  v.setUint16(2, type, true);
  v.setUint32(4, data.length, true);
  out.set(data, 8);
  return out;
}

// MSODrawingGroup BIFF record (0x00EB) — a DggContainer → BStoreContainer with
// one BSE per image. The image bytes sit after a filler prefix; the reader
// magic-scans for them.
export function msoDrawingGroupRec(images: ReadonlyArray<Uint8Array>): Uint8Array {
  const bses = images.map((png) => {
    const d = new Uint8Array(44 + png.length); // 44-byte BSE/blip header filler
    d.set(png, 44);
    return escher(2, 0, 0xf007, d); // BSE
  });
  const bstore = escher(0xf, images.length, 0xf001, concat(bses)); // BStoreContainer
  return rec(0x00eb, escher(0xf, 0, 0xf000, bstore)); // DggContainer
}

// MSODrawing BIFF record (0x00EC) — a SpgrContainer of picture SpContainers,
// each with an OPT `pib` (BLIP index) and a cell anchor.
export function msoDrawingRec(
  pics: ReadonlyArray<{
    blipIndex: number;
    col1?: number;
    row1?: number;
    col2?: number;
    row2?: number;
  }>,
): Uint8Array {
  const sps = pics.map((p) => {
    const opt = new Uint8Array(6);
    const ov = new DataView(opt.buffer);
    ov.setUint16(0, 0x4104, true); // pib | fBid
    ov.setUint32(2, p.blipIndex, true);
    const optRec = escher(3, 1, 0xf00b, opt); // OPT, one property

    const anc = new Uint8Array(18);
    const av = new DataView(anc.buffer);
    av.setUint16(2, p.col1 ?? 0, true);
    av.setUint16(6, p.row1 ?? 0, true);
    av.setUint16(10, p.col2 ?? 2, true);
    av.setUint16(14, p.row2 ?? 3, true);
    const ancRec = escher(0, 0, 0xf010, anc); // ClientAnchor

    return escher(0xf, 0, 0xf004, concat([optRec, ancRec])); // SpContainer
  });
  return rec(0x00ec, escher(0xf, 0, 0xf003, concat(sps))); // SpgrContainer
}

// --- BIFF chart substream builders (XLS-6) -----------------------------------

interface ChartRange {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

// An AI record (0x1051) whose formula is a ptgArea3d over `range`.
function aiRec(id: number, range: ChartRange): Uint8Array {
  const rgce = new Uint8Array(11);
  const rv = new DataView(rgce.buffer);
  rgce[0] = 0x3b; // ptgArea3d
  rv.setUint16(1, 0, true); // ixti
  rv.setUint16(3, range.r0, true);
  rv.setUint16(5, range.r1, true);
  rv.setUint16(7, range.c0, true);
  rv.setUint16(9, range.c1, true);
  const d = new Uint8Array(8 + rgce.length);
  const dv = new DataView(d.buffer);
  d[0] = id; // 0 = name, 1 = values, 2 = categories
  d[1] = 2; // rt = worksheet reference
  dv.setUint16(6, rgce.length, true); // cce
  d.set(rgce, 8);
  return rec(0x1051, d);
}

function seriesTextRec(name: string): Uint8Array {
  const chars = ascii(name);
  const d = new Uint8Array(4 + chars.length);
  d[2] = name.length; // cch
  d.set(chars, 4);
  return rec(0x100d, d);
}

// A nested chart substream (BOF dt=0x20 … EOF) with one series, to embed in a
// sheet's records. The values / categories reference worksheet cells.
export function chartRecords(opts: {
  kind?: 'bar' | 'line' | 'pie' | 'area' | 'scatter';
  values: ChartRange;
  categories?: ChartRange;
  name?: string;
}): Array<Uint8Array> {
  const bof = new Uint8Array(16);
  const bv = new DataView(bof.buffer);
  bv.setUint16(0, 0x0600, true);
  bv.setUint16(2, 0x0020, true); // dt = chart
  const typeCode = {
    bar: 0x1017,
    line: 0x1018,
    pie: 0x1019,
    area: 0x101a,
    scatter: 0x101b,
  }[opts.kind ?? 'bar'];
  const out: Array<Uint8Array> = [
    rec(0x0809, bof),
    rec(typeCode, new Uint8Array(6)), // chart-type group
    rec(0x1003, new Uint8Array(12)), // SERIES
    aiRec(1, opts.values),
  ];
  if (opts.categories) out.push(aiRec(2, opts.categories));
  if (opts.name) out.push(seriesTextRec(opts.name));
  out.push(rec(0x000a, new Uint8Array(0))); // chart EOF
  return out;
}

// MSODrawing (0x00EC) with non-picture shapes — each an SpContainer carrying an
// Sp record (instance = shape type), an optional fill colour, a cell anchor and an
// optional ClientTextbox marker (XLS-7).
export function msoDrawingShapesRec(
  shapes: ReadonlyArray<{
    shapeType: number;
    hasText?: boolean;
    fillRgb?: [number, number, number];
  }>,
): Uint8Array {
  const sps = shapes.map((s) => {
    const children: Array<Uint8Array> = [escher(2, s.shapeType, 0xf00a, new Uint8Array(8))]; // Sp
    if (s.fillRgb) {
      const d = new Uint8Array(6);
      const v = new DataView(d.buffer);
      v.setUint16(0, 0x0181, true); // fillColor property id
      v.setUint32(2, (s.fillRgb[0] | (s.fillRgb[1] << 8) | (s.fillRgb[2] << 16)) >>> 0, true);
      children.push(escher(3, 1, 0xf00b, d)); // OPT, 1 property
    }
    const anc = new Uint8Array(18);
    const av = new DataView(anc.buffer);
    av.setUint16(10, 2, true); // col2
    av.setUint16(14, 3, true); // row2
    children.push(escher(0, 0, 0xf010, anc)); // ClientAnchor
    if (s.hasText) children.push(escher(0, 0, 0xf00d, new Uint8Array(0))); // ClientTextbox
    return escher(0xf, 0, 0xf004, concat(children)); // SpContainer
  });
  return rec(0x00ec, escher(0xf, 0, 0xf003, concat(sps))); // SpgrContainer
}

// A TXO (0x01B6) text-object record + its CONTINUE record carrying the text.
export function txoRecs(text: string): Array<Uint8Array> {
  const td = new Uint8Array(18);
  new DataView(td.buffer).setUint16(10, text.length, true); // cchText
  const chars = ascii(text);
  const cont = new Uint8Array(1 + chars.length); // [grbit=compressed][chars]
  cont.set(chars, 1);
  return [rec(0x01b6, td), rec(0x003c, cont)];
}

export interface XlsSheetInput {
  readonly name: string;
  readonly records: ReadonlyArray<Uint8Array>;
}

// Assemble the full Workbook stream and wrap it in a CFB → `.xls` bytes.
export function buildXls(opts: {
  readonly sheets: ReadonlyArray<XlsSheetInput>;
  readonly sst?: ReadonlyArray<string>;
  readonly date1904?: boolean;
  readonly styleRecords?: ReadonlyArray<Uint8Array>;
  readonly globalRecords?: ReadonlyArray<Uint8Array>;
}): Uint8Array {
  return buildCfb([{ name: 'Workbook', data: buildWorkbookStream(opts) }]);
}

export function buildWorkbookStream(opts: {
  readonly sheets: ReadonlyArray<XlsSheetInput>;
  readonly sst?: ReadonlyArray<string>;
  readonly date1904?: boolean;
  // FONT/FORMAT/XF/PALETTE records (already wrapped) injected into the globals.
  readonly styleRecords?: ReadonlyArray<Uint8Array>;
  // Other already-wrapped workbook-global records (e.g. NAME / Lbl) — appended to
  // the globals substream after the style records, before EOF.
  readonly globalRecords?: ReadonlyArray<Uint8Array>;
}): Uint8Array {
  // BoundSheet8 data is kept patchable; lbPlyPos is filled in once sheet sizes
  // are known. Records are wrapped (length-prefixed) only AFTER patching, since
  // rec() copies the data. The EOF and the (pre-wrapped) style records bracket
  // the unwrapped globals records.
  const boundData = opts.sheets.map((s) => boundSheetData(s.name));
  const globalRecs: Array<{ type: number; data: Uint8Array }> = [
    { type: 0x0809, data: bofData(0x0005) },
  ];
  if (opts.date1904 !== undefined) {
    globalRecs.push({ type: 0x0022, data: dateModeData(opts.date1904) });
  }
  for (const bd of boundData) globalRecs.push({ type: 0x0085, data: bd });
  if (opts.sst && opts.sst.length > 0) globalRecs.push({ type: 0x00fc, data: sstData(opts.sst) });

  const styleRecords = opts.styleRecords ?? [];
  const globalRecords = opts.globalRecords ?? [];
  const extraLen = [...styleRecords, ...globalRecords].reduce((n, r) => n + r.length, 0);
  const eof = rec(0x000a, new Uint8Array(0));
  const globalsLen = globalRecs.reduce((n, r) => n + 4 + r.data.length, 0) + extraLen + eof.length;

  const sheetStreams = opts.sheets.map((s) =>
    concat([rec(0x0809, bofData(0x0010)), ...s.records, rec(0x000a, new Uint8Array(0))]),
  );

  // Patch each BoundSheet8's lbPlyPos to its sheet substream's byte offset.
  let off = globalsLen;
  for (let i = 0; i < boundData.length; i++) {
    const bd = boundData[i]!;
    new DataView(bd.buffer, bd.byteOffset, bd.byteLength).setUint32(0, off, true);
    off += sheetStreams[i]!.length;
  }

  const globalsBytes = concat([
    ...globalRecs.map((r) => rec(r.type, r.data)),
    ...styleRecords,
    ...globalRecords,
    eof,
  ]);
  return concat([globalsBytes, ...sheetStreams]);
}

function concat(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}
