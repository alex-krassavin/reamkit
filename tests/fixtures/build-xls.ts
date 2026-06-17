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
}): Uint8Array {
  return buildCfb([{ name: 'Workbook', data: buildWorkbookStream(opts) }]);
}

export function buildWorkbookStream(opts: {
  readonly sheets: ReadonlyArray<XlsSheetInput>;
  readonly sst?: ReadonlyArray<string>;
  readonly date1904?: boolean;
  // FONT/FORMAT/XF/PALETTE records (already wrapped) injected into the globals.
  readonly styleRecords?: ReadonlyArray<Uint8Array>;
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
  const styleLen = styleRecords.reduce((n, r) => n + r.length, 0);
  const eof = rec(0x000a, new Uint8Array(0));
  const globalsLen = globalRecs.reduce((n, r) => n + 4 + r.data.length, 0) + styleLen + eof.length;

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
