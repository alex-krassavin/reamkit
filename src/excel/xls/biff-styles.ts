// BIFF8 styling (XLS-4) — the FONT / FORMAT / PALETTE / XF records in a `.xls`
// workbook globals, mapped onto the SAME XlsxStyles model the OOXML path uses, so
// the print model renders a legacy workbook's fonts, fills, borders, number
// formats and alignment exactly as it does for xlsx. Unlike xlsx (which has
// separate fill/border tables), BIFF packs each cell's fill and border inline in
// its XF record, and colours are palette indices — so we resolve indices through
// the default 56-colour palette (plus any PALETTE override) and de-duplicate the
// inline fills/borders into the model's shared arrays.

import type {
  XlsxBorder,
  XlsxBorderStyleName,
  XlsxCellAlignment,
  XlsxCellXf,
  XlsxFill,
  XlsxFont,
  XlsxHorizontalAlign,
  XlsxStyles,
  XlsxVerticalAlign,
} from '@/core/spreadsheet-model';

interface BiffRec {
  readonly type: number;
  readonly data: Uint8Array;
}

const REC_FONT = 0x0031;
const REC_FORMAT = 0x041e;
const REC_XF = 0x00e0;
const REC_PALETTE = 0x0092;

// §2.5.178 fls fill patterns → ST_PatternType names (the order matches, and the
// names line up with the print model's PATTERN_DENSITY table).
const FILL_PATTERNS = [
  'none',
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
] as const;

// §2.5.161 border line styles (dg) → ST_BorderStyle names.
const BORDER_STYLES: ReadonlyArray<XlsxBorderStyleName> = [
  'none',
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];

const HALIGN: ReadonlyArray<XlsxHorizontalAlign | undefined> = [
  undefined, // general
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
];
const VALIGN: ReadonlyArray<XlsxVerticalAlign | undefined> = [
  'top',
  'center',
  'bottom',
  'justify',
  'distributed',
];

// The 56-colour default BIFF8 palette (indices 8–63). Indices 0–7 mirror 8–15.
const PALETTE56 = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
];

/**
 * Build a palette-index → `"FFRRGGBB"` (8-hex ARGB, like the xlsx `rgb` attr)
 * resolver from the default 56-colour palette plus any `PALETTE` override.
 * `0x7FFF` / 64 / 65 are automatic or system colours → `undefined` (the renderer
 * default). Exported so the conditional-format reader (XLS-13) resolves its dxf
 * colours the same way.
 *
 * @param records The workbook-globals records (scanned for a `PALETTE` override).
 * @returns A function mapping a colour index (icv) to its hex, or `undefined`.
 */
export function buildBiffPalette(
  records: ReadonlyArray<BiffRec>,
): (icv: number) => string | undefined {
  const palette: Array<string> = [...PALETTE56.slice(0, 8), ...PALETTE56];
  for (const rec of records) if (rec.type === REC_PALETTE) applyPalette(rec.data, palette);
  return (icv: number): string | undefined => {
    if (icv === 0x7fff || icv === 64 || icv === 65) return undefined;
    const rgb = palette[icv];
    return rgb ? `FF${rgb}` : undefined;
  };
}

/**
 * Read a `.xls` workbook globals' FONT / FORMAT / PALETTE / XF records (XLS-4)
 * into the shared `XlsxStyles` model the OOXML path uses. BIFF packs each XF's
 * fill and border inline and stores colours as palette indices; these are
 * resolved through {@link buildBiffPalette} and de-duplicated into the model's
 * shared fill/border arrays.
 *
 * @param records The workbook-globals record stream.
 * @returns The style table (numFmts, fonts, fills, borders, cellXfs).
 */
export function parseBiffStyles(records: ReadonlyArray<BiffRec>): XlsxStyles {
  const color = buildBiffPalette(records);

  // FONT records — index 4 is skipped in BIFF, so the 5th record is font id 5.
  const fonts: Array<XlsxFont> = [];
  let fontIdx = 0;
  const numFmts = new Map<number, string>();

  // Inline fills/borders de-duplicated into shared arrays; ids 0/1 reserved so the
  // print model's "fillId 0=none, 1=gray125" convention holds.
  const fills: Array<XlsxFill> = [{ patternType: 'none' }, { patternType: 'gray125' }];
  const fillIds = new Map<string, number>();
  const borders: Array<XlsxBorder> = [{}];
  const borderIds = new Map<string, number>();
  const cellXfs: Array<XlsxCellXf> = [];

  const internFill = (fill: XlsxFill | undefined): number => {
    if (!fill || !fill.patternType || fill.patternType === 'none') return 0;
    const key = `${fill.patternType}|${fill.fgColorHex ?? ''}|${fill.bgColorHex ?? ''}`;
    let id = fillIds.get(key);
    if (id === undefined) {
      id = fills.length;
      fills.push(fill);
      fillIds.set(key, id);
    }
    return id;
  };
  const internBorder = (border: XlsxBorder | undefined): number => {
    if (!border || Object.keys(border).length === 0) return 0;
    const key = JSON.stringify(border);
    let id = borderIds.get(key);
    if (id === undefined) {
      id = borders.length;
      borders.push(border);
      borderIds.set(key, id);
    }
    return id;
  };

  for (const rec of records) {
    const d = rec.data;
    if (rec.type === REC_FONT) {
      const id = fontIdx < 4 ? fontIdx : fontIdx + 1;
      while (fonts.length < id) fonts.push({}); // fill the skipped index 4
      fonts[id] = parseFont(d, color);
      fontIdx++;
    } else if (rec.type === REC_FORMAT) {
      const ifmt = u16(d, 0);
      const code = readXlUnicode(d, 2);
      if (code) numFmts.set(ifmt, code);
    } else if (rec.type === REC_XF) {
      cellXfs.push(parseXf(d, color, internFill, internBorder));
    }
  }

  return { numFmts, fonts, fills, borders, cellXfs };
}

function parseFont(d: Uint8Array, color: (icv: number) => string | undefined): XlsxFont {
  const twips = u16(d, 0);
  const grbit = u16(d, 2);
  const icv = u16(d, 4);
  const bls = u16(d, 6); // bold weight: ≥ 700 is bold
  const uls = d[10] ?? 0;
  const name = readShortUnicode(d, 14);
  const out: { -readonly [K in keyof XlsxFont]: XlsxFont[K] } = {};
  if (twips > 0) out.sizePt = twips / 20;
  if (bls >= 700) out.bold = true;
  if ((grbit & 0x0002) !== 0) out.italic = true;
  if (uls !== 0) out.underline = true;
  const c = color(icv);
  if (c) out.colorHex = c;
  if (name) out.name = name;
  return out;
}

function parseXf(
  d: Uint8Array,
  color: (icv: number) => string | undefined,
  internFill: (f: XlsxFill | undefined) => number,
  internBorder: (b: XlsxBorder | undefined) => number,
): XlsxCellXf {
  const fontId = u16(d, 0);
  const numFmtId = u16(d, 2);
  const align1 = d[6] ?? 0;
  const rotation = d[7] ?? 0;
  const indentByte = d[8] ?? 0;

  // Alignment.
  const alignment: { -readonly [K in keyof XlsxCellAlignment]: XlsxCellAlignment[K] } = {};
  const h = HALIGN[align1 & 0x07];
  if (h) alignment.horizontal = h;
  const v = VALIGN[(align1 >> 4) & 0x07];
  if (v && v !== 'bottom') alignment.vertical = v;
  if ((align1 & 0x08) !== 0) alignment.wrapText = true;
  const indent = indentByte & 0x0f;
  if (indent > 0) alignment.indent = indent;
  if ((indentByte & 0x10) !== 0) alignment.shrinkToFit = true;
  if (rotation > 0 && rotation !== 255) alignment.textRotation = rotation;
  else if (rotation === 255) alignment.textRotation = 255;

  // Border (styles in brdbkg1, colours in brdbkg2).
  const brd1 = u16(d, 10);
  const brd2 = u32(d, 12);
  const brd3 = u32(d, 16);
  const edge = (style: number, icv: number) => {
    const name = BORDER_STYLES[style];
    if (!name || name === 'none') return undefined;
    const c = color(icv);
    return c ? { style: name, colorHex: c } : { style: name };
  };
  const border: { -readonly [K in keyof XlsxBorder]: XlsxBorder[K] } = {};
  const left = edge(brd1 & 0x0f, brd2 & 0x7f);
  const right = edge((brd1 >> 4) & 0x0f, (brd2 >> 7) & 0x7f);
  const top = edge((brd1 >> 8) & 0x0f, (brd2 >> 16) & 0x7f);
  const bottom = edge((brd1 >> 12) & 0x0f, (brd2 >> 23) & 0x7f);
  if (left) border.left = left;
  if (right) border.right = right;
  if (top) border.top = top;
  if (bottom) border.bottom = bottom;

  // Fill (pattern + colours in brdbkg3): fls bits 26–31, fg bits 0–6, bg bits 7–13.
  const fls = (brd3 >> 26) & 0x3f;
  const patternType = FILL_PATTERNS[fls];
  let fill: XlsxFill | undefined;
  if (patternType && patternType !== 'none') {
    const fg = color(brd3 & 0x7f);
    const bg = color((brd3 >> 7) & 0x7f);
    fill = { patternType, ...(fg ? { fgColorHex: fg } : {}), ...(bg ? { bgColorHex: bg } : {}) };
  }

  const fillId = internFill(fill);
  const borderId = internBorder(Object.keys(border).length > 0 ? border : undefined);
  return {
    numFmtId,
    fontId,
    fillId,
    borderId,
    ...(numFmtId > 0 ? { applyNumberFormat: true } : {}),
    ...(fillId > 1 ? { applyFill: true } : {}),
    ...(borderId > 0 ? { applyBorder: true } : {}),
    ...(Object.keys(alignment).length > 0 ? { applyAlignment: true, alignment } : {}),
  };
}

// §2.5.17 — a PALETTE record replaces the user palette (indices 8 onward). Each
// colour is 4 bytes: R, G, B, reserved.
function applyPalette(d: Uint8Array, palette: Array<string>): void {
  const count = u16(d, 0);
  for (let i = 0; i < count; i++) {
    const o = 2 + i * 4;
    if (o + 3 > d.length) break;
    const hex = `${hex2(d[o]!)}${hex2(d[o + 1]!)}${hex2(d[o + 2]!)}`;
    if (8 + i < palette.length) palette[8 + i] = hex;
  }
}

// §2.5.293 ShortXLUnicodeString (8-bit count) — used for font names.
function readShortUnicode(d: Uint8Array, off: number): string {
  if (off >= d.length) return '';
  const cch = d[off]!;
  const high = (d[off + 1]! & 0x01) !== 0;
  return decodeChars(d, off + 2, cch, high);
}

// §2.5.294 XLUnicodeString (16-bit count) — used for number-format codes.
function readXlUnicode(d: Uint8Array, off: number): string {
  if (off + 3 > d.length) return '';
  const cch = u16(d, off);
  const high = (d[off + 2]! & 0x01) !== 0;
  return decodeChars(d, off + 3, cch, high);
}

function decodeChars(d: Uint8Array, off: number, cch: number, high: boolean): string {
  let s = '';
  if (high) {
    for (let i = 0; i < cch && off + i * 2 + 1 < d.length; i++) {
      s += String.fromCharCode(d[off + i * 2]! | (d[off + i * 2 + 1]! << 8));
    }
  } else {
    for (let i = 0; i < cch && off + i < d.length; i++) s += String.fromCharCode(d[off + i]!);
  }
  return s;
}

function hex2(n: number): string {
  return n.toString(16).padStart(2, '0').toUpperCase();
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
