// BIFF8 reader (MS-XLS) — the record stream inside a legacy `.xls`'s `Workbook`
// OLE stream, mapped onto the same SheetDoc the OOXML xlsx reader produces, so
// the entire render pipeline (projection → PDF/SVG/HTML) works on a 1997–2003
// `.xls` for free. A BIFF stream is a flat sequence of `[type:u16][len:u16][data]`
// records grouped into substreams (a globals substream then one per sheet, each
// BOF…EOF). We read the workbook globals (sheet directory, shared-string table,
// date system) and each sheet's cell records into WorksheetCells.
//
// Scope: cell values + structure (sheets, merges, column widths, shared strings)
// AND cell styling — fonts, fills, borders, number formats and alignment, read
// from the FONT/FORMAT/PALETTE/XF records (biff-styles) into the same XlsxStyles
// model the OOXML path uses. Embedded charts/drawings are not read yet.

import type { BodyElement, Chart, ShapeBlock } from '@/core/document-model';
import type { Sheet, SheetChartRef, SheetDoc, SheetImageRef } from '@/core/ir/sheet';
import type {
  ColumnWidth,
  MergedRange,
  ParsedWorksheet,
  WorksheetCell,
} from '@/core/spreadsheet-model';
import type { EscherAnchor, EscherShape } from '@/excel/xls/escher';

import { ResourceStore, pt } from '@/core/ir';
import { parseBiffChart } from '@/excel/xls/biff-chart';
import { parseBiffStyles } from '@/excel/xls/biff-styles';
import { parseBlipStore, parseSheetPictures, parseSheetShapes } from '@/excel/xls/escher';
import { openCfb } from '@/core/ole/cfb';

// Record type numbers (MS-XLS §2.3).
const REC = {
  BOF: 0x0809,
  EOF: 0x000a,
  BOUNDSHEET8: 0x0085,
  SST: 0x00fc,
  CONTINUE: 0x003c,
  DATEMODE: 0x0022,
  DIMENSIONS: 0x0200,
  NUMBER: 0x0203,
  RK: 0x027e,
  MULRK: 0x00bd,
  LABELSST: 0x00fd,
  LABEL: 0x0204,
  RSTRING: 0x00d6,
  BOOLERR: 0x0205,
  FORMULA: 0x0006,
  STRING: 0x0207,
  MERGECELLS: 0x00e5,
  COLINFO: 0x007d,
  MSODRAWINGGROUP: 0x00eb, // workbook-globals Escher BLIP store
  MSODRAWING: 0x00ec, // per-sheet Escher shapes
  TXO: 0x01b6, // text object (text-box content)
} as const;

// Bombs / corruption guards.
const MAX_SHEETS = 4096;
const MAX_CELLS = 1 << 22;
const MAX_STRINGS = 1 << 21;

interface BiffRecord {
  readonly type: number;
  readonly data: Uint8Array;
}

// bytes → SheetDoc. Throws on a non-OLE / non-BIFF8 input (the reader's sniff
// keeps those away); a structurally odd record is skipped, never fatal.
export function readXlsToSheetDoc(xls: Uint8Array): SheetDoc {
  const cfb = openCfb(xls);
  const wb = cfb.readStream('Workbook') ?? cfb.readStream('Book');
  if (!wb) throw new Error('not an .xls: no Workbook/Book stream');
  const view = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);

  const globals = readSubstream(wb, view, 0);
  if (globals.length === 0 || globals[0]!.type !== REC.BOF) {
    throw new Error('.xls: missing workbook globals BOF');
  }
  // BIFF8 BOF carries vers 0x0600; older BIFF (Excel 95 and earlier) stores
  // strings differently and is out of scope.
  if (globals[0]!.data.length >= 2 && readU16(globals[0]!.data, 0) !== 0x0600) {
    throw new Error('.xls: only BIFF8 (Excel 97-2003) is supported — re-save as .xlsx');
  }

  let date1904 = false;
  const boundSheets: Array<{ name: string; offset: number; type: number }> = [];
  let sharedStrings: ReadonlyArray<string> = [];
  for (let i = 0; i < globals.length; i++) {
    const rec = globals[i]!;
    switch (rec.type) {
      case REC.DATEMODE:
        date1904 = rec.data.length >= 2 && readU16(rec.data, 0) === 1;
        break;
      case REC.BOUNDSHEET8:
        if (boundSheets.length < MAX_SHEETS) boundSheets.push(parseBoundSheet(rec.data));
        break;
      case REC.SST: {
        // The SST + any immediately-following CONTINUE records form one logical
        // byte stream; a single string may straddle the boundaries.
        const blocks = [rec.data];
        while (i + 1 < globals.length && globals[i + 1]!.type === REC.CONTINUE) {
          blocks.push(globals[++i]!.data);
        }
        sharedStrings = readSst(blocks);
        break;
      }
    }
  }

  // Office Drawing (Escher): the workbook-globals MSODrawingGroup holds the image
  // pool (BLIP store); each sheet's MSODrawing references it. Bytes go into one
  // shared resource store (XLS-5).
  const resources = new ResourceStore();
  const blips = parseBlipStore(gatherDrawing(globals, REC.MSODRAWINGGROUP));
  // Embedded charts resolved against each sheet's cells, keyed globally (XLS-6).
  const chartData = new Map<string, Chart>();

  const sheets: Array<Sheet> = [];
  for (const bs of boundSheets) {
    if (bs.type !== 0) continue; // 0 = worksheet (skip chart/macro sheets)
    if (bs.offset + 4 > wb.length) continue;
    const { grid, images, charts, shapes } = readSheet(
      wb,
      view,
      bs.offset,
      blips,
      resources,
      sharedStrings,
      chartData,
    );
    sheets.push({
      name: bs.name,
      grid,
      ...(images.length > 0 ? { images } : {}),
      ...(charts.length > 0 ? { charts } : {}),
      ...(shapes.length > 0 ? { shapes } : {}),
    });
  }
  if (sheets.length === 0) throw new Error('.xls has no worksheets');

  return {
    kind: 'sheet',
    sheets,
    styles: parseBiffStyles(globals),
    sharedStrings,
    definedNames: [],
    date1904,
    resources,
    ...(chartData.size > 0 ? { chartData } : {}),
  };
}

// Concatenate a record's data with its immediately-following CONTINUE records'
// data — an MSODrawingGroup / MSODrawing Escher stream can be split across them.
function gatherDrawing(recs: ReadonlyArray<BiffRecord>, type: number): Uint8Array {
  const parts: Array<Uint8Array> = [];
  let inside = false;
  for (const rec of recs) {
    if (rec.type === type) {
      parts.push(rec.data);
      inside = true;
    } else if (rec.type === REC.CONTINUE && inside) {
      parts.push(rec.data);
    } else {
      inside = false;
    }
  }
  return concatBytes(parts);
}

function concatBytes(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

// Read a BOF…EOF substream starting at a byte offset, returning its records
// (the leading BOF and trailing EOF included). Bounds-checked; stops at EOF or
// the end of the stream.
function readSubstream(wb: Uint8Array, view: DataView, start: number): Array<BiffRecord> {
  const out: Array<BiffRecord> = [];
  let p = start;
  // A worksheet substream can nest a chart substream (its own BOF…EOF), so track
  // BOF/EOF depth and stop only at the matching EOF — not the first nested one.
  let depth = 0;
  while (p + 4 <= wb.length) {
    const type = view.getUint16(p, true);
    const len = view.getUint16(p + 2, true);
    const dataStart = p + 4;
    if (dataStart + len > wb.length) break;
    out.push({ type, data: wb.subarray(dataStart, dataStart + len) });
    p = dataStart + len;
    if (type === REC.BOF) depth++;
    else if (type === REC.EOF && --depth <= 0) break;
    if (out.length > 1_000_000) break; // runaway guard
  }
  return out;
}

// §2.4.28 BoundSheet8 — lbPlyPos (the sheet substream's BOF offset), the sheet
// type, then a ShortXLUnicodeString name.
function parseBoundSheet(data: Uint8Array): { name: string; offset: number; type: number } {
  const offset = readU32(data, 0);
  const type = data[5] ?? 0; // grbit: byte 4 = visibility, byte 5 = dt (sheet type)
  const name = readShortString(data, 6);
  return { name, offset, type };
}

// A worksheet substream → its grid + any embedded pictures. Only value-bearing
// records are read; the column/row count come from the actual cells. Pictures
// come from the sheet's MSODrawing (Escher) referencing the workbook BLIP store.
function readSheet(
  wb: Uint8Array,
  view: DataView,
  start: number,
  blips: ReadonlyArray<Uint8Array | undefined>,
  resources: ResourceStore,
  sharedStrings: ReadonlyArray<string>,
  chartData: Map<string, Chart>,
): {
  grid: ParsedWorksheet;
  images: Array<SheetImageRef>;
  charts: Array<SheetChartRef>;
  shapes: Array<ShapeBlock>;
} {
  const records = readSubstream(wb, view, start);
  const cells: Array<WorksheetCell> = [];
  const merges: Array<MergedRange> = [];
  const columns: Array<ColumnWidth> = [];
  let maxRow = -1;
  let maxColumn = -1;
  const add = (cell: WorksheetCell): void => {
    if (cells.length >= MAX_CELLS) return;
    cells.push(cell);
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.column > maxColumn) maxColumn = cell.column;
  };

  for (let i = 0; i < records.length; i++) {
    const rec = records[i]!;
    const d = rec.data;
    switch (rec.type) {
      case REC.NUMBER:
        add(numCell(readU16(d, 0), readU16(d, 2), readU16(d, 4), String(readF64(d, 6))));
        break;
      case REC.RK:
        add(numCell(readU16(d, 0), readU16(d, 2), readU16(d, 4), String(decodeRk(readU32(d, 6)))));
        break;
      case REC.MULRK: {
        const row = readU16(d, 0);
        const colFirst = readU16(d, 2);
        const count = Math.floor((d.length - 6) / 6);
        for (let k = 0; k < count; k++) {
          const ixfe = readU16(d, 4 + k * 6);
          const rk = readU32(d, 4 + k * 6 + 2);
          add(numCell(row, colFirst + k, ixfe, String(decodeRk(rk))));
        }
        break;
      }
      case REC.LABELSST: {
        const isst = readU32(d, 6);
        add({
          column: readU16(d, 2),
          row: readU16(d, 0),
          type: 's',
          rawValue: String(isst),
          styleIndex: readU16(d, 4),
        });
        break;
      }
      case REC.LABEL:
      case REC.RSTRING:
        add(strCell(readU16(d, 0), readU16(d, 2), readU16(d, 4), readXlString(d, 6)));
        break;
      case REC.BOOLERR: {
        const row = readU16(d, 0);
        const col = readU16(d, 2);
        const styleIndex = readU16(d, 4);
        const val = d[6] ?? 0;
        const isErr = (d[7] ?? 0) !== 0;
        add(
          isErr
            ? { column: col, row, type: 'e', rawValue: errorText(val), styleIndex }
            : { column: col, row, type: 'b', rawValue: val ? '1' : '0', styleIndex },
        );
        break;
      }
      case REC.FORMULA: {
        const cell = formulaCell(d, records[i + 1]);
        if (cell) add(cell);
        break;
      }
      case REC.MERGECELLS: {
        const cmcs = readU16(d, 0);
        for (let k = 0; k < cmcs; k++) {
          const o = 2 + k * 8;
          if (o + 8 > d.length) break;
          merges.push({
            startRow: readU16(d, o),
            endRow: readU16(d, o + 2),
            startColumn: readU16(d, o + 4),
            endColumn: readU16(d, o + 6),
          });
        }
        break;
      }
      case REC.COLINFO: {
        const colFirst = readU16(d, 0);
        const colLast = readU16(d, 2);
        const widthChars = readU16(d, 4) / 256; // §coldx: 1/256 of a character
        columns.push({ min: colFirst + 1, max: colLast + 1, widthChars });
        break;
      }
    }
  }

  const grid: ParsedWorksheet = { cells, maxRow, maxColumn, columns, merges, rowHeights: [] };
  const drawing = gatherDrawing(records, REC.MSODRAWING);
  const images =
    drawing.length > 0 && blips.length > 0 ? buildImages(drawing, blips, resources) : [];

  // Embedded charts: a nested chart substream (BOF dt=0x20 … EOF) is plotted from
  // the sheet's own cells (XLS-6).
  const charts: Array<SheetChartRef> = [];
  for (let i = 0; i < records.length; i++) {
    if (records[i]!.type !== REC.BOF || readU16(records[i]!.data, 2) !== 0x0020) continue;
    let depth = 0;
    const sub: Array<BiffRecord> = [];
    let j = i;
    for (; j < records.length; j++) {
      sub.push(records[j]!);
      if (records[j]!.type === REC.BOF) depth++;
      else if (records[j]!.type === REC.EOF && --depth <= 0) break;
    }
    i = j;
    const chart = parseBiffChart(sub, cells, sharedStrings);
    if (chart) {
      const path = `xls-chart-${chartData.size}`;
      chartData.set(path, chart);
      charts.push({ chartPartPath: path, widthPt: 360, heightPt: 240 });
    }
  }

  // Drawing shapes (autoshapes, text boxes): the Escher non-picture shapes, with
  // each text box's text from its TXO record, associated by order (XLS-7).
  const shapes =
    drawing.length > 0 ? buildShapes(parseSheetShapes(drawing), gatherTxoTexts(records)) : [];

  return { grid, images, charts, shapes };
}

// A sheet's TXO (text-object) record texts, in order — the cch is in the TXO
// record, the characters in the immediately-following CONTINUE record.
function gatherTxoTexts(recs: ReadonlyArray<BiffRecord>): Array<string> {
  const out: Array<string> = [];
  for (let i = 0; i < recs.length; i++) {
    if (recs[i]!.type !== REC.TXO) continue;
    const cch = readU16(recs[i]!.data, 10);
    const next = recs[i + 1];
    if (cch > 0 && next && next.type === REC.CONTINUE && next.data.length > 0) {
      out.push(decodeChars(next.data, 1, cch, (next.data[0]! & 0x01) !== 0));
    } else {
      out.push('');
    }
  }
  return out;
}

// Escher shapes → ShapeBlocks: preset geometry from the shape type, the
// best-effort fill/line colours, the anchor size, and the text box's text (from
// the matching TXO, by order). Text-bearing shapes consume the TXO texts in turn.
function buildShapes(
  escherShapes: ReadonlyArray<EscherShape>,
  txoTexts: ReadonlyArray<string>,
): Array<ShapeBlock> {
  const out: Array<ShapeBlock> = [];
  let textIdx = 0;
  for (const s of escherShapes) {
    const text = s.hasText ? (txoTexts[textIdx++] ?? '') : '';
    const { widthPt, heightPt } = anchorSize(s.anchor);
    out.push({
      width: pt(widthPt),
      height: pt(heightPt),
      geometry: { kind: 'preset', preset: shapePreset(s.shapeType) },
      fill: s.fillColorHex ? { kind: 'solid', colorHex: s.fillColorHex } : { kind: 'none' },
      ...(s.lineColorHex ? { line: { colorHex: s.lineColorHex } } : {}),
      ...(text.length > 0
        ? {
            text: {
              content: [
                {
                  kind: 'paragraph',
                  paragraph: { properties: {}, runs: [{ text, properties: {} }] },
                },
              ] satisfies Array<BodyElement>,
            },
          }
        : {}),
      paragraphProperties: {},
    });
  }
  return out;
}

// §2.5.* MSOSPT shape type → a DrawingML preset-geometry name the renderer draws.
function shapePreset(type: number): string {
  switch (type) {
    case 1:
      return 'rect';
    case 2:
      return 'roundRect';
    case 3:
      return 'ellipse';
    case 4:
      return 'diamond';
    case 5:
      return 'triangle';
    case 6:
      return 'rtTriangle';
    case 7:
      return 'parallelogram';
    case 8:
      return 'trapezoid';
    case 9:
      return 'hexagon';
    case 20:
      return 'line';
    default:
      return 'rect'; // text boxes (202) and the long tail render as a rectangle
  }
}

// Escher picture shapes → SheetImageRefs: pull each shape's BLIP bytes into the
// shared resource store and size it from its cell anchor (approximate — the
// projection inlines the placement anyway).
function buildImages(
  drawing: Uint8Array,
  blips: ReadonlyArray<Uint8Array | undefined>,
  resources: ResourceStore,
): Array<SheetImageRef> {
  const out: Array<SheetImageRef> = [];
  for (const pic of parseSheetPictures(drawing)) {
    const bytes = blips[pic.blipIndex - 1];
    if (!bytes) continue;
    const { widthPt, heightPt } = anchorSize(pic.anchor);
    out.push({ resourceId: resources.put(bytes), widthPt, heightPt });
  }
  return out;
}

// A cell anchor → an approximate point size using default cell metrics (the
// exact column widths / row heights are not yet threaded here). Absent anchor →
// a default thumbnail size.
function anchorSize(anchor: EscherAnchor | undefined): { widthPt: number; heightPt: number } {
  if (!anchor) return { widthPt: 96, heightPt: 72 };
  const DEFAULT_COL_PT = 48;
  const DEFAULT_ROW_PT = 15;
  const cols = anchor.col2 + anchor.dx2 / 1024 - (anchor.col1 + anchor.dx1 / 1024);
  const rows = anchor.row2 + anchor.dy2 / 1024 - (anchor.row1 + anchor.dy1 / 1024);
  return {
    widthPt: Math.max(8, cols * DEFAULT_COL_PT),
    heightPt: Math.max(8, rows * DEFAULT_ROW_PT),
  };
}

function numCell(row: number, column: number, style: number, rawValue: string): WorksheetCell {
  return { column, row, type: 'n', rawValue, styleIndex: style };
}

function strCell(row: number, column: number, style: number, text: string): WorksheetCell {
  return { column, row, type: 'inlineStr', rawValue: '', inlineText: text, styleIndex: style };
}

// §2.4.127 FORMULA — the cached result. When the last two bytes are 0xFFFF the
// result is non-numeric: a string (the following STRING record), a boolean, an
// error, or blank; otherwise the 8 bytes are an IEEE double.
function formulaCell(d: Uint8Array, next: BiffRecord | undefined): WorksheetCell | undefined {
  const row = readU16(d, 0);
  const col = readU16(d, 2);
  const style = readU16(d, 4);
  if (d.length >= 14 && readU16(d, 12) === 0xffff) {
    const kind = d[6] ?? 0;
    if (kind === 0) {
      // String result — text lives in the next STRING record.
      const text = next && next.type === REC.STRING ? readXlString(next.data, 0) : '';
      return strCell(row, col, style, text);
    }
    if (kind === 1)
      return { column: col, row, type: 'b', rawValue: d[8] ? '1' : '0', styleIndex: style };
    if (kind === 2) {
      return { column: col, row, type: 'e', rawValue: errorText(d[8] ?? 0), styleIndex: style };
    }
    return undefined; // 3 = blank string → no cell
  }
  return numCell(row, col, style, String(readF64(d, 6)));
}

// §2.5.198.5 RkNumber — a packed 30-bit number: bit0 ×100, bit1 integer-vs-IEEE.
// Exported for unit tests (the int / ×100 / IEEE paths).
export function decodeRk(rk: number): number {
  const div100 = (rk & 0x01) !== 0;
  let value: number;
  if ((rk & 0x02) !== 0) {
    value = rk >> 2; // signed 30-bit integer (arithmetic shift)
  } else {
    // The 30 high bits are the 30 high bits of a 64-bit double, the rest zero.
    const buf = new ArrayBuffer(8);
    const dv = new DataView(buf);
    dv.setUint32(0, 0, true);
    dv.setUint32(4, rk & 0xfffffffc, true);
    value = dv.getFloat64(0, true);
  }
  return div100 ? value / 100 : value;
}

function errorText(code: number): string {
  switch (code) {
    case 0x00:
      return '#NULL!';
    case 0x07:
      return '#DIV/0!';
    case 0x0f:
      return '#VALUE!';
    case 0x17:
      return '#REF!';
    case 0x1d:
      return '#NAME?';
    case 0x24:
      return '#NUM!';
    case 0x2a:
      return '#N/A';
    default:
      return '#VALUE!';
  }
}

// §2.5.293 ShortXLUnicodeString — an 8-bit character count then a flags byte
// (only fHighByte matters) then the characters. Used for sheet names.
function readShortString(d: Uint8Array, off: number): string {
  if (off >= d.length) return '';
  const cch = d[off]!;
  const highByte = (d[off + 1]! & 0x01) !== 0;
  return decodeChars(d, off + 2, cch, highByte);
}

// §2.5.294 XLUnicodeString (rich-extended) within a single record — a 16-bit
// count, a flags byte, optional rich-run count / phonetic size, then characters
// (skipping the rich-run + phonetic tails). Cross-record continuation (very rare
// for LABEL) is not handled here — the SST path owns that.
function readXlString(d: Uint8Array, off: number): string {
  if (off + 3 > d.length) return '';
  const cch = readU16(d, off);
  const flags = d[off + 2]!;
  const highByte = (flags & 0x01) !== 0;
  const rich = (flags & 0x08) !== 0;
  const phonetic = (flags & 0x04) !== 0;
  let p = off + 3;
  if (rich) p += 2;
  if (phonetic) p += 4;
  return decodeChars(d, p, cch, highByte);
}

function decodeChars(d: Uint8Array, off: number, cch: number, highByte: boolean): string {
  let s = '';
  if (highByte) {
    for (let i = 0; i < cch && off + i * 2 + 1 < d.length; i++) {
      s += String.fromCharCode(d[off + i * 2]! | (d[off + i * 2 + 1]! << 8));
    }
  } else {
    for (let i = 0; i < cch && off + i < d.length; i++) s += String.fromCharCode(d[off + i]!);
  }
  return s;
}

// §2.4.265 SST — `cstTotal`/`cstUnique` then `cstUnique` rich strings, read by a
// continuation-aware reader (a string can split across a CONTINUE boundary, and
// the byte after the boundary is a fresh fHighByte flag for the rest). Exported
// for unit tests of the continuation handling. `blocks` are the SST record's
// data followed by each CONTINUE record's data.
export function readSst(blocks: ReadonlyArray<Uint8Array>): Array<string> {
  const r = new SstReader(blocks);
  r.u32(); // cstTotal (ignored — cstUnique drives the count)
  const cstUnique = Math.min(r.u32(), MAX_STRINGS);
  const out: Array<string> = [];
  for (let i = 0; i < cstUnique && !r.atEnd(); i++) out.push(r.readString());
  return out;
}

// A cursor over the SST's record blocks. Header/flag reads stay within a block
// (the spec never splits them); only the character array crosses boundaries, and
// each crossing re-reads the fHighByte flag.
class SstReader {
  private bi = 0;
  private off = 0;
  constructor(private readonly blocks: ReadonlyArray<Uint8Array>) {}

  atEnd(): boolean {
    this.settle();
    return this.bi >= this.blocks.length;
  }

  // Advance off the end of exhausted blocks (a "skip" crossing — no flag byte).
  private settle(): void {
    while (this.bi < this.blocks.length && this.off >= this.blocks[this.bi]!.length) {
      this.bi++;
      this.off = 0;
    }
  }

  byte(): number {
    this.settle();
    if (this.bi >= this.blocks.length) return 0;
    return this.blocks[this.bi]![this.off++]!;
  }
  u16(): number {
    return this.byte() | (this.byte() << 8);
  }
  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }

  private skip(n: number): void {
    let left = n;
    while (left > 0) {
      this.settle();
      if (this.bi >= this.blocks.length) return;
      const take = Math.min(left, this.blocks[this.bi]!.length - this.off);
      this.off += take;
      left -= take;
    }
  }

  readString(): string {
    const cch = this.u16();
    const flags = this.byte();
    let highByte = (flags & 0x01) !== 0;
    const rich = (flags & 0x08) !== 0;
    const phonetic = (flags & 0x04) !== 0;
    const cRun = rich ? this.u16() : 0;
    const cbExtRst = phonetic ? this.u32() : 0;
    let s = '';
    for (let i = 0; i < cch; i++) {
      if (this.off >= this.blocks[this.bi]!.length) {
        // Character-array boundary: the continuation starts with a fresh flag.
        this.bi++;
        this.off = 0;
        if (this.bi >= this.blocks.length) break;
        highByte = (this.blocks[this.bi]![this.off++]! & 0x01) !== 0;
      }
      const block = this.blocks[this.bi]!;
      if (highByte) {
        s += String.fromCharCode(block[this.off]! | (block[this.off + 1]! << 8));
        this.off += 2;
      } else {
        s += String.fromCharCode(block[this.off]!);
        this.off += 1;
      }
    }
    this.skip(cRun * 4 + cbExtRst); // rich runs + phonetic block (no flag bytes)
    return s;
  }
}

function readU16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
}
function readU32(d: Uint8Array, off: number): number {
  return (
    ((d[off] ?? 0) |
      ((d[off + 1] ?? 0) << 8) |
      ((d[off + 2] ?? 0) << 16) |
      ((d[off + 3] ?? 0) << 24)) >>>
    0
  );
}
function readF64(d: Uint8Array, off: number): number {
  if (off + 8 > d.length) return 0;
  return new DataView(d.buffer, d.byteOffset + off, 8).getFloat64(0, true);
}
