// BIFF8 reader (MS-XLS) — the record stream inside a legacy `.xls`'s `Workbook`
// OLE stream, mapped onto the same SheetDoc the OOXML xlsx reader produces, so
// the entire render pipeline (projection → PDF/SVG/HTML) works on a 1997–2003
// `.xls` for free. A BIFF stream is a flat sequence of `[type:u16][len:u16][data]`
// records grouped into substreams (a globals substream then one per sheet, each
// BOF…EOF). We read the workbook globals (sheet directory, shared-string table,
// date system) and each sheet's cell records into WorksheetCells.
//
// Scope: cell VALUES and structure (sheets, merges, column widths, the shared
// strings). Cell styling (fonts/fills/borders/number formats via the XF table)
// is not read yet — the grid renders with default formatting, a documented
// graceful loss, exactly as the OOXML path started before its styles wave.

import type { Sheet, SheetDoc } from '@/core/ir/sheet';
import type {
  ColumnWidth,
  MergedRange,
  ParsedWorksheet,
  WorksheetCell,
} from '@/core/spreadsheet-model';

import { EMPTY_XLSX_STYLES } from '@/excel/styles-parser';
import { ResourceStore } from '@/core/ir';
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

  const sheets: Array<Sheet> = [];
  for (const bs of boundSheets) {
    if (bs.type !== 0) continue; // 0 = worksheet (skip chart/macro sheets)
    if (bs.offset + 4 > wb.length) continue;
    const grid = readSheet(wb, view, bs.offset);
    sheets.push({ name: bs.name, grid });
  }
  if (sheets.length === 0) throw new Error('.xls has no worksheets');

  return {
    kind: 'sheet',
    sheets,
    styles: EMPTY_XLSX_STYLES,
    sharedStrings,
    definedNames: [],
    date1904,
    resources: new ResourceStore(),
  };
}

// Read a BOF…EOF substream starting at a byte offset, returning its records
// (the leading BOF and trailing EOF included). Bounds-checked; stops at EOF or
// the end of the stream.
function readSubstream(wb: Uint8Array, view: DataView, start: number): Array<BiffRecord> {
  const out: Array<BiffRecord> = [];
  let p = start;
  while (p + 4 <= wb.length) {
    const type = view.getUint16(p, true);
    const len = view.getUint16(p + 2, true);
    const dataStart = p + 4;
    if (dataStart + len > wb.length) break;
    out.push({ type, data: wb.subarray(dataStart, dataStart + len) });
    p = dataStart + len;
    if (type === REC.EOF) break;
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

// A worksheet substream → ParsedWorksheet. Only value-bearing records are read;
// the column count/row count come from the actual cells.
function readSheet(wb: Uint8Array, view: DataView, start: number): ParsedWorksheet {
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
        add(numCell(readU16(d, 0), readU16(d, 2), String(readF64(d, 6))));
        break;
      case REC.RK:
        add(numCell(readU16(d, 0), readU16(d, 2), String(decodeRk(readU32(d, 6)))));
        break;
      case REC.MULRK: {
        const row = readU16(d, 0);
        const colFirst = readU16(d, 2);
        const count = Math.floor((d.length - 6) / 6);
        for (let k = 0; k < count; k++) {
          const rk = readU32(d, 4 + k * 6 + 2);
          add(numCell(row, colFirst + k, String(decodeRk(rk))));
        }
        break;
      }
      case REC.LABELSST: {
        const isst = readU32(d, 6);
        add({ column: readU16(d, 2), row: readU16(d, 0), type: 's', rawValue: String(isst) });
        break;
      }
      case REC.LABEL:
      case REC.RSTRING:
        add(strCell(readU16(d, 0), readU16(d, 2), readXlString(d, 6)));
        break;
      case REC.BOOLERR: {
        const row = readU16(d, 0);
        const col = readU16(d, 2);
        const val = d[6] ?? 0;
        const isErr = (d[7] ?? 0) !== 0;
        add(
          isErr
            ? { column: col, row, type: 'e', rawValue: errorText(val) }
            : { column: col, row, type: 'b', rawValue: val ? '1' : '0' },
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

  return { cells, maxRow, maxColumn, columns, merges, rowHeights: [] };
}

function numCell(row: number, column: number, rawValue: string): WorksheetCell {
  return { column, row, type: 'n', rawValue };
}

function strCell(row: number, column: number, text: string): WorksheetCell {
  return { column, row, type: 'inlineStr', rawValue: '', inlineText: text };
}

// §2.4.127 FORMULA — the cached result. When the last two bytes are 0xFFFF the
// result is non-numeric: a string (the following STRING record), a boolean, an
// error, or blank; otherwise the 8 bytes are an IEEE double.
function formulaCell(d: Uint8Array, next: BiffRecord | undefined): WorksheetCell | undefined {
  const row = readU16(d, 0);
  const col = readU16(d, 2);
  if (d.length >= 14 && readU16(d, 12) === 0xffff) {
    const kind = d[6] ?? 0;
    if (kind === 0) {
      // String result — text lives in the next STRING record.
      const text = next && next.type === REC.STRING ? readXlString(next.data, 0) : '';
      return strCell(row, col, text);
    }
    if (kind === 1) return { column: col, row, type: 'b', rawValue: d[8] ? '1' : '0' };
    if (kind === 2) return { column: col, row, type: 'e', rawValue: errorText(d[8] ?? 0) };
    return undefined; // 3 = blank string → no cell
  }
  return numCell(row, col, String(readF64(d, 6)));
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
