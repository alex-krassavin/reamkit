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
import type {
  Sheet,
  SheetChartRef,
  SheetComment,
  SheetDoc,
  SheetHyperlink,
  SheetImageRef,
} from '@/core/ir/sheet';
import type {
  ColumnWidth,
  DataValidation,
  DataValidationType,
  DefinedName,
  HeaderFooter,
  MergedRange,
  ParsedWorksheet,
  WorksheetCell,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPrintOptions,
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
  HLINK: 0x01b8, // cell/range hyperlink (§2.4.140)
  NAME: 0x0018, // Lbl — workbook defined name (print area/titles, named range) XLS-10
  // Print model (all sheet-scoped) — XLS-9.
  SETUP: 0x00a1, // page setup (paper/scale/orientation/fit/header+footer margin)
  WSBOOL: 0x0081, // sheet flags — fFitToPage at 0x0100 (NOT 0x0085 = BoundSheet8)
  PRINTGRIDLINES: 0x002b,
  HCENTER: 0x0083,
  VCENTER: 0x0084,
  LEFTMARGIN: 0x0026,
  RIGHTMARGIN: 0x0027,
  TOPMARGIN: 0x0028,
  BOTTOMMARGIN: 0x0029,
  HEADER: 0x0014,
  FOOTER: 0x0015,
  HPAGEBREAK: 0x001b, // horizontal (row) breaks
  VPAGEBREAK: 0x001a, // vertical (column) breaks
  // Cell comments (sheet-scoped) — XLS-11.
  NOTE: 0x001c, // Note — the comment's cell + author + idObj
  OBJ: 0x005d, // Obj — the comment's text-box object (FtCmo ot 0x19)
  // Data validation (sheet-scoped) — XLS-12.
  DV: 0x01be, // one validation rule (DVAL 0x01B2 is the table header, not needed)
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
  const lblRecords: Array<Uint8Array> = []; // defined names (Lbl), resolved after the loop
  for (let i = 0; i < globals.length; i++) {
    const rec = globals[i]!;
    switch (rec.type) {
      case REC.NAME:
        lblRecords.push(rec.data);
        break;
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
    const { grid, images, charts, shapes, hyperlinks, comments } = readSheet(
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
      ...(hyperlinks.length > 0 ? { hyperlinks } : {}),
      ...(comments.length > 0 ? { comments } : {}),
    });
  }
  if (sheets.length === 0) throw new Error('.xls has no worksheets');

  return {
    kind: 'sheet',
    sheets,
    styles: parseBiffStyles(globals),
    sharedStrings,
    definedNames: buildDefinedNames(lblRecords, boundSheets),
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
  hyperlinks: Array<SheetHyperlink>;
  comments: Array<SheetComment>;
} {
  const records = readSubstream(wb, view, start);
  const cells: Array<WorksheetCell> = [];
  const merges: Array<MergedRange> = [];
  const columns: Array<ColumnWidth> = [];
  const hyperlinks: Array<SheetHyperlink> = [];
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
      case REC.HLINK: {
        const link = parseHlink(d);
        if (link) hyperlinks.push(link);
        break;
      }
    }
  }

  const dataValidations = parseDataValidations(records);
  const grid: ParsedWorksheet = {
    cells,
    maxRow,
    maxColumn,
    columns,
    merges,
    rowHeights: [],
    ...parsePrintModel(records),
    ...(dataValidations.length > 0 ? { dataValidations } : {}),
  };
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

  return { grid, images, charts, shapes, hyperlinks, comments: parseComments(records) };
}

// HLink flags (the hlink object's hlinkFlags, [MS-OSHARED] §2.3.7.1).
const HLINK_HAS_MONIKER = 0x00000001;
const HLINK_HAS_DISPLAY_NAME = 0x00000010;
const HLINK_HAS_FRAME_NAME = 0x00000080;
const HLINK_MONIKER_AS_STR = 0x00000100;
// URLMoniker CLSID {79EAC9E0-BAF9-11CE-8C82-00AA004BA90B} as on-disk LE bytes.
const URL_MONIKER_CLSID: ReadonlyArray<number> = [
  0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];

// An HLink record (§2.4.140) → a resolved cell/range hyperlink, but only for an
// external URL: a Ref8U range, the StdHlink CLSID, then a Hyperlink Object whose
// moniker is a URLMoniker (the common web-link case). In-workbook location links
// and file/UNC monikers carry no URL and are skipped (missing, never wrong).
function parseHlink(d: Uint8Array): SheetHyperlink | undefined {
  if (d.length < 32) return undefined;
  const ref: MergedRange = {
    startRow: readU16(d, 0),
    endRow: readU16(d, 2),
    startColumn: readU16(d, 4),
    endColumn: readU16(d, 6),
  };
  // ref8 (8) + hlinkClsid (16) → the hlink object at offset 24.
  if (readU32(d, 24) !== 2) return undefined; // streamVersion MUST be 2
  const flags = readU32(d, 28);
  let off = 32;
  if (flags & HLINK_HAS_DISPLAY_NAME) off = skipHlinkString(d, off);
  if (flags & HLINK_HAS_FRAME_NAME) off = skipHlinkString(d, off);
  if ((flags & HLINK_HAS_MONIKER) === 0) return undefined; // location-only link
  if (flags & HLINK_MONIKER_AS_STR) {
    // Moniker stored as a string (relative/UNC path) — keep only an absolute URL.
    const s = readHlinkString(d, off);
    return s && isAbsoluteUrl(s.value) ? { ref, url: s.value } : undefined;
  }
  // oleMoniker: a 16-byte CLSID then the moniker data.
  if (off + 16 > d.length || !clsidEquals(d, off, URL_MONIKER_CLSID)) return undefined;
  const url = readUrlMonikerUrl(d, off + 16);
  return url ? { ref, url } : undefined;
}

// A HyperlinkString (§2.3.7.9): a u32 character count (including the terminating
// NUL) then that many UTF-16LE chars. Returns the string (NUL stripped) + offset.
function readHlinkString(d: Uint8Array, off: number): { value: string; next: number } | undefined {
  if (off + 4 > d.length) return undefined;
  const bytes = readU32(d, off) * 2;
  const start = off + 4;
  if (start + bytes > d.length) return undefined;
  return { value: stripNul(decodeUtf16Le(d, start, bytes)), next: start + bytes };
}
function skipHlinkString(d: Uint8Array, off: number): number {
  return readHlinkString(d, off)?.next ?? d.length; // a bad length aborts the walk
}

// A URLMoniker's data (§2.3.7.6): a u32 byte length then a NUL-terminated UTF-16LE
// URL. The NUL bounds the URL, so the optional trailing serial-GUID tail is ignored.
function readUrlMonikerUrl(d: Uint8Array, off: number): string | undefined {
  if (off + 4 > d.length) return undefined;
  let s = '';
  for (let i = off + 4; i + 1 < d.length; i += 2) {
    const c = d[i]! | (d[i + 1]! << 8);
    if (c === 0) break;
    s += String.fromCharCode(c);
  }
  return s.length > 0 ? s : undefined;
}

function decodeUtf16Le(d: Uint8Array, off: number, bytes: number): string {
  let s = '';
  for (let i = 0; i + 1 < bytes; i += 2) {
    s += String.fromCharCode(d[off + i]! | (d[off + i + 1]! << 8));
  }
  return s;
}
function stripNul(s: string): string {
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) === 0) return s.slice(0, i);
  return s;
}
function isAbsoluteUrl(s: string): boolean {
  return /^[a-z][a-z0-9+.-]*:\/\//i.test(s);
}
function clsidEquals(d: Uint8Array, off: number, clsid: ReadonlyArray<number>): boolean {
  for (let i = 0; i < 16; i++) if (d[off + i] !== clsid[i]) return false;
  return true;
}

// The sheet's print model (XLS-9): the page-setup, print-option, margin,
// header/footer and manual-page-break records scattered through the worksheet
// substream → the same grid fields the OOXML `.xlsx` path fills, so the existing
// print/pagination model renders a legacy workbook identically. Each record is
// optional; an absent one leaves its field undefined (the projection's defaults).
function parsePrintModel(records: ReadonlyArray<BiffRecord>): {
  pageSetup?: XlsxPageSetup;
  fitToPage?: boolean;
  printOptions?: XlsxPrintOptions;
  pageMargins?: XlsxPageMargins;
  headerFooter?: HeaderFooter;
  rowBreaks?: Array<number>;
  colBreaks?: Array<number>;
} {
  let paperSize: number | undefined;
  let orientation: 'portrait' | 'landscape' | undefined;
  let scale: number | undefined;
  let fitToWidth: number | undefined;
  let fitToHeight: number | undefined;
  let fitToPage: boolean | undefined;
  let gridLines: boolean | undefined;
  let horizontalCentered: boolean | undefined;
  let verticalCentered: boolean | undefined;
  let left: number | undefined;
  let right: number | undefined;
  let top: number | undefined;
  let bottom: number | undefined;
  let headerMargin: number | undefined;
  let footerMargin: number | undefined;
  let oddHeader: string | undefined;
  let oddFooter: string | undefined;
  let rowBreaks: Array<number> | undefined;
  let colBreaks: Array<number> | undefined;

  for (const r of records) {
    const d = r.data;
    switch (r.type) {
      case REC.SETUP: {
        if (d.length < 12) break;
        // grbit: fNoPls 0x04 ⇒ paper/scale/orientation undefined; fNoOrient 0x40 ⇒
        // orientation undefined; fPortrait 0x02 ⇒ portrait, else landscape.
        const grbit = readU16(d, 10);
        if ((grbit & 0x0004) === 0 && (grbit & 0x0040) === 0) {
          orientation = (grbit & 0x0002) !== 0 ? 'portrait' : 'landscape';
        }
        const sc = readU16(d, 2);
        if (sc >= 10 && sc <= 400) scale = sc;
        const ps = readU16(d, 0);
        if ((grbit & 0x0004) === 0 && ps > 0) paperSize = ps;
        fitToWidth = readU16(d, 6);
        fitToHeight = readU16(d, 8);
        if (d.length >= 32) {
          headerMargin = readF64(d, 16); // numHdr (inches)
          footerMargin = readF64(d, 24); // numFtr (inches)
        }
        break;
      }
      case REC.WSBOOL:
        if (d.length >= 2) fitToPage = (readU16(d, 0) & 0x0100) !== 0; // fFitToPage
        break;
      case REC.PRINTGRIDLINES:
        if (d.length >= 2) gridLines = (readU16(d, 0) & 0x0001) !== 0;
        break;
      case REC.HCENTER:
        if (d.length >= 2) horizontalCentered = readU16(d, 0) !== 0;
        break;
      case REC.VCENTER:
        if (d.length >= 2) verticalCentered = readU16(d, 0) !== 0;
        break;
      case REC.LEFTMARGIN:
        if (d.length >= 8) left = readF64(d, 0);
        break;
      case REC.RIGHTMARGIN:
        if (d.length >= 8) right = readF64(d, 0);
        break;
      case REC.TOPMARGIN:
        if (d.length >= 8) top = readF64(d, 0);
        break;
      case REC.BOTTOMMARGIN:
        if (d.length >= 8) bottom = readF64(d, 0);
        break;
      case REC.HEADER: {
        const s = readXlString(d, 0); // empty record ⇒ '' ⇒ no header
        if (s.length > 0) oddHeader = s;
        break;
      }
      case REC.FOOTER: {
        const s = readXlString(d, 0);
        if (s.length > 0) oddFooter = s;
        break;
      }
      case REC.HPAGEBREAK:
        rowBreaks = parseBreaks(d);
        break;
      case REC.VPAGEBREAK:
        colBreaks = parseBreaks(d);
        break;
    }
  }

  const pageSetup: XlsxPageSetup = {
    ...(paperSize !== undefined ? { paperSize } : {}),
    ...(orientation !== undefined ? { orientation } : {}),
    ...(scale !== undefined ? { scale } : {}),
    ...(fitToWidth !== undefined ? { fitToWidth } : {}),
    ...(fitToHeight !== undefined ? { fitToHeight } : {}),
  };
  const printOptions: XlsxPrintOptions = {
    ...(gridLines !== undefined ? { gridLines } : {}),
    ...(horizontalCentered !== undefined ? { horizontalCentered } : {}),
    ...(verticalCentered !== undefined ? { verticalCentered } : {}),
  };
  const pageMargins: XlsxPageMargins | undefined =
    left !== undefined && right !== undefined && top !== undefined && bottom !== undefined
      ? {
          leftInches: left,
          rightInches: right,
          topInches: top,
          bottomInches: bottom,
          ...(headerMargin !== undefined ? { headerInches: headerMargin } : {}),
          ...(footerMargin !== undefined ? { footerInches: footerMargin } : {}),
        }
      : undefined;
  const headerFooter: HeaderFooter = {
    ...(oddHeader !== undefined ? { oddHeader } : {}),
    ...(oddFooter !== undefined ? { oddFooter } : {}),
  };

  return {
    ...(Object.keys(pageSetup).length > 0 ? { pageSetup } : {}),
    ...(fitToPage !== undefined ? { fitToPage } : {}),
    ...(Object.keys(printOptions).length > 0 ? { printOptions } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(Object.keys(headerFooter).length > 0 ? { headerFooter } : {}),
    ...(rowBreaks && rowBreaks.length > 0 ? { rowBreaks } : {}),
    ...(colBreaks && colBreaks.length > 0 ? { colBreaks } : {}),
  };
}

// A page-break record (HPageBreak/VPageBreak): a u16 count then that many 6-byte
// entries whose first u16 is the row (or column) that begins the next page.
function parseBreaks(d: Uint8Array): Array<number> {
  const count = readU16(d, 0);
  const out: Array<number> = [];
  for (let k = 0; k < count && 2 + k * 6 + 2 <= d.length; k++) out.push(readU16(d, 2 + k * 6));
  return out;
}

// === Defined names (Lbl records) — XLS-10 ====================================

// The built-in name ids the reader models — the print-model names that the
// projection resolves (others, e.g. _FilterDatabase, carry no render meaning and
// are skipped). The reserved _xlnm. prefix matches the OOXML reader's names.
const BUILTIN_NAMES: ReadonlyMap<number, string> = new Map([
  [0x06, '_xlnm.Print_Area'],
  [0x07, '_xlnm.Print_Titles'],
]);

// One area extracted from a name's parsed formula, 0-based inclusive.
interface NameArea {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

// The workbook's Lbl (defined-name) records → DefinedName[] in the same shape the
// OOXML reader produces, so the projection resolves print areas / titles (and the
// writer round-trips named ranges) identically.
function buildDefinedNames(
  lblRecords: ReadonlyArray<Uint8Array>,
  boundSheets: ReadonlyArray<{ name: string; offset: number; type: number }>,
): Array<DefinedName> {
  // Map a 1-based BoundSheet8 index → the 0-based worksheet index (type 0 only),
  // since SheetDoc.sheets and localSheetId count only worksheets.
  const wsIndex: Array<number | undefined> = [];
  let ws = 0;
  for (const bs of boundSheets) wsIndex.push(bs.type === 0 ? ws++ : undefined);

  const out: Array<DefinedName> = [];
  for (const d of lblRecords) {
    const dn = parseLbl(d, wsIndex);
    if (dn) out.push(dn);
  }
  return out;
}

// One Lbl record (§2.4.150) → a DefinedName, or undefined when it is an unmodelled
// built-in, an unparseable formula, or scoped to a non-worksheet (missing, never
// wrong). Layout: grbit(2) chKey(1) cch(1) cce(2) reserved(2) itab(2) reserved(4)
// then the Name (XLUnicodeStringNoCch: flags(1) + chars) then the rgce formula.
function parseLbl(
  d: Uint8Array,
  wsIndex: ReadonlyArray<number | undefined>,
): DefinedName | undefined {
  if (d.length < 15) return undefined;
  const grbit = readU16(d, 0);
  const builtin = (grbit & 0x0020) !== 0; // fBuiltin
  const cch = d[3]!;
  const cce = readU16(d, 4);
  const itab = readU16(d, 8);
  const nameFlags = d[14]!; // XLUnicodeStringNoCch flags (fHighByte = bit 0)
  const highByte = (nameFlags & 0x01) !== 0;
  const nameBytes = cch * (highByte ? 2 : 1);

  let name: string;
  if (builtin) {
    const id = highByte ? readU16(d, 15) : (d[15] ?? 0xff); // the single built-in id char
    const mapped = BUILTIN_NAMES.get(id);
    if (mapped === undefined) return undefined;
    name = mapped;
  } else {
    name = decodeChars(d, 15, cch, highByte);
    if (name.length === 0) return undefined;
  }

  const rgceOff = 15 + nameBytes;
  const value = formatNameValue(d.subarray(rgceOff, Math.min(d.length, rgceOff + cce)));
  if (value === undefined) return undefined; // unparseable formula → skip

  if (itab > 0) {
    const localSheetId = wsIndex[itab - 1];
    if (localSheetId === undefined) return undefined; // scoped to a non-worksheet
    return { name, value, localSheetId };
  }
  return { name, value };
}

// A name's NameParsedFormula (Rgce) → an A1-style value string (the form the
// OOXML defined-name resolver parses). Walks the ptg stream collecting the cell
// ranges it references; any unrecognized ptg aborts the walk so a name resolves
// to its ranges or to nothing — never to a wrong range.
function formatNameValue(rgce: Uint8Array): string | undefined {
  const areas = extractNameAreas(rgce);
  return areas ? areas.map(formatArea).join(',') : undefined;
}

function extractNameAreas(rgce: Uint8Array): Array<NameArea> | undefined {
  const areas: Array<NameArea> = [];
  let off = 0;
  while (off < rgce.length) {
    const ptg = rgce[off]!;
    const base = ptg & 0x1f;
    // Operand ptgs carry a class in bits 5–6 (ptg ≥ 0x20); operators are < 0x20.
    if (ptg >= 0x20 && base === 0x05) {
      // PtgArea (2-D): an 8-byte RgceArea.
      if (off + 9 > rgce.length) return undefined;
      areas.push(readArea(rgce, off + 1));
      off += 9;
    } else if (ptg >= 0x20 && base === 0x1b) {
      // PtgArea3d: a 2-byte ixti then the 8-byte area.
      if (off + 11 > rgce.length) return undefined;
      areas.push(readArea(rgce, off + 3));
      off += 11;
    } else if (ptg >= 0x20 && base === 0x04) {
      // PtgRef (single cell): a 4-byte RgceLoc.
      if (off + 5 > rgce.length) return undefined;
      areas.push(readCell(rgce, off + 1));
      off += 5;
    } else if (ptg >= 0x20 && base === 0x1a) {
      // PtgRef3d: ixti then the 4-byte loc.
      if (off + 7 > rgce.length) return undefined;
      areas.push(readCell(rgce, off + 3));
      off += 7;
    } else if (ptg >= 0x20 && base === 0x09) {
      off += 3; // PtgMemFunc: ptg + cce; the subexpression follows inline
    } else if (ptg >= 0x20 && base === 0x06) {
      off += 7; // PtgMemArea: ptg + reserved(4) + cce; subexpression inline
    } else if (ptg === 0x10 || ptg === 0x0f || ptg === 0x11 || ptg === 0x15) {
      off += 1; // tUnion (,) / tIsect (space) / tRange (:) / tParen
    } else {
      return undefined; // an unmodelled ptg — abort rather than risk a wrong range
    }
  }
  return areas.length > 0 ? areas : undefined;
}

function readArea(d: Uint8Array, off: number): NameArea {
  return {
    r0: readU16(d, off),
    r1: readU16(d, off + 2),
    c0: readU16(d, off + 4) & 0x3fff,
    c1: readU16(d, off + 6) & 0x3fff,
  };
}
function readCell(d: Uint8Array, off: number): NameArea {
  const r = readU16(d, off);
  const c = readU16(d, off + 2) & 0x3fff;
  return { r0: r, r1: r, c0: c, c1: c };
}

// An area → an A1 token: a whole-row band (full column span) is "r0:r1", a
// whole-column band is "cA:cB", a single cell is "A1", else "A1:B2".
function formatArea(a: NameArea): string {
  const fullWidth = a.c0 === 0 && a.c1 >= 0xff; // repeat-rows band
  const fullHeight = a.r0 === 0 && a.r1 >= 0xffff; // repeat-columns band
  if (fullWidth && !fullHeight) return `${a.r0 + 1}:${a.r1 + 1}`;
  if (fullHeight && !fullWidth) return `${colLetters(a.c0)}:${colLetters(a.c1)}`;
  const tl = `${colLetters(a.c0)}${a.r0 + 1}`;
  return a.r0 === a.r1 && a.c0 === a.c1 ? tl : `${tl}:${colLetters(a.c1)}${a.r1 + 1}`;
}
function colLetters(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

// === Data validation (DV records) — XLS-12 ===================================

// valType (DV flags bits 0–3) → the validation type; typOperator (bits 20–23) →
// the comparison operator (meaningful only for the value-range types).
const DV_TYPES: ReadonlyArray<DataValidationType> = [
  'none',
  'whole',
  'decimal',
  'list',
  'date',
  'time',
  'textLength',
  'custom',
];
const DV_OPERATORS: ReadonlyArray<string> = [
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
];
const DV_OPERATOR_TYPES = new Set<DataValidationType>([
  'whole',
  'decimal',
  'date',
  'time',
  'textLength',
]);

// The sheet's DV records → DataValidation[] in the same shape the OOXML reader
// builds, so a `list` rule paints the projection's in-cell dropdown affordance.
function parseDataValidations(records: ReadonlyArray<BiffRecord>): Array<DataValidation> {
  const out: Array<DataValidation> = [];
  for (const r of records) {
    if (r.type === REC.DV) {
      const dv = parseDv(r.data);
      if (dv) out.push(dv);
    }
  }
  return out;
}

// One DV record (§2.4.82) → a DataValidation, or undefined when structurally short
// (missing, never wrong). Layout: dwDvFlags(u32), four XLUnicodeStrings
// (promptTitle/errorTitle/prompt/error), two DVParsedFormulas (cce + unused +
// rgce), then a SqRefU (count + Ref8U array).
function parseDv(d: Uint8Array): DataValidation | undefined {
  if (d.length < 4) return undefined;
  const flags = readU32(d, 0);
  const type = DV_TYPES[flags & 0x0f];
  if (type === undefined) return undefined;
  const operator = DV_OPERATORS[(flags >> 20) & 0x0f];

  let off = 4;
  const promptTitle = readXlStringEnd(d, off);
  off = promptTitle.next;
  const errorTitle = readXlStringEnd(d, off);
  off = errorTitle.next;
  const prompt = readXlStringEnd(d, off);
  off = prompt.next;
  const error = readXlStringEnd(d, off);
  off = error.next;

  if (off + 4 > d.length) return undefined;
  const cce1 = readU16(d, off);
  off += 4; // cce1 + unused
  const rgce1 = d.subarray(off, Math.min(d.length, off + cce1));
  off += cce1;
  if (off + 4 > d.length) return undefined;
  const cce2 = readU16(d, off);
  off += 4 + cce2; // skip formula2 (unused for list / most rules)

  if (off + 2 > d.length) return undefined;
  const cref = readU16(d, off);
  off += 2;
  const ranges: Array<MergedRange> = [];
  for (let k = 0; k < cref && off + 8 <= d.length; k++) {
    ranges.push({
      startRow: readU16(d, off),
      endRow: readU16(d, off + 2),
      startColumn: readU16(d, off + 4),
      endColumn: readU16(d, off + 6),
    });
    off += 8;
  }
  if (ranges.length === 0) return undefined;

  const formula1 = type === 'list' ? extractListSource(rgce1) : undefined;
  return {
    type,
    ranges,
    ...(operator !== undefined && DV_OPERATOR_TYPES.has(type) ? { operator } : {}),
    ...((flags & 0x100) !== 0 ? { allowBlank: true } : {}),
    showDropDown: (flags & 0x200) !== 0, // fSuppressCombo (ECMA's inverted sense)
    ...((flags & 0x40000) !== 0 ? { showInputMessage: true } : {}),
    ...((flags & 0x80000) !== 0 ? { showErrorMessage: true } : {}),
    ...(promptTitle.value ? { promptTitle: promptTitle.value } : {}),
    ...(errorTitle.value ? { errorTitle: errorTitle.value } : {}),
    ...(prompt.value ? { prompt: prompt.value } : {}),
    ...(error.value ? { error: error.value } : {}),
    ...(formula1 ? { formula1 } : {}),
  };
}

// A `list` validation's source from rgce1: a PtgStr (0x17) literal "a,b,c", or an
// area reference rendered to A1. Undefined when neither (the dropdown still shows).
function extractListSource(rgce: Uint8Array): string | undefined {
  if (rgce.length >= 3 && rgce[0] === 0x17) {
    const cch = rgce[1]!;
    return `"${decodeChars(rgce, 3, cch, (rgce[2]! & 0x01) !== 0)}"`;
  }
  const areas = extractNameAreas(rgce);
  return areas ? areas.map(formatArea).join(',') : undefined;
}

// An XLUnicodeString that also returns its end offset (readXlString returns only
// the text) — used to walk the DV record's four consecutive strings.
function readXlStringEnd(d: Uint8Array, off: number): { value: string; next: number } {
  if (off + 3 > d.length) return { value: '', next: d.length };
  const cch = readU16(d, off);
  const flags = d[off + 2]!;
  const high = (flags & 0x01) !== 0;
  let p = off + 3;
  if ((flags & 0x08) !== 0) p += 2; // rich-run count
  if ((flags & 0x04) !== 0) p += 4; // phonetic size
  return { value: decodeChars(d, p, cch, high), next: p + cch * (high ? 2 : 1) };
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

// Cell comments (XLS-11): a Note (0x1C) gives the cell + author + idObj; its text
// lives in the Txo of the Obj whose FtCmo id matches idObj (ot 0x19 = comment).
// Join on the shared object id (POI's HSSFComment key), not record adjacency, so
// a non-comment Obj interleaved in the drawing block can't desync the text.
function parseComments(records: ReadonlyArray<BiffRecord>): Array<SheetComment> {
  const textByObj = new Map<number, string>();
  let pendingObjId: number | undefined;
  const notes: Array<{ idObj: number; rw: number; col: number; author: string }> = [];

  for (let i = 0; i < records.length; i++) {
    const r = records[i]!;
    const d = r.data;
    if (r.type === REC.OBJ) {
      // FtCmo: ft 0x15 @0, cb @2, ot @4 (0x19 = comment), id @6.
      pendingObjId =
        d.length >= 8 && readU16(d, 0) === 0x0015 && readU16(d, 4) === 0x0019
          ? readU16(d, 6)
          : undefined;
    } else if (r.type === REC.TXO && pendingObjId !== undefined) {
      const cch = readU16(d, 10);
      const next = records[i + 1];
      textByObj.set(
        pendingObjId,
        cch > 0 && next && next.type === REC.CONTINUE && next.data.length > 0
          ? decodeChars(next.data, 1, cch, (next.data[0]! & 0x01) !== 0)
          : '',
      );
      pendingObjId = undefined;
    } else if (r.type === REC.NOTE && d.length >= 11) {
      const cchAuthor = readU16(d, 8);
      notes.push({
        rw: readU16(d, 0),
        col: readU16(d, 2),
        idObj: readU16(d, 6),
        author: decodeChars(d, 11, cchAuthor, (d[10]! & 0x01) !== 0),
      });
    }
  }

  const out: Array<SheetComment> = [];
  for (const n of notes) {
    const text = textByObj.get(n.idObj) ?? '';
    if (text.length === 0) continue; // no resolvable text → skip (missing, never wrong)
    out.push({
      ref: `${colLetters(n.col)}${n.rw + 1}`,
      ...(n.author.length > 0 ? { author: n.author } : {}),
      text,
      threaded: false,
    });
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
