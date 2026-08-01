// Sheet drawings (ECMA-376 §20.5 SpreadsheetDrawingML): the worksheet's
// <drawing r:id> points at xl/drawings/drawingN.xml, whose anchors place
// graphic frames over the grid. v1 extracts CHART frames only (shapes and
// images on sheets stay out of scope) and renders each chart as a block after
// the sheet's table — anchor-ordered, sized from the cell-range anchor.

import { XMLParser } from 'fast-xml-parser';
import type { OpcPackage } from '@/core/opc';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';

import { emuToPt } from '@/core/ir';
import {
  COL_PADDING_TWIPS,
  DEFAULT_COL_TWIPS,
  DEFAULT_ROW_TWIPS,
  TWIPS_PER_EXCEL_CHAR,
  columnTwips,
} from '@/excel/print-model';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** §20.5 — a chart frame anchored over the grid, sized from its cell-range anchor. */
export interface SheetChartRef {
  /**
   * The resolved chart part path (`'xl/charts/chart1.xml'`) — globally unique, used
   * as the FlowDoc charts key and the `ChartBlock.chartRelId`.
   */
  readonly chartPartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
  /**
   * Where the anchor puts the frame, in points from the grid's top-left. A
   * chart is anchored over the sheet like any other drawing; placed in the flow
   * instead it lands at the left margin, below everything, and on whatever page
   * that turns out to be.
   */
  readonly xPt: number;
  readonly yPt: number;
  /** Anchor top row (0-based) — used only to order charts on the sheet. */
  readonly anchorRow: number;
}

/**
 * §20.5.2.1 `xdr:pic` — a picture anchored over the grid. The reader reads the
 * resolved media part's bytes into the SheetDoc resource store.
 */
export interface SheetPicture {
  /** The resolved media part path; the reader loads its bytes into the resource store. */
  readonly imagePartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
  /** Where the anchor puts it, in points from the grid's top-left. */
  readonly xPt: number;
  readonly yPt: number;
  /** Anchor top row (0-based) — used only to order pictures on the sheet. */
  readonly anchorRow: number;
}

/** Both kinds of anchored frame the drawing yields, anchor-ordered. */
export interface SheetDrawing {
  readonly charts: Array<SheetChartRef>;
  readonly pictures: Array<SheetPicture>;
}

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  parseTagValue: false,
});

const TWIPS_PER_PT = 20;

/**
 * Parse a `xl/drawings/drawingN.xml` part (§20.5.2.35/.33/.1 — `xdr:twoCellAnchor`
 * / `oneCellAnchor` / `absoluteAnchor`) into the {@link SheetDrawing} it yields.
 * Each anchor frames either a chart (`graphicFrame`) or a picture (`xdr:pic`); both
 * are sized from the anchor and returned anchor-ordered.
 *
 * @param drawingXml      The drawing part bytes.
 * @param drawingPartPath The drawing part path, for resolving its relationships.
 * @param pkg             The OPC package, used to resolve relId → part path.
 * @param worksheet       The host worksheet, for the column/row track geometry.
 */
export function parseSheetDrawing(
  drawingXml: Uint8Array,
  drawingPartPath: string,
  pkg: OpcPackage,
  worksheet: ParsedWorksheet,
): SheetDrawing {
  const tree = parser.parse(new TextDecoder().decode(drawingXml)) as Record<string, unknown>;
  const root = tree['wsDr'];
  if (!root || typeof root !== 'object') return { charts: [], pictures: [] };
  const rootObj = root as Record<string, unknown>;
  const colWidthPt = makeColWidthPt(worksheet);
  const rowHeightPt = makeRowHeightPt(worksheet);
  const rels = pkg.getPartRelationships(drawingPartPath);
  const partPathOf = (relId: string): string | undefined => {
    const rel = rels.find((r) => r.id === relId);
    return rel ? pkg.resolveRelatedPart(drawingPartPath, rel)?.path : undefined;
  };

  const charts: Array<SheetChartRef> = [];
  const pictures: Array<SheetPicture> = [];
  for (const kind of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'] as const) {
    for (const anchor of asArray(rootObj[kind])) {
      if (!anchor || typeof anchor !== 'object') continue;
      const a = anchor as Record<string, unknown>;
      const chartRelId = chartRelIdOf(a);
      const picRelId = chartRelId ? undefined : picRelIdOf(a);
      if (!chartRelId && !picRelId) continue;
      // §20.1.2.2.8 `cNvPr@hidden` — the object says it is not to be shown.
      if (isHiddenDrawing(a, chartRelId ? 'graphicFrame' : 'pic')) continue;

      const from = cellMarker(a['from']);
      let widthPt = 0;
      let heightPt = 0;
      if (kind === 'twoCellAnchor') {
        const to = cellMarker(a['to']);
        if (!from || !to) continue;
        widthPt = spanPt(from.col, from.colOffPt, to.col, to.colOffPt, colWidthPt);
        heightPt = spanPt(from.row, from.rowOffPt, to.row, to.rowOffPt, rowHeightPt);
      } else {
        const ext = a['ext'];
        if (!ext || typeof ext !== 'object') continue;
        const e = ext as Record<string, unknown>;
        widthPt = emuToPt(num(e['@_cx']) ?? 0);
        heightPt = emuToPt(num(e['@_cy']) ?? 0);
      }
      if (widthPt <= 0 || heightPt <= 0) continue;
      const anchorRow = from?.row ?? 0;
      // The anchor's own origin: every track before it, plus its offset into
      // the one it starts in. §20.5.2.1 `absoluteAnchor` names no cell at all —
      // it carries the position outright, in EMU from the grid origin
      // (absolute-anchor-over-empty-sheet.xlsx puts its logo 351pt across and
      // 520pt down an otherwise empty sheet).
      const posRaw = kind === 'absoluteAnchor' ? a['pos'] : undefined;
      const posNode =
        posRaw && typeof posRaw === 'object' ? (posRaw as Record<string, unknown>) : undefined;
      const xPt = posNode
        ? emuToPt(num(posNode['@_x']) ?? 0)
        : from
          ? spanPt(0, 0, from.col, from.colOffPt, colWidthPt)
          : 0;
      const yPt = posNode
        ? emuToPt(num(posNode['@_y']) ?? 0)
        : from
          ? spanPt(0, 0, from.row, from.rowOffPt, rowHeightPt)
          : 0;

      if (chartRelId) {
        const path = partPathOf(chartRelId);
        if (path) charts.push({ chartPartPath: path, widthPt, heightPt, xPt, yPt, anchorRow });
      } else if (picRelId) {
        const path = partPathOf(picRelId);
        if (path) pictures.push({ imagePartPath: path, widthPt, heightPt, xPt, yPt, anchorRow });
      }
    }
  }
  charts.sort((x, y) => x.anchorRow - y.anchorRow);
  pictures.sort((x, y) => x.anchorRow - y.anchorRow);
  return { charts, pictures };
}

// §20.1.2.2.8 `<xdr:cNvPr hidden="1"/>` — "Specifies whether this DrawingML
// object shall be displayed", default false. Reached through whichever
// non-visual wrapper the object carries (xdr:nvPicPr / xdr:nvGraphicFramePr),
// which removeNSPrefix flattens to nvPicPr / nvGraphicFramePr.
function isHiddenDrawing(anchor: Record<string, unknown>, kind: 'pic' | 'graphicFrame'): boolean {
  const node = anchor[kind];
  if (!node || typeof node !== 'object') return false;
  const nv = (node as Record<string, unknown>)[kind === 'pic' ? 'nvPicPr' : 'nvGraphicFramePr'];
  if (!nv || typeof nv !== 'object') return false;
  const cNvPr = (nv as Record<string, unknown>)['cNvPr'];
  if (!cNvPr || typeof cNvPr !== 'object') return false;
  const hidden = (cNvPr as Record<string, unknown>)['@_hidden'];
  return hidden === '1' || hidden === 'true' || hidden === true;
}

// xdr:pic → xdr:blipFill → a:blip @r:embed (removeNSPrefix → pic/blipFill/blip,
// r:embed → @_embed, mirroring chartRelIdOf's r:id → @_id).
function picRelIdOf(anchor: Record<string, unknown>): string | undefined {
  const pic = anchor['pic'];
  if (!pic || typeof pic !== 'object') return undefined;
  const blipFill = (pic as Record<string, unknown>)['blipFill'];
  if (!blipFill || typeof blipFill !== 'object') return undefined;
  const blip = (blipFill as Record<string, unknown>)['blip'];
  if (!blip || typeof blip !== 'object') return undefined;
  const embed = (blip as Record<string, unknown>)['@_embed'];
  return typeof embed === 'string' && embed !== '' ? embed : undefined;
}

// xdr:graphicFrame → a:graphic → a:graphicData[uri=chart] → c:chart @r:id.
function chartRelIdOf(anchor: Record<string, unknown>): string | undefined {
  const frame = anchor['graphicFrame'];
  if (!frame || typeof frame !== 'object') return undefined;
  const graphic = (frame as Record<string, unknown>)['graphic'];
  if (!graphic || typeof graphic !== 'object') return undefined;
  const data = (graphic as Record<string, unknown>)['graphicData'];
  if (!data || typeof data !== 'object') return undefined;
  const d = data as Record<string, unknown>;
  if (d['@_uri'] !== CHART_URI) return undefined;
  const chart = d['chart'];
  if (!chart || typeof chart !== 'object') return undefined;
  const id = (chart as Record<string, unknown>)['@_id'];
  return typeof id === 'string' && id !== '' ? id : undefined;
}

interface CellMarker {
  readonly col: number;
  readonly colOffPt: number;
  readonly row: number;
  readonly rowOffPt: number;
}

// §20.5.2.24/.25 — <xdr:from>/<xdr:to>: col/colOff(EMU)/row/rowOff(EMU).
function cellMarker(node: unknown): CellMarker | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const m = node as Record<string, unknown>;
  const col = num(m['col']);
  const row = num(m['row']);
  if (col === undefined || row === undefined) return undefined;
  return {
    col,
    row,
    colOffPt: emuToPt(num(m['colOff']) ?? 0),
    rowOffPt: emuToPt(num(m['rowOff']) ?? 0),
  };
}

// Distance between two cell markers along one axis: full tracks in
// [from..to) plus the offset difference.
function spanPt(
  from: number,
  fromOffPt: number,
  to: number,
  toOffPt: number,
  trackPt: (index: number) => number,
): number {
  let span = 0;
  for (let i = from; i < to; i++) span += trackPt(i);
  return span - fromOffPt + toOffPt;
}

/**
 * Build a `col → width-in-points` accessor: a `<col>` override (Excel "chars")
 * where present, else the default — the same conversions the print model uses.
 * Exported so the sheet-shape parser (E-SHEET W2) sizes shape anchors with the same
 * track geometry.
 */
export function makeColWidthPt(ws: ParsedWorksheet): (col: number) => number {
  // §18.3.1.13: a rendered column is `chars × MDW + 5px`, and the 5px is not
  // optional — the grid has always added it. Here it was dropped, so an anchor
  // drifted 3.75pt left for every explicitly-sized column before it, and the
  // drawing and the cell it is anchored to disagreed about where that column
  // starts. shape-macro-ext-ref.xlsx put its macro button 3pt short of the
  // column band its own anchor names.
  const widthPt = (chars: number): number =>
    columnTwips(chars, TWIPS_PER_EXCEL_CHAR) / TWIPS_PER_PT;
  return (col: number): number => {
    for (const c of ws.columns) {
      if (col >= c.min - 1 && col <= c.max - 1) {
        return widthPt(c.widthChars);
      }
    }
    // §18.3.1.81 `<sheetFormatPr defaultColWidth>` governs every column no
    // `<col>` covers — the same fallback the grid uses. Hardcoding Excel's
    // 8.43 characters contradicted the file's own declaration and sized every
    // shape anchor against a track the sheet does not have.
    // §18.3.1.81 — `defaultColWidth` already includes the margin and gridline
    // padding; `baseColWidth` explicitly does not, so a default derived from it
    // takes the padding once more. The anchor tracks have to agree with the
    // grid's columns or a drawing lands beside the cell it is anchored to.
    if (ws.defaultColWidthChars !== undefined) return widthPt(ws.defaultColWidthChars);
    if (ws.baseColWidthChars !== undefined) {
      return (
        (columnTwips(ws.baseColWidthChars, TWIPS_PER_EXCEL_CHAR) + COL_PADDING_TWIPS) / TWIPS_PER_PT
      );
    }
    // DEFAULT_COL_TWIPS is 960 — Excel's 8.43 characters WITH the padding.
    return DEFAULT_COL_TWIPS / TWIPS_PER_PT;
  };
}

/**
 * Build a `row → height-in-points` accessor: the explicit row height where present,
 * else the default. Pairs with {@link makeColWidthPt} for the anchor track geometry.
 */
export function makeRowHeightPt(ws: ParsedWorksheet): (row: number) => number {
  return (row: number): number => {
    for (const r of ws.rowHeights) {
      if (r.row === row) return r.heightPt;
    }
    // §18.3.1.81 `defaultRowHeight` likewise: bnc762542.xlsx declares 12.75pt
    // and we measured its anchored box against 15, which alone made the shape
    // 18% too tall.
    return ws.defaultRowHeightPt ?? DEFAULT_ROW_TWIPS / TWIPS_PER_PT;
  };
}

function num(v: unknown): number | undefined {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (typeof v === 'string' && v !== '') {
    const n = Number(v);
    if (Number.isFinite(n)) return n;
  }
  return undefined;
}

function asArray(v: unknown): Array<unknown> {
  return Array.isArray(v) ? v : v !== undefined ? [v] : [];
}
