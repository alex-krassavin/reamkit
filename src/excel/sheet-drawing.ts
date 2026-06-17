// Sheet drawings (ECMA-376 §20.5 SpreadsheetDrawingML): the worksheet's
// <drawing r:id> points at xl/drawings/drawingN.xml, whose anchors place
// graphic frames over the grid. v1 extracts CHART frames only (shapes and
// images on sheets stay out of scope) and renders each chart as a block after
// the sheet's table — anchor-ordered, sized from the cell-range anchor.

import { XMLParser } from 'fast-xml-parser';
import type { OpcPackage } from '@/core/opc';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';

import { emuToPt } from '@/core/ir';
import { DEFAULT_COL_TWIPS, DEFAULT_ROW_TWIPS, TWIPS_PER_EXCEL_CHAR } from '@/excel/print-model';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

export interface SheetChartRef {
  // The resolved chart part path ('xl/charts/chart1.xml') — globally unique,
  // used as the FlowDoc charts key and the ChartBlock.chartRelId.
  readonly chartPartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
  // Anchor top row (0-based) — used only to order charts on the sheet.
  readonly anchorRow: number;
}

// §20.5.2.1 xdr:pic — a picture anchored over the grid. The reader reads the
// resolved media part's bytes into the SheetDoc resource store.
export interface SheetPicture {
  readonly imagePartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly anchorRow: number;
}

// Both kinds of anchored frame the drawing yields, anchor-ordered.
export interface SheetDrawing {
  readonly charts: Array<SheetChartRef>;
  readonly pictures: Array<SheetPicture>;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  parseTagValue: false,
});

const TWIPS_PER_PT = 20;

// §20.5.2.35/.33/.1 — xdr:twoCellAnchor / oneCellAnchor / absoluteAnchor. Each
// anchor frames either a chart (graphicFrame) or a picture (xdr:pic); both are
// sized from the anchor and returned anchor-ordered.
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

      if (chartRelId) {
        const path = partPathOf(chartRelId);
        if (path) charts.push({ chartPartPath: path, widthPt, heightPt, anchorRow });
      } else if (picRelId) {
        const path = partPathOf(picRelId);
        if (path) pictures.push({ imagePartPath: path, widthPt, heightPt, anchorRow });
      }
    }
  }
  charts.sort((x, y) => x.anchorRow - y.anchorRow);
  pictures.sort((x, y) => x.anchorRow - y.anchorRow);
  return { charts, pictures };
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

// Column width in points: <col> overrides (Excel "chars"), else the default —
// the same conversions the print model uses.
function makeColWidthPt(ws: ParsedWorksheet): (col: number) => number {
  return (col: number): number => {
    for (const c of ws.columns) {
      if (col >= c.min - 1 && col <= c.max - 1) {
        return (c.widthChars * TWIPS_PER_EXCEL_CHAR) / TWIPS_PER_PT;
      }
    }
    return DEFAULT_COL_TWIPS / TWIPS_PER_PT;
  };
}

function makeRowHeightPt(ws: ParsedWorksheet): (row: number) => number {
  return (row: number): number => {
    for (const r of ws.rowHeights) {
      if (r.row === row) return r.heightPt;
    }
    return DEFAULT_ROW_TWIPS / TWIPS_PER_PT;
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
