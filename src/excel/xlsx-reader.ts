// XLSX reader (ir-design §7 + E-SHEET): bytes → SheetDoc → FlowDoc. The reader
// builds the SpreadsheetML IR node (readXlsxToSheetDoc) — a workbook of grid
// sheets with its style table, shared strings, defined names and resolved
// charts — and the print model projects it to flow blocks (projectSheetDoc).
// Document-derived state only; caller conversion options stay with the
// converter/facade.

import type { Chart, DocumentInfo, ShapeBlock } from '@/core/document-model';
import type { CoreProperties, Relationship } from '@/core/opc';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type {
  Sheet,
  SheetChartRef,
  SheetDoc,
  SheetHyperlink,
  SheetImageRef,
  SheetSlicer,
  SheetSlicerItem,
} from '@/core/ir/sheet';
import type {
  ExcelTable,
  MergedRange,
  PivotTable,
  WorksheetCell,
  XlsxStyles,
} from '@/core/spreadsheet-model';
import type { TableFilterColumn } from '@/excel/table-parser';
import type { SlicerCacheDef, SlicerDef } from '@/excel/slicer-parser';

import { FEATURES, ResourceStore } from '@/core/ir';
import { OpcPackage, isOoxmlRel, parseCoreProperties } from '@/core/opc';
import {
  EMPTY_XLSX_STYLES,
  parseAreaRef,
  parseSharedStrings,
  parseWorkbook,
  parseWorksheet,
  parseXlsxStyles,
} from '@/excel';
import { bytesInclude } from '@/core/bytes';
import { parseChart, withChartColorStyle } from '@/core/drawingml/chart-parser';
import { DEFAULT_THEME_PALETTE, makeColorResolver } from '@/core/drawingml/colors';
import { parseTheme } from '@/core/drawingml/theme-parser';
import { parseSheetDrawing } from '@/excel/sheet-drawing';
import { parseTablePartFull } from '@/excel/table-parser';
import { parsePivotTablePart } from '@/excel/pivot-table-parser';
import { parseSlicerCachePart, parseSlicerPart } from '@/excel/slicer-parser';
import { parseSheetShapes } from '@/excel/sheet-shape-parser';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { resolveCellText } from '@/excel/print-model';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';
// MS relationship type tails for slicer parts (E-SHEET SV2). The worksheet
// references its slicer parts; the workbook references the slicer caches.
const SLICER_REL_TAIL = '/slicer';
const SLICER_CACHE_REL_TAIL = '/slicerCache';
// A slicer over a huge column is bounded so a crafted file cannot blow up the box.
const MAX_SLICER_ITEMS = 256;

// A table's location, header depth and value filters — indexed by the table's
// numeric id so a slicer's <tableSlicerCache> can resolve its column's values
// (E-SHEET SV2). Reader-internal; not part of the persisted grid model.
interface TableLoc {
  readonly cells: ReadonlyArray<WorksheetCell>;
  readonly ref: MergedRange;
  readonly headerRows: number;
  readonly filters: ReadonlyArray<TableFilterColumn>;
}

export function readXlsx(xlsx: Uint8Array): ReadResult<FlowDoc> {
  return { doc: projectSheetDoc(readXlsxToSheetDoc(xlsx)), losses: [] };
}

// bytes → SheetDoc: the SpreadsheetML IR node. Everything that needs the OPC
// package (parsing + chart/drawing resolution) happens here; the SheetDoc →
// FlowDoc projection (sheet-to-flow) is then a pure transform.
export function readXlsxToSheetDoc(xlsx: Uint8Array): SheetDoc {
  const pkg = OpcPackage.open(xlsx);
  const workbookData = pkg.getPart(WORKBOOK_PART);
  if (!workbookData) throw new Error('Not a valid xlsx: missing xl/workbook.xml');
  const { sheets, date1904, definedNames } = parseWorkbook(workbookData);
  if (sheets.length === 0) throw new Error('xlsx has no sheets');

  const sharedStringsData = pkg.getPart(SHARED_STRINGS_PART);
  const sharedStrings = sharedStringsData ? parseSharedStrings(sharedStringsData) : [];

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseXlsxStyles(stylesData) : EMPTY_XLSX_STYLES;

  const workbookRels = pkg.getPartRelationships(WORKBOOK_PART);
  // Charts keyed by their part path (globally unique across sheets); the
  // theme-backed resolver mirrors the docx reader's so schemeClr references in
  // charts resolve to the workbook's actual accents.
  const chartData = new Map<string, Chart>();
  // Content-addressed store for sheet pictures (W1); populated below as anchors
  // resolve, then handed to the SheetDoc so the renderer can fetch image bytes.
  const resources = new ResourceStore();
  const palette = buildThemePalette(pkg, workbookRels);
  const resolveColor = makeColorResolver(palette);

  const sheetsOut: Array<Sheet> = [];
  // §SV2 slicer-resolution state: tables indexed by id (a slicer's
  // tableSlicerCache may reference a table on another sheet) and the slicer parts
  // found per OUTPUT sheet — both consumed after the loop, once all tables index.
  const tableIndex = new Map<number, TableLoc>();
  const pendingSlicers: Array<{ outIdx: number; defs: ReadonlyArray<SlicerDef> }> = [];
  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    const worksheet = parseWorksheet(resolved.data);

    // §20.5: the sheet's drawing part — resolve chart frames, pictures and shapes
    // here; the projection emits a block per frame after the grid (W1 pictures,
    // W2 shapes).
    const charts: Array<SheetChartRef> = [];
    const images: Array<SheetImageRef> = [];
    let shapes: Array<ShapeBlock> | undefined;
    if (worksheet.drawingRelId) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const drawingRel = wsRels.find((r) => r.id === worksheet.drawingRelId);
      const drawing = drawingRel ? pkg.resolveRelatedPart(resolved.path, drawingRel) : undefined;
      if (drawing) {
        const { charts: chartRefs, pictures } = parseSheetDrawing(
          drawing.data,
          drawing.path,
          pkg,
          worksheet,
        );
        for (const ref of chartRefs) {
          if (!chartData.has(ref.chartPartPath)) {
            const chartXml = pkg.getPart(ref.chartPartPath);
            const parsed = chartXml ? parseChart(chartXml, resolveColor) : null;
            if (!parsed) continue;
            chartData.set(
              ref.chartPartPath,
              withChartColorStyle(parsed, pkg, ref.chartPartPath, resolveColor),
            );
          }
          charts.push({
            chartPartPath: ref.chartPartPath,
            widthPt: ref.widthPt,
            heightPt: ref.heightPt,
          });
        }
        for (const pic of pictures) {
          const bytes = pkg.getPart(pic.imagePartPath);
          if (!bytes) continue;
          images.push({
            resourceId: resources.put(bytes),
            widthPt: pic.widthPt,
            heightPt: pic.heightPt,
          });
        }
        // §20.5.2.30 xdr:sp shapes (W2). The shared DrawingML readers need the
        // preserveOrder PoNode tree, so shapes parse the drawing a second time —
        // gated on a shape open tag (`:sp>`/`:sp `) so chart/picture-only drawings
        // skip it (xdr:spPr / xdr:grpSp do not match).
        if (bytesInclude(drawing.data, ':sp>') || bytesInclude(drawing.data, ':sp ')) {
          const parsed = parseSheetShapes(drawing.data, worksheet, resolveColor);
          if (parsed.length > 0) shapes = parsed;
        }
      }
    }
    // §18.5: the sheet's table parts — resolve each relationship to its
    // tableN.xml, parse it, and resolve its named style to fill colours against
    // the workbook accent (E-SHEET SC3). The projection applies banded shading.
    let tables: Array<ExcelTable> | undefined;
    if (worksheet.tablePartRelIds && worksheet.tablePartRelIds.length > 0) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const resolvedTables: Array<ExcelTable> = [];
      for (const rid of worksheet.tablePartRelIds) {
        const rel = wsRels.find((r) => r.id === rid);
        const part = rel ? pkg.resolveRelatedPart(resolved.path, rel) : undefined;
        const full = part ? parseTablePartFull(part.data) : undefined;
        if (!full) continue;
        resolvedTables.push(resolveTableStyle(full.table, palette));
        // Index the table by id so a slicer can resolve its column (E-SHEET SV2).
        if (full.id !== undefined) {
          tableIndex.set(full.id, {
            cells: worksheet.cells,
            ref: full.table.ref,
            headerRows: full.table.headerRowCount,
            filters: full.filters,
          });
        }
      }
      if (resolvedTables.length > 0) tables = resolvedTables;
    }

    // §18.3 — the sheet's slicer parts (E-SHEET SV2), resolved after the loop.
    // Record the OUTPUT index this sheet will occupy: continued (skipped) sheets
    // make the output index trail the source index.
    {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const defs: Array<SlicerDef> = [];
      for (const rel of wsRels) {
        if (!rel.type.endsWith(SLICER_REL_TAIL)) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        if (part) defs.push(...parseSlicerPart(part.data));
      }
      if (defs.length > 0) pendingSlicers.push({ outIdx: sheetsOut.length, defs });
    }

    // §18.10: the sheet's pivot tables — referenced ONLY via the worksheet's
    // relationships (there is no element in the sheet XML), so enumerate the rels
    // by type. The output cells are already cached in the grid; PV1 just records
    // the location + named style for PV2 to band (E-PIVOT).
    let pivotTables: Array<PivotTable> | undefined;
    {
      const resolvedPivots: Array<PivotTable> = [];
      for (const rel of pkg.getPartRelationships(resolved.path)) {
        if (!isOoxmlRel(rel.type, 'pivotTable')) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        const parsed = part ? parsePivotTablePart(part.data) : undefined;
        if (parsed) resolvedPivots.push(resolvePivotStyle(parsed, palette));
      }
      if (resolvedPivots.length > 0) pivotTables = resolvedPivots;
    }

    // §18.3.1.47 cell hyperlinks (W3): resolve each relId to its external URL via
    // the worksheet rels; location-only (in-workbook) links carry no URL.
    let hyperlinks: Array<SheetHyperlink> | undefined;
    if (worksheet.hyperlinks && worksheet.hyperlinks.length > 0) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const resolvedLinks: Array<SheetHyperlink> = [];
      for (const h of worksheet.hyperlinks) {
        if (h.relId === undefined) continue; // in-workbook location link → no URL
        const rel = wsRels.find((r) => r.id === h.relId);
        if (!rel || rel.targetMode !== 'External' || !rel.target) continue;
        const area = parseAreaRef(h.ref);
        if (!area) continue;
        resolvedLinks.push({
          ref: {
            startColumn: area.startColumn,
            startRow: area.startRow,
            endColumn: area.endColumn,
            endRow: area.endRow,
          },
          url: rel.target,
        });
      }
      if (resolvedLinks.length > 0) hyperlinks = resolvedLinks;
    }

    const grid =
      tables || pivotTables
        ? { ...worksheet, ...(tables ? { tables } : {}), ...(pivotTables ? { pivotTables } : {}) }
        : worksheet;
    sheetsOut.push({
      name: sheet.name,
      grid,
      ...(charts.length > 0 ? { charts } : {}),
      ...(images.length > 0 ? { images } : {}),
      ...(shapes ? { shapes } : {}),
      ...(hyperlinks ? { hyperlinks } : {}),
    });
  }

  // §SV2 — resolve slicer panels now that every table is indexed. Slicer caches
  // are workbook-scoped (referenced from the workbook rels); each binds by name
  // to a slicer's @cache. A panel renders after its sheet's grid + charts.
  if (pendingSlicers.length > 0) {
    const cacheByName = new Map<string, SlicerCacheDef>();
    for (const rel of workbookRels) {
      if (!rel.type.endsWith(SLICER_CACHE_REL_TAIL)) continue;
      const part = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
      const cache = part ? parseSlicerCachePart(part.data) : undefined;
      if (cache) cacheByName.set(cache.name, cache);
    }
    for (const { outIdx, defs } of pendingSlicers) {
      const sheet = sheetsOut[outIdx];
      if (!sheet) continue;
      const slicers = defs.map((def) =>
        resolveSlicer(
          def,
          cacheByName.get(def.cacheName),
          tableIndex,
          styles,
          sharedStrings,
          date1904,
          palette,
        ),
      );
      sheetsOut[outIdx] = { ...sheet, slicers };
    }
  }

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;
  const info = infoFromCore(coreProps);

  return {
    kind: 'sheet',
    sheets: sheetsOut,
    styles,
    sharedStrings,
    definedNames,
    date1904,
    ...(chartData.size > 0 ? { chartData } : {}),
    resources,
    ...(info ? { info } : {}),
  };
}

// Theme palette: the workbook's theme part merged over the built-in Office
// defaults (the docx reader's pattern). Drives both chart schemeClr resolution
// and the table-style accent (E-SHEET SC3).
function buildThemePalette(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): Map<string, string> {
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (!resolved) continue;
    for (const [slot, hex] of parseTheme(resolved.data)) palette.set(slot, hex);
    break;
  }
  return palette;
}

// Resolve a table's named built-in style to header / band fill colours against
// the workbook theme. The definitions live in Excel (not the file), but the
// name encodes the accent: TableStyle{Light|Medium|Dark}{N}, where the gallery
// is 7 columns wide — column (N-1)%7 picks the accent (0 = the neutral/grey
// column, 1..6 = accent1..6). Medium/Dark take a solid accent header with white
// text; Light a tinted header with black text; the band is a light tint. A
// style-less / unrecognized / TableStyleNone table is left uncoloured.
function resolveTableStyle(t: ExcelTable, palette: ReadonlyMap<string, string>): ExcelTable {
  const m = t.styleName ? /TableStyle(Light|Medium|Dark)(\d+)/i.exec(t.styleName) : null;
  if (!m) return t;
  const kind = m[1]!.toLowerCase();
  const column = (Number(m[2]) - 1) % 7;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  if (kind === 'light') {
    return { ...t, headerHex: lighten(base, 0.6), bandHex: lighten(base, 0.85) };
  }
  // medium / dark: a solid accent header with white text.
  return { ...t, headerHex: base, bandHex: lighten(base, 0.8), headerTextHex: 'FFFFFF' };
}

// Resolve a pivot's named built-in style to header / band colours. Pivot styles
// (PivotStyle{Light|Medium|Dark}{N}) live in Excel, not the file; we approximate
// with the same accent-column heuristic as table styles — the pivot gallery
// differs in exact numbering, refined later (E-PIVOT PV2). A style-less /
// unrecognized pivot is left uncoloured (it then renders as a plain grid).
function resolvePivotStyle(p: PivotTable, palette: ReadonlyMap<string, string>): PivotTable {
  const m = p.styleName ? /PivotStyle(Light|Medium|Dark)(\d+)/i.exec(p.styleName) : null;
  if (!m) return p;
  const kind = m[1]!.toLowerCase();
  const column = (Number(m[2]) - 1) % 7;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  if (kind === 'light') {
    return { ...p, headerHex: lighten(base, 0.6), bandHex: lighten(base, 0.85) };
  }
  // medium / dark: a solid accent header with white text.
  return { ...p, headerHex: base, bandHex: lighten(base, 0.8), headerTextHex: 'FFFFFF' };
}

// §SV2 — resolve a slicer definition + its cache into a renderable panel: the
// caption, the column-button items with selection, and the style accent. A
// native-table cache resolves items from the referenced table column; an
// OLAP/pivot cache (no <tableSlicerCache>) yields a caption-only box.
function resolveSlicer(
  def: SlicerDef,
  cache: SlicerCacheDef | undefined,
  tableIndex: ReadonlyMap<number, TableLoc>,
  styles: XlsxStyles,
  sharedStrings: ReadonlyArray<string>,
  date1904: boolean,
  palette: ReadonlyMap<string, string>,
): SheetSlicer {
  const items = cache
    ? resolveTableSlicerItems(cache, tableIndex, styles, sharedStrings, date1904)
    : [];
  return {
    caption: def.caption,
    columnCount: def.columnCount,
    items,
    ...resolveSlicerStyle(def.styleName, palette),
  };
}

// The distinct values of the slicer's table column become its buttons (first-seen
// row order, header rows skipped, bounded by MAX_SLICER_ITEMS). Selection comes
// from the column's autofilter when present; an unfiltered column shows
// everything as selected (a freshly-saved slicer's default).
function resolveTableSlicerItems(
  cache: SlicerCacheDef,
  tableIndex: ReadonlyMap<number, TableLoc>,
  styles: XlsxStyles,
  sharedStrings: ReadonlyArray<string>,
  date1904: boolean,
): Array<SheetSlicerItem> {
  const loc = cache.tableId !== undefined ? tableIndex.get(cache.tableId) : undefined;
  if (!loc || cache.columnId === undefined) return [];
  const colOffset = cache.columnId - 1; // table column id (1-based) → range offset
  const absCol = loc.ref.startColumn + colOffset;
  if (absCol < loc.ref.startColumn || absCol > loc.ref.endColumn) return [];
  const firstDataRow = loc.ref.startRow + Math.max(1, loc.headerRows);
  const filter = loc.filters.find((f) => f.colId === colOffset);
  const kept = filter ? new Set(filter.values) : undefined;
  // The column's data cells in row order, resolved to display text.
  const byRow = new Map<number, string>();
  for (const c of loc.cells) {
    if (c.column !== absCol || c.row < firstDataRow || c.row > loc.ref.endRow) continue;
    byRow.set(c.row, resolveCellText(c, sharedStrings, styles, date1904));
  }
  const seen = new Set<string>();
  const out: Array<SheetSlicerItem> = [];
  for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
    const label = byRow.get(row)!;
    if (label === '' || seen.has(label)) continue;
    seen.add(label);
    out.push({ label, selected: kept ? kept.has(label) : true });
    if (out.length >= MAX_SLICER_ITEMS) break;
  }
  return out;
}

// SlicerStyle{Light|Dark|Other}{N} → an accent for the header + selected buttons
// (the same accent-column heuristic as table/pivot styles; the slicer gallery's
// exact numbering differs, refined later). White text reads on the accent; a
// style-less slicer falls back to accent1.
function resolveSlicerStyle(
  styleName: string | undefined,
  palette: ReadonlyMap<string, string>,
): { headerHex: string; headerTextHex: string; selectedHex: string; selectedTextHex: string } {
  const m = styleName ? /SlicerStyle[A-Za-z]*?(\d+)/i.exec(styleName) : null;
  const column = m ? (Number(m[1]) - 1) % 7 : 1;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  return { headerHex: base, headerTextHex: 'FFFFFF', selectedHex: base, selectedTextHex: 'FFFFFF' };
}

// Lighten a 6-hex colour toward white by `amount` (0..1).
function lighten(hex: string, amount: number): string {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex);
  const n = m ? parseInt(m[1]!, 16) : 0x4472c4;
  const ch = (shift: number): number => {
    const c = (n >> shift) & 0xff;
    return Math.round(c + (255 - c) * amount);
  };
  const hx = (c: number): string => c.toString(16).padStart(2, '0').toUpperCase();
  return `${hx(ch(16))}${hx(ch(8))}${hx(ch(0))}`;
}

export const xlsxReader: DocumentReader<SheetDoc> = {
  id: 'xlsx',
  // The reader's native tree is the SheetDoc; the facade/Ream project it to a
  // FlowDoc for rendering (E-SHEET SB1).
  produces: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesInclude(bytes, 'xl/workbook.xml'),
  read: (bytes) => ({ doc: readXlsxToSheetDoc(bytes), losses: [] }),
};

function infoFromCore(core: CoreProperties | undefined): DocumentInfo | undefined {
  if (!core) return undefined;
  return {
    ...(core.title ? { title: core.title } : {}),
    ...(core.creator ? { author: core.creator } : {}),
    ...(core.subject ? { subject: core.subject } : {}),
    ...(core.keywords ? { keywords: core.keywords } : {}),
    ...(core.created ? { creationDate: core.created } : {}),
    ...(core.modified ? { modificationDate: core.modified } : {}),
  };
}
