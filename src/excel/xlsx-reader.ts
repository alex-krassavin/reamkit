// XLSX reader (ir-design §7 + E-SHEET): bytes → SheetDoc → FlowDoc. The reader
// builds the SpreadsheetML IR node (readXlsxToSheetDoc) — a workbook of grid
// sheets with its style table, shared strings, defined names and resolved
// charts — and the print model projects it to flow blocks (projectSheetDoc).
// Document-derived state only; caller conversion options stay with the
// converter/facade.

import type { Chart, DocumentInfo } from '@/core/document-model';
import type { CoreProperties, Relationship } from '@/core/opc';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Sheet, SheetChartRef, SheetDoc } from '@/core/ir/sheet';
import type { ExcelTable, PivotTable } from '@/core/spreadsheet-model';

import { FEATURES, ResourceStore } from '@/core/ir';
import { OpcPackage, isOoxmlRel, parseCoreProperties } from '@/core/opc';
import {
  EMPTY_XLSX_STYLES,
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
import { parseTablePart } from '@/excel/table-parser';
import { parsePivotTablePart } from '@/excel/pivot-table-parser';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';

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
  const palette = buildThemePalette(pkg, workbookRels);
  const resolveColor = makeColorResolver(palette);

  const sheetsOut: Array<Sheet> = [];
  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    const worksheet = parseWorksheet(resolved.data);

    // §20.5: the sheet's drawing part — resolve chart frames (and their chart
    // parts) here; the projection emits a block per frame after the grid.
    const charts: Array<SheetChartRef> = [];
    if (worksheet.drawingRelId) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const drawingRel = wsRels.find((r) => r.id === worksheet.drawingRelId);
      const drawing = drawingRel ? pkg.resolveRelatedPart(resolved.path, drawingRel) : undefined;
      if (drawing) {
        for (const ref of parseSheetDrawing(drawing.data, drawing.path, pkg, worksheet)) {
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
        const parsed = part ? parseTablePart(part.data) : undefined;
        if (parsed) resolvedTables.push(resolveTableStyle(parsed, palette));
      }
      if (resolvedTables.length > 0) tables = resolvedTables;
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
        if (parsed) resolvedPivots.push(parsed);
      }
      if (resolvedPivots.length > 0) pivotTables = resolvedPivots;
    }

    const grid =
      tables || pivotTables
        ? { ...worksheet, ...(tables ? { tables } : {}), ...(pivotTables ? { pivotTables } : {}) }
        : worksheet;
    sheetsOut.push({ name: sheet.name, grid, ...(charts.length > 0 ? { charts } : {}) });
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
    resources: new ResourceStore(),
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
