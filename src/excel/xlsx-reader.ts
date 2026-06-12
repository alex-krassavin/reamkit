// XLSX reader (ir-design §7 + E-SHEET): bytes → SheetDoc → FlowDoc. The reader
// builds the SpreadsheetML IR node (readXlsxToSheetDoc) — a workbook of grid
// sheets with its style table, shared strings, defined names and resolved
// charts — and the print model projects it to flow blocks (projectSheetDoc).
// Document-derived state only; caller conversion options stay with the
// converter/facade.

import type { Chart, DocumentInfo } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { CoreProperties, Relationship } from '@/core/opc';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Sheet, SheetChartRef, SheetDoc } from '@/core/ir/sheet';

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
  const resolveColor = buildXlsxColorResolver(pkg, workbookRels);

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
    sheetsOut.push({ name: sheet.name, grid: worksheet, ...(charts.length > 0 ? { charts } : {}) });
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

// Theme palette for chart schemeClr resolution: the workbook's theme part
// merged over the built-in Office defaults (the docx reader's pattern).
function buildXlsxColorResolver(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): ColorResolver {
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (!resolved) continue;
    for (const [slot, hex] of parseTheme(resolved.data)) palette.set(slot, hex);
    break;
  }
  return makeColorResolver(palette);
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
