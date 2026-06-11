// XLSX reader (ir-design §7): bytes → FlowDoc. The grid is projected through
// the Excel print model into flow blocks (SheetDoc — the dedicated grid tree —
// is deliberately deferred tech debt; see handoff v1 §1). Document-derived
// state only; caller conversion options stay with the converter/facade.

import type { BodyElement, Chart, DocumentInfo, SectionProperties } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { CoreProperties, Relationship } from '@/core/opc';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';

import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage, parseCoreProperties } from '@/core/opc';
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
import {
  resolvePrintArea,
  resolvePrintTitleRows,
  sectionFromWorksheet,
  worksheetToBody,
} from '@/excel/print-model';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';

export function readXlsx(xlsx: Uint8Array): ReadResult<FlowDoc> {
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
  const body: Array<BodyElement> = [];
  // Charts keyed by their part path (globally unique across sheets); the
  // theme-backed resolver mirrors the docx reader's so schemeClr references
  // in charts resolve to the workbook's actual accents.
  const charts = new Map<string, Chart>();
  const resolveColor = buildXlsxColorResolver(pkg, workbookRels);

  // Page geometry comes from the first sheet's <pageSetup>/<pageMargins>.
  // The renderer only supports one section, so subsequent sheets share the
  // first sheet's geometry; multi-section support is M2/M4-grade work.
  let firstSheetSection: SectionProperties | undefined;

  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    const worksheet = parseWorksheet(resolved.data);
    if (sheetIdx === 0) {
      firstSheetSection = sectionFromWorksheet(worksheet);
    }

    // Each sheet after the first starts on its own PDF page. We do NOT print the
    // sheet name: LibreOffice Calc / Excel `--convert-to pdf` emit it nowhere in
    // the body (nor the default header) — a synthetic title is pure extra text
    // that diverges from the print golden. So emit an empty page-break-only
    // paragraph (no runs ⇒ no glyphs ⇒ no text) for sheets > 0.
    if (sheetIdx > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: { properties: { pageBreakBefore: true }, runs: [] },
      });
    }

    const printArea = resolvePrintArea(definedNames, sheetIdx);
    const titleRows = resolvePrintTitleRows(definedNames, sheetIdx);
    const gridLines = worksheet.printOptions?.gridLines === true;
    body.push(
      ...worksheetToBody(worksheet, sharedStrings, styles, date1904, {
        ...(printArea ? { printArea } : {}),
        ...(titleRows ? { titleRows } : {}),
        gridLines,
      }),
    );

    // §20.5: the sheet's drawing part — chart frames render as blocks after
    // the sheet's grid, anchor-ordered (positional overlay is out of scope).
    if (worksheet.drawingRelId) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const drawingRel = wsRels.find((r) => r.id === worksheet.drawingRelId);
      const drawing = drawingRel ? pkg.resolveRelatedPart(resolved.path, drawingRel) : undefined;
      if (drawing) {
        for (const ref of parseSheetDrawing(drawing.data, drawing.path, pkg, worksheet)) {
          if (!charts.has(ref.chartPartPath)) {
            const chartData = pkg.getPart(ref.chartPartPath);
            const parsed = chartData ? parseChart(chartData, resolveColor) : null;
            if (!parsed) continue;
            charts.set(
              ref.chartPartPath,
              withChartColorStyle(parsed, pkg, ref.chartPartPath, resolveColor),
            );
          }
          body.push({
            kind: 'chart',
            chart: {
              chartRelId: ref.chartPartPath,
              width: pt(ref.widthPt),
              height: pt(ref.heightPt),
              paragraphProperties: {},
            },
          });
        }
      }
    }
  }

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;
  const info = infoFromCore(coreProps);

  const doc: FlowDoc = {
    kind: 'flow',
    // Same stage-6 contract as docx: the body carries resolved properties.
    // Grid cells are built with direct props only, so resolving over the
    // empty sheet just materializes the defaults.
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    ...(firstSheetSection ? { section: firstSheetSection } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
    ...(charts.size > 0 ? { charts } : {}),
    ...(info ? { info } : {}),
  };
  return { doc, losses: [] };
}

// Theme palette for chart schemeClr resolution: the workbook's theme part
// merged over the built-in Office defaults (the docx reader's pattern).
function buildXlsxColorResolver(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): ColorResolver {
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const rel of workbookRels) {
    if (rel.type !== REL_THEME) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (!resolved) continue;
    for (const [slot, hex] of parseTheme(resolved.data)) palette.set(slot, hex);
    break;
  }
  return makeColorResolver(palette);
}

const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

export const xlsxReader: DocumentReader<FlowDoc> = {
  id: 'xlsx',
  produces: 'flow',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesInclude(bytes, 'xl/workbook.xml'),
  read: (bytes) => readXlsx(bytes),
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
