// SheetDoc → FlowDoc projection (E-SHEET SA2). The print model turns each grid
// sheet into flow blocks (a table + chart frames); sheets after the first start
// on a new page. This is the SAME projection the xlsx reader used to inline —
// relocated behind the SheetDoc boundary, byte-for-byte unchanged — so the
// render path (PDF/SVG/HTML) is identical. A dedicated grid layout would be a
// separate SheetDoc consumer; for now FlowDoc is the one projection.

import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { SheetDoc } from '@/core/ir/sheet';

import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { pt } from '@/core/ir';
import {
  resolvePrintArea,
  resolvePrintTitleRows,
  sectionFromWorksheet,
  slicerTable,
  worksheetToBody,
} from '@/excel/print-model';

export function projectSheetDoc(sheet: SheetDoc): FlowDoc {
  const body: Array<BodyElement> = [];
  // Page geometry comes from the first sheet's <pageSetup>/<pageMargins>; the
  // renderer supports one section, so later sheets share the first's geometry.
  let firstSheetSection: SectionProperties | undefined;

  // Sheet name → grid, so a sparkline whose data range is sheet-qualified
  // (Sheet2!A1:C1) resolves against the right sheet (E-SHEET SC2 tail TC3).
  const sheetGrids = new Map(sheet.sheets.map((s) => [s.name, s.grid]));

  for (let sheetIdx = 0; sheetIdx < sheet.sheets.length; sheetIdx++) {
    const ws = sheet.sheets[sheetIdx]!;
    if (sheetIdx === 0) {
      firstSheetSection = sectionFromWorksheet(ws.grid);
    }

    // Each sheet after the first starts on its own PDF page. We do NOT print the
    // sheet name (Calc/Excel `--convert-to pdf` emit it nowhere), so the page
    // break is an empty page-break-only paragraph (no runs ⇒ no glyphs).
    if (sheetIdx > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: { properties: { pageBreakBefore: true }, runs: [] },
      });
    }

    const printArea = resolvePrintArea(sheet.definedNames, sheetIdx);
    const titleRows = resolvePrintTitleRows(sheet.definedNames, sheetIdx);
    const gridLines = ws.grid.printOptions?.gridLines === true;
    body.push(
      ...worksheetToBody(ws.grid, sheet.sharedStrings, sheet.styles, sheet.date1904, {
        ...(printArea ? { printArea } : {}),
        ...(titleRows ? { titleRows } : {}),
        gridLines,
        sheetGrids,
      }),
    );

    // §20.5: the sheet's chart frames render as blocks after its grid,
    // anchor-ordered (resolved chart data lives in sheet.chartData).
    for (const ref of ws.charts ?? []) {
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

    // W1: anchored pictures render as image blocks after the grid (anchor-ordered;
    // bytes live in sheet.resources). Like charts, placement collapses to inline.
    for (const img of ws.images ?? []) {
      body.push({
        kind: 'image',
        image: {
          resource: img.resourceId,
          width: pt(img.widthPt),
          height: pt(img.heightPt),
          paragraphProperties: {},
        },
      });
    }

    // W2: anchored shapes render as shape blocks after the grid (anchor-ordered;
    // placement collapses to inline, like charts/pictures).
    for (const shape of ws.shapes ?? []) {
      body.push({ kind: 'shape', shape });
    }

    // §SV2: slicer panels render as styled button boxes after the grid + charts.
    for (const slicer of ws.slicers ?? []) {
      body.push({ kind: 'table', table: slicerTable(slicer) });
    }
  }

  return {
    kind: 'flow',
    // Same stage-6 contract as docx: the body carries resolved properties. Grid
    // cells are built with direct props only, so resolving over the empty sheet
    // just materializes the defaults.
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    ...(firstSheetSection ? { section: firstSheetSection } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources: sheet.resources,
    ...(sheet.chartData && sheet.chartData.size > 0 ? { charts: sheet.chartData } : {}),
    ...(sheet.info ? { info: sheet.info } : {}),
  };
}
