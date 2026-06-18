// SheetDoc → FlowDoc projection (E-SHEET SA2). The print model turns each grid
// sheet into flow blocks (a table + chart frames); sheets after the first start
// on a new page. This is the SAME projection the xlsx reader used to inline —
// relocated behind the SheetDoc boundary, byte-for-byte unchanged — so the
// render path (PDF/SVG/HTML) is identical. A dedicated grid layout would be a
// separate SheetDoc consumer; for now FlowDoc is the one projection.

import type { BodyElement, HeaderFooterReference, SectionProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type {
  SheetActiveXControl,
  SheetComment,
  SheetDoc,
  SheetFormControl,
} from '@/core/ir/sheet';

import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { pt } from '@/core/ir';
import { buildHeaderFooterContent } from '@/excel/header-footer';
import {
  resolvePrintArea,
  resolvePrintTitleRows,
  sectionFromWorksheet,
  slicerTable,
  worksheetToBody,
} from '@/excel/print-model';

// Synthetic relationship ids keying the first sheet's header/footer band content
// in FlowDoc.headersFooters (E-SHEET W4).
const HEADER_REL = '_xlsxHeaderDefault';
const FOOTER_REL = '_xlsxFooterDefault';

// Projection knobs (E-SHEET W9). `now` is the injected reference date that drives
// conditional-format `timePeriod` windows and TODAY()/NOW() in `expression`
// rules. Omitted ⇒ those clock-relative constructs no-op, so the projection
// stays deterministic and byte-identical to before.
export interface ProjectSheetOptions {
  readonly now?: Date;
}

export function projectSheetDoc(sheet: SheetDoc, options: ProjectSheetOptions = {}): FlowDoc {
  const body: Array<BodyElement> = [];
  // Page geometry comes from the first sheet's <pageSetup>/<pageMargins>; the
  // renderer supports one section, so later sheets share the first's geometry.
  let firstSheetSection: SectionProperties | undefined;
  // First sheet's expanded header/footer band content (E-SHEET W4), keyed for
  // FlowDoc.headersFooters; the renderer paints it in the page margins.
  const headersFooters = new Map<string, ReadonlyArray<BodyElement>>();

  // Sheet name → grid, so a sparkline whose data range is sheet-qualified
  // (Sheet2!A1:C1) resolves against the right sheet (E-SHEET SC2 tail TC3).
  const sheetGrids = new Map(sheet.sheets.map((s) => [s.name, s.grid]));

  for (let sheetIdx = 0; sheetIdx < sheet.sheets.length; sheetIdx++) {
    const ws = sheet.sheets[sheetIdx]!;
    if (sheetIdx === 0) {
      firstSheetSection = withHeaderFooter(sectionFromWorksheet(ws.grid), ws, headersFooters);
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
        sheetName: ws.name,
        definedNames: sheet.definedNames,
        ...(ws.hyperlinks ? { hyperlinks: ws.hyperlinks } : {}),
        ...(sheet.sharedStringRuns ? { sharedStringRuns: sheet.sharedStringRuns } : {}),
        ...(options.now ? { now: options.now } : {}),
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

    // W7: cell comments / notes are listed in a "Comments" section after the grid
    // (Excel's "print comments at end of sheet"): a heading + one line per comment.
    if (ws.comments && ws.comments.length > 0) {
      body.push(...commentBlocks(ws.comments));
    }

    // W8: form controls are listed in a "Form controls" section after the grid,
    // each with a type-appropriate affordance and its state.
    if (ws.formControls && ws.formControls.length > 0) {
      body.push(...formControlBlocks(ws.formControls));
    }

    // W10: ActiveX controls in an "ActiveX controls" section, same as form
    // controls (type-appropriate affordance + the property bag's visible state).
    if (ws.activeXControls && ws.activeXControls.length > 0) {
      body.push(...activeXBlocks(ws.activeXControls));
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
    ...(headersFooters.size > 0 ? { headersFooters } : {}),
    ...(sheet.info ? { info: sheet.info } : {}),
  };
}

// Cell comments / notes (W7) as a "Comments" section after the grid: a bold
// heading then one paragraph per comment — "<ref> — <author>: <text>". Multi-line
// note text is collapsed to a single line so the listing stays compact.
function commentBlocks(comments: ReadonlyArray<SheetComment>): Array<BodyElement> {
  const out: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: 'Comments', properties: { bold: true } }] },
    },
  ];
  for (const c of comments) {
    const body = c.text.replace(/\s+/g, ' ').trim();
    const label = c.author ? `${c.ref} — ${c.author}: ` : `${c.ref}: `;
    out.push({
      kind: 'paragraph',
      paragraph: {
        properties: {},
        runs: [
          { text: label, properties: { bold: true } },
          ...(body.length > 0 ? [{ text: body, properties: {} }] : []),
        ],
      },
    });
  }
  return out;
}

// Form controls (W8) as a "Form controls" section after the grid: a bold heading
// then one line per control with a type-appropriate ASCII affordance and state —
// a checkbox/option button shows its checked state, a spin/scroll its value.
function formControlBlocks(controls: ReadonlyArray<SheetFormControl>): Array<BodyElement> {
  const out: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: 'Form controls', properties: { bold: true } }] },
    },
  ];
  for (const c of controls) {
    out.push({
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: formControlLabel(c), properties: {} }] },
    });
  }
  return out;
}

function formControlLabel(c: SheetFormControl): string {
  const name = c.name ?? c.objectType ?? 'Control';
  switch ((c.objectType ?? '').toLowerCase()) {
    case 'checkbox':
      return `${c.checked ? '[x]' : '[ ]'} ${name}`;
    case 'radio':
      return `${c.checked ? '(o)' : '( )'} ${name}`;
    case 'buttons':
      return `[ ${name} ]`;
    case 'spin':
    case 'scroll':
      return c.value !== undefined ? `${name} (value ${c.value})` : name;
    case 'drop':
    case 'list':
      return `${name} (list)`;
    default:
      return c.objectType ? `${name} (${c.objectType})` : name;
  }
}

// ActiveX controls (W10) as an "ActiveX controls" section after the grid — one
// line per control with a type-appropriate ASCII affordance and the visible
// state read from its property bag.
function activeXBlocks(controls: ReadonlyArray<SheetActiveXControl>): Array<BodyElement> {
  const out: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: {
        properties: {},
        runs: [{ text: 'ActiveX controls', properties: { bold: true } }],
      },
    },
  ];
  for (const c of controls) {
    out.push({
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: activeXLabel(c), properties: {} }] },
    });
  }
  return out;
}

function activeXLabel(c: SheetActiveXControl): string {
  const label = c.caption && c.caption.length > 0 ? c.caption : c.type;
  const on = c.value === '1' || c.value?.toLowerCase() === 'true';
  switch (c.type) {
    case 'checkbox':
      return `${on ? '[x]' : '[ ]'} ${label}`;
    case 'option':
      return `${on ? '(o)' : '( )'} ${label}`;
    case 'button':
    case 'toggle':
      return `[ ${label} ]`;
    case 'textbox':
      return `[ ${c.value ?? c.caption ?? ''} ]`;
    case 'combo':
    case 'list':
      return `${label} (list)`;
    case 'spin':
    case 'scroll':
      return c.value !== undefined ? `${label} (value ${c.value})` : label;
    case 'label':
      return c.caption ?? 'Label';
    default:
      return label;
  }
}

// Expand the first sheet's <headerFooter> into header/footer bands and attach them
// to its section (creating a minimal section when the sheet has no custom page
// geometry). The section is returned unchanged when there is no header/footer.
function withHeaderFooter(
  section: SectionProperties | undefined,
  ws: SheetDoc['sheets'][number],
  headersFooters: Map<string, ReadonlyArray<BodyElement>>,
): SectionProperties | undefined {
  const hf = ws.grid.headerFooter;
  if (!hf || (!hf.oddHeader && !hf.oddFooter)) return section;
  const headers: Array<HeaderFooterReference> = [];
  const footers: Array<HeaderFooterReference> = [];
  if (hf.oddHeader) {
    const content = buildHeaderFooterContent(hf.oddHeader, ws.name);
    if (content.length > 0) {
      headersFooters.set(HEADER_REL, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      headers.push({ type: 'default', relationshipId: HEADER_REL });
    }
  }
  if (hf.oddFooter) {
    const content = buildHeaderFooterContent(hf.oddFooter, ws.name);
    if (content.length > 0) {
      headersFooters.set(FOOTER_REL, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      footers.push({ type: 'default', relationshipId: FOOTER_REL });
    }
  }
  if (headers.length === 0 && footers.length === 0) return section;
  return { ...(section ?? { headers: [], footers: [] }), headers, footers };
}
