// SheetDoc → FlowDoc projection (E-SHEET SA2). The print model turns each grid
// sheet into flow blocks (a table + chart frames); sheets after the first start
// on a new page. This is the SAME projection the xlsx reader used to inline —
// relocated behind the SheetDoc boundary, byte-for-byte unchanged — so the
// render path (PDF/SVG/HTML) is identical. A dedicated grid layout would be a
// separate SheetDoc consumer; for now FlowDoc is the one projection.

import type {
  BodyElement,
  FloatAnchor,
  HeaderFooterReference,
  Section,
  SectionProperties,
  ShapeBlock,
} from '@/core/document-model';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';
import type { Pt } from '@/core/ir';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir/loss';
import type {
  SheetActiveXControl,
  SheetComment,
  SheetControlBox,
  SheetDoc,
  SheetFormControl,
} from '@/core/ir/sheet';

import { pt } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { buildHeaderFooterContent } from '@/excel/header-footer';
import {
  printableWidthPt,
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

/**
 * Projection knobs (E-SHEET W9).
 */
export interface ProjectSheetOptions {
  /**
   * The injected reference date that drives conditional-format `timePeriod`
   * windows and `TODAY()`/`NOW()` in `expression` rules. Omitted ⇒ those
   * clock-relative constructs no-op, so the projection stays deterministic and
   * byte-identical to before.
   */
  readonly now?: Date;
  /**
   * Sink the projection writes its {@link Loss} entries into — the print
   * model's defence-in-depth caps (grid size, per-sheet text budget, sparkline
   * range) fire on pathological input, and a cap that fires without saying so
   * is a silent wrongness. {@link readXlsx} always supplies one and returns it
   * as the read result's loss report; omitted ⇒ the caps still apply but go
   * unreported.
   */
  readonly losses?: Array<Loss>;
  /**
   * Width of a digit in the font the document will actually be rendered in, at
   * the workbook's default size, in points.
   *
   * §18.3.1.13 measures a column in Maximum Digit Widths of the workbook's own
   * default font — the unit is a property of the font, not a constant. Laying
   * columns out in one font's digit and then drawing the text in another makes
   * every column the wrong width for its contents: with Calibri's 5.25 pt
   * assumed and Roboto's 6.18 pt drawn, a column holds a sixth less than the
   * file says, and the text that does not fit is clipped away.
   *
   * Omitted ⇒ Excel's own 7 px, which is right only if the render font matches
   * the workbook's.
   */
  readonly digitWidthPt?: number;
}

/**
 * Project a {@link SheetDoc} into a {@link FlowDoc} (E-SHEET SA2): each grid sheet
 * becomes flow blocks (a table + chart/picture/shape/slicer frames, then comment
 * / control listings), with sheets after the first starting on a new page. The
 * first sheet's page geometry + header/footer drive the document section.
 *
 * @param sheet   The SpreadsheetML IR tree.
 * @param options Projection knobs (the W9 reference date).
 * @returns The format-neutral flow document the render path consumes.
 */
export function projectSheetDoc(sheet: SheetDoc, options: ProjectSheetOptions = {}): FlowDoc {
  const body: Array<BodyElement> = [];
  // ONE SECTION PER SHEET. A workbook's sheets set their paper independently —
  // tdf171828_fail_to_import_file.xlsx is A4 landscape, then Letter portrait,
  // then A4 landscape again — and reprinting them all on the first sheet's
  // paper is not an approximation, it is the wrong page. The layout has taken
  // `sections` (each with the body index it runs to) since docx needed it; the
  // spreadsheet projection simply never used them.
  // Each sheet's geometry, in order; the body index each sheet ends at is
  // recorded alongside once its blocks are in.
  const sheetSections: Array<SectionProperties> = [];
  const sheetEnds: Array<number> = [];
  // Kept for FlowDoc.section, the single-section field the render path falls
  // back to and which other consumers still read.
  let firstSheetSection: SectionProperties | undefined;
  // First sheet's expanded header/footer band content (E-SHEET W4), keyed for
  // FlowDoc.headersFooters; the renderer paints it in the page margins.
  const headersFooters = new Map<string, ReadonlyArray<BodyElement>>();

  // Sheet name → grid, so a sparkline whose data range is sheet-qualified
  // (Sheet2!A1:C1) resolves against the right sheet (E-SHEET SC2 tail TC3).
  const sheetGrids = new Map(sheet.sheets.map((s) => [s.name, s.grid]));

  // Sheets actually printed so far — not the loop index, which has to keep
  // counting hidden sheets because `localSheetId` on a defined name does.
  let printed = 0;
  for (let sheetIdx = 0; sheetIdx < sheet.sheets.length; sheetIdx++) {
    const ws = sheet.sheets[sheetIdx]!;
    // §18.2.19: a hidden tab is not printed. Excel and LibreOffice both leave it
    // out entirely — tdf171828.xlsx hides its lookup table, and printing it added
    // two pages of working data to the end of the document.
    if (ws.hidden) continue;

    // The grid is projected FIRST, before the section is built: the header band
    // has to be laid out at the sheet's print scale, and the scale only falls
    // out of the projection. Its blocks are pushed below, in their old place.
    const scaleSink = { value: 1 };
    const drawingExtentPt = shapeExtentPt(ws.shapes);
    const printArea = resolvePrintArea(sheet.definedNames, sheetIdx);
    const titleRows = resolvePrintTitleRows(sheet.definedNames, sheetIdx);
    const gridBody = worksheetToBody(ws.grid, sheet.sharedStrings, sheet.styles, sheet.date1904, {
      ...(printArea ? { printArea } : {}),
      ...(titleRows ? { titleRows } : {}),
      gridLines: ws.grid.printOptions?.gridLines === true,
      sheetGrids,
      sheetName: ws.name,
      definedNames: sheet.definedNames,
      ...(ws.hyperlinks ? { hyperlinks: ws.hyperlinks } : {}),
      ...(sheet.sharedStringRuns ? { sharedStringRuns: sheet.sharedStringRuns } : {}),
      ...(options.now ? { now: options.now } : {}),
      ...(options.digitWidthPt !== undefined ? { digitWidthPt: options.digitWidthPt } : {}),
      ...(options.losses ? { losses: options.losses } : {}),
      scaleSink,
      ...(drawingExtentPt ? { drawingExtentPt } : {}),
    });

    // The header/footer band is a document-level resource keyed by a synthetic
    // id, so only the first sheet's can be carried; its geometry, though, is
    // per-sheet.
    const sheetSection =
      printed === 0
        ? withHeaderFooter(
            sectionFromWorksheet(ws.grid),
            ws,
            headersFooters,
            scaleSink.value,
            sheet.styles.fonts[0]?.sizePt,
          )
        : sectionFromWorksheet(ws.grid);
    if (printed === 0) firstSheetSection = sheetSection;
    sheetSections.push(sheetSection);

    // Each sheet after the first starts on its own PDF page. We do NOT print the
    // sheet name (Calc/Excel `--convert-to pdf` emit it nowhere), so the page
    // break is an empty page-break-only paragraph (no runs ⇒ no glyphs).
    if (printed > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: { properties: { pageBreakBefore: true }, runs: [] },
      });
    }

    body.push(...gridBody);

    // §20.5: the sheet's chart frames render as blocks after its grid,
    // anchor-ordered (resolved chart data lives in sheet.chartData).
    for (const ref of ws.charts ?? []) {
      body.push({
        kind: 'chart',
        chart: {
          ...anchorFloat(ref.xPt, ref.yPt, scaleSink.value),
          chartRelId: ref.chartPartPath,
          width: pt(ref.widthPt * scaleSink.value),
          height: pt(ref.heightPt * scaleSink.value),
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
          ...anchorFloat(img.xPt, img.yPt, scaleSink.value),
          resource: img.resourceId,
          width: pt(img.widthPt * scaleSink.value),
          height: pt(img.heightPt * scaleSink.value),
          paragraphProperties: {},
        },
      });
    }

    // W2: anchored shapes render as floating shape blocks over the grid, at the
    // point their `twoCellAnchor` names — scaled with the sheet, since that is
    // what the anchor's tracks were measured in.
    for (const shape of ws.shapes ?? []) {
      body.push({
        kind: 'shape',
        shape: centreShape(
          scaleShape(shape, scaleSink.value),
          ws.grid,
          drawingExtentPt,
          scaleSink.value,
        ),
      });
    }

    // W8/W10: a control that knows where it belongs is DRAWN there, as the
    // widget it is — the ones with no geometry in the file are listed after the
    // grid instead (formControlBlocks / activeXBlocks below).
    const controlBlocks = controlShapeBlocks(ws.formControls, ws.activeXControls);
    // A sheet whose drawings run wider than the page paginates ACROSS them, the
    // way a wide grid paginates across its columns ("down, then over"). With no
    // grid there are no column bands to follow, so the drawings are banded on
    // their own: tdf111980_radioButtons.xlsx has no cells at all and eleven
    // controls spread over 650pt of an A4 page, and everything past the right
    // margin simply fell off the one page we printed.
    const drawingBands =
      gridBody.length === 0
        ? bandDrawings(controlBlocks, printableWidthPt(ws.grid), scaleSink.value)
        : [controlBlocks];
    for (let band = 0; band < drawingBands.length; band++) {
      if (band > 0) {
        body.push({
          kind: 'paragraph',
          paragraph: { properties: { pageBreakBefore: true }, runs: [] },
        });
      }
      for (const shape of drawingBands[band]!) {
        body.push({
          kind: 'shape',
          shape: centreShape(
            scaleShape(shape, scaleSink.value),
            ws.grid,
            drawingExtentPt,
            scaleSink.value,
          ),
        });
      }
    }

    // §SV2: slicer panels render as styled button boxes after the grid + charts.
    for (const slicer of ws.slicers ?? []) {
      body.push({ kind: 'table', table: slicerTable(slicer) });
    }

    // W7: cell comments / notes are listed in a "Comments" section after the
    // grid (Excel's "print comments at end of sheet") — unless the sheet says
    // not to. §18.3.1.63 `cellComments="none"` is the author answering the
    // print dialog's comments question with "(None)", and printing them anyway
    // put four notes on tdf171828.xlsx that no other reader shows.
    //
    // Silence is not that answer. The spec's default is `none`, but a sheet
    // that never opened the dialog says nothing about its notes, and dropping
    // content the document carries on the strength of an unstated default is
    // the losing side of the trade: listed, a note is visibly extra; omitted,
    // it is invisibly gone.
    if (ws.comments && ws.comments.length > 0 && ws.grid.pageSetup?.cellComments !== 'none') {
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
    sheetEnds.push(body.length);
    printed++;
  }

  // Section i covers body[sections[i-1].endIndex .. sections[i].endIndex). A
  // single-sheet workbook keeps `sections` empty so the render path stays on
  // the FlowDoc.section fallback it has always used, byte for byte.
  const sections: Array<Section> =
    sheetSections.length > 1
      ? sheetSections.map((properties, i) => ({ properties, endIndex: sheetEnds[i]! }))
      : [];

  return {
    kind: 'flow',
    // Same stage-6 contract as docx: the body carries resolved properties. Grid
    // cells are built with direct props only, so resolving over the empty sheet
    // just materializes the defaults.
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections,
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

// Excel's Forms widgets, in points, measured off the reference render: the
// option button's disc, the gutter its caption clears, and the grey the group
// box is framed in.
const RADIO_DIAMETER_PT = 8.4;
const RADIO_RING_PT = 0.6;
const RADIO_DOT_INSET_PT = 1.2;
const CHECKBOX_SIDE_PT = 8.4;
const CAPTION_GUTTER_PT = 11.5;
const CONTROL_LINE_PT = 0.75;
const GROUP_BOX_GREY = '808080';
// A Forms control's shape names its own size (`<font size>`); an ActiveX one
// keeps its font in the .bin we do not read, so it takes the 10pt its class
// defaults to — which is what the reference draws these five in.
const DEFAULT_FORM_CONTROL_PT = 8;
const DEFAULT_ACTIVEX_CONTROL_PT = 10;

/** The kinds of widget we draw; everything else falls back to a plain frame. */
type ControlGlyph = 'radio' | 'checkbox' | 'group' | 'button' | 'label' | 'frame';

function formControlGlyph(objectType: string | undefined): ControlGlyph {
  switch ((objectType ?? '').toLowerCase()) {
    case 'radio':
      return 'radio';
    case 'checkbox':
      return 'checkbox';
    case 'gbox':
      return 'group';
    case 'button':
    case 'buttons':
      return 'button';
    case 'label':
      return 'label';
    default:
      return 'frame';
  }
}

function activeXGlyph(type: string): ControlGlyph {
  switch (type) {
    case 'option':
      return 'radio';
    case 'checkbox':
      return 'checkbox';
    case 'button':
    case 'toggle':
      return 'button';
    case 'label':
      return 'label';
    default:
      return 'frame';
  }
}

/** A floating shape at `box`, sized and positioned in sheet points. */
function controlShape(
  box: SheetControlBox,
  dx: number,
  dy: number,
  widthPt: number,
  heightPt: number,
  rest: Omit<ShapeBlock, 'float' | 'width' | 'height' | 'paragraphProperties'>,
): ShapeBlock {
  return {
    float: {
      wrap: 'none',
      posH: { relativeFrom: 'margin', offsetPt: pt(box.xPt + dx) },
      posV: { relativeFrom: 'margin', offsetPt: pt(box.yPt + dy) },
    },
    width: pt(widthPt),
    height: pt(heightPt),
    ...rest,
    paragraphProperties: {},
  };
}

/**
 * A control drawn where the document puts it (E-SHEET W8/W10), as the widget it
 * is: an option button's ring and dot, a check box, a group box's frame, a
 * push button's outline — each with its caption beside or inside it.
 *
 * Listing the captions after the grid instead was a stand-in for having no
 * geometry to place them with. It read as a paragraph of ASCII where
 * tdf111980_radioButtons.xlsx has eleven widgets spread across the sheet, the
 * furthest of them 18cm from where we drew it.
 */
function controlShapes(
  box: SheetControlBox,
  glyph: ControlGlyph,
  caption: string | undefined,
  checked: boolean,
  fontSizePt: number,
): Array<ShapeBlock> {
  const out: Array<ShapeBlock> = [];
  const noFill = { kind: 'none' as const };
  const rect = { kind: 'preset' as const, preset: 'rect' };
  const ellipse = { kind: 'preset' as const, preset: 'ellipse' };
  const grey = { colorHex: GROUP_BOX_GREY, width: pt(CONTROL_LINE_PT) };
  const black = { colorHex: '000000', width: pt(CONTROL_LINE_PT) };

  if (glyph === 'radio' || glyph === 'checkbox') {
    const side = glyph === 'radio' ? RADIO_DIAMETER_PT : CHECKBOX_SIDE_PT;
    const top = Math.max(0, (box.heightPt - side) / 2);
    out.push(
      controlShape(box, 0, top, side, side, {
        geometry: glyph === 'radio' ? ellipse : rect,
        fill: { kind: 'solid', colorHex: 'FFFFFF' },
        line: { colorHex: '000000', width: pt(RADIO_RING_PT) },
      }),
    );
    if (checked) {
      const dot = side - 2 * RADIO_DOT_INSET_PT;
      out.push(
        controlShape(box, RADIO_DOT_INSET_PT, top + RADIO_DOT_INSET_PT, dot, dot, {
          geometry: glyph === 'radio' ? ellipse : rect,
          fill: { kind: 'solid', colorHex: '000000' },
        }),
      );
    }
  } else if (glyph !== 'label') {
    // A group box frames its contents in grey; a button and any control we have
    // no widget for get a plain outline at their own box.
    out.push(
      controlShape(box, 0, 0, box.widthPt, box.heightPt, {
        geometry: rect,
        fill: noFill,
        line: glyph === 'group' ? grey : black,
      }),
    );
  }

  if (caption !== undefined && caption.length > 0) {
    // Beside the glyph for a labelled control, inside the frame for a group box
    // (Excel breaks the top border around it, which needs a clip we do not have)
    // and across the whole box for a button.
    const inset = glyph === 'radio' || glyph === 'checkbox' ? CAPTION_GUTTER_PT : 3;
    const runs = [{ text: caption, properties: { fontSizePt: pt(fontSizePt) } }];
    out.push(
      controlShape(box, 0, 0, box.widthPt, box.heightPt, {
        geometry: rect,
        fill: noFill,
        text: {
          content: [
            {
              kind: 'paragraph',
              paragraph: {
                properties: glyph === 'button' ? { alignment: 'center' as const } : {},
                runs,
              },
            },
          ],
          insetLeft: pt(inset),
          insetTop: pt(1),
          insetRight: pt(1),
          insetBottom: pt(1),
          anchor: glyph === 'group' ? ('t' as const) : ('ctr' as const),
        },
      }),
    );
  }
  return out;
}

/**
 * Split floating drawings into page-wide bands, left to right — the "down, then
 * over" order a wide sheet prints in (§18.3.1.63 `pageOrder`), applied to the
 * drawing layer when there is no grid whose column bands it could follow.
 *
 * Each band is shifted back to the page's left margin, so band 2 starts where
 * the printable width ran out. A drawing straddling the boundary belongs to the
 * band its left edge is in, which is where the reference draws it too — the
 * part past the edge is clipped on that page and continues on the next.
 */
function bandDrawings(
  shapes: ReadonlyArray<ShapeBlock>,
  printable: number,
  scale: number,
): Array<Array<ShapeBlock>> {
  if (shapes.length === 0 || !(printable > 0)) return [shapes as Array<ShapeBlock>];
  const rightOf = (s: ShapeBlock): number => (s.float?.posH?.offsetPt ?? 0) * scale + s.width;
  const extent = Math.max(0, ...shapes.map(rightOf));
  const bandCount = Math.ceil(extent / printable);
  if (bandCount <= 1) return [shapes as Array<ShapeBlock>];
  const bands: Array<Array<ShapeBlock>> = Array.from({ length: bandCount }, () => []);
  for (const shape of shapes) {
    const posH = shape.float?.posH;
    const left = (posH?.offsetPt ?? 0) * scale;
    const right = left + shape.width;
    const first = Math.max(0, Math.min(bandCount - 1, Math.floor(left / printable)));
    // EVERY band the drawing reaches into, not just the one it starts in. A
    // group box 209pt wide anchored 46pt before the boundary is on both pages —
    // clipped at the edge of the first and continuing from the edge of the
    // second, which is exactly what the reference draws. Assigned to one band
    // only, its far half and every caption inside it fell off the document.
    const last = Math.max(first, Math.min(bandCount - 1, Math.ceil(right / printable) - 1));
    for (let band = first; band <= last; band++) {
      if (!shape.float || !posH) {
        bands[band]!.push(shape);
        continue;
      }
      // A continuation carries the drawing's FRAME and not its text. The label
      // is anchored at the drawing's left edge, which is back in the band the
      // drawing started in — it was printed there, and printing it again puts
      // the same caption on two pages. The reference clips the drawing at the
      // band boundary instead, which needs a clip rectangle the page model does
      // not have; dropping the repeat is the part of that we can honour.
      const { text: _text, ...frame } = shape;
      // …and a continuation with neither text, outline nor fill draws nothing
      // at all: the caption's own box is invisible once its text stays behind.
      if (band !== first && !frame.line && frame.fill.kind === 'none') continue;
      bands[band]!.push({
        ...(band === first ? shape : frame),
        float: {
          ...shape.float,
          posH: { ...posH, offsetPt: pt((posH.offsetPt ?? 0) - (band * printable) / scale) },
        },
      });
    }
  }
  return bands.filter((b) => b.length > 0);
}

/**
 * §20.5.2.35 — the float a drawing's anchor asks for: pinned at the point the
 * anchor names, measured from the page margin, wrapping nothing.
 *
 * A chart or a picture placed in the FLOW instead lands at the left margin,
 * below every block before it, and on whatever page that turns out to be:
 * simple-monthly-budget.xlsx anchors its chart beside the summary on page one
 * and we printed it alone on page two, a third of the way in from the left.
 * Shapes have floated at their anchors since W2; these two never did.
 *
 * No anchor (an `absoluteAnchor`, or a reader that gave none) ⇒ no float, and
 * the block keeps the flow placement it has always had.
 */
function anchorFloat(
  xPt: number | undefined,
  yPt: number | undefined,
  scale: number,
): { float?: FloatAnchor } {
  if (xPt === undefined || yPt === undefined) return {};
  return {
    float: {
      wrap: 'none',
      posH: { relativeFrom: 'margin', offsetPt: pt(xPt * scale) },
      posV: { relativeFrom: 'margin', offsetPt: pt(yPt * scale) },
    },
  };
}

/** Every control that carries geometry, as the shapes that draw it. */
function controlShapeBlocks(
  formControls: ReadonlyArray<SheetFormControl> | undefined,
  activeXControls: ReadonlyArray<SheetActiveXControl> | undefined,
): Array<ShapeBlock> {
  const out: Array<ShapeBlock> = [];
  for (const c of formControls ?? []) {
    if (!c.box) continue;
    out.push(
      ...controlShapes(
        c.box,
        formControlGlyph(c.objectType),
        c.name,
        c.checked === true,
        c.fontSizePt ?? DEFAULT_FORM_CONTROL_PT,
      ),
    );
  }
  for (const c of activeXControls ?? []) {
    if (!c.box) continue;
    const on = c.value === '1' || c.value?.toLowerCase() === 'true';
    out.push(
      ...controlShapes(
        c.box,
        activeXGlyph(c.type),
        c.caption ?? c.name,
        on,
        DEFAULT_ACTIVEX_CONTROL_PT,
      ),
    );
  }
  return out;
}

// Form controls (W8) as a "Form controls" section after the grid: a bold heading
// then one line per control with a type-appropriate ASCII affordance and state —
// a checkbox/option button shows its checked state, a spin/scroll its value.
// Only for controls whose drawing carries no geometry; one that knows where it
// belongs is drawn there instead (see {@link controlShapes}).
function formControlBlocks(controls: ReadonlyArray<SheetFormControl>): Array<BodyElement> {
  const listed = controls.filter((c) => c.box === undefined);
  if (listed.length === 0) return [];
  const out: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: 'Form controls', properties: { bold: true } }] },
    },
  ];
  for (const c of listed) {
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
  const listed = controls.filter((c) => c.box === undefined);
  if (listed.length === 0) return [];
  const out: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: {
        properties: {},
        runs: [{ text: 'ActiveX controls', properties: { bold: true } }],
      },
    },
  ];
  for (const c of listed) {
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
/**
 * How far the sheet's drawings reach from its origin, or undefined when it has
 * none. Fit-to-page has to fit them as well as the cells: a drawing anchored
 * past the last value is still printed, and on a sheet whose values sit in a
 * handful of cells it is the drawing that decides the page.
 */
function shapeExtentPt(
  shapes: ReadonlyArray<ShapeBlock> | undefined,
): { widthPt: number; heightPt: number } | undefined {
  if (!shapes || shapes.length === 0) return undefined;
  let widthPt = 0;
  let heightPt = 0;
  for (const shape of shapes) {
    const x = shape.float?.posH?.offsetPt ?? 0;
    const y = shape.float?.posV?.offsetPt ?? 0;
    widthPt = Math.max(widthPt, x + shape.width);
    heightPt = Math.max(heightPt, y + shape.height);
  }
  return widthPt > 0 || heightPt > 0 ? { widthPt, heightPt } : undefined;
}

/**
 * `<printOptions horizontalCentered/verticalCentered>` centres what is PRINTED,
 * and a drawing is printed. We centred the table only, which leaves a drawing
 * at its raw offset — and on a sheet whose cells are all empty there is no
 * table to centre at all: bnc762542.xlsx asks for both axes and its callout sat
 * some 280pt left of where the reference puts it, matching the file's own
 * numbers exactly but ignoring the request that moves them.
 */
function centreShape(
  shape: ShapeBlock,
  worksheet: ParsedWorksheet,
  extent: { widthPt: number; heightPt: number } | undefined,
  scale: number,
): ShapeBlock {
  const options = worksheet.printOptions;
  if (!extent || (!options?.horizontalCentered && !options?.verticalCentered)) return shape;
  const section = sectionFromWorksheet(worksheet);
  const size = section.pageSize;
  const margins = section.margins;
  if (!size || !margins) return shape;
  const dx = options.horizontalCentered
    ? (size.width - margins.left - margins.right - extent.widthPt * scale) / 2
    : 0;
  const dy = options.verticalCentered
    ? (size.height - margins.top - margins.bottom - extent.heightPt * scale) / 2
    : 0;
  if (dx <= 0 && dy <= 0) return shape;
  const posH = shape.float?.posH;
  const posV = shape.float?.posV;
  if (!shape.float) return shape;
  return {
    ...shape,
    float: {
      ...shape.float,
      ...(posH && dx > 0 ? { posH: { ...posH, offsetPt: pt((posH.offsetPt ?? 0) + dx) } } : {}),
      ...(posV && dy > 0 ? { posV: { ...posV, offsetPt: pt((posV.offsetPt ?? 0) + dy) } } : {}),
    },
  };
}

/**
 * A drawing at the sheet's print scale: its box and the point it floats at
 * shrink together, exactly as the grid beneath them does. Returns the shape
 * untouched at 1.0 so an unscaled sheet stays byte-identical.
 */
function scaleShape(shape: ShapeBlock, scale: number): ShapeBlock {
  if (!(scale > 0) || scale >= 0.999) return shape;
  const offset = (o: { readonly offsetPt?: Pt } | undefined): number | undefined =>
    o?.offsetPt === undefined ? undefined : o.offsetPt * scale;
  const posH = shape.float?.posH;
  const posV = shape.float?.posV;
  return {
    ...shape,
    width: pt(shape.width * scale),
    height: pt(shape.height * scale),
    ...(shape.float
      ? {
          float: {
            ...shape.float,
            ...(posH ? { posH: { ...posH, ...withOffset(offset(posH)) } } : {}),
            ...(posV ? { posV: { ...posV, ...withOffset(offset(posV)) } } : {}),
          },
        }
      : {}),
  };
}

function withOffset(value: number | undefined): { offsetPt?: Pt } {
  return value === undefined ? {} : { offsetPt: pt(value) };
}

function withHeaderFooter(
  section: SectionProperties,
  ws: SheetDoc['sheets'][number],
  headersFooters: Map<string, ReadonlyArray<BodyElement>>,
  scale: number,
  basePt: number | undefined,
): SectionProperties {
  const hf = ws.grid.headerFooter;
  if (!hf || (!hf.oddHeader && !hf.oddFooter)) return section;
  const headers: Array<HeaderFooterReference> = [];
  const footers: Array<HeaderFooterReference> = [];
  if (hf.oddHeader) {
    const content = buildHeaderFooterContent(hf.oddHeader, ws.name, scale, basePt);
    if (content.length > 0) {
      headersFooters.set(HEADER_REL, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      headers.push({ type: 'default', relationshipId: HEADER_REL });
    }
  }
  if (hf.oddFooter) {
    const content = buildHeaderFooterContent(hf.oddFooter, ws.name, scale, basePt);
    if (content.length > 0) {
      headersFooters.set(FOOTER_REL, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      footers.push({ type: 'default', relationshipId: FOOTER_REL });
    }
  }
  if (headers.length === 0 && footers.length === 0) return section;
  return { ...section, headers, footers };
}
