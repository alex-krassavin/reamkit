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
import type { ParsedWorksheet, XlsxStyles } from '@/core/spreadsheet-model';
import type { Pt } from '@/core/ir';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir/loss';
import type {
  Sheet,
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
  cellPaintsVisibly,
  printableHeightPt,
  printableWidthPt,
  resolvePrintArea,
  resolvePrintTitleRows,
  sectionFromWorksheet,
  slicerTable,
  worksheetToBody,
} from '@/excel/print-model';

// Synthetic relationship ids keying each sheet's header/footer band content in
// FlowDoc.headersFooters (E-SHEET W4). The sheet's index is appended — the map
// is document-level, and one fixed id meant the last sheet's bands overwrote
// every earlier sheet's.
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
  /**
   * §18.3.1.34 `&F` — the workbook's file name, for a header or footer that
   * prints it. The reader takes bytes and cannot know it, so the caller
   * supplies it; absent, the code is dropped exactly as before.
   */
  readonly fileName?: string;
}

/**
 * Project a {@link SheetDoc} into a {@link FlowDoc} (E-SHEET SA2): each grid sheet
 * becomes flow blocks (a table + chart/picture/shape/slicer frames, then comment
 * / control listings), with sheets after the first starting on a new page. Each
 * sheet gets its own section — its page geometry and its own header/footer.
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
  // Every sheet's expanded header/footer band content (E-SHEET W4), keyed for
  // FlowDoc.headersFooters; the renderer paints it in the page margins.
  const headersFooters = new Map<string, ReadonlyArray<BodyElement>>();

  // Sheet name → grid, so a sparkline whose data range is sheet-qualified
  // (Sheet2!A1:C1) resolves against the right sheet (E-SHEET SC2 tail TC3).
  const sheetGrids = new Map(sheet.sheets.map((s) => [s.name, s.grid]));

  // Sheets actually printed so far — not the loop index, which has to keep
  // counting hidden sheets because `localSheetId` on a defined name does.
  let printed = 0;
  // How many sheets are still candidates to print, so the guard above can tell
  // "this one is empty" from "every one of them is".
  let printableSheets = sheet.sheets.filter((s) => !s.hidden).length;
  // Does anything in this workbook print at all? When nothing does, ONE sheet
  // still prints its page (see the guard below) — and the one worth printing is
  // whichever carries a header or footer, since that is the only thing such a
  // workbook can put on paper. npe.xlsx keeps its footer on the second of two
  // empty tabs, and printing the first gave a page with nothing on it where
  // LibreOffice prints the footer.
  const visible = sheet.sheets.filter((s) => !s.hidden);
  const anyPrints = visible.some((s) => sheetPrintsAnything(s, sheet.styles));
  const fallbackSheet = anyPrints
    ? undefined
    : (visible.find((s) => hasHeaderOrFooter(s)) ?? visible[0]);
  for (let sheetIdx = 0; sheetIdx < sheet.sheets.length; sheetIdx++) {
    const ws = sheet.sheets[sheetIdx]!;
    // §18.2.19: a hidden tab is not printed. Excel and LibreOffice both leave it
    // out entirely — tdf171828.xlsx hides its lookup table, and printing it added
    // two pages of working data to the end of the document.
    if (ws.hidden) continue;
    // …and neither is a sheet with nothing on it. tdf115159.xlsx carries two
    // untouched tabs beside its data, and printing them added a blank page the
    // reference does not produce. A workbook of nothing but empty sheets is
    // still a document, though, and the one it prints is the one carrying a
    // header or footer (see `fallbackSheet`). HeaderFooterComplexFormats.xlsx
    // is three empty tabs whose FIRST has both, and keeping the LAST one
    // standing printed the blank third: a page with nothing on it where
    // LibreOffice prints two lines of formatted text.
    if (!sheetPrintsAnything(ws, sheet.styles) && ws !== fallbackSheet) {
      printableSheets--;
      continue;
    }

    // The grid is projected FIRST, before the section is built: the header band
    // has to be laid out at the sheet's print scale, and the scale only falls
    // out of the projection. Its blocks are pushed below, in their old place.
    const scaleSink = { value: 1 };
    const bandSink = { lefts: [0] };
    const drawingExtentPt = drawingReachPt(ws);
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
      // §18.8.28 — the Normal style's size, which sets the height of a row the
      // sheet gives none. A workbook property, so it is read here rather than
      // asked of the caller.
      ...(sheet.styles.fonts[0]?.sizePt !== undefined
        ? { defaultFontPt: sheet.styles.fonts[0].sizePt }
        : {}),
      ...(options.losses ? { losses: options.losses } : {}),
      scaleSink,
      bandSink,
      ...(drawingExtentPt ? { drawingExtentPt } : {}),
    });

    // Each sheet's header/footer band and page geometry are its own.
    const sheetSection = withHeaderFooter(
      sectionFromWorksheet(ws.grid),
      ws,
      headersFooters,
      scaleSink.value,
      sheet.styles.fonts[0]?.sizePt,
      printed,
      options.fileName,
      sheet.themePalette,
      options.now,
    );
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

    // The sheet's drawings go in BEFORE its grid. They are out-of-flow floats,
    // so they consume no space and the grid still starts at the top — but a
    // float lands on whatever page the layout has reached when it meets the
    // block, and a wide sheet's grid is several pages of column bands. Emitted
    // after them, every chart on the sheet ended up on the LAST of those pages:
    // chart_hyperlink.xlsx anchors two charts under its data in the first band
    // and we printed them alone on the second page.
    //
    // First page of the sheet, then — which is where a drawing anchored in the
    // first band belongs, and that is nearly all of them. One anchored in a
    // later band still lands too early; putting each drawing on its own band's
    // page needs the band boundaries the grid projection keeps to itself.

    // Collected rather than pushed: a drawing anchored in a later column band
    // has to go in beside THAT band's table, not ahead of the whole grid.
    const drawings: Array<BodyElement> = [];

    // §20.5: the sheet's chart frames render as blocks after its grid,
    // anchor-ordered (resolved chart data lives in sheet.chartData).
    for (const ref of ws.charts ?? []) {
      drawings.push({
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
      drawings.push({
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
      drawings.push({
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
    // …and DOWN as well as across. singlecontrol.xlsx anchors its one checkbox
    // 7331pt below the sheet's top — nine pages past the only page its (empty)
    // grid produces — and every trace of it fell off the document. Down first,
    // then over, the way Excel paginates; a sheet whose drawings all fit one
    // page deep bands exactly as before.
    const drawingBands =
      gridBody.length === 0
        ? bandDrawings(controlBlocks, printableHeightPt(ws.grid), scaleSink.value, DOWN).flatMap(
            (row) => bandDrawings(row, printableWidthPt(ws.grid), scaleSink.value),
          )
        : [controlBlocks];
    for (let band = 0; band < drawingBands.length; band++) {
      if (band > 0) {
        drawings.push({
          kind: 'paragraph',
          paragraph: { properties: { pageBreakBefore: true }, runs: [] },
        });
      }
      for (const shape of drawingBands[band]!) {
        drawings.push({
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

    body.push(...withDrawingsByBand(drawings, gridBody, bandSink.lefts));
    // §SV2: slicer panels render as styled button boxes after the grid + charts.
    for (const slicer of ws.slicers ?? []) {
      body.push({ kind: 'table', table: slicerTable(slicer) });
    }

    // W7: cell comments / notes are listed in a "Comments" section after the
    // grid — but only when the sheet asks for it. §18.3.1.63 `cellComments`
    // defaults to `none`, which is the print dialog's own default: a note is an
    // editing annotation, and neither Excel nor LibreOffice puts one on paper
    // unless told to.
    //
    // This started out the other way, on the reasoning that dropping content
    // the document carries is worse than showing a note nobody asked for. Two
    // corpus documents settled it — tdf171828.xlsx says `none` outright and
    // NamedSheetViews.xlsx says nothing at all, and neither reference prints a
    // word of either. A default the format states and both readers honour is
    // not an unstated one. The omission is reported rather than silent.
    const printComments =
      ws.grid.pageSetup?.cellComments === 'atEnd' ||
      ws.grid.pageSetup?.cellComments === 'asDisplayed';
    if (ws.comments && ws.comments.length > 0) {
      if (printComments) body.push(...commentBlocks(ws.comments));
      else
        options.losses?.push({
          severity: 'dropped',
          feature: 'comments',
          detail:
            `${ws.comments.length} cell note(s) not printed — ` +
            `<pageSetup cellComments> is "${ws.grid.pageSetup?.cellComments ?? 'none'}"`,
          where: `sheet "${ws.name}"`,
        });
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
  // §18.3.* — an embedded file draws as an icon and a caption, and both live in
  // a metafile we do not decode. Saying so beats a page that quietly lost two
  // attachments (bug64512_embed.xlsx).
  if (sheet.embeddedObjects) {
    options.losses?.push({
      severity: 'dropped',
      feature: 'shapes',
      detail: `${String(sheet.embeddedObjects)} embedded object(s) not rendered — an OLE embedding draws as an icon from a metafile preview`,
    });
  }
  // …and a picture that IS a metafile is the same gap without the OLE wrapper:
  // WithDrawing.xlsx anchors five and three of them are WMF/EMF, which both
  // references draw and no page of ours does.
  if (sheet.metafilePictures) {
    options.losses?.push({
      severity: 'dropped',
      feature: 'images',
      detail: `${String(sheet.metafilePictures)} picture(s) not rendered — a WMF/EMF/PICT metafile is replayed, not embedded; the anchor keeps its space`,
    });
  }
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
      const mark = side - 2 * RADIO_DOT_INSET_PT;
      if (glyph === 'radio') {
        out.push(
          controlShape(box, RADIO_DOT_INSET_PT, top + RADIO_DOT_INSET_PT, mark, mark, {
            geometry: ellipse,
            fill: { kind: 'solid', colorHex: '000000' },
          }),
        );
      } else {
        // A ticked check box is a CROSS in the box, in both Excel and Calc.
        // Filling the square the way a radio fills its ring reads as a colour
        // swatch: the difference between checked and unchecked is then the
        // amount of black, not a mark, and singlecontrol.xlsx's one control
        // came out a solid blob where the reference draws ☒.
        for (const flipV of [false, true]) {
          out.push(
            controlShape(box, RADIO_DOT_INSET_PT, top + RADIO_DOT_INSET_PT, mark, mark, {
              geometry: { kind: 'preset', preset: 'line' },
              fill: noFill,
              line: { colorHex: '000000', width: pt(RADIO_RING_PT) },
              ...(flipV ? { transform: { flipV: true } } : {}),
            }),
          );
        }
      }
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
/**
 * Whether a sheet has anything to put on a page.
 *
 * A cell that carries a value or inline text, any drawing, control or slicer,
 * or a printed note. A tab nobody touched has none of those, and Excel and
 * LibreOffice both leave it out of the print — a header or footer does not
 * rescue it.
 *
 * @param ws The sheet.
 * @returns True when the sheet would print something.
 */
/**
 * Whether a sheet carries a header or footer with anything in it.
 *
 * @param ws The sheet.
 * @returns True when either band has a non-empty format string.
 */
function hasHeaderOrFooter(ws: Sheet): boolean {
  const hf = ws.grid.headerFooter;
  return Boolean(hf && ((hf.oddHeader?.length ?? 0) > 0 || (hf.oddFooter?.length ?? 0) > 0));
}

function sheetPrintsAnything(ws: Sheet, styles: XlsxStyles): boolean {
  if (ws.grid.cells.some((c) => c.rawValue !== '' || c.inlineText !== undefined)) return true;
  // A cell that paints is on the page with nothing in it, and a sheet of them
  // is a sheet: 48779.xlsx is one workbook of three tabs whose whole content is
  // A1 filled red, and judged on values alone all three read as empty — the
  // last one standing printed, and it was one of the blank two.
  if (ws.grid.cells.some((c) => cellPaintsVisibly(c, styles))) return true;
  if ((ws.charts?.length ?? 0) > 0) return true;
  if ((ws.images?.length ?? 0) > 0) return true;
  if ((ws.shapes?.length ?? 0) > 0) return true;
  if ((ws.slicers?.length ?? 0) > 0) return true;
  if ((ws.comments?.length ?? 0) > 0) return true;
  if ((ws.formControls?.length ?? 0) > 0) return true;
  if ((ws.activeXControls?.length ?? 0) > 0) return true;
  // A header or footer is NOT content. I let one count when this rule went in;
  // 47737.xlsx has a sheet whose every cell is valueless and whose only text is
  // `<oddHeader>Agency Footnotes</oddHeader>`, and Excel refuses to print it at
  // all ("We didn't find anything to print"). LibreOffice agrees. The guard at
  // the call site still keeps a workbook of nothing but empty sheets from
  // printing zero pages.
  return false;
}

/**
 * Put each drawing beside the column band it is anchored in.
 *
 * A wide sheet's grid prints as one table per band, each on its own page, and a
 * float lands on whatever page the layout has reached when it meets the block.
 * Emitted ahead of the whole grid, a drawing anchored in the second band was
 * printed on the FIRST page at an offset a page and a half wide — off the paper
 * entirely. shape-macro-ext-ref.xlsx anchors its chart and its macro button at
 * x=838pt and x=891pt of a 480pt page and lost both.
 *
 * Each drawing goes in front of its band's table with its horizontal anchor
 * measured from that band's left edge instead of the sheet's.
 *
 * @param drawings The sheet's floats, in the order they were projected.
 * @param gridBody The grid's body elements — one table per band.
 * @param lefts    Each band's left edge in points (from the projection's sink).
 * @returns The two merged; drawings first when the sheet does not band.
 */
function withDrawingsByBand(
  drawings: ReadonlyArray<BodyElement>,
  gridBody: ReadonlyArray<BodyElement>,
  lefts: ReadonlyArray<number>,
): Array<BodyElement> {
  const tables = gridBody.filter((el) => el.kind === 'table').length;
  // One band, or a grid whose tables and bands do not line up: nothing to
  // distribute against, so keep the order the sheet had before bands existed.
  if (drawings.length === 0 || lefts.length < 2 || tables !== lefts.length) {
    return [...drawings, ...gridBody];
  }
  const bandOf = (x: number): number => {
    let band = 0;
    while (band + 1 < lefts.length && x >= lefts[band + 1]!) band++;
    return band;
  };
  const byBand: Array<Array<BodyElement>> = lefts.map(() => []);
  for (const el of drawings) {
    const x = floatLeftPt(el);
    const band = x === undefined ? 0 : bandOf(x);
    byBand[band]!.push(x === undefined ? el : shiftFloatLeft(el, lefts[band]!));
  }
  // AFTER the band's table, not before it: a band table breaks to its own page
  // on its FIRST ROW, so a float emitted ahead of it is still on the page
  // before. The layout is on band k's page from the moment table k starts until
  // table k+1 breaks, and the end of table k is inside that window.
  const out: Array<BodyElement> = [];
  let seen = 0;
  for (const el of gridBody) {
    out.push(el);
    if (el.kind === 'table') out.push(...byBand[seen++]!);
  }
  return out;
}

/** A float's horizontal anchor, for the block kinds a sheet anchors. */
function floatLeftPt(el: BodyElement): number | undefined {
  const float =
    el.kind === 'chart'
      ? el.chart.float
      : el.kind === 'image'
        ? el.image.float
        : el.kind === 'shape'
          ? el.shape.float
          : undefined;
  return float?.posH?.offsetPt;
}

/** The same block with its horizontal anchor measured from `by` points later. */
function shiftFloatLeft(el: BodyElement, by: number): BodyElement {
  if (by === 0) return el;
  const moved = <T extends { float?: FloatAnchor }>(block: T): T => ({
    ...block,
    float: {
      ...block.float!,
      posH: { ...block.float!.posH!, offsetPt: pt((block.float!.posH!.offsetPt ?? 0) - by) },
    },
  });
  if (el.kind === 'chart') return { ...el, chart: moved(el.chart) };
  if (el.kind === 'image') return { ...el, image: moved(el.image) };
  if (el.kind === 'shape') return { ...el, shape: moved(el.shape) };
  return el;
}

interface BandAxis {
  /** The shape's anchor along this axis, as the float stores it (unscaled). */
  readonly anchorOf: (s: ShapeBlock) => number | undefined;
  /** How far the shape reaches along this axis, in points. */
  readonly sizeOf: (s: ShapeBlock) => number;
  /** The same float with its anchor moved back by `by` (unscaled). */
  readonly shifted: (s: ShapeBlock, by: number) => ShapeBlock;
}

const ACROSS: BandAxis = {
  anchorOf: (s) => s.float?.posH?.offsetPt,
  sizeOf: (s) => s.width,
  shifted: (s, by) => ({
    ...s,
    float: {
      ...s.float!,
      posH: { ...s.float!.posH!, offsetPt: pt((s.float!.posH!.offsetPt ?? 0) - by) },
    },
  }),
};

const DOWN: BandAxis = {
  anchorOf: (s) => s.float?.posV?.offsetPt,
  sizeOf: (s) => s.height,
  shifted: (s, by) => ({
    ...s,
    float: {
      ...s.float!,
      posV: { ...s.float!.posV!, offsetPt: pt((s.float!.posV!.offsetPt ?? 0) - by) },
    },
  }),
};

function bandDrawings(
  shapes: ReadonlyArray<ShapeBlock>,
  printable: number,
  scale: number,
  axis: BandAxis = ACROSS,
): Array<Array<ShapeBlock>> {
  if (shapes.length === 0 || !(printable > 0)) return [shapes as Array<ShapeBlock>];
  const rightOf = (s: ShapeBlock): number => (axis.anchorOf(s) ?? 0) * scale + axis.sizeOf(s);
  const extent = Math.max(0, ...shapes.map(rightOf));
  const bandCount = Math.ceil(extent / printable);
  if (bandCount <= 1) return [shapes as Array<ShapeBlock>];
  const bands: Array<Array<ShapeBlock>> = Array.from({ length: bandCount }, () => []);
  for (const shape of shapes) {
    const anchor = axis.anchorOf(shape);
    const left = (anchor ?? 0) * scale;
    const right = left + axis.sizeOf(shape);
    const first = Math.max(0, Math.min(bandCount - 1, Math.floor(left / printable)));
    // EVERY band the drawing reaches into, not just the one it starts in. A
    // group box 209pt wide anchored 46pt before the boundary is on both pages —
    // clipped at the edge of the first and continuing from the edge of the
    // second, which is exactly what the reference draws. Assigned to one band
    // only, its far half and every caption inside it fell off the document.
    const last = Math.max(first, Math.min(bandCount - 1, Math.ceil(right / printable) - 1));
    for (let band = first; band <= last; band++) {
      if (!shape.float || anchor === undefined) {
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
      bands[band]!.push(axis.shifted(band === first ? shape : frame, (band * printable) / scale));
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
        // Its CAPTION, never its name: the name is an internal id, and drawing
        // it put "CheckBox28" over the form's own text forty times
        // (45540_form_Header.xlsx).
        c.caption,
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
        // Its Caption property, and nothing else. `<control name>` is the
        // shape's identifier — Excel shows it in the name box, never on the
        // page — and falling back to it printed "CheckBox28" over the form's own
        // text forty times (45540_form_Header.xlsx, whose boxes have no
        // caption at all).
        c.caption,
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
function drawingReachPt(ws: Sheet): { widthPt: number; heightPt: number } | undefined {
  let widthPt = 0;
  let heightPt = 0;
  for (const shape of ws.shapes ?? []) {
    widthPt = Math.max(widthPt, (shape.float?.posH?.offsetPt ?? 0) + shape.width);
    heightPt = Math.max(heightPt, (shape.float?.posV?.offsetPt ?? 0) + shape.height);
  }
  // Charts and pictures are printed too, and only the SHAPES were measured —
  // so a sheet whose only drawing is a chart reported no extent at all, and
  // 57362.xlsx's chart hung 350pt off the right edge of a page the grid alone
  // said needed no splitting.
  for (const c of ws.charts ?? []) {
    widthPt = Math.max(widthPt, (c.xPt ?? 0) + c.widthPt);
    heightPt = Math.max(heightPt, (c.yPt ?? 0) + c.heightPt);
  }
  for (const p of ws.images ?? []) {
    widthPt = Math.max(widthPt, (p.xPt ?? 0) + p.widthPt);
    heightPt = Math.max(heightPt, (p.yPt ?? 0) + p.heightPt);
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

/**
 * The sheet's header/footer bands, registered under ids of this sheet's own.
 *
 * The map is document-level and the ids used to be fixed, so a second sheet
 * overwrote the first's bands and every section pointed at the last one — in
 * practice only the first sheet's header survived, because only its section
 * carried the references. formats.xlsx puts the sheet name and page number on
 * all five of its pages and we printed them on one.
 */
function withHeaderFooter(
  section: SectionProperties,
  ws: SheetDoc['sheets'][number],
  headersFooters: Map<string, ReadonlyArray<BodyElement>>,
  scale: number,
  basePt: number | undefined,
  sheetIdx: number,
  fileName?: string,
  themePalette?: ReadonlyMap<string, string>,
  now?: Date,
): SectionProperties {
  const headerRel = `${HEADER_REL}${sheetIdx}`;
  const footerRel = `${FOOTER_REL}${sheetIdx}`;
  const hf = ws.grid.headerFooter;
  if (!hf || (!hf.oddHeader && !hf.oddFooter)) return section;
  const headers: Array<HeaderFooterReference> = [];
  const footers: Array<HeaderFooterReference> = [];
  if (hf.oddHeader) {
    const content = buildHeaderFooterContent(
      hf.oddHeader,
      ws.name,
      scale,
      basePt,
      fileName,
      themePalette,
      now,
    );
    if (content.length > 0) {
      headersFooters.set(headerRel, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      headers.push({ type: 'default', relationshipId: headerRel });
    }
  }
  if (hf.oddFooter) {
    const content = buildHeaderFooterContent(
      hf.oddFooter,
      ws.name,
      scale,
      basePt,
      fileName,
      themePalette,
      now,
    );
    if (content.length > 0) {
      headersFooters.set(footerRel, resolveBodyStyles(content, EMPTY_STYLE_SHEET));
      footers.push({ type: 'default', relationshipId: footerRel });
    }
  }
  if (headers.length === 0 && footers.length === 0) return section;
  return { ...section, headers, footers };
}
