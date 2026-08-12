// E-PDF — shared FlowDoc construction for the two reconstruction paths (the
// tagged fast-path EP3 and the heuristic layout path EP4). A reconstructed PDF
// carries a body of paragraphs/tables/images over the empty style sheet, with
// any lifted image bytes (EP6) in the resource store the writers embed from.

import { displayOf } from './display';
import type {
  BodyElement,
  CustomPathCmd,
  FloatAnchor,
  ParagraphProperties,
  Section,
  SectionProperties,
  ShapeFill,
  ShapeLine,
  TextOutline,
} from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { FontRegistry } from '@/core/font';
import type { Loss, Pt } from '@/core/ir';

import type { PdfImage } from './images';
import type { PdfPage } from './document';
import type { PdfVector } from './vector';
import type { TextMarkup } from './annot-draw';
import type { TextRun } from './content';
import { ResourceStore, pt } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

/**
 * A reconstruction's document plus the losses incurred reading it (e.g. an
 * undecodable image colour space) — surfaced through the reader's `LossReport`.
 */
export interface Reconstruction {
  readonly doc: FlowDoc;
  readonly losses: ReadonlyArray<Loss>;
}

/**
 * The corner a page's marks are measured from, in the SHOWN page's own y-up
 * frame — see `./display`, which puts every mark into it.
 *
 * The shown page starts at its own origin, so `left` is zero and `top` is its
 * height; the two are kept as a pair because a caller that has not been through
 * `display` (none, today) would state something else.
 */
export interface PageFrame {
  /** Left edge — subtract it to get an offset from the page. */
  readonly left: number;
  /** Top edge — subtract from it to flip into a top-down frame. */
  readonly top: number;
}

/**
 * Build a paragraph {@link BodyElement} from a single plain-text string,
 * optionally at the given outline (heading) level. Empty text yields a
 * paragraph with no runs. The link-free counterpart of {@link paragraphFromRuns}.
 */
export function paragraphBlock(text: string, outlineLevel?: number): BodyElement {
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: { properties, runs: text.length > 0 ? [{ text, properties: {} }] : [] },
  };
}

/** One piece of reconstructed text, carrying any hyperlink (E-PDF EP8). */
export interface TextSpan {
  readonly text: string;
  readonly href?: string;
  /** The size the glyphs were SHOWN at (§9.3.1 Tf), so the run keeps it. */
  readonly sizePt?: number;
  /** §8.6.8 — the colour they were painted in, when it is not plain black. */
  readonly colorHex?: string;
  /** §9.6.2 — the face's own name, for a document that embeds its programs. */
  readonly fontName?: string;
  /** §9.3.6 — a line round the glyphs, when the page asked for one. */
  readonly outline?: TextOutline;
  /** §9.8.1 — the face was a bold one. */
  readonly bold?: boolean;
  /** §9.8.1 — the face was a slanted one. */
  readonly italic?: boolean;
  /**
   * §17.3.2.42 — the glyphs stood OFF the baseline the line is set on, and
   * smaller: a footnote mark, an exponent, an index. The page states it by
   * placement, and a document by the property.
   */
  readonly script?: 'superscript' | 'subscript';
  /** §12.5.6.10 — a text-markup annotation marks these words. */
  readonly markup?: TextMarkup;
}

/**
 * Build a paragraph {@link BodyElement} from positioned {@link TextSpan}s,
 * coalescing consecutive spans that share an `href` into one run (so a link
 * survives as its own run) and squashing whitespace. With no hrefs this
 * collapses to a single run — the same shape {@link paragraphBlock} produces.
 */
export function paragraphFromRuns(
  spans: ReadonlyArray<TextSpan>,
  outlineLevel?: number,
  placement?: Pick<
    ParagraphProperties,
    'alignment' | 'spacingBefore' | 'indentLeft' | 'indentFirstLine'
  >,
): BodyElement {
  const merged: Array<{
    text: string;
    href?: string;
    sizePt?: number;
    colorHex?: string;
    fontName?: string;
    outline?: TextOutline;
    bold?: boolean;
    italic?: boolean;
    script?: 'superscript' | 'subscript';
    markup?: TextMarkup;
  }> = [];
  for (const s of spans) {
    const last = merged[merged.length - 1];
    if (
      last &&
      last.href === s.href &&
      last.sizePt === s.sizePt &&
      last.colorHex === s.colorHex &&
      last.fontName === s.fontName &&
      last.outline?.colorHex === s.outline?.colorHex &&
      last.outline?.widthPt === s.outline?.widthPt &&
      last.bold === s.bold &&
      last.italic === s.italic &&
      last.script === s.script &&
      sameMarkup(last.markup, s.markup)
    ) {
      last.text += s.text;
    } else
      merged.push({
        text: s.text,
        ...(s.href !== undefined ? { href: s.href } : {}),
        ...(s.sizePt !== undefined ? { sizePt: s.sizePt } : {}),
        ...(s.colorHex !== undefined ? { colorHex: s.colorHex } : {}),
        ...(s.fontName !== undefined ? { fontName: s.fontName } : {}),
        ...(s.outline !== undefined ? { outline: s.outline } : {}),
        ...(s.bold !== undefined ? { bold: s.bold } : {}),
        ...(s.italic !== undefined ? { italic: s.italic } : {}),
        ...(s.script !== undefined ? { script: s.script } : {}),
        ...(s.markup !== undefined ? { markup: s.markup } : {}),
      });
  }
  const runs = merged
    .map((m) => ({ ...m, text: m.text.replace(/\s+/g, ' ') }))
    .filter((m) => m.text.length > 0);
  // Trim the paragraph's outer whitespace.
  if (runs.length > 0) {
    runs[0]!.text = runs[0]!.text.replace(/^ /, '');
    runs[runs.length - 1]!.text = runs[runs.length - 1]!.text.replace(/ $/, '');
  }
  // §17.3.1 — how the page SET it, beside what it says: a PDF states no
  // alignment and no paragraph spacing, and both are read back off where the
  // lines were placed (see `layout.ts`).
  const properties: ParagraphProperties = {
    ...(outlineLevel !== undefined ? { outlineLevel } : {}),
    ...placement,
  };
  return {
    kind: 'paragraph',
    paragraph: {
      properties,
      runs: runs
        .filter((r) => r.text.length > 0)
        .map((r) => ({
          text: r.text,
          // §9.3.1/§8.6.8/§9.8.1 — the size, colour and face the page showed it
          // in. Dropped, a form set in 7pt was rebuilt at the 11pt default and
          // grew by half again, its blue field labels and red warning came back
          // black, and 160F-2019.pdf's title, set in Arial-BoldMT, came back
          // light.
          properties: {
            ...(r.sizePt !== undefined ? { fontSizePt: pt(r.sizePt) } : {}),
            ...(r.colorHex !== undefined ? { colorHex: r.colorHex } : {}),
            // §17.3.2.26 `w:rFonts` — the face by name, which is how the layout
            // finds a program the document itself carries.
            ...(r.fontName !== undefined ? { fontFamily: { ascii: r.fontName } } : {}),
            // §9.3.6 / §21.1.2.3.9 — the page drew a line round these glyphs.
            ...(r.outline !== undefined ? { textOutline: r.outline } : {}),
            ...(r.bold ? { bold: true } : {}),
            ...(r.italic ? { italic: true } : {}),
            // §17.3.2.42 `w:vertAlign` — the page set these glyphs off the
            // line's own baseline and smaller. The SIZE that carries is the
            // line's, not the mark's: a document states the nominal size and
            // the layout shrinks a script, so keeping the drawn 7pt under a
            // superscript would draw it at five.
            ...(r.script !== undefined ? { verticalAlign: r.script } : {}),
            // §12.5.6.10 — a highlight, an underline or a strikeout stated
            // ABOUT these words rather than painted among them, so it re-sets
            // with them: §17.3.2.32 `w:shd`, §17.3.2.40 `w:u`, §17.3.2.37
            // `w:strike`.
            ...(r.markup?.highlightHex !== undefined
              ? { shadingColorHex: r.markup.highlightHex }
              : {}),
            ...(r.markup?.underline !== undefined ? { underline: r.markup.underline } : {}),
            ...(r.markup?.underlineHex !== undefined
              ? { underlineColorHex: r.markup.underlineHex }
              : {}),
            ...(r.markup?.strike === true ? { strike: true } : {}),
          },
          ...(r.href ? { href: r.href } : {}),
        })),
    },
  };
}

/** Whether two runs are marked the same way, so they may join into one. */
function sameMarkup(a: TextMarkup | undefined, b: TextMarkup | undefined): boolean {
  return (
    a?.highlightHex === b?.highlightHex &&
    a?.underline === b?.underline &&
    a?.underlineHex === b?.underlineHex &&
    a?.strike === b?.strike
  );
}

/**
 * Store a {@link PdfImage}'s bytes (content-addressed dedup) and build the image
 * {@link BodyElement} that references them, sized in points from the placement
 * CTM. `alt` becomes the block's alt text when given.
 */
export function imageBlock(
  image: PdfImage,
  resources: ResourceStore,
  alt?: string,
  frame?: PageFrame,
  zOrder?: number,
  behind = false,
): BodyElement {
  const resource = resources.put(image.bytes);
  // §20.4.2.3 — anchored where the page placed it, for the same reason a lifted
  // path is: a picture is not a paragraph and has no turn in a reading order.
  // Stacked in flow, 22060_A1_01_Plans.pdf's four floor plans made two pages of
  // a sheet that is one.
  const float: FloatAnchor | undefined =
    frame !== undefined
      ? {
          wrap: 'none',
          ...(behind ? { behind: true } : {}),
          ...(zOrder !== undefined ? { zOrder } : {}),
          posH: { relativeFrom: 'page', offsetPt: pt(image.x - frame.left) },
          posV: {
            relativeFrom: 'page',
            offsetPt: pt(Math.max(0, frame.top - image.y - image.heightPt)),
          },
        }
      : undefined;
  return {
    kind: 'image',
    image: {
      ...(float ? { float } : {}),
      resource,
      width: pt(image.widthPt),
      height: pt(image.heightPt),
      // §20.1.7.6 — a turn is stated in sixtieth-thousandths of a degree, and
      // clockwise, which is the other way round from the page's own axis.
      ...(image.rotationDeg !== undefined
        ? { rotation60k: Math.round(-image.rotationDeg * 60000) }
        : {}),
      // §20.1.8.55 — the box above is what the clip left showing, so the source
      // must be cut to match it or the whole picture squeezes into it.
      ...(image.crop ? { crop: image.crop } : {}),
      // §20.1.8.4 `a:alphaModFix` — the page asked for the picture to be seen
      // through, and a format that can say so should say so.
      ...(image.alpha !== undefined ? { alpha: image.alpha } : {}),
      paragraphProperties: {},
      ...(alt ? { altText: alt } : {}),
    },
  };
}

/**
 * A line of text as an anchored box, standing where the page set it.
 *
 * The flowed reconstruction reads a document OUT of a page: paragraphs in
 * reading order, re-flowable, free to land wherever the next medium puts them.
 * A form is not that document. 160F-2019.pdf is a grid of ruled boxes with a
 * label in each, and a label means nothing an inch from the box it labels — the
 * artwork is placed absolutely, so text that flows beside it lines up with none
 * of it.
 *
 * @param spans       The line's runs.
 * @param box         Its page-space rectangle (y-up, as PDF measures).
 * @param frame       The page's own corner, to measure the box off.
 * @param zOrder      Its place in the page's painting order.
 * @param rotation60k §20.1.7.6 — how far the box turns about its own centre,
 *                    for a baseline the page did not set flat.
 * @returns A shape carrying the text, anchored where the glyphs were.
 */
export function positionedText(
  spans: ReadonlyArray<TextSpan>,
  box: { x: number; y: number; width: number; height: number },
  frame: PageFrame,
  zOrder: number,
  rotation60k?: number,
): BodyElement {
  const paragraph = paragraphFromRuns(spans);
  return {
    kind: 'shape',
    shape: {
      float: {
        wrap: 'none',
        zOrder,
        posH: { relativeFrom: 'page', offsetPt: pt(box.x - frame.left) },
        posV: {
          relativeFrom: 'page',
          offsetPt: pt(Math.max(0, frame.top - box.y - box.height)),
        },
      },
      width: pt(Math.max(1, box.width)),
      height: pt(Math.max(1, box.height)),
      ...(rotation60k !== undefined ? { transform: { rotation60k } } : {}),
      geometry: { kind: 'preset', preset: 'rect' },
      fill: { kind: 'none' },
      // A box drawn round a line of a form would be a box the page never had:
      // the shape is here to place the words, not to be seen.
      text: {
        content: [paragraph],
        insetLeft: pt(0),
        insetTop: pt(0),
        insetRight: pt(0),
        insetBottom: pt(0),
      },
      paragraphProperties: {},
    },
  };
}

/** Collapse losses sharing a `detail` message (the same colour space dropped on many pages). */
export function dedupeLosses(losses: ReadonlyArray<Loss>): Array<Loss> {
  const byDetail = new Map<string, Loss>();
  for (const loss of losses) if (!byDetail.has(loss.detail)) byDetail.set(loss.detail, loss);
  return [...byDetail.values()];
}

/**
 * Turn a lifted {@link PdfVector} path (filled EP10 / stroked EP11) into a
 * custom-geometry shape {@link BodyElement}. Page-space points (y-up) become
 * path-space (bbox-relative, y-down); the shape is sized from the bounding box
 * (plus the stroke thickness). A fill becomes a solid fill, a stroke the outline.
 *
 * Given the page's frame the shape is ANCHORED where the page drew it, behind
 * the text, rather than taking a place of its own in the flow. A drawing is not
 * a paragraph: 22060_A1_01_Plans.pdf is one A3 sheet of vectors, and stacking
 * its forty-nine paths one under another spilled it onto a second page.
 */
export function shapeBlock(
  v: PdfVector,
  frame?: PageFrame,
  zOrder?: number,
  behind = false,
): BodyElement {
  const w = v.maxX - v.minX;
  const h = v.maxY - v.minY;
  const fx = (x: number): number => x - v.minX;
  const fy = (y: number): number => v.maxY - y; // flip to top-left origin
  const commands: Array<CustomPathCmd> = v.segs.map((s): CustomPathCmd => {
    switch (s.op) {
      case 'move':
        return { cmd: 'move', x: fx(s.x), y: fy(s.y) };
      case 'line':
        return { cmd: 'line', x: fx(s.x), y: fy(s.y) };
      case 'cubic':
        return {
          cmd: 'cubic',
          x1: fx(s.x1),
          y1: fy(s.y1),
          x2: fx(s.x2),
          y2: fy(s.y2),
          x: fx(s.x),
          y: fy(s.y),
        };
      case 'close':
        return { cmd: 'close' };
    }
  });
  // §8.4.3.2 — the pen is as wide as the page says. A width of zero asks for the
  // thinnest line the device can draw, which on a fixed page is a hairline.
  //
  // This used to be raised to half a point, because the same number also sized
  // the SHAPE BOX and a flat line needs a box to draw in. The two are not the
  // same thing: Brotli-Prototype-FileA.pdf draws its elevations with a 0.12pt
  // pen, and at half a point every clapboard line came out four times too heavy
  // — a drawing that reads grey in every viewer arrived black.
  const HAIRLINE_PT = 0.1;
  const stated = v.lineWidth ?? 0.75;
  const pen = v.strokeHex !== undefined ? (stated > 0 ? stated : HAIRLINE_PT) : 0;
  // A stroked line can be geometrically flat (a horizontal rule has h≈0); give
  // the shape box at least half a point so the line has room to draw.
  const thick = v.strokeHex !== undefined ? Math.max(pen, 0.5) : 0;
  // §11.6.4.4 — a band the page meant to be read THROUGH is not the same mark
  // as one that hides what it covers: 22060_A1_01_Plans.pdf marks its
  // evacuation routes at `ca` 0.6 over the floor plan they run across.
  const alpha = v.alpha !== undefined ? { alpha: v.alpha } : {};
  const fill: ShapeFill =
    v.gradient !== undefined
      ? { kind: 'gradient', gradient: v.gradient, ...alpha }
      : v.fillHex !== undefined
        ? { kind: 'solid', colorHex: v.fillHex, ...alpha }
        : { kind: 'none' };
  const line: ShapeLine | undefined =
    v.strokeHex !== undefined
      ? { width: pt(pen), colorHex: v.strokeHex, fill: 'solid' }
      : undefined;
  // §20.4.2.3 — anchored to the PAGE at the position it was drawn at, y flipped
  // from PDF's upward axis.
  //
  // `behind` is about the TEXT, and `zOrder` about the other marks: a legend's
  // white box still covers the floor plan it backs, because both are behind and
  // the order the page painted them in still ranks them. What `behind` decides
  // is whether artwork may cover WORDS, and in a flowing reading it may not —
  // there the words have moved and the artwork has not, so anything over them
  // covers text it never covered. annotation-tx3.pdf is a form field filled
  // pale blue with four lines typed in it, and the fill buried all four.
  //
  // §11.3.5 also puts a mark that only DARKENS behind the text wherever it is
  // read: no anchor blends, so a highlighter over its words would bury them,
  // and under them it comes to the same picture.
  const float: FloatAnchor | undefined =
    frame !== undefined
      ? {
          wrap: 'none',
          ...(behind || v.darkens === true ? { behind: true } : {}),
          ...(zOrder !== undefined ? { zOrder } : {}),
          posH: { relativeFrom: 'page', offsetPt: pt(v.minX - frame.left) },
          posV: { relativeFrom: 'page', offsetPt: pt(Math.max(0, frame.top - v.maxY)) },
        }
      : undefined;
  return {
    kind: 'shape',
    shape: {
      ...(float ? { float } : {}),
      width: pt(Math.max(w, thick)),
      height: pt(Math.max(h, thick)),
      geometry: { kind: 'custom', custom: { pathWidth: w, pathHeight: h, commands } },
      fill,
      ...(line ? { line } : {}),
      paragraphProperties: {},
    },
  };
}

/**
 * Derive the {@link SectionProperties} geometry from the source pages so a
 * reconstructed PDF re-renders at its real page size and orientation rather than
 * the layout engine's `A4` default — without it an `A3` source paginates onto
 * several `A4` pages, a wide/landscape page is letterboxed, and so on. PDF
 * user-space units are points, matching `Pt`, so the `MediaBox` extents map
 * straight to the page size. Uses the first {@link PdfPage}'s box (the
 * near-universal uniform-size case); a PDF whose pages differ in size reflows to
 * this single geometry — a known approximation, still far better than a fixed
 * `A4`. Returns `undefined` when there is no usable first-page box.
 */
export function sectionFromPdfPages(pages: ReadonlyArray<PdfPage>): SectionProperties | undefined {
  const first = pages[0];
  if (!first) return undefined;
  // §14.11.1 — the page as it is SHOWN. A landscape sheet drawn sideways in a
  // portrait box with `/Rotate 270` is a landscape page, and read as its box
  // says every one of Brotli-Prototype-FileA.pdf's twenty-five came back
  // portrait with its words running down the page.
  const shown = displayOf(first);
  const width = shown.width;
  const height = shown.height;
  if (!(width > 0 && height > 0)) return undefined;
  return {
    pageSize: {
      width: pt(width),
      height: pt(height),
      orientation: width > height ? 'landscape' : 'portrait',
    },
    // A PDF page has no margin model — text is positioned absolutely anywhere on
    // the MediaBox — so the page box is the content box. Zero margins keep the
    // reflow from inventing a 1-inch inset the source never had, never clip a
    // small page (the layout default of 1 inch can exceed a tiny MediaBox), and
    // reduce spurious over-pagination.
    margins: { top: pt(0), right: pt(0), bottom: pt(0), left: pt(0) },
    headers: [],
    footers: [],
  };
}

/**
 * The margins the SOURCE used, measured off where its words actually sit.
 *
 * A PDF states none — text is placed anywhere on the MediaBox — so the reader
 * used to leave them at zero rather than invent an inch. But the words
 * themselves say where the margin was: the leftmost glyph on the page is the
 * left margin, and reflowing inside it keeps the measure the author set instead
 * of running the text from edge to edge.
 *
 * Measured on the MEDIAN page rather than the extreme one, so a single full-
 * bleed rule or a page number in the corner does not collapse the margin for
 * the whole document, and clamped so a strange page cannot leave no text area
 * at all.
 *
 * @param section  The section the page box gave, or `undefined`.
 * @param shown    Each page as it is shown, for its own width and height.
 * @param pageRuns Each page's runs, already placed on the shown page.
 * @returns The section with measured margins, or `section` when nothing is
 *          measurable.
 */
/** What the measure gives back, so the widest line still fits when re-set. */
const SLACK = 0.01;

/** How far a face's ascender stands above its baseline, as a fraction of the size. */
const ASCENDER = 0.8;

/** And its descender below — the two together are a little over one em. */
const DESCENDER = 0.22;

export function withMeasuredMargins(
  section: SectionProperties | undefined,
  shown: ReadonlyArray<{ width: number; height: number }>,
  pageRuns: ReadonlyArray<ReadonlyArray<TextRun>>,
): SectionProperties | undefined {
  if (!section?.pageSize) return section;
  const width = section.pageSize.width as number;
  const height = section.pageSize.height as number;
  const lefts: Array<number> = [];
  const rights: Array<number> = [];
  const tops: Array<number> = [];
  const bottoms: Array<number> = [];
  pageRuns.forEach((runs, i) => {
    const page = shown[i];
    if (!page || runs.length === 0) return;
    let minX = Infinity;
    let maxX = -Infinity;
    let minY = Infinity;
    let maxY = -Infinity;
    // The faces of the topmost and bottommost lines, for the room their
    // ascenders and descenders take beyond the baseline.
    let topSize = 0;
    let bottomSize = 0;
    for (const r of runs) {
      if (!Number.isFinite(r.x) || !Number.isFinite(r.y)) continue;
      minX = Math.min(minX, r.x);
      maxX = Math.max(maxX, r.endX);
      if (r.y < minY) {
        minY = r.y;
        bottomSize = r.fontSizePt;
      }
      if (r.y > maxY) {
        maxY = r.y;
        topSize = r.fontSizePt;
      }
    }
    if (!Number.isFinite(minX) || !Number.isFinite(minY)) return;
    lefts.push(minX);
    rights.push(page.width - maxX);
    // Runs carry a BASELINE, and a margin is to the top of the LINE.
    //
    // Measured to the baseline, every converted PDF came back a whole ascender
    // too low: on annotation-stamp.pdf the word "Stamp" slid under the stamp
    // anchored above it, and the same shift put a label inside its own drawing
    // on four more files of the corpus. The reader cannot know which face the
    // layout will re-set the line in, so the ascender is estimated at four
    // fifths of the size — near enough for the faces documents use, and wrong
    // by a fraction of a line where it is wrong at all, instead of by a line.
    tops.push(page.height - maxY - topSize * ASCENDER);
    bottoms.push(minY - bottomSize * DESCENDER);
  });
  if (lefts.length === 0) return section;
  const median = (xs: Array<number>): number => {
    const s = [...xs].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)] ?? 0;
  };
  // Never more than a third of the sheet, never negative: a margin that eats
  // the text area is worse than none.
  const clamp = (v: number, span: number): Pt => pt(Math.max(0, Math.min(v, span / 3)));
  return {
    ...section,
    margins: {
      left: clamp(median(lefts), width),
      // The right margin gives back a little of what it measured. The page was
      // set in faces this reader does not have, and re-setting it in
      // substitutes cannot come out narrower everywhere — so a measure exactly
      // as wide as the widest line wraps that line's last word onto the next.
      // basicapi.pdf's contents line runs 504.5pt across a 504.5pt measure, and
      // its page number came back at the head of the line below.
      right: clamp(median(rights) - width * SLACK, width),
      // Down the page the TIGHTEST page decides, not the middle one. A margin
      // is a wall the text may not cross, and the pages differ: the last one
      // ends early, and taking the middle of two puts the wall above the line
      // the first page ends on. ZapfDingbats.pdf's second sheet stops five rows
      // short of its first, and its first sheet lost a row to a page of its own.
      top: clamp(Math.min(...tops), height),
      // …and it gives back a little of what it measured, for the same reason
      // the right margin does: the page was set in faces this reader does not
      // have, and re-set in substitutes it cannot come out shorter everywhere.
      // A measure exactly as deep as the text block drops its last line onto a
      // sheet of its own.
      bottom: clamp(Math.min(...bottoms) - height * SLACK, height),
    },
  };
}

/**
 * Assemble the final {@link FlowDoc} for a reconstruction: the body elements
 * with their styles resolved against the empty style sheet, the lifted-image
 * resource store, and the optional page {@link SectionProperties}. Shared by
 * both reconstruction paths (the tagged fast-path EP3 and the heuristic layout
 * path EP4).
 */
export function buildFlowDoc(
  body: ReadonlyArray<BodyElement>,
  resources: ResourceStore = new ResourceStore(),
  section?: SectionProperties,
  embeddedFonts?: ReadonlyMap<string, FontRegistry>,
  sections: ReadonlyArray<Section> = [],
  headersFooters?: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
): FlowDoc {
  return {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    // §17.6 — a document whose pages differ in size is several sections; one
    // page size for all of them is the ordinary case and states none.
    sections,
    ...(headersFooters && headersFooters.size > 0 ? { headersFooters } : {}),
    ...(section ? { section } : {}),
    ...(embeddedFonts && embeddedFonts.size > 0 ? { embeddedFonts } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
}
