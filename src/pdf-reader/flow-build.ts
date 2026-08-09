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
  SectionProperties,
  ShapeFill,
  ShapeLine,
  TextOutline,
} from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { FontRegistry } from '@/core/font';
import type { Loss } from '@/core/ir';

import type { PdfImage } from './images';
import type { PdfPage } from './document';
import type { PdfVector } from './vector';
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
      last.italic === s.italic
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
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
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
          },
          ...(r.href ? { href: r.href } : {}),
        })),
    },
  };
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
export function shapeBlock(v: PdfVector, frame?: PageFrame, zOrder?: number): BodyElement {
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
  // Not `behind`: a path is not always under the pictures. 22060_A1_01_Plans.pdf
  // backs its legend with a white box painted OVER a floor plan, and forced
  // behind it, the plan and the title block read straight through the legend.
  // `zOrder` carries the order the page painted in, which is the only thing
  // that decides this.
  const float: FloatAnchor | undefined =
    frame !== undefined
      ? {
          wrap: 'none',
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
): FlowDoc {
  return {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    ...(section ? { section } : {}),
    ...(embeddedFonts && embeddedFonts.size > 0 ? { embeddedFonts } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
}
