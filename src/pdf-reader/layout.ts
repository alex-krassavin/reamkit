// E-PDF EP4 — heuristic reconstruction for UNTAGGED PDFs. With no structure tree
// there is only positioned content (EP2/EP6/EP8): glyphs with an (x, y), a size
// and any hyperlink, plus images with a page rectangle. We recover reading order
// the way a human eye does — split a clean two-column page at its central gutter
// (EP17), then within each column cluster runs sharing a baseline into lines,
// order a line's runs left-to-right inserting spaces across gaps, group lines
// into paragraphs by their vertical spacing, and interleave the column's images
// by their top edge. Each run's href is carried through as a span so links
// survive. Headings are guessed from a font size well above the document's
// median. Untagged recovery is inherently approximate (quality is a metric, not
// a guarantee).

import {
  buildFlowDoc,
  dedupeLosses,
  imageBlock,
  paragraphFromRuns,
  positionedText,
  sectionFromPdfPages,
  shapeBlock,
  withMeasuredMargins,
} from './flow-build';
import { displayOf, placeImages, placeRuns, placeVectors } from './display';
import { collectEmbeddedFonts } from './embedded-fonts';
import { collectPageImages } from './images';
import { extractPageText } from './text';
import { collectPageVectors } from './vector';
import { markDrawnRules } from './text-rules';
import { isRightToLeft } from './content';
import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { Loss, Pt } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile, PdfPage } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import { FEATURES, ResourceStore, pt } from '@/core/ir';

interface Line {
  readonly y: number; // baseline (page space, y-up)
  readonly fontSize: number;
  readonly text: string; // joined text, for emptiness/heading checks
  readonly spans: ReadonlyArray<TextSpan>;
  /** Leftmost glyph origin, and how far the line reaches — placed reconstruction. */
  readonly x: number;
  readonly width: number;
}

/**
 * Heuristically reconstruct an untagged PDF into a {@link Reconstruction}
 * (E-PDF EP4). With no structure tree there is only positioned content, so
 * reading order is recovered the way a human eye does: split a clean two-column
 * page at its central gutter (EP17), then within each column cluster runs
 * sharing a baseline into lines, order each line left-to-right inserting spaces
 * across gaps, group lines into paragraphs by their vertical spacing, and
 * interleave the column's images and filled vector paths (EP10) by their top
 * edge. Each run's `href` is carried through as a span so links survive, and a
 * font size well above the document median is guessed as a heading. The result
 * is inherently approximate.
 *
 * @param file The PDF to reconstruct.
 * @returns The reconstructed {@link FlowDoc} plus any read-time losses.
 */
export function reconstructByLayout(
  file: PdfFile,
  mode: 'flow' | 'positional' = 'flow',
): Reconstruction {
  const pages = file.pages();
  // §14.11.1 — every mark is lifted into the page's SHOWN frame, so nothing
  // downstream has to know the page was ever turned.
  const shown = pages.map((page) => displayOf(page));
  const pageRuns = pages.map((page, i) => placeRuns(extractPageText(file, page), shown[i]!));
  const medianFont =
    median(
      pageRuns
        .flat()
        .map((r) => r.fontSizePt)
        .filter((s) => s > 0),
    ) || 12;

  const resources = new ResourceStore();
  const losses: Array<Loss> = [];
  // §8.6.6.2 — type filled with a tiling pattern keeps the pattern's colour at
  // the pattern's own density and loses its shape: a run carries one colour, not
  // a content stream, so a hatch that alternates ink and paper becomes the flat
  // tint the two average to.
  if (pageRuns.some((page) => page.some((r) => r.fillPatternName !== undefined))) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'text filled with a tiling pattern is drawn as a flat tint of the pattern’s colour, not as the pattern',
    });
  }
  const body: Array<BodyElement> = [];
  pages.forEach((page, i) => {
    const runs = pageRuns[i]!;
    const display = shown[i]!;
    // EP17 — detect a clean two-column split (a central vertical gutter no run
    // crosses); fall back to a single column. Each column is grouped and read
    // independently, then the left column's blocks precede the right's.
    const pageWidth = display.width;
    const gutter = detectGutter(runs, pageWidth);
    // Blocks carry a column key so the final sort reads column-by-column: left
    // column top-to-bottom, then right column.
    const blocks: Array<{ col: number; top: number; el: BodyElement }> = [];
    const addColumn = (allRuns: ReadonlyArray<TextRun>, col: number): void => {
      // §9.6.5 — a Type 3 run's marks are its glyph PROCEDURES, which the path
      // and picture passes lift. Re-setting its codes in a substitute face
      // would draw a second, smaller copy of a drawing.
      //
      // §9.3.6 — and a run the page painted nowhere is not drawn either. Both
      // are kept for the FLOWING reading, which is a document being read rather
      // than a page being reproduced: a scanned page's every word lives in its
      // invisible layer.
      const colRuns =
        mode === 'positional'
          ? allRuns.filter((r) => r.type3 !== true && r.invisible !== true)
          : allRuns;
      if (mode === 'positional') {
        // Every line stands where the page set it. Lines are NOT grouped into
        // paragraphs here: a paragraph is a thing that reflows, and nothing in
        // a placed page does.
        //
        // Runs sharing a baseline make a line, and a TURNED baseline is not the
        // upright one however close their y's fall. Each angle is grouped in
        // its own frame and comes back carrying it: 160F-2019.pdf sets "Nature"
        // on its side down the middle of a column, and read flat it joined the
        // row it happened to cross.
        for (const [angle, runs] of byAngle(colRuns)) {
          for (const line of groupIntoLines(rotate(runs, -angle), true)) {
            if (line.text.length === 0) continue;
            const box = turnedBox(line, angle, pageWidth);
            placed.push({
              key: [Number.MAX_SAFE_INTEGER, placed.length],
              col,
              // The box's own top on the PAGE — `line.y` is measured in the
              // turned frame, and the blocks are ordered by where they stand.
              top: box.y + box.height,
              make: (z: number): BodyElement =>
                positionedText(line.spans, box, frame, z, rotation60kOf(angle)),
            });
          }
        }
        return;
      }
      const lines = groupIntoLines(colRuns).filter((l) => l.text.length > 0);
      // The column the paragraphs were set in — its own edges, not the page's,
      // so a two-column page judges each side against the side it belongs to.
      const measure =
        lines.length > 0
          ? {
              left: Math.min(...lines.map((l) => l.x)),
              right: Math.max(...lines.map((l) => l.x + l.width)),
            }
          : undefined;
      for (const para of groupIntoParagraphs(lines, measure, display.height)) {
        blocks.push({
          col,
          top: para.top,
          el: paragraphFromRuns(para.spans, headingLevel(para.fontSize, medianFont), {
            ...(para.alignment !== undefined ? { alignment: para.alignment } : {}),
            ...(para.spacingBefore !== undefined ? { spacingBefore: pt(para.spacingBefore) } : {}),
          }),
        });
      }
    };
    const placed: Array<{
      key: ReadonlyArray<number>;
      col: number;
      top: number;
      make: (z: number) => BodyElement;
    }> = [];
    const colOf = (centerX: number): number => (gutter !== undefined && centerX >= gutter ? 1 : 0);
    // The shown page has its own corner: the turn has already been applied, so
    // what is left is a box that starts at the origin.
    const frame = { left: 0, top: display.height };
    const raw = collectPageImages(file, page);
    const imgs = { images: placeImages(raw.images, display), losses: raw.losses };
    losses.push(...imgs.losses);
    // Filled vector paths (EP10) are ANCHORED where the page drew them — they
    // are artwork, not paragraphs, and a sheet of them has no reading order to
    // take a place in. They still sort by top edge, so their z-order is the
    // order the page painted them in.
    // A white box drawn over a picture is not invisible paint but the thing
    // that hides it, so the paths are filtered against what is already placed.
    const covered = imgs.images.map((img) => ({
      minX: img.x,
      minY: img.y,
      maxX: img.x + img.widthPt,
      maxY: img.y + img.heightPt,
    }));
    const lifted = collectPageVectors(file, page, covered);
    losses.push(...lifted.losses);
    // A PDF has no underline: it draws a thin bar under the words. Read onto
    // the runs BEFORE they are grouped, so the mark travels with them and the
    // bar is not placed a second time where the words no longer are.
    const placedVectors = placeVectors(lifted.vectors, display);
    const ruled = markDrawnRules(runs, placedVectors);
    const vectors = placedVectors.filter((v) => !ruled.consumed.has(v));
    if (gutter !== undefined) {
      addColumn(
        ruled.runs.filter((r) => r.x < gutter),
        0,
      );
      addColumn(
        ruled.runs.filter((r) => r.x >= gutter),
        1,
      );
    } else {
      addColumn(ruled.runs, 0);
    }

    // §20.4.2.3 `relativeHeight` — pictures and paths share one z-order, and
    // it is the page's own painting order (§8.5.3), not one kind before the
    // other. 22060_A1_01_Plans.pdf backs a legend with a white box painted over
    // a floor plan AND draws a key icon over a red swatch: pictures under paths
    // loses the key, paths under pictures loses the legend.
    // In a FLOWING reading the words are re-set and the artwork is not, so no
    // mark may cover them: the page's own painting order still ranks the marks
    // against each other, but all of them sit under the text.
    const under = mode !== 'positional';
    const marks = [
      ...imgs.images.map((img) => ({
        key: img.orderKey,
        col: colOf(img.x + img.widthPt / 2),
        top: img.y + img.heightPt,
        make: (z: number): BodyElement => imageBlock(img, resources, undefined, frame, z, under),
      })),
      ...vectors.map((v) => ({
        key: v.orderKey,
        col: colOf((v.minX + v.maxX) / 2),
        top: v.maxY,
        make: (z: number): BodyElement => shapeBlock(v, frame, z, under),
      })),
      ...placed,
    ].sort((a, b) => compareOrder(a.key, b.key));
    marks.forEach((mark, z) => {
      blocks.push({ col: mark.col, top: mark.top, el: mark.make(z) });
    });
    blocks.sort((a, b) => a.col - b.col || b.top - a.top);
    // Each source page after the first opens an output page of its own. Flowed,
    // the layout repaginates and this hardly shows; PLACED, every mark is
    // anchored to "the page", so without it all twenty-five pages of
    // Brotli-Prototype-FileA.pdf stack onto one.
    if (i > 0 && blocks.length > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: {
          properties: { pageBreakBefore: true, spacingLine: pt(0), spacingLineRule: 'exact' },
          runs: [],
        },
      });
    }
    for (const block of blocks) body.push(block.el);
  });
  const section = sectionFromPdfPages(pages);
  return {
    doc: buildFlowDoc(
      body,
      resources,
      // A placed reading anchors everything to the page, so its margins must
      // stay at zero or the anchors move. A FLOWING one is a document being
      // re-set, and a document with no margins prints its words against the
      // edge of the paper — which is what every converted PDF looked like.
      mode === 'positional' ? section : withMeasuredMargins(section, shown, pageRuns),
      collectEmbeddedFonts(file, pages),
    ),
    losses: dedupeLosses(losses),
  };
}

// EP17 — a two-column gutter: the centre of the widest vertical whitespace band
// that no run's horizontal extent crosses. Conservative on purpose — it fires
// only on a genuine two-column page (a full-width line spans the centre, leaving
// no gap there, so title/single-column pages keep their existing reading order).
function detectGutter(runs: ReadonlyArray<TextRun>, pageWidth: number): number | undefined {
  if (runs.length < 30 || pageWidth <= 0) return undefined;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const intervals = runs
    .map((r): [number, number] => [r.x, Math.max(r.endX, r.x + 1)])
    .sort((a, b) => a[0] - b[0]);
  const minX = intervals[0]![0];
  const maxX = Math.max(...intervals.map((iv) => iv[1]));
  const span = maxX - minX;
  if (span < pageWidth * 0.5) return undefined; // text doesn't span enough of the page
  let curEnd = intervals[0]![1];
  let gapMid = 0;
  let gapW = 0;
  for (const [l, r] of intervals) {
    if (l - curEnd > gapW) {
      gapW = l - curEnd;
      gapMid = (curEnd + l) / 2;
    }
    if (r > curEnd) curEnd = r;
  }
  const frac = (gapMid - minX) / span;
  if (gapW < fontSize * 3 || frac < 0.35 || frac > 0.65) return undefined; // not a central gutter
  const left = runs.filter((r) => r.x < gapMid).length;
  if (left < runs.length * 0.25 || left > runs.length * 0.75) return undefined; // unbalanced
  return gapMid;
}

/**
 * The gap, in ems, past which the placed reader cuts a line rather than write a
 * space across it.
 *
 * It began at four ems, from a measurement on 160F-2019.pdf: the gaps between
 * words there run 0.00–3.13 em and the gaps between COLUMNS 4.41–43.66 em, with
 * nothing in between. That told a column from a word, which was the question at
 * the time. It is not the question here.
 *
 * A placed piece is set down at the x it was measured at, so cutting costs
 * nothing — and a space costs whatever the page's gap was not. A quarter of an
 * em of type standing in for one em of pen leaves everything after it three
 * quarters of an em short, and the error runs on down the line. So the reader
 * cuts wherever {@link lineSpans} would otherwise write a space, and a placed
 * page never stands a space in for a gap it measured. Only the placed reader
 * does this — a flowing paragraph is meant to be read across.
 */
const SPACE_GAP_EM = 0.25;

/** Runs by the direction of their baseline, upright first, each angle rounded. */
function byAngle(runs: ReadonlyArray<TextRun>): Array<[number, Array<TextRun>]> {
  const groups = new Map<number, Array<TextRun>>();
  for (const run of runs) {
    // Rounded to the degree: a page that sets a label on its side sets every
    // glyph of it at the same angle, give or take the arithmetic.
    const angle = Math.round(run.angleDeg ?? 0);
    const group = groups.get(angle);
    if (group) group.push(run);
    else groups.set(angle, [run]);
  }
  return [...groups].sort((a, b) => Math.abs(a[0]) - Math.abs(b[0]));
}

/** The same runs seen from a frame turned by `deg`, where their baseline is flat. */
function rotate(runs: ReadonlyArray<TextRun>, deg: number): Array<TextRun> {
  if (deg === 0) return [...runs];
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return runs.map((r) => ({
    ...r,
    x: r.x * cos - r.y * sin,
    y: r.x * sin + r.y * cos,
    endX: r.endX * cos - r.endY * sin,
    endY: r.endX * sin + r.endY * cos,
  }));
}

/**
 * A turned line's page-space box: the rectangle the renderer is to place and
 * then spin about its own centre, which is how a shape's rotation works
 * (§20.1.7.6 `a:xfrm rot`).
 *
 * The box is measured in the line's own frame and only its CENTRE is carried
 * back into page space — an axis-aligned box of the same size, centred there
 * and turned, puts the words back along the baseline they were set on.
 */
function turnedBox(
  line: Line,
  angleDeg: number,
  pageWidth: number,
): { x: number; y: number; width: number; height: number } {
  const height = line.fontSize * 1.25;
  // §9.4.4 — a baseline is not a box: the line reaches about a fifth of its
  // size below and the rest above.
  const bottom = line.y - line.fontSize * 0.25;
  // A width that falls short makes the line WRAP, and a wrapped line in a
  // placed page walks down over its neighbours. Upright, there is a page edge
  // to reach for; turned, the box's own size decides where the spin puts it, so
  // the slack is a fifth of the words plus an em rather than the whole page.
  // A right-to-left line is set from the box's RIGHT edge, so slack on the
  // right pushes it off the words it was measured from. It gets its own width
  // and no more: ArabicCIDTrueType.pdf's every line stood a hundred and fifty
  // points right of where the page draws it.
  const slack = line.width * 1.2 + line.fontSize;
  const width = isRightToLeft(line.text)
    ? line.width
    : angleDeg === 0
      ? Math.max(line.width, pageWidth - line.x)
      : slack;
  if (angleDeg === 0) return { x: line.x, y: bottom, width, height };
  const rad = (angleDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const cu = line.x + width / 2;
  const cv = bottom + height / 2;
  return {
    x: cu * cos - cv * sin - width / 2,
    y: cu * sin + cv * cos - height / 2,
    width,
    height,
  };
}

/** §20.1.7.6 — a shape turns CLOCKWISE in 1/60000°, and PDF measures the other way. */
function rotation60kOf(angleDeg: number): number | undefined {
  if (angleDeg === 0) return undefined;
  return Math.round((((-angleDeg % 360) + 360) % 360) * 60000);
}

/**
 * The step, in ems, past which a run does not share its neighbour's baseline
 * but stands above or below it — a superscript, or the next row of a form.
 *
 * Measured across the readable pdfjs corpus: 123 runs sit exactly on their
 * line's baseline, three within 0.02 em of it (rounding), one between 0.02 and
 * 0.05, and thirty at 0.05 em or more. A twentieth of an em sits in that gap.
 */
const BASELINE_STEP_EM = 0.05;

// Cluster runs that share a baseline (within half a line's height) into lines,
// top of the page first; within a line, order by x and build link-aware spans.
// With `split`, a cluster is cut wherever a column-wide gap opens or a baseline
// steps, so each piece keeps its own x and its own y instead of being dragged
// against its neighbour.
function groupIntoLines(runs: ReadonlyArray<TextRun>, split = false): Array<Line> {
  const sorted = [...runs].sort((a, b) => b.y - a.y || a.x - b.x);
  const clusters: Array<{ y: number; fontSize: number; runs: Array<TextRun> }> = [];
  for (const run of sorted) {
    const last = clusters[clusters.length - 1];
    const tol = Math.max(1, (run.fontSizePt || 10) * 0.5);
    if (last && Math.abs(last.y - run.y) <= tol) {
      last.runs.push(run);
      last.fontSize = Math.max(last.fontSize, run.fontSizePt || 0);
    } else {
      clusters.push({ y: run.y, fontSize: run.fontSizePt || 10, runs: [run] });
    }
  }
  return clusters.flatMap((c) => {
    const ordered = c.runs.sort((a, b) => a.x - b.x);
    const fontSize = c.fontSize || 10;
    if (!split) return [lineOf(ordered, c.y, fontSize)];
    const pieces: Array<Array<TextRun>> = [[]];
    for (const run of ordered) {
      const prev = pieces[pieces.length - 1]!;
      const last = prev[prev.length - 1];
      const size = run.fontSizePt || fontSize;
      // A gap the page put there is a placement, and so is a step off the
      // baseline: 160F-2019.pdf sets its footnote marks a size smaller and
      // three quarters of an em up, and read as one line they came down flat.
      //
      // The gap that splits is the one a SPACE would have to stand in for. A
      // placed piece is set down at its measured x, so cutting costs nothing
      // and a space costs whatever the page's gap was not: a quarter of an em
      // of type standing in for one em of pen leaves everything after it three
      // quarters of an em short, and the error runs on down the line.
      // A right-to-left piece is set from its RIGHT edge, so a piece that holds
      // more than one word carries the whole line's width error into where the
      // last of them lands. Each run gets its own box and its own right edge.
      // An OVERLAP is a placement by the same argument, and the argument runs
      // the same way in both directions: a run that starts a quarter of an em
      // before the one before it ended was set down ON it, not after it, and
      // flowing the two end to end moves the second one out by the whole
      // overlap. ContentStream*Type3.pdf stamps its inner word three times at
      // half its own width — 36pt of step under 64.8pt of type — and read as one
      // line the three came out side by side, half again as wide as the page
      // sets them.
      const steps =
        last !== undefined &&
        (Math.abs(run.x - last.endX) > size * SPACE_GAP_EM ||
          Math.abs(run.y - prev[0]!.y) > size * BASELINE_STEP_EM ||
          isRightToLeft(run.text) ||
          isRightToLeft(last.text));
      if (steps) pieces.push([]);
      pieces[pieces.length - 1]!.push(run);
    }
    // Each piece stands on its own baseline, at its own size — a mark lifted
    // out of a line of eleven-point text is not eleven points tall.
    return pieces.map((piece) =>
      lineOf(piece, piece[0]!.y, Math.max(...piece.map((r) => r.fontSizePt || 0)) || fontSize),
    );
  });
}

/** One run of runs, left to right on a shared baseline, as a {@link Line}. */
function lineOf(runs: ReadonlyArray<TextRun>, y: number, fontSize: number): Line {
  const ordered = lineSpans(runs, fontSize);
  // §9.4 — the runs came off the page in the order they were PAINTED, which is
  // left to right whatever the script. `logicalOrder` turned each run's own
  // letters back the right way round; the runs themselves are still in visual
  // order, and a line of them reads as the sentence backwards.
  // ArabicCIDTrueType.pdf's every line came out with its words in reverse.
  const spans = ordered.every((s) => s.text.trim() === '' || isRightToLeft(s.text))
    ? [...ordered].reverse()
    : ordered;
  const x = runs[0]!.x;
  return {
    x,
    width: runs[runs.length - 1]!.endX - x,
    y,
    fontSize,
    text: spans
      .map((s) => s.text)
      .join('')
      .replace(/\s+/g, ' ')
      .trim(),
    spans,
  };
}

// A line's runs as spans, inserting a (link-free) space where a horizontal gap
// suggests one.
function lineSpans(runs: ReadonlyArray<TextRun>, fontSize: number): Array<TextSpan> {
  const spans: Array<TextSpan> = [];
  let prevEnd: number | undefined;
  for (const run of runs) {
    if (prevEnd !== undefined && run.x - prevEnd > fontSize * 0.25) spans.push({ text: ' ' });
    // §9.3.1/§8.6.8 — the size and colour the page showed the glyphs at. The
    // tagged path has carried these since it learned to; this one never did, so
    // every line it read came back at the 11pt default in black. Placed, that
    // is not a wrong shade but a wrong SHAPE: 160F-2019.pdf's footnotes are set
    // in 7pt nine and a half apart, and drawn at eleven they climbed over each
    // other.
    spans.push({
      text: run.text,
      sizePt: run.fontSizePt,
      ...(run.colorHex !== '000000' ? { colorHex: run.colorHex } : {}),
      ...(run.fontName !== undefined ? { fontName: run.fontName } : {}),
      ...(run.outlineHex !== undefined
        ? { outline: { colorHex: run.outlineHex, widthPt: pt(run.outlineWidthPt ?? 1) } }
        : {}),
      ...(run.bold ? { bold: true } : {}),
      ...(run.italic ? { italic: true } : {}),
      ...(run.markup !== undefined ? { markup: run.markup } : {}),
      ...(run.href !== undefined ? { href: run.href } : {}),
    });
    prevEnd = run.endX;
  }
  return spans;
}

// Group consecutive lines into paragraphs: a vertical gap well over a single
// line's leading starts a new paragraph. `top` is the paragraph's first (highest) line.
function groupIntoParagraphs(
  lines: ReadonlyArray<Line>,
  column?: { left: number; right: number },
  pageHeight = 0,
): Array<{
  spans: Array<TextSpan>;
  fontSize: number;
  top: number;
  alignment?: 'center' | 'right';
  spacingBefore?: number;
}> {
  const groups: Array<Array<Line>> = [];
  const gaps: Array<number> = [];
  let prev: Line | undefined;
  for (const line of lines) {
    const gap = prev !== undefined ? prev.y - line.y : 0;
    const opened = prev !== undefined && gap > line.fontSize * 1.5;
    if (
      groups.length === 0 ||
      opened ||
      (prev !== undefined && endedParagraph(prev, line, column))
    ) {
      groups.push([]);
      gaps.push(prev === undefined ? 0 : gap);
    }
    groups[groups.length - 1]!.push(line);
    prev = line;
  }
  return groups.map((g, i) => {
    const first = g[0]!;
    const fontSize = Math.max(...g.map((l) => l.fontSize));
    // The gap that OPENED this paragraph, less the line it would have taken
    // anyway, is the space its author put before it. Under a third of a line it
    // is just leading, and a paragraph is not spaced by rounding error.
    //
    // Bounded by a third of the sheet, not by three lines: the gap is MEASURED,
    // and three lines is a guess overriding a measurement.
    // annotation-square-circle-without-appearance.pdf sets its two labels two
    // hundred points apart, each over its own pair of drawings, and capped at
    // thirty the second label came back inside the first drawing.
    const opened = (gaps[i] ?? 0) - fontSize * 1.2;
    const most = pageHeight > 0 ? pageHeight / 3 : fontSize * 3;
    const spacingBefore = opened > fontSize * 0.3 ? Math.min(opened, most) : undefined;
    return {
      spans: g.flatMap((l, k) => (k > 0 ? [{ text: ' ' }, ...l.spans] : [...l.spans])),
      fontSize,
      top: first.y,
      ...(spacingBefore !== undefined ? { spacingBefore } : {}),
      ...alignmentOf(g, column),
    };
  });
}

/**
 * Whether a line ENDED a paragraph, rather than wrapping into the next.
 *
 * Leading alone cannot tell the two apart: five labels stacked at 15pt with a
 * 12pt face look exactly like five wrapped lines, and alphatrans.pdf's five are
 * read as one paragraph and re-wrapped into two. But a wrapping engine pulls
 * the next word UP — so a line that stops well short of the measure stopped
 * because its author stopped it, and the line after it begins something new.
 * The same rule separates two paragraphs set with no extra space between them,
 * which used to run together for the same reason.
 *
 * Only where both lines start at the same edge. Where they do not, the block is
 * placed rather than set — a centred title's every line is short of the measure
 * and none of them ends anything.
 *
 * @param prev   The line before: where it starts, how wide it is, its face.
 * @param next   The line after — only where it starts matters.
 * @param column The measure both were set in, when it is known.
 * @returns Whether the first line ended a paragraph.
 */
export function endedParagraph(
  prev: { x: number; width: number; fontSize: number },
  next: { x: number },
  column: { left: number; right: number } | undefined,
): boolean {
  if (!column) return false;
  const width = column.right - column.left;
  if (!(width > 0)) return false;
  // A first-line indent is a quarter-inch or so; beyond that the two lines are
  // not part of one setting.
  if (Math.abs(prev.x - next.x) > Math.max(prev.fontSize, 4)) return false;
  // A quarter of the measure: less than that is the ragged edge every
  // unjustified paragraph has, and breaking on it would cut prose into lines.
  return column.right - (prev.x + prev.width) > width * 0.25;
}

/**
 * §17.3.1.13 — where a paragraph sits across its column, which is the only
 * witness a PDF leaves of how it was set: every line is placed absolutely and
 * nothing says "centred".
 *
 * A paragraph whose lines are inset by about as much on each side is centred; a
 * one-line paragraph pushed to the right edge is right-aligned. Everything else
 * is left alone — a justified paragraph and a ragged-right one look the same
 * from here, and guessing between them would re-set the body of every document.
 */
function alignmentOf(
  lines: ReadonlyArray<Line>,
  column: { left: number; right: number } | undefined,
): { alignment?: 'center' | 'right' } {
  if (!column || lines.length === 0) return {};
  const width = column.right - column.left;
  if (!(width > 0)) return {};
  const insets = lines.map((l) => ({
    lead: l.x - column.left,
    trail: column.right - (l.x + l.width),
  }));
  // A tenth of the measure is the smallest inset worth calling a placement:
  // below it every ragged line would read as placed.
  const meaningful = width * 0.1;
  const even = width * 0.06;
  // Judged line by line and only then as a whole. Taking the smallest inset
  // over the whole paragraph makes a CENTRED block read as a full one the
  // moment any of its lines nearly fills the measure — and a two-line title
  // whose first line runs the width is exactly that.
  if (
    insets.every((i) => Math.abs(i.lead - i.trail) <= even) &&
    Math.max(...insets.map((i) => Math.min(i.lead, i.trail))) >= meaningful
  ) {
    return { alignment: 'center' };
  }
  // Flush right: every line ends at the measure and at least one starts well
  // inside it. A justified paragraph fails this on its last line, which is the
  // only place the two differ at all.
  if (
    insets.every((i) => i.trail <= even) &&
    Math.max(...insets.map((i) => i.lead)) >= meaningful
  ) {
    return { alignment: 'right' };
  }
  return {};
}

// A line markedly larger than the body text reads as a heading.
function headingLevel(fontSize: number, medianFont: number): number | undefined {
  if (fontSize >= medianFont * 1.5) return 0;
  if (fontSize >= medianFont * 1.25) return 1;
  return undefined;
}

function median(values: ReadonlyArray<number>): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
}

/**
 * Compare two painting-order keys (§8.5.3): position by position, and the
 * shorter one first where they agree — a form's call comes before the marks it
 * makes inside it.
 */
function compareOrder(a: ReadonlyArray<number>, b: ReadonlyArray<number>): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    if (a[i] !== b[i]) return a[i]! - b[i]!;
  }
  return a.length - b.length;
}
