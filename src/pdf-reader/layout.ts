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
} from './flow-build';
import { collectPageImages } from './images';
import { extractPageText } from './text';
import { collectPageVectors } from './vector';
import type { BodyElement } from '@/core/document-model';
import type { Loss } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile, PdfPage } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import { ResourceStore } from '@/core/ir';

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
  const pageRuns = pages.map((page) => extractPageText(file, page));
  const medianFont =
    median(
      pageRuns
        .flat()
        .map((r) => r.fontSizePt)
        .filter((s) => s > 0),
    ) || 12;

  const resources = new ResourceStore();
  const losses: Array<Loss> = [];
  const body: Array<BodyElement> = [];
  pages.forEach((page, i) => {
    const runs = pageRuns[i]!;
    const [px0, , px1] = page.mediaBox;
    // EP17 — detect a clean two-column split (a central vertical gutter no run
    // crosses); fall back to a single column. Each column is grouped and read
    // independently, then the left column's blocks precede the right's.
    const pageWidth = Math.abs(px1 - px0);
    const gutter = detectGutter(runs, pageWidth);
    // Blocks carry a column key so the final sort reads column-by-column: left
    // column top-to-bottom, then right column.
    const blocks: Array<{ col: number; top: number; el: BodyElement }> = [];
    const addColumn = (colRuns: ReadonlyArray<TextRun>, col: number): void => {
      const lines = groupIntoLines(colRuns).filter((l) => l.text.length > 0);
      if (mode === 'positional') {
        // Every line stands where the page set it. Lines are NOT grouped into
        // paragraphs here: a paragraph is a thing that reflows, and nothing in
        // a placed page does.
        for (const line of lines) {
          placed.push({
            key: [Number.MAX_SAFE_INTEGER, placed.length],
            col,
            top: line.y,
            make: (z: number): BodyElement =>
              positionedText(
                line.spans,
                {
                  x: line.x,
                  // §9.4.4 — a baseline is not a box: the line reaches about a
                  // fifth of its size below and the rest above.
                  y: line.y - line.fontSize * 0.25,
                  // To the page's edge, not to the estimated end of the words.
                  // The box paints nothing, so its width costs nothing — but a
                  // width that falls short makes the line WRAP, and a wrapped
                  // line in a placed page walks down over its neighbours.
                  width: Math.max(line.width, pageWidth - line.x),
                  height: line.fontSize * 1.25,
                },
                pageHeightOf(page),
                z,
              ),
          });
        }
        return;
      }
      for (const para of groupIntoParagraphs(lines)) {
        blocks.push({
          col,
          top: para.top,
          el: paragraphFromRuns(para.spans, headingLevel(para.fontSize, medianFont)),
        });
      }
    };
    const placed: Array<{
      key: ReadonlyArray<number>;
      col: number;
      top: number;
      make: (z: number) => BodyElement;
    }> = [];
    if (gutter !== undefined) {
      addColumn(
        runs.filter((r) => r.x < gutter),
        0,
      );
      addColumn(
        runs.filter((r) => r.x >= gutter),
        1,
      );
    } else {
      addColumn(runs, 0);
    }
    const colOf = (centerX: number): number => (gutter !== undefined && centerX >= gutter ? 1 : 0);
    const pageHeight = pageHeightOf(page);
    const imgs = collectPageImages(file, page);
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
    const vectors = collectPageVectors(file, page, covered);

    // §20.4.2.3 `relativeHeight` — pictures and paths share one z-order, and
    // it is the page's own painting order (§8.5.3), not one kind before the
    // other. 22060_A1_01_Plans.pdf backs a legend with a white box painted over
    // a floor plan AND draws a key icon over a red swatch: pictures under paths
    // loses the key, paths under pictures loses the legend.
    const marks = [
      ...imgs.images.map((img) => ({
        key: img.orderKey,
        col: colOf(img.x + img.widthPt / 2),
        top: img.y + img.heightPt,
        make: (z: number): BodyElement => imageBlock(img, resources, undefined, pageHeight, z),
      })),
      ...vectors.map((v) => ({
        key: v.orderKey,
        col: colOf((v.minX + v.maxX) / 2),
        top: v.maxY,
        make: (z: number): BodyElement => shapeBlock(v, pageHeight, z),
      })),
      ...placed,
    ].sort((a, b) => compareOrder(a.key, b.key));
    marks.forEach((mark, z) => {
      blocks.push({ col: mark.col, top: mark.top, el: mark.make(z) });
    });
    blocks.sort((a, b) => a.col - b.col || b.top - a.top);
    for (const block of blocks) body.push(block.el);
  });
  return {
    doc: buildFlowDoc(body, resources, sectionFromPdfPages(pages)),
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
  const right = (r: TextRun): number =>
    r.x + Math.max(1, r.text.length) * (r.fontSizePt || fontSize) * 0.5;
  const intervals = runs.map((r): [number, number] => [r.x, right(r)]).sort((a, b) => a[0] - b[0]);
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

// Cluster runs that share a baseline (within half a line's height) into lines,
// top of the page first; within a line, order by x and build link-aware spans.
function groupIntoLines(runs: ReadonlyArray<TextRun>): Array<Line> {
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
  return clusters.map((c) => {
    const ordered = c.runs.sort((a, b) => a.x - b.x);
    const fontSize = c.fontSize || 10;
    const spans = lineSpans(ordered, fontSize);
    const first = ordered[0]!;
    const last = ordered[ordered.length - 1]!;
    const x = first.x;
    // Half an em per glyph, the same estimate `lineSpans` measures gaps by.
    const width = last.x + last.text.length * (last.fontSizePt || fontSize) * 0.5 - x;
    return {
      x,
      width,
      y: c.y,
      fontSize,
      text: spans
        .map((s) => s.text)
        .join('')
        .replace(/\s+/g, ' ')
        .trim(),
      spans,
    };
  });
}

// A line's runs as spans, inserting a (link-free) space where a horizontal gap
// suggests one. Run widths are estimated (half-em per char) — glyph metrics are
// not kept.
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
      ...(run.href !== undefined ? { href: run.href } : {}),
    });
    prevEnd = run.x + run.text.length * (run.fontSizePt || fontSize) * 0.5;
  }
  return spans;
}

// Group consecutive lines into paragraphs: a vertical gap well over a single
// line's leading starts a new paragraph. `top` is the paragraph's first (highest) line.
function groupIntoParagraphs(
  lines: ReadonlyArray<Line>,
): Array<{ spans: Array<TextSpan>; fontSize: number; top: number }> {
  const groups: Array<Array<Line>> = [];
  let prevY: number | undefined;
  for (const line of lines) {
    const gap = prevY !== undefined ? prevY - line.y : 0;
    if (groups.length === 0 || (prevY !== undefined && gap > line.fontSize * 1.5)) groups.push([]);
    groups[groups.length - 1]!.push(line);
    prevY = line.y;
  }
  return groups.map((g) => ({
    spans: g.flatMap((l, i) => (i > 0 ? [{ text: ' ' }, ...l.spans] : [...l.spans])),
    fontSize: Math.max(...g.map((l) => l.fontSize)),
    top: g[0]!.y,
  }));
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

/** The page's height in points — the flip between PDF's y-up frame and ours. */
function pageHeightOf(page: PdfPage): number {
  return Math.abs(page.mediaBox[3] - page.mediaBox[1]);
}
