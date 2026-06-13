// E-PDF EP4 — heuristic reconstruction for UNTAGGED PDFs. With no structure tree
// there is only positioned content (EP2/EP6/EP8): glyphs with an (x, y), a size
// and any hyperlink, plus images with a page rectangle. We recover reading order
// the way a human eye does — cluster runs sharing a baseline into lines, order a
// line's runs left-to-right inserting spaces across gaps, group lines into
// paragraphs by their vertical spacing, then interleave the page's images by
// their top edge. Each run's href is carried through as a span so links survive.
// Headings are guessed from a font size well above the document's median.
// Untagged recovery is inherently approximate (quality is a metric, not a
// guarantee).

import {
  buildFlowDoc,
  dedupeLosses,
  imageBlock,
  paragraphFromRuns,
  shapeBlock,
} from './flow-build';
import { collectPageImages } from './images';
import { extractPageText } from './text';
import { collectPageVectors } from './vector';
import type { BodyElement } from '@/core/document-model';
import type { Loss } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import { ResourceStore } from '@/core/ir';

interface Line {
  readonly y: number; // baseline (page space, y-up)
  readonly fontSize: number;
  readonly text: string; // joined text, for emptiness/heading checks
  readonly spans: ReadonlyArray<TextSpan>;
}

export function reconstructByLayout(file: PdfFile): Reconstruction {
  const pages = file.pages();
  const pageLines = pages.map((page) =>
    groupIntoLines(extractPageText(file, page)).filter((l) => l.text.length > 0),
  );
  const medianFont = median(pageLines.flat().map((l) => l.fontSize)) || 12;

  const resources = new ResourceStore();
  const losses: Array<Loss> = [];
  const body: Array<BodyElement> = [];
  pages.forEach((page, i) => {
    // Position-keyed blocks for this page: paragraphs at their top line, images
    // at their top edge. One descending-y sort puts them in reading order.
    const blocks: Array<{ top: number; el: BodyElement }> = [];
    for (const para of groupIntoParagraphs(pageLines[i]!)) {
      blocks.push({
        top: para.top,
        el: paragraphFromRuns(para.spans, headingLevel(para.fontSize, medianFont)),
      });
    }
    const imgs = collectPageImages(file, page);
    losses.push(...imgs.losses);
    for (const img of imgs.images) {
      blocks.push({ top: img.y + img.heightPt, el: imageBlock(img, resources) });
    }
    // Filled vector paths (EP10) interleave by their top edge, like images.
    for (const v of collectPageVectors(file, page)) {
      blocks.push({ top: v.maxY, el: shapeBlock(v) });
    }
    blocks.sort((a, b) => b.top - a.top);
    for (const block of blocks) body.push(block.el);
  });
  return { doc: buildFlowDoc(body, resources), losses: dedupeLosses(losses) };
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
    return {
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
    spans.push(run.href !== undefined ? { text: run.text, href: run.href } : { text: run.text });
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
