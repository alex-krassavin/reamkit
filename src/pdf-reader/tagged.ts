// E-PDF EP3 — tagged fast-path reconstruction. Walks the structure tree (EP3
// struct-tree.ts), pulls each element's text from the per-page MCID → text map
// the content interpreter produced (EP2), and rebuilds a FlowDoc: headings
// (H1–H6 → outline level), paragraphs, tables (Table → TR → TH/TD), list items
// (each LI → its label + body text), and figures (EP6 — each /Figure's MCID
// resolves to a lifted image, carrying its /Alt). The honest inverse of the
// tagged PDF Ream writes.

import {
  buildFlowDoc,
  dedupeLosses,
  imageBlock,
  paragraphBlock,
  paragraphFromRuns,
  sectionFromPdfPages,
  shapeBlock,
  withMeasuredMargins,
} from './flow-build';
import { displayOf, placeRuns, placeVectors } from './display';
import { collectEmbeddedFonts } from './embedded-fonts';
import { collectPageImages } from './images';
import { collectPageVectors } from './vector';
import { readStructTree } from './struct-tree';
import { extractPageText } from './text';
import type { BodyElement, Table, TableCell, TableRow } from '@/core/document-model';
import type { Pt } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import type { PdfImage } from './images';
import type { StructNode } from './struct-tree';
import { ResourceStore, pt } from '@/core/ir';

// Printable width assumed for a synthesized table grid (6.5"). The structure
// tree carries no column widths, so the PROPORTIONS come from where the glyphs
// actually sit on the page (see `gridFromLefts`) and only the total is assumed.
const ASSUMED_CONTENT_WIDTH_PT = 468;

// A column narrower than this holds no word, only a stack of single letters.
const MIN_COLUMN_PT = 6;

/**
 * Reconstruct a {@link Reconstruction} from a tagged PDF's logical structure
 * (E-PDF EP3 — the honest inverse of the tagged PDF Ream writes). Walks the
 * structure tree ({@link readStructTree}), pulls each element's text from the
 * per-page MCID → text map the content interpreter produced (EP2), and rebuilds
 * headings (`H1`–`H6` → outline level), paragraphs, tables (`Table` → `TR` →
 * `TH`/`TD`), list items (each `LI` → label + body) and figures (EP6 — each
 * `/Figure`'s MCID resolves to a lifted image carrying its `/Alt`). Images no
 * `/Figure` claims are appended in page + top-down order so nothing is lost.
 *
 * @returns The reconstructed document and image losses, or `undefined` when the
 *          PDF carries no structure tree or yields no body content.
 */
export function reconstructTaggedPdf(file: PdfFile): Reconstruction | undefined {
  const root = readStructTree(file);
  if (!root) return undefined;

  const pages = file.pages();
  // §14.11.1 — the pages as they are SHOWN, and every run placed on them. The
  // tree says what the words ARE; where they sit is still the page's to say,
  // and it is the only witness of the margins the author set.
  const shown = pages.map((page) => displayOf(page));
  const placedRuns = pages.map((page, i) => placeRuns(extractPageText(file, page), shown[i]!));
  // Per page: MCID → its runs, in show order (runs carry any hyperlink, EP8).
  const pageRuns = placedRuns.map((runs) => {
    const byMcid = new Map<number, Array<TextRun>>();
    for (const run of runs) {
      if (run.mcid === undefined) continue;
      const list = byMcid.get(run.mcid);
      if (list) list.push(run);
      else byMcid.set(run.mcid, [run]);
    }
    return byMcid;
  });
  // Every run the tree actually reached, for the check at the end: a tree that
  // reaches almost none of a page's words is not a description of this
  // document's text.
  const claimed = new Set<TextRun>();
  const runsOfMcid = (page: number, mcid: number): Array<TextRun> => {
    const runs = pageRuns[page]?.get(mcid) ?? [];
    for (const run of runs) claimed.add(run);
    return runs;
  };

  // Per page: the lifted images, indexed by their owning MCID (a /Figure's).
  const resources = new ResourceStore();
  const pageImages = pages.map((page) => collectPageImages(file, page));
  const imageLosses = dedupeLosses(pageImages.flatMap((p) => p.losses));
  const imagesByMcid = pageImages.map((p) => {
    const byMcid = new Map<number, Array<PdfImage>>();
    for (const img of p.images) {
      if (img.mcid === undefined) continue;
      const list = byMcid.get(img.mcid);
      if (list) list.push(img);
      else byMcid.set(img.mcid, [img]);
    }
    return byMcid;
  });
  const emitted = new Set<PdfImage>();
  const imagesForNode = (node: StructNode): Array<PdfImage> =>
    node.mcids.flatMap(({ page, mcid }) => imagesByMcid[page]?.get(mcid) ?? []);

  const textOf = (node: StructNode): string =>
    squash(
      node.mcids
        .map(({ page, mcid }) =>
          runsOfMcid(page, mcid)
            .map((r) => r.text)
            .join(''),
        )
        .join(' '),
    );

  // The node's own runs as link-carrying spans, with a space between MCIDs.
  const spansOf = (node: StructNode): Array<TextSpan> => {
    const spans: Array<TextSpan> = [];
    node.mcids.forEach(({ page, mcid }, i) => {
      if (i > 0) spans.push({ text: ' ' });
      for (const run of runsOfMcid(page, mcid)) {
        spans.push({
          text: run.text,
          sizePt: run.fontSizePt,
          // Black is the default; carrying it would put a colour on every run.
          ...(run.colorHex !== '000000' ? { colorHex: run.colorHex } : {}),
          ...(run.fontName !== undefined ? { fontName: run.fontName } : {}),
          ...(run.outlineHex !== undefined
            ? { outline: { colorHex: run.outlineHex, widthPt: pt(run.outlineWidthPt ?? 1) } }
            : {}),
          ...(run.bold ? { bold: true } : {}),
          ...(run.italic ? { italic: true } : {}),
          ...(run.href !== undefined ? { href: run.href } : {}),
        });
      }
    });
    return spans;
  };

  // All text under a node, in reading order (a list item's label + body).
  const collectText = (node: StructNode): string =>
    squash([textOf(node), ...node.children.map(collectText)].join(' '));

  function emit(node: StructNode, out: Array<BodyElement>): void {
    if (node.type === 'Table') {
      const table = buildTable(node);
      if (table) out.push(table);
      return;
    }
    if (node.type === 'Figure') {
      for (const img of imagesForNode(node)) {
        emitted.add(img);
        out.push(imageBlock(img, resources, node.alt));
      }
      return;
    }
    if (node.type === 'LI') {
      const text = collectText(node);
      if (text.length > 0) out.push(paragraphBlock(text, undefined));
      return;
    }
    if (node.children.length === 0) {
      if (textOf(node).length > 0) {
        out.push(paragraphFromRuns(spansOf(node), headingLevel(node.type)));
      }
      return;
    }
    for (const child of node.children) emit(child, out);
  }

  function buildTable(tableNode: StructNode): BodyElement | undefined {
    const raw: Array<RawRow> = [];
    const collectRows = (n: StructNode): void => {
      for (const child of n.children) {
        if (child.type === 'TR') raw.push(buildRow(child));
        else if (child.type === 'THead' || child.type === 'TBody' || child.type === 'TFoot') {
          collectRows(child);
        }
      }
    };
    collectRows(tableNode);
    if (raw.length === 0) return undefined;
    const laid = layOutColumns(raw);
    const table: Table = { properties: {}, grid: laid.grid, rows: laid.rows };
    return { kind: 'table', table };
  }

  /**
   * The span of page x every glyph under a node covers — measured, since the
   * interpreter advances the text matrix by the font's own widths (§9.4.4).
   */
  function edgesOf(node: StructNode): { left: number; right: number } | undefined {
    let left: number | undefined;
    let right: number | undefined;
    const visit = (n: StructNode): void => {
      for (const { page, mcid } of n.mcids) {
        for (const run of runsOfMcid(page, mcid)) {
          if (left === undefined || run.x < left) left = run.x;
          if (right === undefined || run.endX > right) right = run.endX;
        }
      }
      for (const child of n.children) visit(child);
    };
    visit(node);
    return left !== undefined && right !== undefined ? { left, right } : undefined;
  }

  function buildRow(trNode: StructNode): RawRow {
    const cells: Array<RawCell> = [];
    let allHeader = false;
    for (const cell of trNode.children) {
      if (cell.type !== 'TH' && cell.type !== 'TD') continue;
      if (cells.length === 0) allHeader = true;
      if (cell.type !== 'TH') allHeader = false;
      const content: Array<BodyElement> = [];
      for (const child of cell.children) emit(child, content);
      // Text sitting directly on the TD, with no child element to carry it:
      // taken as SPANS so the cell keeps its size and any link, not as bare text.
      if (content.length === 0) content.push(paragraphFromRuns(spansOf(cell)));
      const edges = edgesOf(cell);
      cells.push({
        content,
        span: cell.colSpan ?? 1,
        ...(edges ? { x: edges.left, right: edges.right } : {}),
      });
    }
    return { header: allHeader && cells.length > 0, cells };
  }

  const body: Array<BodyElement> = [];
  emit(root, body);
  // Artwork sits UNDER the text the tree placed: `zOrder` starts below zero so
  // a lifted rule never covers the words it rules off.
  let zOrder = -1_000_000;

  // A structure tree names the document's WORDS, and says nothing about the
  // lines drawn around them. 160F-2019.pdf is a form: every rule, every box and
  // every tinted field is a painted path, and reading the tree alone gave its
  // text with no form under it at all. Those paths are lifted and anchored
  // where the page drew them, exactly as the untagged path does — the tree
  // supplies the reading order, the page supplies its own artwork.
  pages.forEach((page, index) => {
    // §14.11.1/§14.11.2 — the page as it is SHOWN, corner and turn together.
    const display = shown[index]!;
    const frame = { left: 0, top: display.height };
    const covered = (pageImages[index]?.images ?? []).map((img) => ({
      minX: img.x,
      minY: img.y,
      maxX: img.x + img.widthPt,
      maxY: img.y + img.heightPt,
    }));
    const lifted = collectPageVectors(file, page, covered);
    imageLosses.push(...lifted.losses);
    for (const v of placeVectors(lifted.vectors, display)) {
      body.push(shapeBlock(v, frame, zOrder++));
    }
  });

  // Images not claimed by a /Figure (untagged figures, third-party PDFs) still
  // belong in the document — append them in page + top-down order so nothing is
  // silently lost.
  const orphans: Array<{ page: number; img: PdfImage }> = [];
  pageImages.forEach((p, page) => {
    for (const img of p.images) if (!emitted.has(img)) orphans.push({ page, img });
  });
  orphans.sort((a, b) => a.page - b.page || b.img.y - a.img.y);
  for (const { img } of orphans) body.push(imageBlock(img, resources));

  if (body.length === 0) return undefined;
  // A tree that names words the page never marked describes nothing.
  //
  // annotation-choice-widget.pdf carries a structure tree and not one of its
  // runs carries an MCID, so every node came back empty and the file converted
  // to its list boxes with no text in them at all — while the artwork alone
  // kept `body` non-empty, so the tagged reading still won. Marked content is
  // what joins the tree to the page; where the join is missing the heuristic
  // reading has the whole page to work from, and the words come back.
  //
  // Half, not all: a header, a footer and a page number are Artifacts by
  // design and belong to no element, and a document is not untagged for
  // leaving them out.
  const onPage = placedRuns.flat().reduce((n, r) => n + r.text.length, 0);
  const reached = [...claimed].reduce((n, r) => n + r.text.length, 0);
  if (onPage > 0 && reached * 2 < onPage) return undefined;
  return {
    doc: buildFlowDoc(
      body,
      resources,
      // A tagged reading re-sets the words exactly as an untagged one does, so
      // it needs the same margins: measured off where the source put them.
      // Without this every tagged PDF came back with its text against all four
      // edges of the paper.
      withMeasuredMargins(sectionFromPdfPages(pages), shown, placedRuns),
      collectEmbeddedFonts(file, pages),
    ),
    losses: imageLosses,
  };
}

/** A cell before its column span is known: content, the tagged span, its page x. */
interface RawCell {
  readonly content: Array<BodyElement>;
  /** `/ColSpan` as the structure tree states it — a floor, never a ceiling. */
  readonly span: number;
  /** Leftmost glyph in page space; absent when the cell holds no text. */
  readonly x?: number;
  /** Where its glyphs stop, estimated — the last column has no start after it. */
  readonly right?: number;
}

interface RawRow {
  readonly header: boolean;
  readonly cells: Array<RawCell>;
}

/** Two starts within this many points are the same column. */
const COLUMN_TOLERANCE_PT = 2;

/**
 * Lay the tagged cells onto a column grid read from the page.
 *
 * A structure tree states no column widths, and states `/ColSpan` only when its
 * producer bothered: 160F-2019.pdf tags twenty-six columns and then writes rows
 * of three plain `TD`s that visually run half the page. Believed literally,
 * those three sat in columns 0–2 while twenty-three stood empty beside them,
 * and 153 characters were asked to fit in 7.7pt — one page reconstructed as
 * five, a word to a line.
 *
 * So the grid comes from where the cells actually START. Every distinct start
 * across the table is a column boundary; a cell runs from its own boundary to
 * the next cell's in its row, which is its span; and the width of a column is
 * the distance to the next boundary. The tagged `/ColSpan` still sets a floor,
 * so a producer that did the work is never contradicted.
 *
 * Falls back to the tagged spans over an equal grid when the page says too
 * little — fewer than two distinct starts, or a table whose cells hold no text.
 *
 * @param raw The rows as tagged, each cell carrying its page x where it has one.
 * @returns The rows with spans resolved, and one width per column.
 */
function layOutColumns(raw: ReadonlyArray<RawRow>): {
  rows: Array<TableRow>;
  grid: Array<Pt>;
} {
  const bounds = columnBounds(raw);
  if (bounds.length < 2) return equalGrid(raw);

  const indexOf = (x: number): number => {
    let best = 0;
    for (let i = 0; i < bounds.length; i++) if (x >= bounds[i]! - COLUMN_TOLERANCE_PT) best = i;
    return best;
  };

  const rows = raw.map(({ header, cells }) => {
    const out: Array<TableCell> = [];
    let col = 0;
    cells.forEach((cell, i) => {
      const start = cell.x !== undefined ? Math.max(col, indexOf(cell.x)) : col;
      // The cell reaches the next cell that knows where it starts; the last of
      // a row reaches the end of the grid, which is what fills the row out.
      const nextX = cells.slice(i + 1).find((c) => c.x !== undefined)?.x;
      const end = nextX !== undefined ? Math.max(start + 1, indexOf(nextX)) : bounds.length;
      const span = Math.max(cell.span, end - start, 1);
      out.push({ properties: span > 1 ? { colSpan: span } : {}, content: cell.content });
      col = start + span;
    });
    return { properties: header ? { isHeader: true } : {}, cells: out };
  });

  // A boundary is a START, so every column but the last is measured by the one
  // after it. The last has nothing after it and is measured to where the text
  // stops instead — the mean of the others would make a two-column table equal
  // however far apart its two columns actually are.
  const widths = bounds.map((x, i) => (i + 1 < bounds.length ? bounds[i + 1]! - x : 0));
  const tableRight = Math.max(
    ...raw.flatMap((r) => r.cells.map((c) => c.right ?? 0)),
    bounds[bounds.length - 1]!,
  );
  widths[widths.length - 1] = Math.max(0, tableRight - bounds[bounds.length - 1]!);

  const total = widths.reduce((a, b) => a + Math.max(MIN_COLUMN_PT, b), 0);
  const scale = total > 0 ? ASSUMED_CONTENT_WIDTH_PT / total : 1;
  return { rows, grid: widths.map((w) => pt(Math.max(1, Math.max(MIN_COLUMN_PT, w) * scale))) };
}

/** Every distinct cell start across the table, ascending — the column boundaries. */
function columnBounds(raw: ReadonlyArray<RawRow>): Array<number> {
  const xs = raw
    .flatMap((r) => r.cells.map((c) => c.x))
    .filter((x): x is number => x !== undefined)
    .sort((a, b) => a - b);
  const bounds: Array<number> = [];
  for (const x of xs) {
    const last = bounds[bounds.length - 1];
    if (last === undefined || x - last > COLUMN_TOLERANCE_PT) bounds.push(x);
  }
  return bounds;
}

/** The old reading: the tagged spans, over columns of equal width. */
function equalGrid(raw: ReadonlyArray<RawRow>): { rows: Array<TableRow>; grid: Array<Pt> } {
  const rows = raw.map(({ header, cells }) => ({
    properties: header ? { isHeader: true } : {},
    cells: cells.map((c) => ({
      properties: c.span > 1 ? { colSpan: c.span } : {},
      content: c.content,
    })),
  }));
  const numCols = Math.max(
    1,
    ...rows.map((r) => r.cells.reduce((s, c) => s + (c.properties.colSpan ?? 1), 0)),
  );
  const w = pt(Math.max(1, ASSUMED_CONTENT_WIDTH_PT / numCols));
  return { rows, grid: Array.from({ length: numCols }, () => w) };
}

function squash(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// H1–H6 → outline level 0–5 (the FlowDoc heading representation, §17.3.1.20).
function headingLevel(type: string): number | undefined {
  const m = /^H([1-6])$/.exec(type);
  return m ? Number(m[1]) - 1 : undefined;
}
