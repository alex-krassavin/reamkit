// E-PDF EP3 — tagged fast-path reconstruction. Walks the structure tree (EP3
// struct-tree.ts), pulls each element's text from the per-page MCID → text map
// the content interpreter produced (EP2), and rebuilds a FlowDoc: headings
// (H1–H6 → outline level), paragraphs, tables (Table → TR → TH/TD), list items
// (each LI → its label + body text), and figures (EP6 — each /Figure's MCID
// resolves to a lifted image, carrying its /Alt). The honest inverse of the
// tagged PDF Ream writes.

import {
  ASCENDER,
  CARRIER_LINE_PT,
  buildFlowDoc,
  dedupeLosses,
  floatOntoSheet,
  imageBlock,
  paragraphBlock,
  paragraphFromRuns,
  sectionFromPdfPages,
  sectionOnSheet,
  shapeBlock,
  spaceAfter,
  withMeasuredMargins,
} from './flow-build';
import { displayOf, placeRuns, placeVectors, textFrameOf, wordsTurnOf } from './display';
import { collectEmbeddedFonts } from './embedded-fonts';
import { collectFaceFamilies } from './font';
import { collectPageImages } from './images';
import { collectPageVectors } from './vector';
import { UNMAPPED, endedParagraph } from './layout';
import { markDrawnRules } from './text-rules';
import { readStructTree } from './struct-tree';
import { extractPageText } from './text';
import type { BodyElement, Table, TableCell, TableRow } from '@/core/document-model';
import type { Loss, Pt } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import type { PdfImage } from './images';
import type { StructNode } from './struct-tree';
import { FEATURES, ResourceStore, pt } from '@/core/ir';

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
  const sheets = pages.map((page) => displayOf(page));
  const extracted = pages.map((page) => extractPageText(file, page));
  const onSheets = extracted.map((runs, i) => placeRuns(runs, sheets[i]!));
  // §17.6.20 — a document whose words run DOWN its sheets is read in the frame
  // where they stand upright and set back on the sheets turned, as the untagged
  // reading does a page at a time. A tagged reading is one section, so it
  // turns when every page that has words has them running down it.
  const worded = onSheets.filter((runs) => runs.some((r) => r.text.trim() !== ''));
  const turned = worded.length > 0 && worded.every((runs) => wordsTurnOf(runs) === 270);
  const shown = turned ? sheets.map((sheet) => textFrameOf(sheet)) : sheets;
  const placedRuns = turned ? extracted.map((runs, i) => placeRuns(runs, shown[i]!)) : onSheets;
  // A PDF has no underline: it draws a thin bar under the words, and the tree
  // says nothing about it. Read onto the runs before they are gathered, so it
  // travels with them, and the bar itself is not placed a second time.
  const vectorLosses: Array<Loss> = [];
  const pageVectors = pages.map((page, i) => {
    // A white box drawn over a picture is what HIDES it, so the paths are
    // filtered against the pictures already lifted off the same page.
    const covered = collectPageImages(file, page).images.map((img) => ({
      minX: img.x,
      minY: img.y,
      maxX: img.x + img.widthPt,
      maxY: img.y + img.heightPt,
    }));
    const lifted = collectPageVectors(file, page, covered);
    vectorLosses.push(...lifted.losses);
    return placeVectors(lifted.vectors, shown[i]!);
  });
  const ruled = placedRuns.map((runs, i) => markDrawnRules(runs, pageVectors[i] ?? []));
  // Per page: MCID → its runs, in show order (runs carry any hyperlink, EP8).
  const pageRuns = ruled.map(({ runs }) => {
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

  // The node's own runs as link-carrying spans. Between two MCIDs stands the
  // space the page shows there: one where the second starts a line or stands
  // clear of the first, none where it closes up on it. A space every time put
  // one before the full stop that ends a formula — bug1937438_mml_from_latex.pdf
  // read "𝑥 ∈ ℝ ." — and doubled the one a phrase already ended with.
  const spansOf = (node: StructNode): Array<TextSpan> => {
    const spans: Array<TextSpan> = [];
    let last: TextRun | undefined;
    for (const { page, mcid } of node.mcids) {
      const runs = runsOfMcid(page, mcid);
      const first = runs[0];
      if (last !== undefined && first !== undefined && spacedApart(last, first)) {
        spans.push(spaceAfter(spans[spans.length - 1]));
      }
      for (const run of runs) spans.push(spanOf(run));
      last = runs[runs.length - 1] ?? last;
    }
    return spans;
  };

  /** One run as the span that carries everything the page showed it with. */
  const spanOf = (run: TextRun): TextSpan => ({
    // §9.10.2 — a glyph the face maps to no character says "no text here", and
    // writing it on would put bytes no reader can show into the document.
    text: run.text.replaceAll(UNMAPPED, ''),
    sizePt: run.fontSizePt,
    // Black is the default; carrying it would put a colour on every run.
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

  // All text under a node, in reading order (a list item's label + body).
  const collectText = (node: StructNode): string =>
    squash([textOf(node), ...node.children.map(collectText)].join(' '));

  // Where each emitted paragraph was SET on the page, for the space between
  // them: the tree says what the words are and never how far apart they stood.
  const setting = new Map<BodyElement, Setting>();
  // …and the page each element begins on, which the tree does not say either.
  const pageOf = new Map<BodyElement, number>();
  /** The lowest baseline a page's own paragraphs reach, page space (y up). */
  const lowestOf = (own: ReadonlyArray<BodyElement>): number | undefined => {
    const bottoms = own.flatMap((el) => {
      const set = setting.get(el);
      return set ? [set.bottom] : [];
    });
    return bottoms.length > 0 ? Math.min(...bottoms) : undefined;
  };

  /** The topmost and bottommost baseline under a node, and its largest face. */
  function baselinesOf(node: StructNode): Setting | undefined {
    let top: number | undefined;
    let bottom: number | undefined;
    let size = 0;
    const visit = (n: StructNode): void => {
      for (const { page, mcid } of n.mcids) {
        for (const run of runsOfMcid(page, mcid)) {
          if (top === undefined || run.y > top) top = run.y;
          if (bottom === undefined || run.y < bottom) bottom = run.y;
          if (run.fontSizePt > size) size = run.fontSizePt;
        }
      }
      for (const child of n.children) visit(child);
    };
    visit(node);
    return top !== undefined && bottom !== undefined && size > 0
      ? { top, bottom, size }
      : undefined;
  }

  /**
   * The element's own lines, gathered into the paragraphs they were SET as.
   *
   * A tree names elements, not lines, and a producer may put a whole page under
   * one: annotation-underline.pdf marks every glyph on its page with the same
   * id, so a heading, a blank line and a body line came back as one paragraph
   * and re-wrapped into a single run-together line. The page still says where
   * one ended — a line stopping well short of the measure stopped because its
   * author stopped it — and that is the rule the heuristic reading already uses
   * ({@link endedParagraph}). A properly tagged paragraph's inner lines all
   * reach the measure, so nothing there changes.
   */
  function settingsOf(node: StructNode): Array<{ spans: Array<TextSpan>; set: Setting }> {
    const spans = spansOf(node);
    const lines = linesOf(node);
    if (lines.length < 2) {
      const set = baselinesOf(node);
      return set ? [{ spans, set }] : [];
    }
    const measure = {
      left: Math.min(...lines.map((l) => l.x)),
      right: Math.max(...lines.map((l) => l.x + l.width)),
    };
    const groups: Array<Array<(typeof lines)[number]>> = [];
    let prev: (typeof lines)[number] | undefined;
    for (const line of lines) {
      const gap = prev === undefined ? 0 : prev.y - line.y;
      const opened = prev !== undefined && gap > line.fontSize * 1.5;
      if (groups.length === 0 || opened || (prev && endedParagraph(prev, line, measure))) {
        groups.push([]);
      }
      groups[groups.length - 1]!.push(line);
      prev = line;
    }
    // A line the page ended with a space has its break written already, and
    // joined with a second one chrome-text-selection-markedContent.pdf read
    // "in the  range of".
    const ends = (spans: ReadonlyArray<TextSpan>): boolean => /\s$/u.test(spans.at(-1)?.text ?? '');
    const opens = (spans: ReadonlyArray<TextSpan>): boolean => /^\s/u.test(spans[0]?.text ?? '');
    return groups.map((g) => ({
      spans: g.flatMap((l, i) =>
        i > 0 && !ends(g[i - 1]!.spans) && !opens(l.spans)
          ? [spaceAfter(g[i - 1]!.spans.at(-1)), ...l.spans]
          : [...l.spans],
      ),
      set: {
        top: g[0]!.y,
        bottom: g[g.length - 1]!.y,
        size: Math.max(...g.map((l) => l.fontSize)),
      },
    }));
  }

  /** The node's runs clustered onto the baselines they were shown on. */
  function linesOf(
    node: StructNode,
  ): Array<{ y: number; x: number; width: number; fontSize: number; spans: Array<TextSpan> }> {
    const rows: Array<{ y: number; size: number; runs: Array<TextRun> }> = [];
    const visit = (n: StructNode): void => {
      for (const { page, mcid } of n.mcids) {
        for (const run of runsOfMcid(page, mcid)) {
          if (run.angleDeg !== undefined) return;
          // A superscript sits off the baseline it belongs to and is set
          // SMALLER, so the tolerance has to come from the LINE's face and not
          // the mark's: measured against its own 7pt, bug2013793.pdf's "240th"
          // made a line of its own out of two raised "th"s and cut the
          // paragraph in half around it.
          const row = rows.find(
            (r) => Math.abs(r.y - run.y) <= Math.max(r.size, run.fontSizePt) * 0.5,
          );
          if (row) {
            row.runs.push(run);
            row.size = Math.max(row.size, run.fontSizePt);
            // The baseline is the one most of the line stands on, which is the
            // lowest of them: a raised mark never lowers it.
            row.y = Math.min(row.y, run.y);
          } else rows.push({ y: run.y, size: run.fontSizePt, runs: [run] });
        }
      }
      for (const child of n.children) visit(child);
    };
    visit(node);
    return rows
      .sort((a, b) => b.y - a.y)
      .map(({ y, runs }) => {
        const ordered = [...runs].sort((a, b) => a.x - b.x);
        const x = Math.min(...ordered.map((r) => r.x));
        return {
          y,
          x,
          width: Math.max(...ordered.map((r) => r.endX)) - x,
          fontSize: Math.max(...ordered.map((r) => r.fontSizePt)),
          spans: ordered.map((r) => spanOf(r)),
        };
      })
      .filter((l) => l.spans.some((sp) => sp.text.trim().length > 0));
  }

  function emit(node: StructNode, out: Array<BodyElement>): void {
    // The page the node's words begin on, which is the page what it becomes
    // begins on (see `byPage`).
    const on = (el: BodyElement): BodyElement => {
      const page = firstPageOf(node);
      if (page !== undefined) pageOf.set(el, page);
      return el;
    };
    if (node.type === 'Table') {
      const table = buildTable(node);
      if (table) out.push(on(table));
      return;
    }
    if (node.type === 'Figure') {
      for (const img of imagesForNode(node)) {
        emitted.add(img);
        out.push(on(imageBlock(img, resources, node.alt)));
      }
      return;
    }
    if (node.type === 'LI') {
      const text = collectText(node);
      if (text.length > 0) out.push(on(paragraphBlock(text, undefined)));
      return;
    }
    if (node.children.length === 0) {
      if (textOf(node).length > 0) {
        for (const part of settingsOf(node)) {
          const el = paragraphFromRuns(part.spans, headingLevel(node.type));
          setting.set(el, part.set);
          out.push(on(el));
        }
      }
      return;
    }
    for (const child of node.children) emit(child, out);
  }

  /** The first page any marked content under a node stands on. */
  function firstPageOf(node: StructNode): number | undefined {
    let first: number | undefined;
    const visit = (n: StructNode): void => {
      for (const { page } of n.mcids) if (first === undefined || page < first) first = page;
      for (const child of n.children) visit(child);
    };
    visit(node);
    return first;
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

  const named: Array<BodyElement> = [];
  emit(root, named);
  // §14.8 — a tree names a document's elements in reading order and says
  // nothing of the pages they stand on, and read as one flow a document came
  // back with its pages run together: bug793632.pdf is four pages of a line
  // each — three of front matter and the first of the body — and came back as
  // one. Each element stands on the page its words begin on, and each source
  // page opens an output page of its own, as the heuristic reading's do.
  const byPage: Array<Array<BodyElement>> = pages.map(() => []);
  let current = 0;
  for (const el of named) {
    // A tree may name something on an earlier page after something on a
    // later one; it keeps its place in the reading order rather than go back.
    current = Math.max(current, pageOf.get(el) ?? current);
    byPage[current]!.push(el);
  }
  // Artwork sits UNDER the text the tree placed: `zOrder` starts below zero so
  // a lifted rule never covers the words it rules off.
  let zOrder = -1_000_000;

  // A structure tree names the document's WORDS, and says nothing about the
  // lines drawn around them. 160F-2019.pdf is a form: every rule, every box and
  // every tinted field is a painted path, and reading the tree alone gave its
  // text with no form under it at all. Those paths are lifted and anchored
  // where the page drew them, exactly as the untagged path does — the tree
  // supplies the reading order, the page supplies its own artwork.
  imageLosses.push(...vectorLosses);
  pages.forEach((_page, index) => {
    // §14.11.1/§14.11.2 — the page as it is SHOWN, corner and turn together.
    const frame = { left: 0, top: shown[index]!.height };
    const taken = ruled[index]?.consumed;
    const drawn: Array<BodyElement> = [];
    for (const v of pageVectors[index] ?? []) {
      // A bar that became a run's underline is not also a bar on the page.
      if (taken?.has(v) === true) continue;
      const shape = shapeBlock(v, frame, zOrder++, true);
      // …and on a turned sheet it stands where the page drew it, turned with it.
      drawn.push(turned ? floatOntoSheet(shape, shown[index]!.height) : shape);
    }
    // First on its page: anchored there, whatever the page's words run on to.
    byPage[index]!.unshift(...drawn);
  });

  // Images not claimed by a /Figure (untagged figures, third-party PDFs) still
  // belong in the document — each at the end of its own page, top-down, so
  // nothing is silently lost.
  pageImages.forEach((p, page) => {
    const left = p.images.filter((img) => !emitted.has(img)).sort((a, b) => b.y - a.y);
    byPage[page]!.push(...left.map((img) => imageBlock(img, resources)));
  });

  if (byPage.every((own) => own.length === 0)) return undefined;
  // A tagged reading re-sets the words exactly as an untagged one does, so it
  // needs the same margins: measured off where the source put them. Without
  // this every tagged PDF came back with its text against all four edges of
  // the paper.
  const measured = withMeasuredMargins(
    sectionFromPdfPages(pages, shown[0]),
    shown,
    placedRuns,
    pageImages.map((p) => p.images),
  );
  const body: Array<BodyElement> = [];
  const top = measured?.margins?.top ?? 0;
  const bottom = measured?.margins?.bottom ?? 0;
  // Taken before the spacing below re-writes the paragraphs it spaces.
  const ends = byPage.map((own) => lowestOf(own));
  byPage.forEach((own, index) => {
    const height = shown[index]?.height ?? 0;
    spaceParagraphs(own, setting, height);
    // A page OPENS an output page of its own where the page before it ended
    // short of its foot: a title page, the end of a chapter, a sheet of front
    // matter a line long — a break its author made. A page the words filled
    // was broken by the paper, and where the words re-set, the break falls
    // wherever the new setting puts it: bug1997343.pdf's first page is two
    // columns of a paper read as one, and broken again where the page broke,
    // a sheet of it ran on to a page of its own. A page with nothing on it is
    // still a page, and so is the one after it.
    const ended = ends[index - 1];
    const opens =
      index > 0 &&
      (own.length === 0 ||
        ended === undefined ||
        ended - bottom > (height - top - bottom) * SHORT_PAGE_SHARE);
    // §17.3.1.33 — a page that opens begins its text where the page began it,
    // not against the top margin, which is measured to the highest ink of any.
    const lead = own.findIndex((el) => !floats(el));
    const set = lead >= 0 ? setting.get(own[lead]!) : undefined;
    if ((index === 0 || opens) && set !== undefined && measured?.margins) {
      const gap = height - top - (set.top + set.size * ASCENDER);
      if (gap > 1) own[lead] = spacedBefore(own[lead]!, gap);
    }
    // A break stands before the paragraph that carries it, so a blank FIRST
    // sheet holds a carrier of its own for the second to break from.
    if (opens) body.push(pageBreak(true));
    else if (index === 0 && own.length === 0 && pages.length > 1) body.push(pageBreak(false));
    body.push(...own);
  });
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
  if (placedRuns.some((page) => page.some((r) => r.text.includes(UNMAPPED)))) {
    imageLosses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail:
        'some glyphs map to no character — the font states no /ToUnicode and its program says nothing either, so that text is unrecoverable',
    });
  }
  const onPage = placedRuns.flat().reduce((n, r) => n + r.text.length, 0);
  const reached = [...claimed].reduce((n, r) => n + r.text.length, 0);
  if (onPage > 0 && reached * 2 < onPage) return undefined;
  return {
    doc: buildFlowDoc(
      body,
      resources,
      // The margins measured above, on the sheet the pages are shown on.
      sectionOnSheet(measured, shown[0]),
      collectEmbeddedFonts(file, pages, imageLosses),
      [],
      undefined,
      collectFaceFamilies(file, pages),
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

/**
 * Whether the page shows a space between one run and the next: the second
 * starts another line, or stands clear of the first by more than a kern —
 * and neither already carries the space.
 */
function spacedApart(prev: TextRun, next: TextRun): boolean {
  if (/\s$/u.test(prev.text) || /^\s/u.test(next.text)) return false;
  const size = Math.max(prev.fontSizePt, next.fontSizePt, 1);
  if (Math.abs(prev.y - next.y) > size * 0.5) return true;
  return next.x - prev.endX > size * MCID_SPACE_EM;
}

/** A gap between two stretches of marked content this wide, in ems, is a word space. */
const MCID_SPACE_EM = 0.15;

function squash(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * How much of a page's text area the page may leave empty below its last line
 * and still be FULL: a page that ends higher than this ended on purpose.
 */
const SHORT_PAGE_SHARE = 0.25;

/** Whether an element FLOATS — a drawing anchored to its page, which takes no room. */
function floats(el: BodyElement): boolean {
  return (
    (el.kind === 'image' && el.image.float !== undefined) ||
    (el.kind === 'shape' && el.shape.float !== undefined)
  );
}

/** A paragraph given the space the page left above it; anything else as it is. */
function spacedBefore(el: BodyElement, before: number): BodyElement {
  if (el.kind !== 'paragraph') return el;
  const properties = { ...el.paragraph.properties, spacingBefore: pt(before) };
  return { ...el, paragraph: { ...el.paragraph, properties } };
}

/**
 * An empty paragraph that takes no room: the carrier of a page break, or of
 * nothing at all on a blank first sheet the next page's break has to follow.
 */
function pageBreak(breaks: boolean): BodyElement {
  return {
    kind: 'paragraph',
    paragraph: {
      properties: {
        ...(breaks ? { pageBreakBefore: true } : {}),
        spacingLine: CARRIER_LINE_PT,
        spacingLineRule: 'exact',
      },
      runs: [],
    },
  };
}

// H1–H6 → outline level 0–5 (the FlowDoc heading representation, §17.3.1.20).
function headingLevel(type: string): number | undefined {
  const m = /^H([1-6])$/.exec(type);
  return m ? Number(m[1]) - 1 : undefined;
}

/** Where a tagged paragraph stood: its outer baselines and its largest face. */
interface Setting {
  /** The topmost baseline under the node, in page space (y up). */
  readonly top: number;
  /** The bottommost. */
  readonly bottom: number;
  /** The largest face any of its runs was shown in, in points. */
  readonly size: number;
}

/**
 * §17.3.1.33 `w:spacing` — the space the SOURCE left before each paragraph.
 *
 * A structure tree names the words and says nothing about how far apart they
 * stood, so every tagged PDF came back at one flat leading: on
 * annotation-polyline-polygon-without-appearance.pdf the two labels, set a
 * third of a page apart above their own drawings, arrived as two lines
 * touching. The page still says it — the gap between the last baseline of one
 * paragraph and the first of the next, less the line it would have taken
 * anyway. This is the rule the heuristic reading already uses, applied to the
 * paragraphs the tree named.
 *
 * @param body       The body elements, in order; amended in place.
 * @param setting    Where each paragraph was set, for those that were.
 * @param pageHeight The shown page's height, which bounds any one gap.
 */
function spaceParagraphs(
  body: Array<BodyElement>,
  setting: Map<BodyElement, Setting>,
  pageHeight: number,
): void {
  let prev: Setting | undefined;
  body.forEach((el, i) => {
    const here = setting.get(el);
    if (!here) return;
    // Only DOWN the page: a paragraph the tree put after one that stands below
    // it is not spaced by the distance between them, it is out of order.
    const gap = prev !== undefined ? prev.bottom - here.top : 0;
    prev = here;
    if (!(gap > 0)) return;
    const opened = gap - here.size * 1.2;
    // Under a third of a line is leading, not spacing.
    if (!(opened > here.size * 0.3)) return;
    if (el.kind !== 'paragraph') return;
    // No one gap may take more than a third of the sheet: the tree is trusted
    // for the ORDER of its paragraphs, and a gap larger than that is a page
    // this reading has no other way to see.
    const most = pageHeight > 0 ? pageHeight / 3 : here.size * 3;
    body[i] = {
      ...el,
      paragraph: {
        ...el.paragraph,
        properties: {
          ...el.paragraph.properties,
          spacingBefore: pt(Math.min(opened, most)),
        },
      },
    };
  });
}
