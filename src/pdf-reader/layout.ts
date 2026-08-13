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
import { matrixBlocks } from './math-rows';
import { isRightToLeft } from './content';
import type {
  BodyElement,
  ParagraphProperties,
  Run,
  SectionProperties,
  TabStop,
  Table,
  TableCell,
} from '@/core/document-model';
import type { Loss, Pt } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile, PdfPage } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import { FEATURES, ResourceStore, pt } from '@/core/ir';

/** The relationships the reconstruction files its running head and foot under. */
const FOOTER_PART = 'pdf-running-foot';
const HEADER_PART = 'pdf-running-head';

/** §9.10.2 — a glyph the face maps to no character (see `./font`). */
export const UNMAPPED = '\uFFFD';

interface Line {
  readonly y: number; // baseline (page space, y-up)
  readonly fontSize: number;
  readonly text: string; // joined text, for emptiness/heading checks
  readonly spans: ReadonlyArray<TextSpan>;
  /** Leftmost glyph origin, and how far the line reaches — placed reconstruction. */
  readonly x: number;
  readonly width: number;
  /** Whether a TAB stands inside it — a gap no word space could be. */
  readonly tabbed?: boolean;
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
  const allRuns = pages.map((page, i) => placeRuns(extractPageText(file, page), shown[i]!));
  // §17.6.13 — what the document repeats at the foot of its pages is a running
  // foot, not a paragraph of the body. Lifted off before anything else reads
  // the page: it must not measure the margins either, and a page number in the
  // text block is a page number in the wrong place.
  const foot = mode === 'positional' ? undefined : runningFoot(allRuns, shown, 'foot');
  const head = mode === 'positional' ? undefined : runningFoot(allRuns, shown, 'head');
  const pageRuns =
    foot || head
      ? allRuns.map((runs, i) =>
          runs.filter((r) => foot?.lift[i]?.has(r) !== true && head?.lift[i]?.has(r) !== true),
        )
      : allRuns;
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
  // §9.10.2 — glyphs whose face maps them to nothing a reader can show. The
  // words are unrecoverable, and a page that silently comes back blank is the
  // one loss this reader must never take without saying so:
  // arial_unicode_ab_cidfont.pdf is four Arabic letters and nothing else.
  if (pageRuns.some((page) => page.some((r) => r.text.includes(UNMAPPED)))) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail:
        'some glyphs map to no character — the font states no /ToUnicode and its program says nothing either, so that text is unrecoverable',
    });
  }
  if (pageRuns.some((page) => page.some((r) => r.fillPatternName !== undefined))) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'text filled with a tiling pattern is drawn as a flat tint of the pattern’s colour, not as the pattern',
    });
  }
  // §8.7.4.5 — and type filled with a SHADING pattern keeps the middle of the
  // sweep, since a run carries one colour: the colour survives, its shape does
  // not. ShowText-ShadingPattern.pdf sets two of its four lines in a blue-to-red
  // gradient and both came back black, which said nothing at all.
  if (pageRuns.some((page) => page.some((r) => r.gradientFill === true))) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'text filled with a shading pattern is drawn in the middle colour of the gradient, not as the gradient',
    });
  }
  // EP17 — the gutters of every page, and then the DOCUMENT's own. A page is
  // set the way its document is set: bug1997343.pdf's second sheet carries a
  // figure across both columns and a float beside it, which leaves too few
  // clean lines for the vote to answer, and read as one column its citations
  // ran into its theorems. Where a page says nothing, the answer the rest of
  // the document gave is put to it, and kept only if its own lines agree.
  const perPage = pages.map((_, i) => detectGutters(pageRuns[i]!, shown[i]!.width));
  const shared = commonGutters(perPage);
  const pageGutters = perPage.map((own, i) =>
    own.length > 0 ? own : shared && fitsGutters(pageRuns[i]!, shared) ? shared : own,
  );
  const body: Array<BodyElement> = [];
  // §17.6 — where the pages differ in size the document is several sections,
  // each ending at the body index its last page's blocks end at.
  const sectionEnds: Array<{
    at: number;
    from: number;
    to: number;
    columns: number;
    spacePt: number;
    continuous: boolean;
  }> = [];
  let sectionFrom = 0;
  let lastSize = '';
  // §17.6.4 — the columns the pages are SET in, which change where a page's
  // gutters do: a paper's title stands over the two columns of its body, and a
  // section is what carries a column setup.
  let curColumns = 1;
  let curSpace = 0;
  let pendingContinuous = false;
  pages.forEach((page, i) => {
    const runs = pageRuns[i]!;
    const display = shown[i]!;
    // EP17 — the page's gutters, and so its columns. Each column is grouped and
    // read independently, and its blocks precede the next column's.
    const gutters = pageGutters[i]!;
    const pageWidth = display.width;
    // Whether this page steps between its words or writes spaces of its own,
    // which decides how wide a gap has to be to mean one.
    const stepped = stepsBetweenWords(runs);
    // Blocks carry a column key so the final sort reads column-by-column: left
    // column top-to-bottom, then right column.
    const blocks: Array<{ band: number; col: number; top: number; el: BodyElement }> = [];
    // EP17 — a full-width line cuts the page in two: what is above it is read
    // before it and what is below after, so a paper's columns do not start at
    // the top of the sheet.
    const split = gutters.length > 0 ? assignColumns(runs, gutters) : undefined;
    // Where the page's text starts and ends, which is the measure a line that
    // spans the page is set across.
    const textEdges = pageTextEdges(runs);
    const bandEpsilon = (median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10) / 2;
    const bandAt = (top: number): number => bandOf(split?.breaks ?? [], top, bandEpsilon);
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
          for (const line of groupIntoLines(rotate(runs, -angle), true, stepped)) {
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
      // §22 — a MATRIX the page drew. It reaches the sheet as numbers on two
      // baselines with stretched brackets between them, and read line by line —
      // which is all a page says — bug1997343.pdf's product of three matrices
      // came back as "1 2 1 1 1 3", "( )( ) = ( )", "3 4 0 1 3 7". Lifted off
      // before the prose is read, so its numbers are not read twice.
      const columnSize = median(colRuns.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
      const maths = matrixBlocks(colRuns, columnSize);
      const inMath = new Set(maths.flatMap((m) => [...m.used]));
      const lines = groupIntoLines(
        inMath.size > 0 ? colRuns.filter((r) => !inMath.has(r)) : colRuns,
        false,
        stepped,
      ).filter((l) => l.text.length > 0);
      // The column the paragraphs were set in — its own edges, not the page's,
      // so a two-column page judges each side against the side it belongs to.
      //
      // The column, not the lines IN it: a measure taken from the very lines
      // being judged is one no line can be inset from, and a title standing
      // alone over the page was never centred because it WAS the measure.
      // bug1997343.pdf sets "A Two Column Example" in the middle of the sheet
      // and we set it flush left.
      const measure = measureOf(col, gutters, textEdges) ?? measureOfLines(lines);
      for (const found of maths) {
        // A display stands where the page stood it — centred in its column, as
        // this one is, or flush with the text.
        const { alignment } = alignmentOf(
          [
            {
              y: found.top,
              fontSize: columnSize,
              text: '',
              spans: [],
              x: found.x,
              width: found.width,
            },
          ],
          measure,
        );
        blocks.push({
          band: bandAt(found.top),
          col,
          top: found.top,
          el: {
            kind: 'paragraph',
            paragraph: {
              properties: alignment ? { alignment } : {},
              runs: [{ text: '', properties: {}, math: found.math }],
            },
          },
        });
      }
      for (const para of groupIntoParagraphs(lines, measure, display.height)) {
        blocks.push({
          band: bandAt(para.top),
          col,
          top: para.top,
          el: paragraphFromRuns(para.spans, headingLevel(para.fontSize, medianFont), {
            ...(para.alignment !== undefined ? { alignment: para.alignment } : {}),
            ...(para.spacingBefore !== undefined ? { spacingBefore: pt(para.spacingBefore) } : {}),
            ...(para.indentLeft !== undefined ? { indentLeft: pt(para.indentLeft) } : {}),
            ...(para.indentFirstLine !== undefined
              ? { indentFirstLine: pt(para.indentFirstLine) }
              : {}),
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

    const colOf = (centerX: number): number => gutters.filter((g) => centerX >= g.mid).length;
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
    // A page RULED into columns is a table, and its ROWS are what it says; a
    // page SET in columns is prose, and its columns are. Read by column, a
    // table comes back one column at a time with every row torn up.
    const ruledIntoColumns = mode !== 'positional' && looksRuled(ruled.runs, gutters, textEdges);
    const asTable =
      ruledIntoColumns && textEdges
        ? tableFrom(ruled.runs, gutters, textEdges, stepped)
        : undefined;
    if (asTable) {
      for (const block of asTable) {
        blocks.push({ band: bandAt(block.top), col: 0, top: block.top, el: block.el });
      }
    } else if (split && !ruledIntoColumns) {
      // A run the rules pass rebuilt is not the one the split was measured on,
      // so its column is looked up by where it stands.
      const columnFor = (r: TextRun): number => split.columnOf.get(r) ?? colOf(r.x);
      const columns = Array.from({ length: gutters.length + 1 }, (_, n) => n);
      for (const col of [SPANNING_COLUMN, ...columns]) {
        addColumn(
          ruled.runs.filter((r) => columnFor(r) === col),
          col,
        );
      }
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
      blocks.push({ band: bandAt(mark.top), col: mark.col, top: mark.top, el: mark.make(z) });
    });
    blocks.sort(
      (a, b) => a.band - b.band || columnOrder(a.col) - columnOrder(b.col) || b.top - a.top,
    );
    // §14.11.2 — a page whose SIZE differs from the one before it opens a
    // section of its own, because a section is what carries a page size.
    // function_based_shading_cmyk.pdf is 290×290 and then 1880×1260, and read
    // as one size the second sheet's six squares were cut down to the one that
    // fitted. A section break already forces a page, so the break paragraph
    // below is for the pages that stay inside one.
    const size = `${shown[i]!.width.toFixed(2)}x${shown[i]!.height.toFixed(2)}`;
    const opensSection = i > 0 && size !== lastSize;
    if (opensSection) {
      sectionEnds.push({
        at: body.length,
        from: sectionFrom,
        to: i,
        columns: curColumns,
        spacePt: curSpace,
        continuous: pendingContinuous,
      });
      sectionFrom = i;
      pendingContinuous = false;
    }
    lastSize = size;
    // Each source page after the first opens an output page of its own. Flowed,
    // the layout repaginates and this hardly shows; PLACED, every mark is
    // anchored to "the page", so without it all twenty-five pages of
    // Brotli-Prototype-FileA.pdf stack onto one.
    if (i > 0 && !opensSection && blocks.length > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: {
          properties: { pageBreakBefore: true, spacingLine: pt(0), spacingLineRule: 'exact' },
          runs: [],
        },
      });
    }
    // A page set in columns is REPRODUCED in them: the reading order alone
    // leaves a two-column paper re-set as one long column, which is not the
    // page the file draws. The count changes at a band that spans — the title
    // over the columns, the footer under them — and each change is a section of
    // its own, continuous, so no page is opened for it.
    // The columns the page was READ in, which a ruled page has none of: its
    // gutters are a table's, and the rows were taken whole.
    const columnsHere = ruledIntoColumns ? 1 : gutters.length + 1;
    const spacePt = columnsHere > 1 ? median(gutters.map((g) => g.to - g.from)) : 0;
    for (const block of blocks) {
      const count = block.col === SPANNING_COLUMN ? 1 : columnsHere;
      if (count !== curColumns) {
        sectionEnds.push({
          at: body.length,
          from: sectionFrom,
          to: i + 1,
          columns: curColumns,
          spacePt: curSpace,
          continuous: pendingContinuous,
        });
        sectionFrom = i;
        pendingContinuous = true;
        curColumns = count;
        curSpace = spacePt;
      }
      body.push(block.el);
    }
  });
  // A placed reading anchors everything to the page, so its margins must stay
  // at zero or the anchors move. A FLOWING one is a document being re-set, and
  // a document with no margins prints its words against the edge of the paper
  // — which is what every converted PDF looked like.
  const setUp = (from: number, to: number): SectionProperties | undefined => {
    const own = sectionFromPdfPages(pages.slice(from, to));
    return mode === 'positional'
      ? own
      : withMeasuredMargins(own, shown.slice(from, to), pageRuns.slice(from, to));
  };
  sectionEnds.push({
    at: body.length,
    from: sectionFrom,
    to: pages.length,
    columns: curColumns,
    spacePt: curSpace,
    continuous: pendingContinuous,
  });
  const sections =
    sectionEnds.length > 1
      ? sectionEnds.flatMap((end) => {
          const base = setUp(end.from, end.to);
          if (!base) return [];
          const properties: SectionProperties = {
            ...base,
            ...(end.columns > 1 && mode !== 'positional'
              ? { columns: { count: end.columns, spacePt: end.spacePt } }
              : {}),
            ...(end.continuous ? { sectionStart: 'continuous' as const } : {}),
          };
          return [{ properties, endIndex: end.at }];
        })
      : [];
  // The band is built from the page that showed it first, and referenced by
  // every section: the foot runs through the document, not through a section.
  const stepped0 = stepsBetweenWords(allRuns[0] ?? []);
  const edges0 = pageTextEdges(allRuns[0] ?? []);
  const band = foot ? footerBand(foot.band, stepped0, edges0, foot.numbered) : [];
  const headBand = head ? footerBand(head.band, stepped0, edges0, head.numbered) : [];
  const withFooter = (properties: SectionProperties | undefined): SectionProperties | undefined =>
    properties
      ? {
          ...properties,
          ...(band.length > 0
            ? { footers: [{ type: 'default' as const, relationshipId: FOOTER_PART }] }
            : {}),
          ...(headBand.length > 0
            ? { headers: [{ type: 'default' as const, relationshipId: HEADER_PART }] }
            : {}),
        }
      : properties;
  return {
    doc: buildFlowDoc(
      body,
      resources,
      withFooter(setUp(0, pages.length)),
      collectEmbeddedFonts(file, pages, losses),
      sections.map((s) => ({ ...s, properties: withFooter(s.properties) ?? s.properties })),
      band.length > 0 || headBand.length > 0
        ? new Map([
            ...(band.length > 0 ? ([[FOOTER_PART, band]] as const) : []),
            ...(headBand.length > 0 ? ([[HEADER_PART, headBand]] as const) : []),
          ])
        : undefined,
    ),
    losses: dedupeLosses(losses),
  };
}

// EP17 — the gutters of a page set in columns: the vertical bands the fewest
// lines cross.
//
// It used to be the widest band NO run crossed, which asked a page to be two
// columns and nothing else. Almost none are: comments.pdf is a conference paper
// — a full-width title, a full-width author block, then two columns of body —
// and no such band exists on it, because the title crosses everything. Read
// flat, its columns were joined line by line: "Abstract and is used for the
// application logic of browser-based productivity Dynamic languages such as
// JavaScript are more difficult to com-…".
//
// So the measure is how many runs cross each x. Inside a column that is every
// line of it; in the gutter it is only the handful of full-width lines, and the
// two counts are far enough apart to tell one from the other. Those full-width
// lines are then what cuts the page into BANDS (see `assignColumns`).
//
// There may be more than one. A page is not always two columns and a middle:
// chrome-text-selection-markedContent.pdf is an analyst's report — two columns
// of comment and a sidebar of figures down the right — and asked for the ONE
// best band it took the body's gutter and read the sidebar as part of the
// text, so the page opened with the guidance box from the foot of the margin.
interface Gutter {
  /** The middle of the band, which is what a run is placed left or right of. */
  readonly mid: number;
  /** The band itself — a run crossing all of it is a full-width line. */
  readonly from: number;
  readonly to: number;
}

function detectGutters(runs: ReadonlyArray<TextRun>, pageWidth: number): Array<Gutter> {
  if (runs.length < 30 || pageWidth <= 0) return [];
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const spans = runs.flatMap((r) => {
    const ink = runInk(r);
    return ink ? [ink] : [];
  });
  if (spans.length === 0) return [];
  const minX = Math.min(...spans.map((iv) => iv[0]));
  const maxX = Math.max(...spans.map((iv) => iv[1]));
  const span = maxX - minX;
  if (span < pageWidth * 0.5) return []; // text doesn't span enough of the page
  // First the strict reading: every band NO run crosses. It is exactly right
  // where it fires, and it fires on a page set in columns and nothing else —
  // including a sparse one, where the crossing count below is zero nearly
  // everywhere and says nothing.
  const sorted = [...spans].sort((a, b) => a[0] - b[0]);
  let curEnd = sorted[0]![1];
  const empty: Array<Gutter> = [];
  for (const [l, r] of sorted) {
    if (l - curEnd >= fontSize * 3) empty.push({ mid: (curEnd + l) / 2, from: curEnd, to: l });
    if (r > curEnd) curEnd = r;
  }
  // And then the question is asked a LINE at a time: at the gutter, most lines
  // have ink on both sides of it and cross none of it, while the few that do
  // cross are the full-width ones. A page set in one column has no such place —
  // every line crosses its middle.
  const rows = rowsOf(runs, fontSize).map((row) =>
    row
      .flatMap((r) => {
        const ink = runInk(r);
        return ink ? [ink] : [];
      })
      .sort((a, b) => a[0] - b[0]),
  );
  const voted: Array<Gutter> = [];
  for (const x of sample(minX + span * 0.1, minX + span * 0.9, 1)) {
    let columned = 0;
    let crossing = 0;
    for (const row of rows) {
      if (row.some(([l, r]) => l < x && r > x)) {
        crossing++;
        continue;
      }
      // Ink on both sides is not enough: what separates two columns is a GAP,
      // and a gap the width of a word space separates two words.
      const before = row.filter(([, r]) => r <= x);
      const after = row.filter(([l]) => l >= x);
      if (before.length === 0 || after.length === 0) continue;
      const gap = Math.min(...after.map(([l]) => l)) - Math.max(...before.map(([, r]) => r));
      if (gap >= fontSize * MIN_GUTTER_EM) columned++;
    }
    // Enough lines have to be split at the SAME x, or it is not a gutter: a
    // form's label-and-value rows have a wide gap on every line and it is in a
    // different place on each, so no single x splits many of them. And more
    // lines must be split here than reach across it — on a page set in one
    // column every line reaches across the middle.
    if (columned < MIN_COLUMNED_ROWS || crossing >= columned) continue;
    const last = voted[voted.length - 1];
    // Every x that answers is part of a band, and the band is the gutter.
    if (last && x - last.to <= 1.5)
      voted[voted.length - 1] = { ...last, to: x, mid: (last.from + x) / 2 };
    else voted.push({ mid: x, from: x, to: x });
  }
  // Both answers, together. Neither alone is the page: the empty band is exact
  // where it fires and silent where a title crosses it, and the vote needs a
  // dozen lines split at the same place, which the sidebar of a sparse page
  // never has. chrome-text-selection-markedContent.pdf is two columns of
  // comment and a margin of figures — the empty band separates the margin, the
  // vote separates the columns, and one without the other reads the sidebar as
  // part of the text.
  // The empty band is exact — its middle is a place no ink is — so the vote
  // only ADDS gutters, and never moves one the strict reading already found:
  // a band as wide as both answers has its middle wherever that falls, which
  // on this page was inside a word, and the line was then read straight across.
  const near = (band: Gutter): boolean =>
    empty.some((e) => band.mid > e.from - fontSize && band.mid < e.to + fontSize);
  return separating(
    [...empty, ...voted.filter((v) => !near(v))].sort((a, b) => a.mid - b.mid),
    spans,
  );
}

/**
 * The candidate gutters that actually separate something, left to right.
 *
 * A gutter with nothing on one side of it is a margin, and two of them with a
 * word between are one column and a stray. Each band is kept only where the
 * region since the last kept one holds a real share of the page's runs — and
 * the last one only if something follows it.
 *
 * @param bands The candidates, in order.
 * @param spans Every run's horizontal extent.
 * @returns The gutters worth splitting on.
 */
function separating(
  bands: ReadonlyArray<Gutter>,
  spans: ReadonlyArray<readonly [number, number]>,
): Array<Gutter> {
  const need = spans.length * MIN_COLUMN_SHARE;
  const out: Array<Gutter> = [];
  let from = -Infinity;
  for (const band of bands) {
    const inside = spans.filter(([l, r]) => (l + r) / 2 > from && (l + r) / 2 < band.mid).length;
    if (inside < need) continue;
    out.push(band);
    from = band.mid;
  }
  // The rightmost column has to hold something too.
  const tail = spans.filter(([l, r]) => (l + r) / 2 > from).length;
  if (out.length > 0 && tail < need) out.pop();
  return out;
}

/** How much of a page's text the narrowest column holds before it is a column. */
const MIN_COLUMN_SHARE = 0.08;

/**
 * The gutters MOST of the document's pages agree on.
 *
 * A document is set one way: §17.6.4 puts the column setup on the section, and
 * a paper does not change it from sheet to sheet. So a page that says nothing
 * on its own — one carrying a figure across both columns, with too few clean
 * lines left for the vote — can be asked the question the rest of the document
 * already answered.
 *
 * @param perPage Each page's own answer, in page order.
 * @returns The answer given by more pages than any other, or `undefined` where
 *          no two pages agree.
 */
function commonGutters(
  perPage: ReadonlyArray<ReadonlyArray<Gutter>>,
): ReadonlyArray<Gutter> | undefined {
  const seen = new Map<string, { gutters: ReadonlyArray<Gutter>; pages: number }>();
  for (const gutters of perPage) {
    if (gutters.length === 0) continue;
    // To the point: two pages set alike put their gutters within a point or two
    // of each other, not at the same fraction of a millimetre.
    const key = gutters.map((g) => Math.round(g.mid / 2)).join(',');
    const had = seen.get(key);
    if (had) had.pages++;
    else seen.set(key, { gutters, pages: 1 });
  }
  let best: { gutters: ReadonlyArray<Gutter>; pages: number } | undefined;
  for (const entry of seen.values()) if (!best || entry.pages > best.pages) best = entry;
  // One page is evidence enough: what protects the others is their own veto
  // (see {@link fitsGutters}), and a paper of two sheets has only one to give.
  return best?.gutters;
}

/**
 * Whether a page's own lines agree with gutters the document states.
 *
 * The document's answer is a suggestion, not a licence: a sheet whose lines run
 * straight through the gutter is set in one column whatever its neighbours do.
 * The bar is lower than the vote's own — the evidence from the other pages is
 * already in — but it is still the page that decides.
 *
 * @param runs    The page's runs.
 * @param gutters The document's gutters.
 */
function fitsGutters(runs: ReadonlyArray<TextRun>, gutters: ReadonlyArray<Gutter>): boolean {
  if (runs.length < 30) return false;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize).map((row) =>
    row
      .flatMap((r) => {
        const ink = runInk(r);
        return ink ? [ink] : [];
      })
      .sort((a, b) => a[0] - b[0]),
  );
  for (const gutter of gutters) {
    let columned = 0;
    let crossing = 0;
    for (const row of rows) {
      if (row.some(([l, r]) => l < gutter.mid && r > gutter.mid)) {
        crossing++;
        continue;
      }
      const before = row.filter(([, r]) => r <= gutter.mid);
      const after = row.filter(([l]) => l >= gutter.mid);
      if (before.length === 0 || after.length === 0) continue;
      const gap = Math.min(...after.map(([l]) => l)) - Math.max(...before.map(([, r]) => r));
      if (gap >= fontSize * MIN_GUTTER_EM) columned++;
    }
    // A page NO line crosses is a page the gutter fits, however few of its
    // lines happen to reach both sides of it: bug1997343.pdf's second sheet
    // sets a figure across the top and then a column at a time, so two of its
    // thirty-three rows have ink on both sides — and not one runs through.
    if (crossing === 0) continue;
    if (columned < MIN_SHARED_ROWS || crossing >= columned) return false;
  }
  return true;
}

/** How many lines a page must split at a gutter the DOCUMENT already states. */
const MIN_SHARED_ROWS = 3;

/** How many lines have to be split at the same place before the page is in columns. */
const MIN_COLUMNED_ROWS = 12;

/**
 * The page's runs grouped by baseline.
 *
 * Swept in order rather than bucketed: a superscript sits a few points above
 * the baseline it belongs to, and a bucket boundary between the two would leave
 * it a line of its own — bug1885505.pdf's author block came back with the
 * asterisks and daggers standing alone on five separate lines.
 */
function rowsOf(runs: ReadonlyArray<TextRun>, fontSize: number): Array<Array<TextRun>> {
  const tolerance = Math.max(fontSize * 0.6, 1);
  const rows: Array<Array<TextRun>> = [];
  let row: Array<TextRun> = [];
  let rowY = Number.POSITIVE_INFINITY;
  for (const run of [...runs].sort((a, b) => b.y - a.y)) {
    if (row.length > 0 && rowY - run.y > tolerance) {
      rows.push(row);
      row = [];
    }
    if (row.length === 0) rowY = run.y;
    row.push(run);
  }
  if (row.length > 0) rows.push(row);
  return rows;
}

/**
 * A run's INK — where its letters are, which a trailing space is not.
 *
 * A run carries the advance it stepped, and the space that ends a line of a
 * column is part of it: chrome-text-selection-markedContent.pdf sets two
 * columns 21 points apart and its left column's lines end ", " — five of those
 * points — so every measurement of the gutter came back a third short and the
 * page was read straight across.
 *
 * @param run The run.
 * @returns Its ink, or `undefined` for a run that is nothing but space.
 */
function runInk(run: TextRun): [number, number] | undefined {
  const from = Math.min(run.x, run.endX);
  const to = Math.max(run.endX, run.x + 1);
  const chars = [...run.text];
  if (chars.length === 0) return [from, to];
  let head = 0;
  while (head < chars.length && SPACE.test(chars[head]!)) head++;
  if (head === chars.length) return undefined;
  let tail = 0;
  while (tail < chars.length - head && SPACE.test(chars[chars.length - 1 - tail]!)) tail++;
  if (head === 0 && tail === 0) return [from, to];
  const step = (to - from) / chars.length;
  return [from + head * step, to - tail * step];
}

const SPACE = /\s/u;

/** Where the page's text starts and ends, ignoring what is only a space. */
function pageTextEdges(runs: ReadonlyArray<TextRun>): { left: number; right: number } | undefined {
  let left = Number.POSITIVE_INFINITY;
  let right = Number.NEGATIVE_INFINITY;
  for (const run of runs) {
    const ink = runInk(run);
    if (!ink) continue;
    left = Math.min(left, ink[0]);
    right = Math.max(right, ink[1]);
  }
  return right > left ? { left, right } : undefined;
}

/**
 * The measure a column is set across: from the gutter on its left to the one
 * on its right, and to the page's own text edge where there is none.
 *
 * A line that spans the page (see {@link SPANNING_COLUMN}) is set across all of
 * them, which is what makes a centred title centred.
 *
 * @param col     The column, or {@link SPANNING_COLUMN}.
 * @param gutters The page's gutters, left to right.
 * @param edges   Where the page's text starts and ends.
 */
function measureOf(
  col: number,
  gutters: ReadonlyArray<Gutter>,
  edges: { left: number; right: number } | undefined,
): { left: number; right: number } | undefined {
  if (!edges) return undefined;
  if (col === SPANNING_COLUMN || gutters.length === 0) return edges;
  return {
    left: col === 0 ? edges.left : (gutters[col - 1]?.to ?? edges.left),
    right: col >= gutters.length ? edges.right : (gutters[col]?.from ?? edges.right),
  };
}

/** The measure the lines themselves reach across, where the page states none. */
function measureOfLines(lines: ReadonlyArray<Line>): { left: number; right: number } | undefined {
  if (lines.length === 0) return undefined;
  return {
    left: Math.min(...lines.map((l) => l.x)),
    right: Math.max(...lines.map((l) => l.x + l.width)),
  };
}

/** The x's to measure at, from `a` to `b` inclusive. */
function sample(a: number, b: number, step: number): Array<number> {
  const out: Array<number> = [];
  for (let x = a; x <= b; x += Math.max(step, 0.5)) out.push(x);
  return out;
}

/**
 * EP17 — which column each run belongs to, and where the page's bands break.
 *
 * The decision is made a LINE at a time, not a run at a time: a title is drawn
 * as several runs and only one of them may reach across the gutter, so judging
 * each on its own would leave the rest of the title standing in a column.
 *
 * @param runs   The page's runs.
 * @param gutter The band from {@link detectGutter}.
 * @returns The column of each run (0, 1, or {@link SPANNING_COLUMN}) and the
 *          baselines of the full-width lines, which separate the bands.
 */
function assignColumns(
  runs: ReadonlyArray<TextRun>,
  gutters: ReadonlyArray<Gutter>,
): { columnOf: Map<TextRun, number>; breaks: Array<number> } {
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize);
  const columnOf = new Map<TextRun, number>();
  const breaks: Array<number> = [];
  const columnAt = (x: number): number => gutters.filter((g) => x >= g.mid).length;
  for (const row of rows) {
    // A space is not ink, and the space that ends a column's line stands in
    // the gutter (see {@link runInk}).
    const inked = row
      .flatMap((run) => {
        const ink = runInk(run);
        return ink ? [{ run, from: ink[0], to: ink[1] }] : [];
      })
      .sort((a, b) => a.from - b.from);
    if (inked.length === 0) continue;
    const leftmost = inked[0]!.from;
    const rightmost = Math.max(...inked.map((r) => r.to));
    // The gutters this row's ink lies across; the others it is on one side of.
    const straddled = gutters.filter((g) => leftmost < g.mid && rightmost > g.mid);
    if (straddled.length === 0) {
      const col = columnAt(leftmost);
      for (const run of row) columnOf.set(run, col);
      continue;
    }
    // Ink on both sides. Several columns, or one line reaching across? The
    // gutter is what tells them apart: a line set across the page may have a
    // WORD space over the middle, and a word space is nothing like a column
    // gap. comments.pdf centres its author block, and split on the word gaps
    // that happened to fall in the middle the names came apart into both
    // columns.
    const gapAt = (mid: number): number => {
      let cur = inked[0]!.to;
      let gap = 0;
      for (const { from, to } of inked.slice(1)) {
        if (from > cur && cur <= mid && from >= mid) gap = from - cur;
        cur = Math.max(cur, to);
      }
      return gap;
    };
    if (straddled.every((g) => gapAt(g.mid) >= fontSize * MIN_GUTTER_EM)) {
      for (const { run, from } of inked) columnOf.set(run, columnAt(from));
      continue;
    }
    for (const run of row) columnOf.set(run, SPANNING_COLUMN);
    breaks.push(Math.max(...row.map((r) => r.y)));
  }
  return { columnOf, breaks: breaks.sort((a, b) => b - a) };
}

/**
 * A line that spans the page belongs to no column: it stands between the bands
 * it separates.
 *
 * Which side of them? A spanning line is what BREAKS a band, and a band runs
 * from one break to the next, so the line is always at the FOOT of its own —
 * the columns of that band are the ones above it. Read ahead of them,
 * bug1997343.pdf's page number came out between the date and the abstract.
 */
const SPANNING_COLUMN = -1;

/** Where a column reads in its band: the spanning line at the foot of it. */
function columnOrder(col: number): number {
  return col === SPANNING_COLUMN ? Number.MAX_SAFE_INTEGER : col;
}

/**
 * How wide, in ems, the gap over the middle has to be for a line to be two
 * lines. A word space is a quarter of an em (see `SPACE_GAP_EM`) and a
 * justified one no more than half; a gutter is an em and more.
 *
 * It stood at one and a half, which is wider than some magazines set:
 * chrome-text-selection-markedContent.pdf puts fourteen and a half points
 * between columns of eleven-point type — 1.31 em — and every line of it was
 * read straight across, the left column's sentence running into the right
 * column's.
 */
const MIN_GUTTER_EM = 1;

/** Which band a mark at this height belongs to — how many breaks stand above it. */
function bandOf(breaks: ReadonlyArray<number>, top: number, epsilon: number): number {
  let n = 0;
  for (const y of breaks) if (y > top + epsilon) n++;
  return n;
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
function groupIntoLines(runs: ReadonlyArray<TextRun>, split = false, stepped = false): Array<Line> {
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
    if (!split) return [lineOf(ordered, c.y, fontSize, stepped)];
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
      lineOf(
        piece,
        piece[0]!.y,
        Math.max(...piece.map((r) => r.fontSizePt || 0)) || fontSize,
        stepped,
      ),
    );
  });
}

/**
 * Where a line's INK is, which is not how far its pen travelled.
 *
 * A run of spaces advances the pen and marks nothing. basicapi.pdf sets its
 * page number as thirty-one spaces and "page 1 / 3" in ONE run, reaching 635pt
 * across a 595pt sheet — and the measure taken off that line was wide enough
 * that the centred title in the same column no longer looked centred.
 *
 * The blanks are deducted at the face's own space width, which the run carries
 * (§9.4.4); where the face states none, at a quarter of the size.
 */
function inkSpan(runs: ReadonlyArray<TextRun>): { x: number; width: number } {
  const marked = runs.filter((r) => r.text.trim().length > 0);
  if (marked.length === 0) {
    const x = runs[0]!.x;
    return { x, width: runs[runs.length - 1]!.endX - x };
  }
  const first = marked[0]!;
  const last = marked[marked.length - 1]!;
  const space = (r: TextRun): number =>
    r.spaceWidthPt !== undefined && r.spaceWidthPt > 0
      ? r.spaceWidthPt
      : (r.fontSizePt || 10) * 0.25;
  const lead = (/^\s*/u.exec(first.text)?.[0].length ?? 0) * space(first);
  const trail = (/\s*$/u.exec(last.text)?.[0].length ?? 0) * space(last);
  const x = first.x + lead;
  return { x, width: Math.max(0, last.endX - trail - x) };
}

/** One run of runs, left to right on a shared baseline, as a {@link Line}. */
function lineOf(runs: ReadonlyArray<TextRun>, y: number, fontSize: number, stepped: boolean): Line {
  // A page may do both. A LaTeX document writes its prose with spaces in it and
  // sets its mathematics by stepping — TeX's thin space is a sixth of an em and
  // its medium one two ninths, both under the quarter a drawn-space page needs
  // — so bug1997343.pdf came back with "f(x) = sinx+cosx" where the file sets
  // "f(x) = sin x + cos x". A line with no space in it anywhere was stepped
  // across, whatever the rest of the page does.
  const steppedLine = stepped || (runs.length > 1 && !runs.some((r) => SPACE.test(r.text)));
  const ordered = lineSpans(runs, fontSize, steppedLine);
  // §9.4 — the runs came off the page in the order they were PAINTED, which is
  // left to right whatever the script. `logicalOrder` turned each run's own
  // letters back the right way round; the runs themselves are still in visual
  // order, and a line of them reads as the sentence backwards.
  // ArabicCIDTrueType.pdf's every line came out with its words in reverse.
  const spans = ordered.every((s) => s.text.trim() === '' || isRightToLeft(s.text))
    ? [...ordered].reverse()
    : ordered;
  const ink = inkSpan(runs);
  return {
    x: ink.x,
    width: ink.width,
    y,
    fontSize,
    ...(tabbed(runs, fontSize) ? { tabbed: true as const } : {}),
    text: spans
      .map((s) => s.text)
      .join('')
      .replace(/\s+/g, ' ')
      .trim(),
    spans,
  };
}

/**
 * Whether a TAB stands inside the line — a gap no word space could be.
 *
 * A line with one is a line the page SET OUT, not a line of prose: a contents
 * entry with its page number at the measure, a two-column list, a label and its
 * value. It does not run on into the line below it, and read as prose it does:
 * bug1997343.pdf's contents came back as "2 Document structures 1 2.1
 * Mathematics ............. 1", two entries in one line, where the file sets
 * one to a line.
 *
 * A leader says the same thing (see {@link carriesLeader}) but only where the
 * entry is dotted; a top-level entry is spaced, and nothing else marks it.
 *
 * @param runs     The line's runs.
 * @param fontSize The line's size.
 */
function tabbed(runs: ReadonlyArray<TextRun>, fontSize: number): boolean {
  const inked = runs
    .flatMap((r) => {
      const ink = runInk(r);
      return ink ? [ink] : [];
    })
    .sort((a, b) => a[0] - b[0]);
  let cur = Number.NEGATIVE_INFINITY;
  for (const [from, to] of inked) {
    if (cur > Number.NEGATIVE_INFINITY && from - cur >= fontSize * TAB_GAP_EM) return true;
    cur = Math.max(cur, to);
  }
  return false;
}

/**
 * How wide a gap has to be, in ems, before a word space could not have stood
 * there. Justification stretches a space to about half an em; two and a half is
 * a jump nothing but a tab makes.
 */
const TAB_GAP_EM = 2.5;

/**
 * The gap between two runs that means a WORD SPACE stood there.
 *
 * A page that draws its own spaces has already said where its words divide, and
 * a gap between two of its runs is a COLUMN or a placement — 160F-2019.pdf is
 * ruled into fields a quarter-inch apart and a generous threshold keeps them
 * apart. A page that draws none has said nothing, and every word boundary on it
 * is a gap: bigboundingbox.pdf steps 0.226 em between words and never writes a
 * space, so at a quarter em its every line ran together — "OrangeDemoInc.",
 * "Whenpayingbycheck,pleasecompletethispaymentadvice".
 *
 * The two want different thresholds, and the page says which it is (see
 * {@link stepsBetweenWords}). The tight one still clears the gaps a producer
 * leaves INSIDE a word when it splits one for kerning, which measure eight
 * hundredths of an em at their widest across this corpus.
 */
function spaceGap(prev: TextRun, fontSize: number, stepped: boolean): number {
  return (prev.fontSizePt || fontSize) * (stepped ? STEPPED_SPACE_EM : DRAWN_SPACE_EM);
}

/**
 * §17.3.1.25 — one character of a LEADER, the dotted rule that carries the eye
 * across a table of contents.
 *
 * A leader is drawn one character at a time with a step about as wide as a word
 * space, so every threshold that tells a space from a kern says "space" between
 * every dot. bug886717.pdf's contents came back as
 * "Abstract . . . . . . . . . . . . 3", four times as long as the page sets it,
 * and its forty entries spilled onto a second page. What tells a leader from
 * words is that it is the SAME character over and over.
 */
function isLeader(text: string): boolean {
  return text.length === 1 && LEADER_CHARS.has(text);
}

const LEADER_CHARS = new Set(['.', '\u00b7', '_', '-', '\u2010', '\u2013']);

/** A page that writes its own spaces: only a wide gap means anything more. */
const DRAWN_SPACE_EM = 0.25;

/** A page that writes none: the step between its words is all there is. */
const STEPPED_SPACE_EM = 0.12;

/**
 * Whether this page STEPS between its words rather than writing spaces.
 *
 * Counted rather than guessed: bigboundingbox.pdf writes a space in one run in
 * a hundred, TAMReview.pdf in a third of them, and no page does a little of
 * both. A page with almost no text says nothing either way and keeps the
 * cautious reading.
 */
function stepsBetweenWords(runs: ReadonlyArray<TextRun>): boolean {
  if (runs.length < 8) return false;
  const drawn = runs.filter((r) => /\s/u.test(r.text)).length;
  return drawn / runs.length < 0.05;
}

// A line's runs as spans, inserting a (link-free) space where a horizontal gap
// suggests one.
function lineSpans(
  runs: ReadonlyArray<TextRun>,
  fontSize: number,
  stepped: boolean,
): Array<TextSpan> {
  const spans: Array<TextSpan> = [];
  // §17.3.2.42 — the line's OWN baseline, which a script stands off. Taken from
  // the runs set at the line's size: the marks are the ones that moved.
  const body = runs.filter((r) => (r.fontSizePt || fontSize) > fontSize * SCRIPT_SIZE);
  const baseline = median((body.length > 0 ? body : runs).map((r) => r.y));
  let prev: TextRun | undefined;
  for (const run of runs) {
    if (
      prev !== undefined &&
      run.x - prev.endX > spaceGap(prev, fontSize, stepped) &&
      !(isLeader(prev.text) && prev.text === run.text)
    ) {
      spans.push({ text: ' ' });
    }
    // §9.3.1/§8.6.8 — the size and colour the page showed the glyphs at. The
    // tagged path has carried these since it learned to; this one never did, so
    // every line it read came back at the 11pt default in black. Placed, that
    // is not a wrong shade but a wrong SHAPE: 160F-2019.pdf's footnotes are set
    // in 7pt nine and a half apart, and drawn at eleven they climbed over each
    // other.
    spans.push({
      text: run.text.replaceAll(UNMAPPED, ''),
      sizePt: run.fontSizePt,
      ...(run.colorHex !== '000000' ? { colorHex: run.colorHex } : {}),
      ...(run.fontName !== undefined ? { fontName: run.fontName } : {}),
      ...(run.outlineHex !== undefined
        ? { outline: { colorHex: run.outlineHex, widthPt: pt(run.outlineWidthPt ?? 1) } }
        : {}),
      ...(run.bold ? { bold: true } : {}),
      ...(run.italic ? { italic: true } : {}),
      ...script(run, baseline, fontSize),
      ...(run.markup !== undefined ? { markup: run.markup } : {}),
      ...(run.href !== undefined ? { href: run.href } : {}),
    });
    prev = run;
  }
  return spans;
}

/**
 * §17.3.2.42 — a run set off the line's baseline, and smaller, as the script it
 * is.
 *
 * A PDF states no such property: an exponent is a smaller face set a little
 * higher, and an index a smaller face set a little lower. Read flat, the whole
 * of mathematics comes back on one line — bug1997343.pdf sets `n^p = n mod p`
 * and we read "np", and every prime on the page landed beside its letter
 * instead of over it.
 *
 * The size the span keeps is the LINE's, not the mark's: a document states the
 * nominal size and the layout shrinks a script, so the drawn seven points under
 * a superscript would come out at five.
 *
 * @param run      The run.
 * @param baseline The line's own baseline.
 * @param fontSize The line's size.
 */
function script(
  run: TextRun,
  baseline: number,
  fontSize: number,
): { script?: 'superscript' | 'subscript'; sizePt?: number } {
  const size = run.fontSizePt || fontSize;
  if (size > fontSize * SCRIPT_SIZE) return {};
  const rise = run.y - baseline;
  if (rise > fontSize * SCRIPT_RISE) return { script: 'superscript', sizePt: fontSize };
  if (rise < -fontSize * SCRIPT_DROP) return { script: 'subscript', sizePt: fontSize };
  return {};
}

/** How much smaller than its line a run must be set to be a script of it. */
const SCRIPT_SIZE = 0.85;

/** How far above the baseline a superscript stands, and below it a subscript. */
const SCRIPT_RISE = 0.15;
const SCRIPT_DROP = 0.08;

/**
 * §17.3.1.25 — whether a line carries a LEADER, which makes it an entry in a
 * directory rather than a line of prose.
 *
 * A contents line runs the full measure — the dots are there to make it — so
 * the ragged-edge test that ends every other paragraph never fires on one, and
 * bug886717.pdf's forty entries came back as a single reflowing paragraph.
 */
function carriesLeader(line: Line): boolean {
  return LEADER_RUN.test(line.text);
}

/** Three of the same leader character in a row is a leader and not punctuation. */
const LEADER_RUN = /([.\u00b7_])\1{2,}/u;

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
  indentLeft?: number;
  indentFirstLine?: number;
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
      (prev !== undefined &&
        (endedParagraph(prev, line, column) || carriesLeader(prev) || prev.tabbed === true))
    ) {
      groups.push([]);
      gaps.push(prev === undefined ? 0 : gap);
    }
    groups[groups.length - 1]!.push(line);
    prev = line;
  }
  // §17.3.1.12 — where the COLUMN's own text begins, which is what an indented
  // paragraph is indented from. Not the leftmost line: a marginal note set
  // outside the measure is not the measure, and bug1997343.pdf puts one forty-
  // five points left of its body. Not the median either — a page can be half
  // list — but the low end of the run of line starts, which the body holds
  // whatever else is on the page.
  const starts = lines.map((l) => l.x).sort((a, b) => a - b);
  const columnLeft = starts[Math.floor(starts.length * COLUMN_LEFT_QUANTILE)] ?? 0;
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
      spans: joinLines(g),
      fontSize,
      top: first.y,
      ...(spacingBefore !== undefined ? { spacingBefore } : {}),
      ...alignmentOf(g, column),
      ...indentOf(g, columnLeft, alignmentOf(g, column).alignment),
    };
  });
}

/**
 * A paragraph's lines as one run of spans, joined the way the page broke them.
 *
 * A line that ends in a hyphen was broken THERE, and the break is not part of
 * the text: re-set at another measure the word has to come back together. Which
 * hyphen decides what is left of it — U+00AD is the discretionary one, put in
 * to mark a place a word MAY break, and it goes with the break; a plain hyphen
 * belongs to the word ("two-" and "column" are "two-column") and stays.
 *
 * Read as prose with a space between every line, bug1997343.pdf came back
 * "typical two-column docu ment incorporating tables, figures and mathemat
 * ics" — the soft hyphens dropped by the page and a space in their place.
 *
 * @param lines The paragraph's lines, in order.
 */
function joinLines(lines: ReadonlyArray<Line>): Array<TextSpan> {
  const out: Array<TextSpan> = [];
  lines.forEach((line, i) => {
    if (i > 0) {
      const prev = out[out.length - 1];
      const ends = prev?.text ?? '';
      const soft = ends.endsWith(SOFT_HYPHEN);
      const hard = HYPHENS.has(ends.slice(-1));
      if (soft && prev) out[out.length - 1] = { ...prev, text: ends.slice(0, -1) };
      else if (!hard) out.push({ text: ' ' });
    }
    out.push(...line.spans);
  });
  return out;
}

/** §17.3.3.29 — the hyphen that is a PLACE a word may break, not a hyphen. */
const SOFT_HYPHEN = '\u00ad';

/** The hyphens that belong to the word they end. */
const HYPHENS = new Set(['-', '\u2010', '\u2011']);

/** How many lines' worth of indent still reads as a first line, not a placement. */
const INDENT_LINES = 3;

/** Where in the run of line starts the column's own left edge is looked for. */
const COLUMN_LEFT_QUANTILE = 0.15;

/**
 * §17.3.1.12 `w:ind` — how far a paragraph is set in from its column, and where
 * its first line begins.
 *
 * A PDF states neither: every line is placed absolutely, and read flat every
 * paragraph came back against the left edge. That is most of what a list looks
 * like — bug1997343.pdf sets "• They may be unordered bullet lists" ten points
 * in and its nested "1. lists may also be nested" twenty more, and we set all
 * of them flush left — and it is the whole of a first-line indent, which is how
 * most of the world's prose marks a new paragraph.
 *
 * The first line is measured against the REST of the paragraph, which is what
 * `indentFirstLine` means: positive is a first line set in (a new paragraph),
 * negative a hanging one (a list, its marker standing out to the left).
 *
 * A paragraph that is centred or set to the right is placed, not indented, and
 * keeps neither.
 *
 * @param lines      The paragraph's lines.
 * @param columnLeft Where the column's own text begins.
 * @param alignment  What {@link alignmentOf} made of it.
 */
function indentOf(
  lines: ReadonlyArray<Line>,
  columnLeft: number,
  alignment: 'center' | 'right' | undefined,
): { indentLeft?: number; indentFirstLine?: number } {
  if (alignment !== undefined || lines.length === 0) return {};
  const size = Math.max(...lines.map((l) => l.fontSize)) || 10;
  const rest = lines.slice(1);
  const body = rest.length > 0 ? Math.min(...rest.map((l) => l.x)) : lines[0]!.x;
  const left = body - columnLeft;
  const first = lines[0]!.x - body;
  const enough = size * INDENT_EM;
  return {
    ...(Math.abs(left) >= enough ? { indentLeft: left } : {}),
    ...(Math.abs(first) >= enough ? { indentFirstLine: first } : {}),
  };
}

/**
 * How far, in ems, a paragraph has to be set in before it is indented rather
 * than merely started. Half an em clears the rounding a producer leaves at the
 * head of a line and is well under the smallest indent anybody sets.
 */
const INDENT_EM = 0.5;

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
  const step = Math.max(prev.fontSize, 4);
  const shift = next.x - prev.x;
  // A line that begins LEFT of the one before it, or a whole measure to the
  // right of it, is not part of the same setting — the block is placed.
  if (shift < -step || shift > step * INDENT_LINES) return false;
  // A modest indent to the right is the oldest mark in typography for a new
  // paragraph, and it used to CANCEL the test below: bug1997343.pdf sets
  // "…figures and mathematics." and then indents "Apart from two commands at
  // the start…", and the two came back as one paragraph. It is read together
  // with the short line that precedes it — a full line followed by an indented
  // one is a list item and its own continuation, not two paragraphs.
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
  least?: number,
): { alignment?: 'center' | 'right' } {
  if (!column || lines.length === 0) return {};
  const width = column.right - column.left;
  if (!(width > 0)) return {};
  const insets = lines.map((l) => ({
    lead: l.x - column.left,
    trail: column.right - (l.x + l.width),
  }));
  // A tenth of the measure is the smallest inset worth calling a placement:
  // below it every ragged line would read as placed. A caller with no rag to
  // guard against — a CELL holds one line — may ask for less.
  const meaningful = Math.min(width * 0.1, least ?? width);
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
  // Flush right: every line ends at the measure, at least one starts well
  // inside it, and they start at DIFFERENT places — a block set to the right is
  // ragged on its left, and one that is merely indented is not. A justified
  // paragraph fails the first test on its last line, which is the only place
  // the two differ at all.
  const leads = insets.map((i) => i.lead);
  if (
    insets.every((i) => i.trail <= even) &&
    Math.max(...leads) >= meaningful &&
    Math.max(...leads) - Math.min(...leads) > even
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

/**
 * §17.6.13 — the running foot a document repeats at the bottom of its pages.
 *
 * A page number, a title, a URL: it stands below the text block, in the margin,
 * and it is not part of what the page SAYS. Read as body it goes wherever the
 * reflow puts it — bug1997343.pdf's "1" came out on a sheet of its own between
 * the two the paper has, and TAMReview.pdf's "Sprouts — http://…" landed in the
 * middle of the abstract.
 *
 * What makes it a running foot is that it RUNS: the same place, page after
 * page, cut off from the text above it by more than a line of white. One page
 * proves nothing, so two are asked for.
 *
 * @param pageRuns Each page's runs, placed.
 * @param shown    Each page's shown geometry.
 * @returns The runs to lift off each page and the band to put them in, or
 *          `undefined` where the document repeats nothing.
 */
function runningFoot(
  pageRuns: ReadonlyArray<ReadonlyArray<TextRun>>,
  shown: ReadonlyArray<{ height: number }>,
  where: 'head' | 'foot',
):
  | {
      lift: ReadonlyArray<ReadonlySet<TextRun>>;
      band: ReadonlyArray<TextRun>;
      /** Whether a number in it is the PAGE's number rather than part of the text. */
      numbered: boolean;
    }
  | undefined {
  const feet = pageRuns.map((runs, i) => edgeLine(runs, shown[i]?.height ?? 0, where));
  const found = feet.filter((f) => f !== undefined);
  if (found.length < 2 || found.length < pageRuns.length * FOOT_SHARE) return undefined;
  // The same place on every page: a foot that wanders is a last paragraph.
  const ys = found.map((f) => f.y);
  const mid = median(ys);
  if (ys.some((y) => Math.abs(y - mid) > FOOT_DRIFT)) return undefined;
  // Whether the foot says something DIFFERENT on each page, which is what a
  // page number is. ZapfDingbats.pdf signs every sheet "© RenderX 2000", and
  // read as a number the year came out as the page: "© RenderX 1".
  const texts = found.map((f) =>
    f.runs
      .map((r) => r.text)
      .join('')
      .trim(),
  );
  return {
    lift: feet.map((f) => new Set(f?.runs ?? [])),
    band: found[0]!.runs,
    numbered: new Set(texts).size > 1,
  };
}

/** How many of a document's pages must carry the foot before it is running. */
const FOOT_SHARE = 0.6;

/** How far, in points, a running foot may drift from page to page. */
const FOOT_DRIFT = 4;

/**
 * The lines at the very bottom of a page, where they stand ALONE below the text
 * block.
 *
 * Alone means the white above them is more than the page's own leading — a last
 * paragraph is a line's gap from the one before it, a running foot is several.
 *
 * A foot need not be ONE line. ZapfDingbats.pdf signs each page twice, the
 * publisher's line and the suite's title on one baseline and the build stamp
 * thirty points below it, and taking only the bottom line left the other in the
 * body: after a table that fills the sheet it had nowhere to go but a page of
 * its own, and a two-page document came out as four.
 */
function edgeLine(
  runs: ReadonlyArray<TextRun>,
  pageHeight: number,
  where: 'head' | 'foot',
): { y: number; runs: ReadonlyArray<TextRun> } | undefined {
  if (runs.length < 4 || pageHeight <= 0) return undefined;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize);
  if (rows.length < 4) return undefined;
  // `rowsOf` runs down the page, so the foot grows upward from its last row and
  // the head downward from its first; the gap is to the row on the text's side
  // of the group either way.
  const at = (i: number): ReadonlyArray<TextRun> =>
    where === 'foot' ? rows[rows.length - 1 - i]! : rows[i]!;
  const edgeY = (row: ReadonlyArray<TextRun>): number =>
    where === 'foot' ? Math.max(...row.map((r) => r.y)) : Math.min(...row.map((r) => r.y));
  const inside = (row: ReadonlyArray<TextRun>): number =>
    where === 'foot' ? Math.min(...row.map((r) => r.y)) : Math.max(...row.map((r) => r.y));
  const y = edgeY(at(0));
  let found: ReadonlyArray<TextRun> | undefined;
  let group: Array<TextRun> = [];
  // Only the foot grows: the lines that OPEN a page are far more often the
  // text itself, and asked for a head of up to three ZapfDingbats.pdf gave up
  // its red note, its running head and the first line of its title.
  const most = where === 'foot' ? FOOT_ROWS : 1;
  for (let i = 0; i < Math.min(most, rows.length - 2); i++) {
    const row = at(i);
    // In the margin, not in the text: an eighth of the sheet at its own end.
    const edge = edgeY(row);
    if (where === 'foot' ? edge > pageHeight * FOOT_BAND : edge < pageHeight * (1 - FOOT_BAND)) {
      break;
    }
    group = [...group, ...row];
    const next = at(i + 1);
    const gap = where === 'foot' ? inside(next) - edgeY(row) : edgeY(row) - inside(next);
    if (gap < fontSize * FOOT_GAP_EM) continue;
    const text = group
      .map((r) => r.text)
      .join('')
      .trim();
    // …and short lines at that.
    if (text.length > 0 && text.length <= FOOT_CHARS * (i + 1)) found = group;
  }
  return found ? { y, runs: found } : undefined;
}

/** How much white, in ems, stands between the text block and a running foot. */
const FOOT_GAP_EM = 2;

/** How many lines a running foot may hold before it is a paragraph. */
const FOOT_ROWS = 3;

/** How far up the sheet a running foot may sit. */
const FOOT_BAND = 0.12;

/** A running foot is a line, not a paragraph. */
const FOOT_CHARS = 90;

/**
 * §17.16.5.35 — the number in a run, made the field it stands for.
 *
 * A foot reads "1" or "Chapter 3 — 47", and the number is not the text: it is
 * the number of the page it is drawn on, which is what lets ONE band serve
 * every page. The run is cut around it and the middle becomes the field.
 *
 * @param run The footer's run.
 * @returns The run, or the two or three it is cut into.
 */
function pageNumbered(run: Run): Array<Run> {
  const found = /(^|\s)(\d{1,4})(\s|$)/u.exec(run.text);
  if (!found || run.field !== undefined) return [run];
  const at = found.index + found[1]!.length;
  const number = found[2]!;
  const before = run.text.slice(0, at);
  const after = run.text.slice(at + number.length);
  return [
    ...(before === '' ? [] : [{ ...run, text: before }]),
    { ...run, text: number, field: 'PAGE' as const },
    ...(after === '' ? [] : [{ ...run, text: after }]),
  ];
}

/**
 * The band itself: the foot's own line, with the number in it made a field.
 *
 * §17.16.5.35 — a page number is not the text "1"; it is the number of the page
 * it is drawn on, which is why the same band serves every page.
 *
 * A foot is written in REGIONS, the way a spreadsheet's is: something at the
 * left of the sheet, something at the middle, something against the far edge.
 * ZapfDingbats.pdf signs each page "© RenderX 2000" at the left and "XSL
 * Formatting Objects Test Suite" at the right, and the two hundred points
 * between them came back as one word space, the two texts crowding each other
 * at the left. They stand on TAB STOPS instead (§17.3.1.38), which is what the
 * two hundred points are.
 *
 * @param runs     The foot's runs.
 * @param stepped  Whether the page steps between its words.
 * @param measure  The measure it was set across, for its alignment.
 * @param numbered Whether a number in it is the page's own.
 */
function footerBand(
  runs: ReadonlyArray<TextRun>,
  stepped: boolean,
  measure: { left: number; right: number } | undefined,
  numbered: boolean,
): Array<BodyElement> {
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const lines = rowsOf(runs, fontSize)
    .map((row) => bandLine(row, fontSize, stepped, measure))
    .filter((line) => line !== undefined);
  // A line apiece, in the order the page shows them: ZapfDingbats.pdf signs
  // each sheet with the publisher's line and the build stamp under it.
  return lines.map(({ spans, properties }) => {
    const el = paragraphFromRuns(spans, undefined, properties);
    if (el.kind !== 'paragraph') return el;
    return {
      kind: 'paragraph',
      paragraph: {
        ...el.paragraph,
        runs: numbered ? el.paragraph.runs.flatMap((run) => pageNumbered(run)) : el.paragraph.runs,
      },
    };
  });
}

/**
 * One line of a running head or foot: its spans, and the placement that puts
 * them where the page had them.
 *
 * A wide gap inside such a line is not a word space, it is the space BETWEEN
 * REGIONS — and each region after the first stands on a tab stop, at the middle
 * of the band or against its far edge, whichever it was written at.
 *
 * @param row      The line's runs.
 * @param fontSize The band's size, for the runs that state none.
 * @param stepped  Whether the page steps between its words.
 * @param measure  The measure the band was set across.
 * @returns The line, or `undefined` where it holds no text.
 */
function bandLine(
  row: ReadonlyArray<TextRun>,
  fontSize: number,
  stepped: boolean,
  measure: { left: number; right: number } | undefined,
): { spans: ReadonlyArray<TextSpan>; properties: ParagraphProperties } | undefined {
  const ordered = [...row].sort((a, b) => a.x - b.x);
  const pieces: Array<Array<TextRun>> = [[]];
  for (const run of ordered) {
    const last = pieces[pieces.length - 1]!;
    const prev = last[last.length - 1];
    if (prev && run.x - prev.endX > (run.fontSizePt || fontSize) * BAND_REGION_EM) pieces.push([]);
    pieces[pieces.length - 1]!.push(run);
  }
  const y = Math.max(...ordered.map((r) => r.y));
  const lines = pieces
    .filter((piece) => piece.length > 0)
    .map((piece) =>
      lineOf(piece, y, Math.max(...piece.map((r) => r.fontSizePt || fontSize)), stepped),
    )
    .filter((line) => line.text.length > 0);
  if (lines.length === 0) return undefined;
  // One region is a line of its own, placed the way the page placed it.
  if (lines.length === 1 || lines.length > BAND_REGIONS || !measure) {
    const { alignment } = alignmentOf(lines.slice(0, 1), measure);
    const spans = lines.flatMap((line, i) =>
      i === 0 ? line.spans : [{ text: ' ' }, ...line.spans],
    );
    return { spans, properties: alignment ? { alignment } : {} };
  }
  const width = measure.right - measure.left;
  const spans: Array<TextSpan> = [...lines[0]!.spans];
  const tabs: Array<TabStop> = [];
  for (const line of lines.slice(1)) {
    // Against the far edge, or somewhere in the middle: which one it is is
    // where the page put it.
    const flush = measure.right - (line.x + line.width) <= width * BAND_RIGHT_SHARE;
    tabs.push({
      positionPt: pt(0),
      relativeTo: flush ? 'right' : 'center',
      alignment: flush ? 'right' : 'center',
    });
    spans.push({ text: '\t' }, ...line.spans);
  }
  return { spans, properties: { tabs } };
}

/** A gap this wide, in ems, stands BETWEEN the regions of a head or foot. */
const BAND_REGION_EM = 4;

/** How many regions a band line may hold — left, centre, right. */
const BAND_REGIONS = 3;

/** How near the far edge a region has to end to be set against it. */
const BAND_RIGHT_SHARE = 0.05;

/**
 * Whether the page is RULED into columns rather than SET in them.
 *
 * A page of two columns and a page of two columns of a table look alike from
 * here: both have gutters, and both put their lines on one baseline grid — a
 * paper's columns are set on the same grid as a matter of course, so "the rows
 * line up" says nothing. What separates them is the CELL: a line of prose fills
 * its measure and a cell does not. Across this corpus a paper's lines cover 84
 * to 95 per cent of their column, and ZapfDingbats.pdf's cells cover 45.
 *
 * Read by column, its five hundred entries came back one column at a time with
 * every row torn into three — "1 a17", "[x2711]", "2 a18" — over five pages.
 * Read by ROW, which is what a ruled page says, each entry is a line of its own.
 *
 * @param runs    The page's runs.
 * @param gutters The page's gutters.
 * @param edges   Where the page's text starts and ends.
 */
function looksRuled(
  runs: ReadonlyArray<TextRun>,
  gutters: ReadonlyArray<Gutter>,
  edges: { left: number; right: number } | undefined,
): boolean {
  // One gutter says nothing: a two-column index is short entries in two
  // columns, exactly like a table of two, and read across it interleaves two
  // lists that have nothing to do with each other — freeculture.pdf's index is
  // that page. Two gutters and more is a ruling.
  if (gutters.length < MIN_RULED_GUTTERS || !edges) return false;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize).filter((row) => row.length > 0);
  if (rows.length < MIN_TABLE_ROWS) return false;
  const bounds = columnBounds(runs, gutters, edges);
  const regions = bounds.slice(0, -1).map((lo, i) => [lo, bounds[i + 1]!] as const);
  const regionOf = (x: number): number => {
    for (let i = regions.length - 1; i >= 0; i--) if (x >= regions[i]![0]) return i;
    return 0;
  };
  const fills: Array<number> = [];
  let aligned = 0;
  for (const row of rows) {
    const byRegion = regions.map((): Array<TextRun> => []);
    for (const run of row) byRegion[regionOf(run.x)]!.push(run);
    byRegion.forEach((inRegion, i) => {
      if (inRegion.length === 0) return;
      const from = Math.min(...inRegion.map((r) => r.x));
      const to = Math.max(...inRegion.map((r) => r.endX));
      const width = regions[i]![1] - regions[i]![0];
      if (width > 0) fills.push((to - from) / width);
    });
    if (byRegion.every((inRegion) => inRegion.length > 0)) aligned++;
  }
  if (aligned < rows.length * TABLE_ALIGNED_SHARE) return false;
  fills.sort((a, b) => a - b);
  return (fills[Math.floor(fills.length / 2)] ?? 1) <= TABLE_CELL_FILL;
}

/**
 * §17.4.38 — the page's rows as the TABLE they are.
 *
 * The regions between the gutters are the columns and the rows are the rows;
 * a cell is what one row leaves in one region, and an empty cell is empty. The
 * grid is measured, so the columns come out where the page put them.
 *
 * The lines the page hangs ABOVE its ruling come back as themselves. A line
 * that crosses every column is not a row of the table — ZapfDingbats.pdf heads
 * each sheet with two red lines of provenance that run wider than the frame
 * drawn under them, and squeezed into a cell they wrapped and cost the sheet a
 * row.
 *
 * @param runs    The page's runs.
 * @param gutters The page's gutters.
 * @param edges   Where the page's text starts and ends.
 * @param stepped Whether the page steps between its words.
 * @returns The blocks, the table among them, each with where it stands.
 */
function tableFrom(
  runs: ReadonlyArray<TextRun>,
  gutters: ReadonlyArray<Gutter>,
  edges: { left: number; right: number },
  stepped: boolean,
): Array<{ el: BodyElement; top: number }> | undefined {
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const all = rowsOf(runs, fontSize).filter((row) => row.length > 0);
  if (all.length === 0) return undefined;
  const bounds = columnBounds(runs, gutters, edges);
  const regionsAt = (
    from: ReadonlyArray<number>,
  ): { regions: ReadonlyArray<readonly [number, number]>; of: (x: number) => number } => {
    const regions = from.slice(0, -1).map((lo, i) => [lo, from[i + 1]!] as const);
    return {
      regions,
      of: (x: number): number => {
        for (let i = regions.length - 1; i >= 0; i--) if (x >= regions[i]![0]) return i;
        return 0;
      },
    };
  };
  const first = regionsAt(bounds);
  // A line the page hangs above its ruling is not a row of it: read as one it
  // is cut to the table's measure, and ZapfDingbats.pdf's red lines of
  // provenance — which run wider than the frame beneath them — wrapped and cost
  // each sheet a row. Only at the TOP: a line across the middle of a table is a
  // heading INSIDE it, and it belongs to the ruling.
  const oneWideLine = (row: ReadonlyArray<TextRun>): boolean => {
    const to = spanEnd(row, 0, first.regions.length, first.of);
    // It crosses a boundary, and nothing else stands in that row: a head with a
    // page number at the far end is two cells and belongs to the ruling.
    return to > 0 && row.every((run) => first.of(run.x) <= to);
  };
  let start = 0;
  while (start < all.length - 1 && oneWideLine(all[start]!)) start++;
  const above = all.slice(0, start);
  const rows = all.slice(start);
  // The first column begins where its own CELLS begin, not where the page's
  // widest line does. ZapfDingbats.pdf sets its running head sixteen points
  // left of the table under it, and started there the whole sheet — three
  // groups, five hundred entries — stood that far left of where the file has it.
  const own = rows
    .filter((row) => spanEnd(row, 0, first.regions.length, first.of) === 0)
    .flatMap((row) => row.filter((run) => first.of(run.x) === 0))
    .map((run) => runInk(run)?.[0])
    .filter((x): x is number => x !== undefined);
  const left = own.length > 0 ? Math.min(...own) : bounds[0]!;
  const { regions, of: regionOf } = regionsAt([left, ...bounds.slice(1)]);
  // Where each column's cells USUALLY start, which is the line a placed cell is
  // placed against. Measured to the column's edge instead, an ordinary line
  // that happens to end near the far edge reads as centred — the lead being
  // only the width of whatever hangs left of the column.
  const flush = regions.map((region, i) => {
    const starts: Array<number> = [];
    for (const row of rows) {
      const at = row
        .filter((run) => regionOf(run.x) === i)
        .map((run) => runInk(run)?.[0])
        .filter((x): x is number => x !== undefined);
      if (at.length > 0) starts.push(Math.min(...at));
    }
    if (starts.length === 0) return region[0];
    // The near-leftmost, not the leftmost: one line hanging left of the column
    // — a head, a heading — would otherwise stand for all of them.
    starts.sort((a, b) => a - b);
    return starts[Math.floor(starts.length * COLUMN_LEFT_QUANTILE)] ?? region[0];
  });
  const table: Table = {
    // The grid is MEASURED — each column is as wide as the band the page drew
    // it in — so a cell's padding would be width the page never spent. Left at
    // the usual eighth of an inch a side, the five columns of ZapfDingbats.pdf
    // came to fifty points more than the sheet holds, and the whole table slid
    // out of the frame drawn around it.
    properties: {
      defaultCellMargins: { left: pt(0), right: pt(0) },
      layout: 'fixed',
      // …and it stands in from the margin by as much as it stands in from the
      // page's text, so the lines that hang to its left still do.
      ...(left > edges.left ? { indentPt: pt(left - edges.left) } : {}),
    },
    grid: regions.map((r) => pt(Math.max(r[1] - r[0], 1))),
    rows: rows.map((row, r) => {
      const byRegion = regions.map((): Array<TextRun> => []);
      for (const run of row) byRegion[regionOf(run.x)]!.push(run);
      const y = Math.max(...row.map((r2) => r2.y));
      const cells: Array<TableCell> = [];
      // A line that RUNS ACROSS the columns is one cell as wide as it is.
      // ZapfDingbats.pdf is five hundred entries in three groups, each group a
      // pair of narrow bands — and the prose above them runs across all six.
      // Cut at the band edges it wrapped where the page never did, and every
      // wrapped line cost the groups beside it an entry: two pages came out as
      // four, and the frame and the title's grey panel went with them.
      for (let i = 0; i < regions.length; ) {
        const to = spanEnd(row, i, regions.length, regionOf);
        const inSpan = byRegion.slice(i, to + 1).flat();
        const width = to - i + 1;
        const size =
          inSpan.length > 0 ? Math.max(...inSpan.map((r) => r.fontSizePt || fontSize)) : fontSize;
        const line = inSpan.length > 0 ? lineOf(inSpan, y, size, stepped) : undefined;
        // A cell's line stands where the page stood it. A cell holds no rag to
        // guard against, so half an em clear on BOTH sides is placement and not
        // an accident: ZapfDingbats.pdf centres its title over the first group,
        // inside the grey panel drawn behind it, and set flush left it came out
        // of that panel at the wrong end.
        // Placed in its cell, and standing in from where this column's lines
        // start: a line that merely reaches the far edge of a wide column is
        // not centred, and one that hangs left of its column — the running head
        // over ZapfDingbats.pdf's first group — is not centred either.
        const placed =
          line && line.x - flush[i]! >= size / 4
            ? alignmentOf([line], { left: regions[i]![0], right: regions[to]![1] }, size / 2)
            : {};
        cells.push({
          properties: width > 1 ? { colSpan: width } : {},
          content:
            line === undefined
              ? [{ kind: 'paragraph' as const, paragraph: { properties: {}, runs: [] } }]
              : [
                  paragraphFromRuns(
                    line.spans,
                    undefined,
                    placed.alignment ? { alignment: placed.alignment } : {},
                  ),
                ],
        });
        i = to + 1;
      }
      // The row stands as far from the next as the page put it. A row laid out
      // by the height of its own text closes up wherever the page left air —
      // ZapfDingbats.pdf keeps thirty points around its title, and set solid
      // the title rose out of the grey panel drawn behind it. `atLeast`, so a
      // cell that needs more than the page gave it still gets it.
      const next = rows[r + 1];
      const pitch = next ? y - Math.max(...next.map((r2) => r2.y)) : 0;
      return {
        properties: pitch > 0 ? { height: pt(pitch), heightRule: 'atLeast' as const } : {},
        cells,
      };
    }),
  };
  return [
    ...above.map((row) => {
      const y = Math.max(...row.map((r) => r.y));
      const size = Math.max(...row.map((r) => r.fontSizePt || fontSize));
      return {
        el: paragraphFromRuns(lineOf(row, y, size, stepped).spans),
        top: y,
      };
    }),
    { el: { kind: 'table' as const, table }, top: Math.max(...rows[0]!.map((r) => r.y)) },
  ];
}

/**
 * Where one column of a ruled page ends and the next begins.
 *
 * The gutter is a BAND, and its middle is only a guess at the line inside it: a
 * gutter is where the fewest lines cross, not where none do (see
 * `detectGutters`), so a page whose prose overhangs the first column by a few
 * points has that prose crossing the middle. Cut there, ZapfDingbats.pdf's
 * heading and its lead paragraph either wrapped inside a cell too narrow for
 * them or swallowed the glyph standing beside them. The boundary is put past
 * the crossing ink instead, as far as the band allows — where the columns
 * really do divide.
 *
 * @param runs    The page's runs.
 * @param gutters The page's gutters.
 * @param edges   Where the page's text starts and ends.
 * @returns The column boundaries, left edge first and right edge last.
 */
function columnBounds(
  runs: ReadonlyArray<TextRun>,
  gutters: ReadonlyArray<Gutter>,
  edges: { left: number; right: number },
): Array<number> {
  const inks = runs
    .map((r) => runInk(r))
    .filter((ink): ink is [number, number] => ink !== undefined);
  // A column measured to the last hair of its longest line has no room for
  // that line in another face: the head of ZapfDingbats.pdf's sheet runs 152.8
  // points across a 152.8-point column, and re-set in a substitute it wrapped.
  // So the boundary clears the crossing ink by a quarter of an em — never past
  // the band, which is the other column's.
  const clearance = (median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10) / 4;
  return [
    edges.left,
    ...gutters.map((g) => {
      let at = g.mid;
      // Only ink that ENDS inside the band overhangs it. Ink that runs out the
      // far side is a line spanning the whole page — a heading, a rule of
      // asterisks — and it crosses every gutter there is.
      for (const [from, to] of inks) {
        if (from < g.mid && to > at && to <= g.to) at = Math.min(to + clearance, g.to);
      }
      return at;
    }),
    edges.right,
  ];
}

/**
 * The last column a cell starting at `from` covers — the one where no run
 * reaches any further right.
 *
 * A run whose INK ends past a column boundary was drawn as one line across
 * both, so both belong to one cell; and a cell widened that way may pick up a
 * run that crosses the next boundary in turn.
 *
 * @param row       The row's runs.
 * @param from      The column the cell starts in.
 * @param count     How many columns the table has.
 * @param regionOf  Which column an x falls in.
 */
function spanEnd(
  row: ReadonlyArray<TextRun>,
  from: number,
  count: number,
  regionOf: (x: number) => number,
): number {
  let to = from;
  for (let grew = true; grew && to < count - 1; ) {
    grew = false;
    for (const run of row) {
      const ink = runInk(run);
      if (ink === undefined) continue;
      const start = regionOf(ink[0]);
      if (start < from || start > to) continue;
      // The ink must reach INTO the next column, not merely touch its edge:
      // a boundary sits in the middle of the gap between them.
      const end = regionOf(ink[1] - 1);
      if (end > to) {
        to = Math.min(end, count - 1);
        grew = true;
      }
    }
  }
  return to;
}

/** Below this many gutters a page in columns is read as columns. */
const MIN_RULED_GUTTERS = 2;

/** Below this many rows a page is not ruled into anything. */
const MIN_TABLE_ROWS = 6;

/** How many of a table's rows must carry a cell in every one of its columns. */
const TABLE_ALIGNED_SHARE = 0.6;

/**
 * How much of its column a CELL covers, at the median. A line of prose fills
 * its measure — 84 to 95 per cent across this corpus — and a cell does not.
 */
const TABLE_CELL_FILL = 0.65;
