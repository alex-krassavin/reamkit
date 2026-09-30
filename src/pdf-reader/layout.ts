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
  BASELINE_AT,
  CARRIER_LINE_PT,
  GUESSED_MARGIN,
  MEASURE_LINES,
  NATURAL_LINE_EM,
  buildFlowDoc,
  dedupeLosses,
  figureBlock,
  floatOntoSheet,
  imageBlock,
  imageMember,
  labelMember,
  membersBox,
  paragraphFromRuns,
  positionedText,
  sectionFromPdfPages,
  sectionOnSheet,
  shapeBlock,
  spaceAfter,
  textSizeOf,
  tooSmallToRead,
  vectorMember,
  withMeasuredMargins,
} from './flow-build';
import {
  displayOf,
  placeImages,
  placeRuns,
  placeVectors,
  textFrameOf,
  wordsTurnOf,
} from './display';
import { collectEmbeddedFonts } from './embedded-fonts';
import { collectFaceFamilies } from './font';
import { collectPageImages } from './images';
import { extractPageText } from './text';
import { faceOutlinesOf, kernedFaces, pageSpacing } from './face-outlines';
import { collectPageVectors } from './vector';
import { markDrawnRules } from './text-rules';
import { joinLetters, pageFigures, tracedRun } from './figures';
import { regionsOf } from './regions';
import { punctuationOf } from './glyph-shapes';
import { matrixBlocks } from './math-rows';
import { pageNumberingOf, runOf } from './page-numbers';
import { isRightToLeft } from './content';
import type { ShownCodes } from './face-outlines';
import type { PageFigure } from './figures';
import type { PageNumbering } from './page-numbers';
import type { SideBySide } from './regions';
import type { PdfVector } from './vector';
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
import type { FigureMember, Reconstruction, TextSpan } from './flow-build';
import { FEATURES, ResourceStore, pt } from '@/core/ir';

/** The relationships the reconstruction files its running head and foot under. */
export const FOOTER_PART = 'pdf-running-foot';
export const HEADER_PART = 'pdf-running-head';

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
  /**
   * Whether it is a line of CODE: every word of it set in a typewriter face.
   * Such a line is ended where the page ends it, and re-set as its own.
   */
  readonly code?: boolean;
  /**
   * Where the pieces after each such gap begin, in page space. A line the page
   * SET OUT stands on stops, and a space in place of them puts the second piece
   * against the first: an invoice's "Bill to" block ran into its own address as
   * "548 Market Street walonade@icloud.com's Organization".
   */
  readonly stops?: ReadonlyArray<number>;
  /**
   * The ink each tab-separated piece covers, in page space: the first piece
   * and one after every stop. What tells a column of figures set against its
   * right edge from one set from its left.
   */
  readonly pieces?: ReadonlyArray<Extent>;
}

/** An extent across the page, from where the ink starts to where it ends. */
interface Extent {
  readonly from: number;
  readonly to: number;
}

/**
 * The last line a region set: its baseline and the exact box it stands in,
 * which is what the next region's first paragraph is spaced from. A box of no
 * height is an edge — the foot of a band of blocks side by side.
 */
interface LineBox {
  readonly y: number;
  readonly lineHeight: number;
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
  const sheets = pages.map((page) => displayOf(page));
  // §9.9 — the codes each font paints, for the faces a writer may embed.
  const painted: ShownCodes = new Map();
  const extracted = typewritten(pages.map((page) => extractPageText(file, page, painted)));
  // §9.4.3 — how the page spaces each face: between words, and inside them.
  const spacing = pageSpacing(extracted);
  // §9.9 — the faces the document carries, and the space each is set with
  // there (see `fittedSpace`).
  const outlines = faceOutlinesOf(painted, spacing);
  const faceSpaces: FaceSpaces = new Map(
    [...outlines].flatMap(([name, face]): Array<[string, number]> => {
      const advance = face.glyphs.get(' ')?.advance;
      return advance !== undefined ? [[name, advance]] : [];
    }),
  );
  const onSheets = extracted.map((runs, i) => placeRuns(runs, sheets[i]!));
  // §17.6.20 — …and a page whose words run DOWN its sheet is read in the frame
  // where they stand upright, and set back on the sheet turned: a viewer turns
  // the words with the page, and so does the document (`sectionOnSheet`). A
  // placed reading has no frame to read in: every line of it already stands
  // where the page shows it, turned boxes and all.
  const turns = onSheets.map((runs) => wordsTurnOf(runs));
  const shown = sheets.map((sheet, i) =>
    mode !== 'positional' && turns[i] === 270 ? textFrameOf(sheet) : sheet,
  );
  const allRuns = shown.map((d, i) => (d.sheet ? placeRuns(extracted[i]!, d) : onSheets[i]!));
  // §17.6.13 — what the document repeats at the foot of its pages is a running
  // foot, not a paragraph of the body. Lifted off before anything else reads
  // the page: it must not measure the margins either, and a page number in the
  // text block is a page number in the wrong place.
  // …on a page read the way it is shown. A header and a footer stay across the
  // sheet whichever way a section's lines run (§17.6.20), so the band of a page
  // whose words run down it stays in its text and turns with it.
  // …and a mark set off the sheet (see `offSheet` below) is no band either.
  const offSheet = (r: TextRun, i: number): boolean => r.y < 0 || r.y > shown[i]!.height;
  const bandRuns = allRuns.map((runs, i) =>
    shown[i]!.sheet ? [] : runs.filter((r) => !offSheet(r, i)),
  );
  const foot = mode === 'positional' ? undefined : runningFoot(bandRuns, shown, 'foot');
  const head = mode === 'positional' ? undefined : runningFoot(bandRuns, shown, 'head');
  // §17.6.12 — the page numbers the band prints, and the sequences they run
  // in: front matter numbered i, ii, iii before a body that starts again at 1
  // is two sections, each numbering its pages its own way.
  const numberedBand = foot?.numbered === true ? foot : head?.numbered === true ? head : undefined;
  const numbering = numberingOf(numberedBand);
  // …and what is set too small to read is a mark the producer left on the
  // sheet, not a line of it: TCPDF signs the last page of everything it makes
  // in one-point type in the very corner of the paper. Read as text it was the
  // page's leftmost and lowest line — the margins, the measure and every
  // indent were taken from it. It is put back where the page had it, a box of
  // its own that nothing else is measured against.
  const textSize = textSizeOf(allRuns);
  // …and so is what an annotation's appearance writes (§12.5.5): a button's
  // caption, a field's value, a tick. Its box is placed where the page has
  // it, and read into the page's lines the words left it — evaljs.pdf's
  // "Execute" stood at the margin, two hundred points from its button.
  // …and so is what stands OFF the sheet: a run whose baseline lies past the
  // page's foot or its head shows a sliver of its glyphs at the edge, if
  // anything. freeculture.pdf sets a printer's bar of ZapfDingbats at forty
  // points nine points under the crop of every page of its front matter; read
  // into the page's lines, it came back four lines of bars at the head of a
  // page of their own, and every page after it one sheet late.
  const stamp = (r: TextRun, i: number): boolean =>
    mode !== 'positional' &&
    (r.annotation === true || tooSmallToRead(r, textSize) || offSheet(r, i));
  const stamps = allRuns.map((runs, i) => runs.filter((r) => stamp(r, i)));
  const readRuns =
    foot || head || stamps.some((s) => s.length > 0)
      ? allRuns.map((runs, i) =>
          runs.filter(
            (r) => foot?.lift[i]?.has(r) !== true && head?.lift[i]?.has(r) !== true && !stamp(r, i),
          ),
        )
      : allRuns;
  // Every page's pictures and paths, lifted before its text is read.
  // A white box drawn over a picture is not invisible paint but the thing that
  // hides it, so the paths are filtered against the pictures already placed.
  const art = pages.map((page, i) => {
    const raw = collectPageImages(file, page);
    const images = placeImages(raw.images, shown[i]!);
    const covered = images.map((img) => ({
      minX: img.x,
      minY: img.y,
      maxX: img.x + img.widthPt,
      maxY: img.y + img.heightPt,
    }));
    const lifted = collectPageVectors(file, page, covered);
    return {
      images,
      imageLosses: raw.losses,
      vectors: placeVectors(lifted.vectors, shown[i]!),
      vectorLosses: lifted.losses,
    };
  });
  // §8.5 — the figures the pages draw (see `./figures`). The words set on a
  // figure are its labels: the page's text is read without them, and they are
  // set with the drawing they label. Read as text, comments.pdf's state machine
  // cut its page into so many columns that the page was taken for a table.
  // …in a FLOWING reading of a page read upright: a placed one anchors every
  // mark where it stands already.
  const figures = art.map((a, i) =>
    mode === 'positional' || shown[i]!.sheet !== undefined
      ? []
      : pageFigures(a.vectors, a.images, readRuns[i]!, shown[i]!),
  );
  const pageRuns = figures.some((f) => f.length > 0)
    ? readRuns.map((runs, i) => {
        const labels = new Set(figures[i]!.flatMap((f) => f.labels));
        return labels.size > 0 ? runs.filter((r) => !labels.has(r)) : runs;
      })
    : readRuns;

  // Every page's pictures, kept for the margins the section is measured to.
  const pageMarks: Array<
    ReadonlyArray<{ x: number; y: number; widthPt: number; heightPt: number }>
  > = pages.map(() => []);
  const medianFont =
    median(
      pageRuns
        .flat()
        .map((r) => r.fontSizePt)
        .filter((s) => s > 0),
    ) || 12;

  const resources = new ResourceStore();
  const losses: Array<Loss> = [];
  // Words that run UP a sheet, or stand on their heads, have no section to be
  // set in that way — Word and LibreOffice both lay a section's text down the
  // sheet or across it and no other way — and are set across it.
  if (mode !== 'positional' && turns.some((t) => t === 90 || t === 180)) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'a page shows its words running up the sheet or upside down; a document sets a section’s text across the sheet or down it, so they are set across it',
    });
  }
  // §8.6.6.2 — type filled with a tiling pattern keeps the pattern's colour at
  // the pattern's own density and loses its shape: a run carries one colour, not
  // a content stream, so a hatch that alternates ink and paper becomes the flat
  // tint the two average to.
  // §9.10.2 — glyphs whose face maps them to nothing a reader can show. The
  // words are unrecoverable, and a page that silently comes back blank is the
  // one loss this reader must never take without saying so:
  // arial_unicode_ab_cidfont.pdf is four Arabic letters and nothing else.
  // A figure's traced labels are among them: drawn, and still not text.
  if (readRuns.some((page) => page.some((r) => r.text.includes(UNMAPPED)))) {
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
  // …and not the gutters a page's code listings show (see `withoutListings`),
  // nor those of a table set over text in columns (`withoutTables`): the
  // text's own gutters are the page's where it has any, and a page that is its
  // table is ruled by it.
  const gutterRuns = pageRuns.map((runs, i) => {
    const listed = withoutListings(runs);
    const prose = withoutTables(listed);
    return prose !== listed && detectGutters(prose, shown[i]!.width).length > 0 ? prose : listed;
  });
  const perPage = pages.map((_, i) => detectGutters(gutterRuns[i]!, shown[i]!.width));
  const shared = commonGutters(perPage);
  const pageGutters = perPage.map((own, i) =>
    own.length > 0 ? own : shared && fitsGutters(gutterRuns[i]!, shared) ? shared : own,
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
  // Each page's sheet: its size, and whether it is set turned.
  const sheetSizes: Array<string> = [];
  // The first paragraph each page's text begins with: where in the body it is,
  // and the baseline and box the page set its first line at.
  const leads: Array<{ page: number; at: number; baseline: number; lineHeight: number }> = [];
  // §17.6.4 — the columns the pages are SET in, which change where a page's
  // gutters do: a paper's title stands over the two columns of its body, and a
  // section is what carries a column setup.
  let curColumns = 1;
  let curSpace = 0;
  let pendingContinuous = false;
  // Where each page's last columns end, down the page (see `balancedEnd`).
  const columnFeet: Array<ReadonlyArray<number> | undefined> = [];
  // Where in the body a page turns from one of its columns to the next: where
  // the column before the turn ends and the one after it, up the page, and
  // whether a line across the page closes their band (see `columnTurns` below).
  const columnTurns: Array<{ at: number; foot: number; next: number; closed: boolean }> = [];
  pages.forEach((page, i) => {
    const runs = pageRuns[i]!;
    const display = shown[i]!;
    // §17.6.20 — a page read in its text frame is set back on its sheet turned,
    // and Word turns a section's paragraphs and not its tables: a table set
    // there lies flat in the corner of the sheet while the text around it runs
    // down. Such a page is read as the lines and stops it shows.
    const turned = display.sheet !== undefined;
    // EP17 — the page's gutters, and so its columns. Each column is grouped and
    // read independently, and its blocks precede the next column's.
    const gutters = pageGutters[i]!;
    const pageWidth = display.width;
    // Whether this page steps between its words or writes spaces of its own,
    // which decides how wide a gap has to be to mean one.
    const stepped = stepsBetweenWords(runs);
    // Blocks carry a column key so the final sort reads column-by-column: left
    // column top-to-bottom, then right column.
    const blocks: Array<Block> = [];
    // EP17 — a full-width line cuts the page in two: what is above it is read
    // before it and what is below after, so a paper's columns do not start at
    // the top of the sheet.
    // Where the page's text starts and ends, which is the measure a line that
    // spans the page is set across.
    const textEdges = pageTextEdges(runs);
    // …but only a page set in columns of PROSE is read down them. An invoice
    // breaks a dozen lines at the same x and is not: read by column its every
    // amount was taken off the line it belongs to and carried to the end of the
    // document — "Total excluding tax" on one sheet and "$100.00" on the next.
    // Read across, each label keeps its figure.
    const inColumns = proseColumns(runs, gutters, textEdges);
    // EP17 — a full-width line cuts the page in two: what is above it is read
    // before it and what is below after, so a paper's columns do not start at
    // the top of the sheet.
    const split = gutters.length > 0 && inColumns ? assignColumns(runs, gutters) : undefined;
    // The page's figures, each read in the column it stands in. One that
    // reaches across a gutter spans the columns, and cuts the page in two
    // where it begins, as a line set across the page does.
    const figs = figures[i]!;
    const spansGutter = (f: PageFigure): boolean =>
      gutters.some((g) => f.minX < g.mid && f.maxX > g.mid);
    const breaks = split
      ? [...split.breaks, ...figs.filter(spansGutter).map((f) => f.maxY)].sort((a, b) => b - a)
      : [];
    const bandEpsilon = (median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10) / 2;
    const bandAt = (top: number): number => bandOf(breaks, top, bandEpsilon);
    const addColumn = (
      allRuns: ReadonlyArray<TextRun>,
      col: number,
      columnFigures: ReadonlyArray<PageFigure> = [],
    ): void => {
      // §9.6.5 — a Type 3 run's marks are its glyph PROCEDURES, which the path
      // and picture passes lift. Re-setting its codes in a substitute face
      // would draw a second, smaller copy of a drawing.
      //
      // §9.3.6 — and a run the page painted nowhere is not drawn either. Both
      // are kept for the FLOWING reading, which is a document being read rather
      // than a page being reproduced: a scanned page's every word lives in its
      // invisible layer.
      // A Type 3 glyph is drawn in EVERY reading, because the page's marks are
      // its only typeface — so its code is not re-set in a substitute anywhere.
      // bug1245391_reduced.pdf sets three Chinese characters as 89×85 bitmaps
      // and they came back doubled: the file's own light strokes with a
      // substitute's heavy ones over them.
      //
      // §9.3.6 — a run the page painted NOWHERE is a different thing, and a
      // flowing reading keeps it: a document is being read rather than a page
      // reproduced, and a scanned page's every word lives in that layer.
      const colRuns = allRuns.filter(
        (r) => r.type3 !== true && (mode !== 'positional' || r.invisible !== true),
      );
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
        for (const [angle, turned] of byAngle(colRuns)) {
          for (const runs of byPainter(turned)) {
            for (const line of groupIntoLines(rotate(runs, -angle), true, stepped)) {
              if (line.text.length === 0) continue;
              const box = turnedBox(line, angle, pageWidth);
              placed.push({
                key: placedKey(runs, placed.length),
                col,
                // The box's own top on the PAGE — `line.y` is measured in the
                // turned frame, and the blocks are ordered by where they stand.
                top: box.y + box.height,
                make: (z: number): BodyElement =>
                  positionedText(line.spans, box, frame, z, rotation60kOf(angle)),
              });
            }
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
      const textRuns = inMath.size > 0 ? colRuns.filter((r) => !inMath.has(r)) : colRuns;
      const lines = groupIntoLines(textRuns, false, stepped).filter((l) => l.text.length > 0);
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
      // The column read region by region: a stretch that reads straight down,
      // then a band of blocks side by side, each spaced from the one before.
      // A figure stands between the text over it and the text under it, and
      // the text under it is spaced from its foot.
      let above: LineBox | undefined;
      const slabs: Array<{ runs: ReadonlyArray<TextRun>; figure?: PageFigure }> = [];
      let rest = textRuns;
      for (const figure of columnFigures) {
        slabs.push({ runs: rest.filter((r) => r.y > figure.maxY), figure });
        rest = rest.filter((r) => r.y <= figure.maxY);
      }
      slabs.push({ runs: rest });
      for (const slab of slabs) {
        const regions = turned
          ? [{ kind: 'flow' as const, runs: slab.runs }]
          : regionsOf(slab.runs);
        for (const region of regions) {
          if (region.kind === 'side') {
            const made = sideBySide(region, measure, above);
            if (made === undefined) continue;
            blocks.push({ band: bandAt(made.top), col, top: made.top, el: made.el });
            above = made.below;
            continue;
          }
          const regionLines = groupIntoLines(
            region.runs,
            false,
            stepped,
            faceSpaces,
            columnRules(vectors),
          ).filter((l) => l.text.length > 0);
          // A sheet of a line or two shows no measure (see `MEASURE_LINES`):
          // its longest line reaches the edge only because it IS the edge, and
          // the .docx is re-set across the sheet's own width instead — at the
          // least all but the third a guessed margin may take. A line whose
          // next word would have fit short of THAT was ended, not broken: run
          // together, checkbox-bad-appearance.pdf's "Checkbox 1 - not checked"
          // and "✔ Checkbox 2 - Checked" came back side by side on one line.
          const reach =
            lines.length < MEASURE_LINES && gutters.length === 0
              ? pageWidth * (1 - GUESSED_MARGIN)
              : undefined;
          const paras = groupIntoParagraphs(
            regionLines,
            measure,
            display.height,
            above,
            false,
            reach,
          );
          for (const set of setParagraphs(
            paras,
            measure && {
              left: measure.left,
              // The table runs to where the PAGE's text ends, not to where this
              // column's longest line does: the last column of a payment history
              // begins at its heading and its figures reach past it, and measured
              // to the line the column came out a point wide.
              right: Math.max(measure.right, textEdges?.right ?? measure.right),
            },
            pageWidth,
          )) {
            blocks.push({
              band: bandAt(set.top),
              col,
              top: set.top,
              el: set.el,
              ...(set.foot !== undefined ? { foot: set.foot } : {}),
            });
          }
          const last = paras[paras.length - 1];
          if (last !== undefined) above = { y: last.bottom, lineHeight: last.lineHeight };
        }
        if (slab.figure !== undefined) {
          const made = figureIn(slab.figure, measure, above, col === SPANNING_COLUMN);
          blocks.push({ band: bandAt(made.top), col, top: made.top, el: made.el });
          above = made.below;
        }
      }
    };
    /** A figure's parts in the order the page painted them, its words over its drawing. */
    const figureMembers = (figure: PageFigure): Array<FigureMember> => {
      const painted = [...figure.vectors].sort((a, b) => compareOrder(a.orderKey, b.orderKey));
      const drawn = [
        ...joinLetters(painted).map((v) => ({ key: v.orderKey, member: vectorMember(v) })),
        ...figure.images.map((img) => ({ key: img.orderKey, member: imageMember(img, resources) })),
      ].sort((a, b) => compareOrder(a.key, b.key));
      const labels: Array<FigureMember> = [];
      // Lettering that states no character is drawn with the figure's paths.
      const written = figure.labels.filter((r) => !tracedRun(r));
      for (const [angle, turnedRuns] of byAngle(written)) {
        for (const layer of byPainter(turnedRuns)) {
          for (const line of groupIntoLines(rotate(layer, -angle), true, stepped)) {
            if (line.text.length === 0) continue;
            // Upright, the box is the words' own and an em more. It is told
            // not to wrap (see `labelMember`), and LibreOffice wraps it all
            // the same where the words fill it: comments.pdf's "Guard" came
            // back "Guar".
            const box =
              angle === 0
                ? {
                    x: line.x,
                    y: line.y - line.fontSize * 0.25,
                    width: line.width + line.fontSize,
                    height: line.fontSize * 1.25,
                  }
                : turnedBox(line, angle, pageWidth);
            labels.push(labelMember(line.spans, box, rotation60kOf(angle)));
          }
        }
      }
      return [...drawn.map((d) => d.member), ...labels];
    };
    /**
     * §20.5.2.17 — a figure as the paragraph it stands in (see `figureBlock`):
     * spaced from the text over it as a paragraph is, and set in from its
     * column's edge as far as the page set it.
     *
     * A figure across the columns is spaced from the text nearest over it in
     * any column: what its own column holds over it is the line before the
     * columns began.
     */
    const figureIn = (
      figure: PageFigure,
      measure: { left: number; right: number } | undefined,
      over: LineBox | undefined,
      across = false,
    ): { top: number; below: LineBox; el: BodyElement } => {
      const members = figureMembers(figure);
      const box = membersBox(members);
      const nearest = across
        ? runs
            .map((r) => r.y - r.fontSizePt * 0.25)
            .filter((edge) => edge > box.top)
            .reduce<number | undefined>((low, edge) => Math.min(low ?? edge, edge), undefined)
        : undefined;
      const above: LineBox | undefined =
        nearest !== undefined ? { y: nearest, lineHeight: 0 } : over;
      const opened =
        above !== undefined ? above.y - (1 - BASELINE_AT) * above.lineHeight - box.top : 0;
      const indent = box.left - (measure?.left ?? box.left);
      const properties: ParagraphProperties = {
        ...(opened > SPACING_NOISE_PT
          ? { spacingBefore: pt(Math.min(opened, display.height / 3)) }
          : {}),
        ...(Math.abs(indent) > SPACING_NOISE_PT ? { indentLeft: pt(indent) } : {}),
      };
      return {
        top: box.top,
        below: { y: box.bottom, lineHeight: 0 },
        el: figureBlock(members, properties),
      };
    };
    /**
     * A region's paragraphs as the elements they are set in: the lines that
     * stand on the same stops as a table (§17.4.38), the rest as paragraphs
     * placed and spaced the way the page placed and spaced them.
     */
    // The paragraphs set farther below the text above them than spacing may
    // say, with their lines (see `setOff`).
    const far = new WeakMap<BodyElement, ReadonlyArray<Line>>();
    const setParagraphs = (
      paras: ReturnType<typeof groupIntoParagraphs>,
      tableMeasure: { left: number; right: number } | undefined,
      sheetWidth = 0,
    ): Array<{ top: number; el: BodyElement; foot?: number }> => {
      // §17.4.38 — consecutive lines set out on the SAME stops are a table:
      // "Description / Qty / Unit price / Tax / Amount" and the row under it.
      // Written as tabbed paragraphs the picture is right and the document is
      // not — nothing downstream can read a column out of it, and a reader that
      // re-wraps one cell drags the whole line with it.
      const asRows = turned
        ? new Map<number, BodyElement | null>()
        : tabbedRows(paras, tableMeasure);
      const out: Array<{ top: number; el: BodyElement; foot?: number }> = [];
      for (const [at, para] of paras.entries()) {
        const table = asRows.get(at);
        if (table !== undefined) {
          if (table !== null) out.push({ top: para.top, el: table });
          continue;
        }
        // A cell's lines keep to the cell: there is no sheet for them to run on into.
        const overflow = sheetWidth > 0 ? runOn(para, tableMeasure, sheetWidth) : 0;
        const el = paragraphFromRuns(para.spans, headingLevel(para.fontSize, medianFont), {
          ...(para.stops !== undefined && para.stops.length > 0
            ? {
                tabs: para.stops
                  .filter((x) => x > 0)
                  .map((x) => ({ positionPt: pt(x), alignment: 'left' as const })),
              }
            : {}),
          ...(para.alignment !== undefined ? { alignment: para.alignment } : {}),
          ...(para.spacingBefore !== undefined ? { spacingBefore: pt(para.spacingBefore) } : {}),
          spacingLine: pt(para.lineHeight),
          spacingLineRule: 'exact',
          ...(para.indentLeft !== undefined ? { indentLeft: pt(para.indentLeft) } : {}),
          ...(para.indentFirstLine !== undefined
            ? { indentFirstLine: pt(para.indentFirstLine) }
            : {}),
          ...(overflow > 0 ? { indentRight: pt(-overflow) } : {}),
        });
        if (sheetWidth > 0 && para.far !== undefined) far.set(el, para.far);
        out.push({
          top: para.top,
          el,
          foot: para.bottom - (1 - BASELINE_AT) * para.lineHeight,
        });
      }
      return out;
    };
    /**
     * §17.4.38 — a band of blocks side by side, as the borderless table a
     * writer sets such a band in: one row, a cell a block, each cell's lines
     * read down on their own and spaced from what stands above the band.
     */
    const sideBySide = (
      region: SideBySide,
      measure: { left: number; right: number } | undefined,
      above: LineBox | undefined,
    ): { top: number; below: LineBox; el: BodyElement } | undefined => {
      const left = measure?.left ?? Math.min(...region.cells.map((c) => c.from));
      const right = Math.max(
        measure?.right ?? 0,
        textEdges?.right ?? 0,
        ...region.cells.map((c) => c.to),
      );
      const cells = region.cells.map((cell) => {
        const lines = groupIntoLines(cell.runs, false, stepped, faceSpaces).filter(
          (l) => l.text.length > 0,
        );
        return { cell, lines };
      });
      if (cells.some((c) => c.lines.length === 0)) return undefined;
      // Where the band begins: the top of the highest first line's box. Each
      // cell's first line is spaced down from it, so every block starts where
      // the page started it, not at the top of the row.
      const top = Math.max(...cells.map((c) => c.lines[0]!.y));
      const edge: LineBox = above ?? {
        y: Math.max(
          ...cells.map((c) => c.lines[0]!.y + BASELINE_AT * NATURAL_LINE_EM * c.lines[0]!.fontSize),
        ),
        lineHeight: 0,
      };
      // Each cell begins where its ink does, the first at the measure.
      const starts = cells.map((c, k) => (k === 0 ? left : c.cell.from));
      let foot = Infinity;
      const row = cells.map(({ cell, lines }, k) => {
        const own = { left: cell.from, right: cell.to };
        // A block beside another is a stack of lines — a label over its value,
        // an address — set as narrow as it is: run together, its lines re-wrap
        // wherever a substitute's widths put the break. A block of PROSE is
        // not: its lines run to the edge the column breaks them at, and set a
        // line apiece, each one a word wider in a substitute's widths left
        // that word on a line of its own — comments.pdf's two columns, read
        // as blocks side by side, came back half again as long, a word
        // standing alone under every line.
        const paras = groupIntoParagraphs(lines, own, display.height, edge, !isProse(lines, own));
        const last = paras[paras.length - 1]!;
        foot = Math.min(foot, last.bottom - (1 - BASELINE_AT) * last.lineHeight);
        const inset = cell.from - starts[k]!;
        const content = setParagraphs(paras, own).map(({ el }) =>
          inset > 0 && el.kind === 'paragraph'
            ? {
                ...el,
                paragraph: {
                  ...el.paragraph,
                  properties: {
                    ...el.paragraph.properties,
                    indentLeft: pt((el.paragraph.properties.indentLeft ?? 0) + inset),
                  },
                },
              }
            : el,
        );
        const width = (starts[k + 1] ?? right) - starts[k]!;
        return { properties: { width: pt(Math.max(width, 1)) }, content };
      });
      const grid = row.map((c) => c.properties.width);
      return {
        top,
        below: { y: foot, lineHeight: 0 },
        el: {
          kind: 'table',
          table: {
            properties: {
              defaultCellMargins: { left: pt(0), right: pt(0) },
              layout: 'fixed',
              widthType: 'dxa',
              widthPt: pt(grid.reduce((sum, w) => sum + w, 0)),
            },
            grid,
            rows: [{ properties: {}, cells: row }],
          },
        },
      };
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
    for (const runs of byPainter(stamps[i] ?? [])) {
      for (const line of groupIntoLines(runs, true, stepped)) {
        if (line.text.length === 0) continue;
        const box = turnedBox(line, 0, pageWidth);
        placed.push({
          key: placedKey(runs, placed.length),
          col: colOf(box.x + box.width / 2),
          top: box.y + box.height,
          make: (z: number): BodyElement => positionedText(line.spans, box, frame, z),
        });
      }
    }
    const lifted = art[i]!;
    losses.push(...lifted.imageLosses);
    // …and kept, because a margin is measured to the page's INK and a picture
    // is ink (see `withMeasuredMargins`).
    pageMarks[i] = lifted.images;
    losses.push(...lifted.vectorLosses);
    // What a figure draws is set with the figure (see `figureIn`).
    const figured = new Set<object>(figures[i]!.flatMap((f) => [...f.vectors, ...f.images]));
    const imgs = {
      images: figured.size > 0 ? lifted.images.filter((img) => !figured.has(img)) : lifted.images,
    };
    // Filled vector paths (EP10) are ANCHORED where the page drew them — they
    // are artwork, not paragraphs, and a sheet of them has no reading order to
    // take a place in. They still sort by top edge, so their z-order is the
    // order the page painted them in.
    // A PDF has no underline: it draws a thin bar under the words. Read onto
    // the runs BEFORE they are grouped, so the mark travels with them and the
    // bar is not placed a second time where the words no longer are.
    const placedVectors =
      figured.size > 0 ? lifted.vectors.filter((v) => !figured.has(v)) : lifted.vectors;
    const drawnRules = markDrawnRules(runs, placedVectors);
    const strayGlyphs = strayMarks(placedVectors, runs);
    // …and the marks the page draws for want of a character, read where their
    // shape says what they are: the hyphen of an invoice number, the colon
    // after a label. A flowing reading drops the drawing, so this is the only
    // way the character is written at all.
    const ruled =
      mode !== 'positional'
        ? { ...drawnRules, runs: readDrawnMarks(drawnRules.runs, strayGlyphs) }
        : drawnRules;
    const vectors = placedVectors
      .filter((v) => !ruled.consumed.has(v))
      // §9.6.6 — a glyph the file states no character for is DRAWN, which is
      // right for a page whose words are all drawn and wrong for one mark
      // inside a line of ordinary text. Stripe declares U+0000 for every piece
      // of punctuation it sets: the colon of "Kazakhstan VAT: 86-1696045" came
      // back as a floating shape, and a flowing document has nowhere to float
      // it — it landed a word away from the line it belongs to, on its own.
      // Where the line around it is readable the mark is dropped instead: the
      // words keep their places, and the loss report already says a character
      // was unrecoverable.
      // …in the FLOWING reading only. A placed one anchors every mark to the
      // page, so a drawn glyph lands exactly where the file draws it.
      .filter((v) => !(mode !== 'positional' && strayGlyphs.has(v)));
    // …and what is left of the traced glyphs is drawn a word at a time.
    const drawn = drawnWords(vectors);
    // A page RULED into columns is a table, and its ROWS are what it says; a
    // page SET in columns is prose, and its columns are. Read by column, a
    // table comes back one column at a time with every row torn up.
    const ruledIntoColumns =
      mode !== 'positional' && !turned && looksRuled(ruled.runs, gutters, textEdges);
    const asTable =
      ruledIntoColumns && textEdges
        ? tableFrom(ruled.runs, gutters, textEdges, stepped)
        : undefined;
    if (asTable) {
      for (const block of asTable) {
        blocks.push({ band: bandAt(block.top), col: 0, top: block.top, el: block.el });
      }
      // A table's rows are read whole and spaced by their own pitch, so a
      // figure among them stands at its top, spaced from nothing.
      for (const figure of figs) {
        const made = figureIn(figure, undefined, undefined);
        blocks.push({ band: bandAt(made.top), col: 0, top: made.top, el: made.el });
      }
    } else if (split && !ruledIntoColumns) {
      // A run the rules pass rebuilt is not the one the split was measured on,
      // so its column is looked up by where it stands.
      const columnFor = (r: TextRun): number => split.columnOf.get(r) ?? colOf(r.x);
      const figureColumn = (f: PageFigure): number =>
        spansGutter(f) ? SPANNING_COLUMN : colOf((f.minX + f.maxX) / 2);
      const columns = Array.from({ length: gutters.length + 1 }, (_, n) => n);
      for (const col of [SPANNING_COLUMN, ...columns]) {
        addColumn(
          ruled.runs.filter((r) => columnFor(r) === col),
          col,
          figs.filter((f) => figureColumn(f) === col),
        );
      }
      spaceUnderSpans(blocks);
    } else {
      addColumn(ruled.runs, 0, figs);
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
    // A RULE is not artwork, it is punctuation: the line under a table's
    // headings, the line over a total. Anchored to the page at the y it was
    // drawn at, it stays there while the words around it re-set — a receipt
    // came back with a black line struck through "Max plan - 5x" and another
    // through "Payment history". Given to the paragraph it separates it moves
    // with it (§17.3.1.24). A placed reading keeps its anchor: nothing moves
    // there.
    // Before the rules are given to the paragraphs they separate: a line that
    // is placed keeps the rule over it where the page drew it.
    setOff(blocks, far, (line, k) =>
      positionedText(line.spans, turnedBox(line, 0, pageWidth), frame, FAR_Z + k),
    );
    const givenAway =
      mode !== 'positional' ? ruleBorders(vectors, blocks, display.width, colOf) : undefined;
    const marks = [
      ...imgs.images.map((img) => ({
        key: img.orderKey,
        col: colOf(img.x + img.widthPt / 2),
        top: img.y + img.heightPt,
        make: (z: number): BodyElement => imageBlock(img, resources, undefined, frame, z, under),
      })),
      ...drawn
        .filter((v) => givenAway?.has(v) !== true)
        .map((v) => ({
          key: v.orderKey,
          col: colOf((v.minX + v.maxX) / 2),
          top: v.maxY,
          // A RULE is not artwork, it is punctuation: the line under a table's
          // headings, the line over a total. Anchored to the page at the y it was
          // drawn at, it stays where the page had it while the words around it
          // re-set — and a receipt came back with a black line struck through
          // "Max plan - 5x" and another through "Payment history". Set in the
          // flow instead, it stands between the blocks it separates wherever they
          // end up. A placed reading keeps its anchor: there nothing moves.
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
    // …and so does a page whose lines run another way than the one before
    // (§17.6.20): a section is what carries the direction too.
    const size = `${shown[i]!.width.toFixed(2)}x${shown[i]!.height.toFixed(2)}${turned ? ' down' : ''}`;
    // …and where the pages' numbering starts again, which only a section does.
    const opensSection =
      i > 0 && (size !== lastSize || numbering?.runs.some((run) => run.from === i) === true);
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
    sheetSizes[i] = size;
    // Each source page after the first opens an output page of its own. Flowed,
    // the layout repaginates and this hardly shows; PLACED, every mark is
    // anchored to "the page", so without it all twenty-five pages of
    // Brotli-Prototype-FileA.pdf stack onto one.
    //
    // …and so does a page with nothing on it to read. A blank sheet is still a
    // sheet: doc_actions.pdf is three of them, and came back as one.
    // Where this page's break stands in the body, where it has one.
    const breakAt = i > 0 && !opensSection ? body.length : undefined;
    if (i > 0 && !opensSection) {
      body.push({
        kind: 'paragraph',
        paragraph: {
          properties: {
            pageBreakBefore: true,
            spacingLine: CARRIER_LINE_PT,
            spacingLineRule: 'exact',
          },
          runs: [],
        },
      });
    } else if (i === 0 && blocks.length === 0) {
      // A break is before the paragraph that carries it, and the first
      // paragraph of a document breaks from nothing: a blank first sheet has
      // to hold something of its own for the second to begin after it.
      body.push({
        kind: 'paragraph',
        paragraph: {
          properties: { spacingLine: CARRIER_LINE_PT, spacingLineRule: 'exact' },
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
    // …and a page whose regions are not PROSE has none either. An invoice is
    // one column of text with its amounts set against the right margin: a
    // dozen lines split at the same x, which is what a gutter looks like from
    // the outside, and nothing at all like two columns from the inside. Set in
    // two, its "Total excluding tax" was indented past the width of the strip
    // it had been given and came back one letter per line down the sheet.
    const columnsHere =
      ruledIntoColumns || !proseColumns(runs, gutters, textEdges) ? 1 : gutters.length + 1;
    const spacePt = columnsHere > 1 ? median(gutters.map((g) => g.to - g.from)) : 0;
    columnFeet[i] = columnsHere > 1 ? lastColumnFeet(blocks) : undefined;
    let led = false;
    // Whether a block of this page stands in the body yet.
    let begun = false;
    // Where each column of a band ends, up the page: the foot of its lowest
    // block, where that is a paragraph's.
    const footOf = (band: number, col: number): number | undefined => {
      const column = blocks.filter((b) => b.band === band && b.col === col);
      if (column.length === 0) return undefined;
      return column.reduce((low, b) => (b.top < low.top ? b : low)).foot;
    };
    let prev: Block | undefined;
    for (const block of blocks) {
      if (
        columnsHere > 1 &&
        prev !== undefined &&
        block.band === prev.band &&
        prev.col !== SPANNING_COLUMN &&
        block.col > prev.col
      ) {
        const ends = footOf(prev.band, prev.col);
        const next = footOf(block.band, block.col);
        const closed = blocks.some((b) => b.band === block.band && b.col === SPANNING_COLUMN);
        if (ends !== undefined && next !== undefined) {
          columnTurns.push({ at: body.length, foot: ends, next, closed });
        }
      }
      prev = block;
      const count = block.col === SPANNING_COLUMN ? 1 : columnsHere;
      if (count !== curColumns) {
        // §17.18.77 — a section that opens a page opens it itself, not
        // continuously after a break: Word does not break before a paragraph
        // that carries the section before it, and comments.pdf's twelfth page
        // came back under the eleventh on one sheet.
        const opensPage = breakAt !== undefined && !begun;
        if (opensPage) {
          // The section before it ends on its own page's last paragraph where
          // that is the section's own. Anywhere else the break stays as its
          // carrier, a line of no height — which, on a page the words fill to
          // the foot, is a line that goes over onto a sheet of its own.
          const last = body[breakAt - 1];
          const own = sectionEnds.every((e) => e.at < breakAt);
          if (last?.kind === 'paragraph' && own) body.pop();
          else {
            body[breakAt] = {
              kind: 'paragraph',
              paragraph: {
                properties: { spacingLine: CARRIER_LINE_PT, spacingLineRule: 'exact' },
                runs: [],
              },
            };
          }
        }
        sectionEnds.push({
          at: body.length,
          from: sectionFrom,
          to: opensPage ? i : i + 1,
          columns: curColumns,
          spacePt: curSpace,
          continuous: pendingContinuous,
        });
        sectionFrom = i;
        pendingContinuous = !opensPage;
        curColumns = count;
        curSpace = spacePt;
      }
      if (!led && mode !== 'positional') {
        const lead = leadingLine(block.el);
        if (lead !== undefined)
          leads.push({ page: i, at: body.length, baseline: block.top, lineHeight: lead });
        // A table or a line of text is where the page's text begins; a mark
        // anchored to the page takes no room and does not.
        led = lead !== undefined || block.el.kind === 'table' || block.el.kind === 'paragraph';
      }
      body.push(block.el);
      begun = true;
    }
  });
  // A placed reading anchors everything to the page, so its margins must stay
  // at zero or the anchors move. A FLOWING one is a document being re-set, and
  // a document with no margins prints its words against the edge of the paper
  // — which is what every converted PDF looked like.
  // Measured in the frame the pages were READ in; `sectionOnSheet` sets a
  // turned one back on its sheet once everything measured against it is done.
  const measured = (
    own: SectionProperties | undefined,
    from: number,
    to: number,
  ): SectionProperties | undefined =>
    withMeasuredMargins(
      own,
      shown.slice(from, to),
      pageRuns.slice(from, to),
      pageMarks.slice(from, to),
      // The band is the upright pages' own, and a turned page kept its own.
      shown[from]?.sheet ? undefined : foot?.band,
    );
  // §17.6.11 — the head and foot of the text block are the SHEET's, not a
  // section's: a section a change of columns opens is measured on the pages
  // it touches, one of them perhaps, and what one page happens to set first
  // is no margin. comments.pdf's eleventh page opens on a chart, and its top
  // margin was measured to the caption under it, 352 points down; Word sets
  // the page from there, and three pages ran over. Every page cut from the
  // same sheet gives its head and foot; across it, a section keeps its own.
  const sheetRun = (from: number): [number, number] => {
    let start = from;
    while (start > 0 && sheetSizes[start - 1] === sheetSizes[from]) start--;
    let end = from + 1;
    while (end < pages.length && sheetSizes[end] === sheetSizes[from]) end++;
    return [start, end];
  };
  const setUps = new Map<string, SectionProperties | undefined>();
  const setUp = (from: number, to: number): SectionProperties | undefined => {
    const key = `${String(from)}:${String(to)}`;
    if (setUps.has(key)) return setUps.get(key);
    const own = sectionFromPdfPages(pages.slice(from, to), shown[from]);
    let section = mode === 'positional' ? own : measured(own, from, to);
    const [start, end] = sheetRun(from);
    if (mode !== 'positional' && section?.margins && (start < from || end > to)) {
      const sheet = measured(own, start, end)?.margins;
      if (sheet) {
        section = {
          ...section,
          margins: { ...section.margins, top: sheet.top, bottom: sheet.bottom },
        };
      }
    }
    setUps.set(key, section);
    return section;
  };
  // §17.6.4 — a last page that sets its columns BALANCED, level with each
  // other above the foot of the sheet, ends their section before the document
  // does. A word processor balances the columns of a section another follows
  // on the same page, and runs the last section's first column to the foot of
  // the sheet: comments.pdf's references stand nine to a column, and came
  // back all nineteen down the left one, the right one empty.
  const lastFeet = columnFeet[pages.length - 1];
  const floor = setUp(sectionFrom, pages.length)?.margins?.bottom;
  if (
    mode !== 'positional' &&
    curColumns > 1 &&
    lastFeet !== undefined &&
    floor !== undefined &&
    balancedEnd(lastFeet, floor, medianFont)
  ) {
    sectionEnds.push({
      at: body.length,
      from: sectionFrom,
      to: pages.length,
      columns: curColumns,
      spacePt: curSpace,
      continuous: pendingContinuous,
    });
    body.push({
      kind: 'paragraph',
      paragraph: {
        properties: { spacingLine: CARRIER_LINE_PT, spacingLineRule: 'exact' },
        runs: [],
      },
    });
    curColumns = 1;
    curSpace = 0;
    pendingContinuous = true;
  }
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
          const base = sectionOnSheet(setUp(end.from, end.to), shown[end.from]);
          if (!base) return [];
          const properties: SectionProperties = {
            ...base,
            ...numberedFrom(numbering, end.from),
            ...(end.columns > 1 && mode !== 'positional'
              ? { columns: { count: end.columns, spacePt: end.spacePt } }
              : {}),
            ...(end.continuous ? { sectionStart: 'continuous' as const } : {}),
          };
          return [{ properties, endIndex: end.at }];
        })
      : [];
  // §17.3.1.33 — a page's text begins where the page began it, not against the
  // top margin: the margin is measured to the highest ink, and an invoice
  // paints a band across the top of the sheet thirty points above its title.
  // Set against the margin the whole page rose by that much.
  for (const lead of leads) {
    const end = sectionEnds.find((e) => e.at > lead.at);
    const top = (end ? setUp(end.from, end.to) : setUp(0, pages.length))?.margins?.top;
    const page = shown[lead.page];
    const el = body[lead.at];
    if (top === undefined || page === undefined) continue;
    const before = page.height - top - lead.baseline - BASELINE_AT * lead.lineHeight;
    if (before <= SPACING_NOISE_PT) continue;
    if (el?.kind === 'paragraph') {
      body[lead.at] = {
        ...el,
        paragraph: {
          ...el.paragraph,
          properties: { ...el.paragraph.properties, spacingBefore: pt(before) },
        },
      };
    } else if (el?.kind === 'shape') {
      body[lead.at] = {
        ...el,
        shape: {
          ...el.shape,
          paragraphProperties: { ...el.shape.paragraphProperties, spacingBefore: pt(before) },
        },
      };
    }
  }
  // §17.3.3.1 — a column the page ends short of the foot of its text ends
  // there. A word processor runs a column to the foot of the text before it
  // turns to the next one, and the foot is the section's, the lowest any of
  // its pages sets: canvas.pdf ends its first sheet's left column fifty points
  // above the foot its second sheet sets, and the head of the right column,
  // "http://blog.nihilogic.dk/" and "Compositing", came back under the left
  // one. A break to the next column says where the page ended it — only
  // where the column ends lines short of the foot, so that one set a line or
  // two longer than the page set it still turns before the foot does.
  //
  // …and a band a line across the page closes is a section of its own, whose
  // columns a word processor balances, level with each other: where the page
  // did not, the break keeps them as the page ended them. canvas.pdf ends its
  // second sheet's left column five lines below its right one, over the line
  // that names its source, and balanced, "Text" and its bar came back at the
  // foot of the left column.
  const line = medianFont * NATURAL_LINE_EM;
  for (const turn of columnTurns) {
    if (turn.closed) {
      if (Math.abs(turn.foot - turn.next) <= line * BALANCED_LINES) continue;
    } else {
      const end = sectionEnds.find((e) => e.at > turn.at);
      const bottom = (end ? setUp(end.from, end.to) : undefined)?.margins?.bottom;
      if (bottom === undefined || turn.foot - bottom < line * SHORT_LINES) continue;
    }
    const brk: Run = { text: '\n', properties: {}, columnBreak: true };
    // Before the next column's first line, where that is a paragraph's: at the
    // end of the column's last, a writer that splits a paragraph at the break
    // opens the next column with an empty line. A drawing anchored to the page
    // takes no room in the column it is written in, and the line is looked for
    // past it: canvas.pdf's right column opens on one, and broken after the
    // left column's last line instead, the right column stood a line low in
    // LibreOffice.
    let first = turn.at;
    while (floating(body[first])) first++;
    const next = body[first];
    const last = body[turn.at - 1];
    if (next?.kind === 'paragraph') {
      body[first] = {
        ...next,
        paragraph: { ...next.paragraph, runs: [brk, ...next.paragraph.runs] },
      };
    } else if (last?.kind === 'paragraph') {
      body[turn.at - 1] = {
        ...last,
        paragraph: { ...last.paragraph, runs: [...last.paragraph.runs, brk] },
      };
    }
  }
  // §17.6.20 — what a page read in its text frame anchors, it anchors on the
  // SHEET: Word stands a drawing at its offsets there, however the section's
  // lines run, so each one is carried onto the sheet and turned with the page.
  let sectionStart = 0;
  for (const end of sectionEnds) {
    const sheet = shown[end.from]?.sheet;
    if (sheet) {
      for (let k = sectionStart; k < end.at; k++) body[k] = floatOntoSheet(body[k]!, sheet.width);
    }
    sectionStart = end.at;
  }
  // The band is built from the page that showed it first, and referenced by
  // every section: the foot runs through the document, not through a section.
  const stepped0 = stepsBetweenWords(allRuns[0] ?? []);
  const edges0 = pageTextEdges(allRuns[0] ?? []);
  // The numeral a band's own first page prints its number as — for the band
  // the numbering was read from.
  const numeralOf = (of: typeof foot): string | undefined => {
    if (of === undefined || of !== numberedBand || numbering === undefined) return undefined;
    const first = of.lift.findIndex((set) => set.size > 0);
    return first >= 0 ? numbering.numbers[first]?.text : undefined;
  };
  const band = foot ? footerBand(foot.band, stepped0, edges0, foot.numbered, numeralOf(foot)) : [];
  const headBand = head
    ? footerBand(head.band, stepped0, edges0, head.numbered, numeralOf(head))
    : [];
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
      withFooter(numberedAs(sectionOnSheet(setUp(0, pages.length), shown[0]), numbering)),
      collectEmbeddedFonts(file, pages, losses),
      sections.map((s) => ({ ...s, properties: withFooter(s.properties) ?? s.properties })),
      band.length > 0 || headBand.length > 0
        ? new Map([
            ...(band.length > 0 ? ([[FOOTER_PART, band]] as const) : []),
            ...(headBand.length > 0 ? ([[HEADER_PART, headBand]] as const) : []),
          ])
        : undefined,
      collectFaceFamilies(file, pages),
      outlines,
      kernedFaces(spacing),
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
  const bands: Array<Gutter> = [];
  for (const [l, r] of sorted) {
    if (l - curEnd >= fontSize * 3) bands.push({ mid: (curEnd + l) / 2, from: curEnd, to: l });
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
  // How many lines a band actually SEPARATES — ink on both sides of it, with a
  // gap no word space explains.
  const straddling = (band: Gutter): number => {
    let n = 0;
    for (const row of rows) {
      const before = row.filter(([, r]) => r <= band.from);
      const after = row.filter(([l]) => l >= band.to);
      if (before.length === 0 || after.length === 0) continue;
      const gap = Math.min(...after.map(([l]) => l)) - Math.max(...before.map(([, r]) => r));
      if (gap >= fontSize * MIN_GUTTER_EM) n++;
    }
    return n;
  };
  // A band no ink crosses is only a gutter if it has text on BOTH sides of it
  // over and over. An invoice has one such band and it separates nothing: its
  // lower half sets the totals on the right and leaves the left empty, so no
  // run crosses the middle of the sheet — read as a column boundary, "Total
  // excluding tax" was laid into a strip ten points wide and came back one
  // letter per line, straight down the page. Two lines that happen to stand
  // apart are not a page set in columns; a handful, repeated, is.
  const empty = bands.filter((b) => straddling(b) >= MIN_EMPTY_BAND_ROWS);
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
  return withoutNarrowColumns(
    separating(
      [...empty, ...voted.filter((v) => !near(v))].sort((a, b) => a.mid - b.mid),
      spans,
    ),
    empty,
    rows,
    [minX, maxX],
    fontSize,
  );
}

/**
 * The gutters left when no column beside a gap the page's lines run across is
 * narrower than a page's column can be.
 *
 * A table's columns stand apart as a page's do, line after line, and a page
 * set in columns of tables has gaps between its tables' columns as well as
 * its own: canvas.pdf sets two columns of Name, Type and Default, and read at
 * every gap its sheets came back in three and five columns, a word wide each,
 * on seven pages where it has two. What gives such a gap away is that the
 * page's other lines — the headings over each table, the prose between them —
 * run across it, and that a column it leaves is a few ems wide: such a gap
 * goes, the one with the less white beside the narrowest column first, until
 * no column is left so narrow. A gap NO line crosses stays however narrow the
 * columns beside it — a page that is one table from edge to edge is read by
 * its columns, as a table — and so does a page's one gap.
 *
 * The gaps that stay are then measured afresh, as the white every line split
 * there leaves: a gutter voted across a column of short cells reaches back
 * into that column, and a section set with it stood its columns sixty-five
 * points further apart than the page does.
 *
 * @param gutters  The gutters found, left to right.
 * @param firm     The gutters no line crosses, which stay.
 * @param rows     Each line's ink across the page, left to right.
 * @param extent   Where the page's ink begins and ends.
 * @param fontSize The page's body size.
 */
function withoutNarrowColumns(
  gutters: ReadonlyArray<Gutter>,
  firm: ReadonlyArray<Gutter>,
  rows: ReadonlyArray<ReadonlyArray<readonly [number, number]>>,
  extent: readonly [number, number],
  fontSize: number,
): Array<Gutter> {
  const kept = [...gutters];
  const least = fontSize * NARROWEST_COLUMN_EM;
  const white = (g: Gutter): number => g.to - g.from;
  let dropped = false;
  // A page split once is two columns however narrow they are: a column a few
  // ems wide is a table's only where there is another column to be part of.
  while (kept.length > 1) {
    const edges = [extent[0], ...kept.flatMap((g) => [g.from, g.to]), extent[1]];
    // The narrowest column with a gap beside it that may go, and that gap.
    let narrowest: { width: number; gap: Gutter } | undefined;
    for (let k = 0; k <= kept.length; k++) {
      const width = edges[2 * k + 1]! - edges[2 * k]!;
      if (width >= least || (narrowest && narrowest.width <= width)) continue;
      const beside = [kept[k - 1], kept[k]].filter(
        (g): g is Gutter => g !== undefined && !firm.includes(g),
      );
      if (beside.length === 0) continue;
      const gap = beside.reduce((a, b) => (white(b) < white(a) ? b : a));
      narrowest = { width, gap };
    }
    if (!narrowest) break;
    kept.splice(kept.indexOf(narrowest.gap), 1);
    dropped = true;
  }
  if (!dropped) return kept;
  return kept.map((g) => {
    if (firm.includes(g)) return g;
    let from = -Infinity;
    let to = Infinity;
    for (const row of rows) {
      // The row's widest white across the band, where the row is split there.
      let widest: readonly [number, number] | undefined;
      let reach = row[0]?.[1] ?? -Infinity;
      for (const [l, r] of row.slice(1)) {
        const overlaps = l > g.from && reach < g.to;
        if (l - reach >= fontSize * MIN_GUTTER_EM && overlaps) {
          if (!widest || l - reach > widest[1] - widest[0]) widest = [reach, l];
        }
        reach = Math.max(reach, r);
      }
      if (!widest) continue;
      from = Math.max(from, widest[0]);
      to = Math.min(to, widest[1]);
    }
    return to > from ? { from, to, mid: (from + to) / 2 } : g;
  });
}

/**
 * How narrow, in ems of the page's body, a column of the page may be: narrower
 * is a table's. The columns canvas.pdf's tables leave are two to six ems wide,
 * and eight is four or five words of prose.
 */
const NARROWEST_COLUMN_EM = 8;

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
 * How many lines a band NO run crosses must separate before it is a gutter
 * rather than an empty half of a page. Lower than the vote's bar, because a
 * band nothing crosses is stronger evidence than a band most lines avoid — but
 * not zero, which is what an invoice's blank left half offers.
 */
const MIN_EMPTY_BAND_ROWS = 8;

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
export function pageTextEdges(
  runs: ReadonlyArray<TextRun>,
): { left: number; right: number } | undefined {
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
  // A listing is one block (see `isCode`), and so is a table (`isTableRow`):
  // where one row of the block reaches across the page, every row of it
  // does. A listing's short lines stand left of the gutter and read as the
  // left column's — comments.pdf's second listing came back with its tail,
  // "...", "side_exit_1:", under the text in the columns — and a table's
  // rows are cut at the gutter where the white between two of its columns
  // stands over it: Figure 13 came back as two tables, one a column.
  // …where one row DOES reach across, or the rows run deeper than two
  // columns' rows fall together by chance: two columns that each set a
  // table are rows of cells on the same baselines too, for as long as their
  // pitches agree. canvas.pdf sets one down either side of the sheet, three
  // rows abreast, and read as one the two came back in a row.
  const inRows = rows.filter((row) => row.some((r) => r.text.trim() !== ''));
  const across = (block: ReadonlyArray<ReadonlyArray<TextRun>>): boolean =>
    block.some((row) => row.some((r) => columnOf.get(r) === SPANNING_COLUMN));
  const blocks = [
    ...blocksOf(inRows, isCode, 1).filter(across),
    ...blocksOf(inRows, isTableRow, LEAST_TABLE_ROWS).filter(
      (block) => across(block) || block.length >= SPANNING_TABLE_ROWS,
    ),
  ];
  for (const block of blocks) {
    for (const row of block) {
      if (row.some((r) => columnOf.get(r) !== SPANNING_COLUMN)) {
        breaks.push(Math.max(...row.map((r) => r.y)));
      }
      for (const run of row) columnOf.set(run, SPANNING_COLUMN);
    }
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

/**
 * §12.5.5 — the runs the page's own content paints, and then the ones an
 * annotation's appearance does: two layers, grouped into lines apart.
 */
function byPainter(runs: ReadonlyArray<TextRun>): Array<ReadonlyArray<TextRun>> {
  const own = runs.filter((r) => r.annotation !== true);
  const annotated = runs.filter((r) => r.annotation === true);
  return [own, annotated].filter((group) => group.length > 0);
}

/**
 * §8.5.3 / §12.5.5 — where a placed line stands among the page's marks, which
 * are ordered as they were painted (see `compareOrder`): the page's own words
 * over everything its content drew, and under every annotation, whose
 * appearances paint after it (`[MAX_SAFE_INTEGER, index, …]`); an
 * annotation's words over all of it. Keyed as one layer, the placed lines
 * fell among the appearances by their count: evaljs.pdf's "Execute" went
 * under the grey of the button it names.
 *
 * @param runs  The layer the line was grouped from (see `byPainter`).
 * @param count How many lines are placed already.
 */
function placedKey(runs: ReadonlyArray<TextRun>, count: number): ReadonlyArray<number> {
  return runs[0]?.annotation === true
    ? [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER, count]
    : [Number.MAX_SAFE_INTEGER - 1, count];
}

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
function groupIntoLines(
  runs: ReadonlyArray<TextRun>,
  split = false,
  stepped = false,
  spaces?: FaceSpaces,
  rules: ReadonlyArray<ColumnRule> = [],
): Array<Line> {
  const sorted = [...runs].sort((a, b) => b.y - a.y || a.x - b.x);
  const clusters: Array<{ y: number; fontSize: number; runs: Array<TextRun> }> = [];
  for (const run of sorted) {
    const last = clusters[clusters.length - 1];
    const tol = Math.max(1, (run.fontSizePt || 10) * 0.5);
    if (last && Math.abs(last.y - run.y) <= tol) {
      last.runs.push(run);
      // The line stands on its type's baseline, not on a mark set over it. Its
      // runs are read from the top of the page down, and a footnote mark is
      // the first of them: comments.pdf's "…back to a double.¹ Clearly, a"
      // stood on the mark's baseline, 3.8 points up, and so ten points over
      // the line after it stood thirteen — a paragraph's gap, and "JavaScript
      // VM that wants to be fast…" came back a paragraph of its own.
      if ((run.fontSizePt || 0) > last.fontSize) last.y = run.y;
      last.fontSize = Math.max(last.fontSize, run.fontSizePt || 0);
    } else {
      clusters.push({ y: run.y, fontSize: run.fontSizePt || 10, runs: [run] });
    }
  }
  for (const c of clusters) c.runs.sort((a, b) => a.x - b.x);
  const stops = sharedStops(
    clusters.map((c) => c.runs),
    rules,
  );
  return clusters.flatMap((c) => {
    const ordered = c.runs;
    const fontSize = c.fontSize || 10;
    if (!split) return [lineOf(ordered, c.y, fontSize, stepped, stops, spaces)];
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
        stops,
        spaces,
      ),
    );
  });
}

/**
 * Whether a block's lines are PROSE: many of them, across a measure a sentence
 * is set to, and most running out to the edge that measure breaks them at, as
 * a column of text does — where a stack of labels and values, or an address,
 * is short lines of their own lengths.
 *
 * @param lines The block's lines, top first.
 * @param own   The block's own left and right edges.
 * @returns True for a column of running text.
 */
function isProse(lines: ReadonlyArray<Line>, own: { left: number; right: number }): boolean {
  if (lines.length < PROSE_LINES) return false;
  // …and a measure wide enough to set a sentence in: an invoice's stack of
  // labels and values is lines of a few words, which end as near its edge as
  // a column's do because the edge is the longest of them.
  const size = median(lines.map((l) => l.fontSize)) || 10;
  if (own.right - own.left < size * PROSE_MEASURE_EM) return false;
  // The last line of a paragraph ends where its words do, so it is not asked.
  const asked = lines.slice(0, -1);
  const full = asked.filter((l) => l.x + l.width >= own.right - l.fontSize * FULL_LINE_EM);
  return full.length >= asked.length * PROSE_FULL_SHARE;
}

/** How many lines a block needs before it can be told for prose. */
const PROSE_LINES = 5;

/** The narrowest measure, in ems, a column of prose is set to. */
const PROSE_MEASURE_EM = 20;

/** How near the block's edge, in ems, a line has to end to have run out to it. */
const FULL_LINE_EM = 2;

/** The share of a block's lines that run out to its edge in a column of prose. */
const PROSE_FULL_SHARE = 0.6;

/**
 * §17.3.1.38 — the stops a block of lines is set out on: where text resumes,
 * after a gap wider than a word space, at the same place on more than one line
 * — and on at least one of them after a gap only a tab makes.
 *
 * A tab's gap is as wide as what stands before it leaves it. An invoice sets
 * "Date due", "Date of issue" and "Invoice number" over one column of values
 * at 108pt, and only the shortest label leaves a gap no space could make: the
 * other two came back with their values a word's width after the label
 * instead of in the column.
 *
 * …or a gap a rule is drawn down. A table ruled into columns needs no white
 * wider than a space between them, the rule says where one ends: comments.pdf's
 * Figure 9 sets its codes "xx1", "000" in a column all as wide as each other,
 * twelve points from the types beside them and a rule between, and the two
 * columns came back one, "xx1 number", with the rule struck through it.
 *
 * @param lines Each line's runs, left to right.
 * @param rules The rules drawn down the block (see {@link columnRules}).
 * @returns The x of every stop two lines share.
 */
function sharedStops(
  lines: ReadonlyArray<ReadonlyArray<TextRun>>,
  rules: ReadonlyArray<ColumnRule> = [],
): Array<number> {
  const seen: Array<{ x: number; line: number; tab: boolean }> = [];
  lines.forEach((runs, line) => {
    let end: number | undefined;
    for (const run of runs) {
      if (run.text.replaceAll(UNMAPPED, '').trim() === '') continue;
      const size = run.fontSizePt || 10;
      if (end !== undefined && run.x - end >= size * STOP_GAP_EM) {
        const from = end;
        const ruled = rules.some(
          (r) => r.x > from && r.x < run.x && r.minY <= run.y + size * 0.5 && r.maxY >= run.y,
        );
        seen.push({ x: run.x, line, tab: ruled || run.x - end >= size * TAB_GAP_EM });
      }
      end = Math.max(end ?? run.endX, run.endX);
    }
  });
  const stops: Array<number> = [];
  for (const tab of seen) {
    if (!tab.tab || stops.some((x) => Math.abs(x - tab.x) <= STOP_SLACK_PT)) continue;
    const lined = new Set(
      seen.filter((o) => Math.abs(o.x - tab.x) <= STOP_SLACK_PT).map((o) => o.line),
    );
    if (lined.size >= 2) stops.push(tab.x);
  }
  return stops;
}

/** The least gap, in ems, that lands a run on a stop the lines around it share. */
const STOP_GAP_EM = 0.5;

/** A rule drawn down a block of lines: where across it stands, and how far down. */
type ColumnRule = { x: number; minY: number; maxY: number };

/**
 * The rules a page draws down its lines: thin upright paths, a line's height
 * long at the least — the edges of a table's columns.
 *
 * @param vectors The page's paths.
 */
function columnRules(vectors: ReadonlyArray<PdfVector>): Array<ColumnRule> {
  return vectors
    .filter(
      (v) =>
        v.glyph !== true &&
        v.maxX - v.minX <= COLUMN_RULE_THIN_PT &&
        v.maxY - v.minY >= COLUMN_RULE_LEAST_PT,
    )
    .map((v) => ({ x: (v.minX + v.maxX) / 2, minY: v.minY, maxY: v.maxY }));
}

/** No thicker than this, a path drawn down the page is a rule. */
const COLUMN_RULE_THIN_PT = 1;

/** …and no shorter than this: a rule stands at least a line of small type tall. */
const COLUMN_RULE_LEAST_PT = 5;

/** How far from a shared stop a run may start and still stand on it. */
const STOP_SLACK_PT = 0.75;

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

/**
 * A document's runs as its listings are read (see {@link isCode}).
 *
 * A typewriter face is a listing's only where the document is set in another:
 * a letter typed throughout, a screenplay, a page set in Courier for its even
 * widths, is PROSE in that face, and its lines run on and its double spaces
 * are spaces. Where most of what the document shows is typewritten, its runs
 * are read as any other face's.
 *
 * @param pages Each page's runs.
 * @returns The same runs, or copies without `fixedPitch` where the typewriter
 *          face is the document's own.
 */
function typewritten(
  pages: ReadonlyArray<ReadonlyArray<TextRun>>,
): ReadonlyArray<ReadonlyArray<TextRun>> {
  let typed = 0;
  let all = 0;
  for (const runs of pages) {
    for (const run of runs) {
      const letters = run.text.replace(/\s/gu, '').length;
      all += letters;
      if (run.fixedPitch === true) typed += letters;
    }
  }
  if (typed === 0 || typed * 2 <= all) return pages;
  return pages.map((runs) =>
    runs.map((run) => {
      if (run.fixedPitch !== true) return run;
      const { fixedPitch: _typed, ...prose } = run;
      return prose;
    }),
  );
}

/**
 * Whether a line is CODE: every word of it set in a typewriter face (§9.8.2).
 *
 * A listing is set line by line with its indents and its comments lined up in
 * columns of the typewriter's cells, and re-set as prose it is no program:
 * comments.pdf's "v0 := ld state[748]" and "st sp[0], v0" ran together as
 * one line of a paragraph, and so did line 4 of Figure 1 and its line 5.
 */
function isCode(runs: ReadonlyArray<TextRun>): boolean {
  const inked = runs.filter((r) => r.text.trim() !== '');
  return inked.length > 0 && inked.every((r) => r.fixedPitch === true);
}

/**
 * Whether the white between two runs of a typewriter face LINES UP what
 * follows it: wider than a cell and a half, where a word space is one cell.
 * A listing puts its comments in a column that way, and a space in their
 * place set each comment against its own line's code.
 */
function aligned(before: TextRun, after: TextRun): boolean {
  if (before.fixedPitch !== true || after.fixedPitch !== true) return false;
  const letters = [...before.text].length;
  if (letters === 0) return false;
  const cell = (before.endX - before.x) / letters;
  return cell > 0 && after.x - before.endX >= cell * ALIGNED_CELLS;
}

/** How many of a typewriter's cells of white line up what follows them. */
const ALIGNED_CELLS = 1.5;

/**
 * A page's runs less its code listings: the rows every word of which is set
 * in a typewriter face (see {@link isCode}).
 *
 * A listing says nothing about the page's columns. comments.pdf sets two
 * across the head of its third page, the code at the left and each comment
 * lined up at one x: the white between them — thirty-five lines of it at the
 * same place — was taken for the gutter of a page in two columns, while the
 * real one, under the listings, was crossed by every comment and not found.
 * The code came back as a column of prose and the comments as another.
 */
function withoutListings(runs: ReadonlyArray<TextRun>): ReadonlyArray<TextRun> {
  if (!runs.some((r) => r.fixedPitch === true)) return runs;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const listed = new Set(rowsOf(runs, fontSize).filter(isCode).flat());
  return listed.size > 0 ? runs.filter((r) => !listed.has(r)) : runs;
}

/**
 * A page's runs less its tables (see {@link isTableRow}): the rows of every
 * block of three and more table rows. The same runs where it has none.
 *
 * The white between a table's columns says nothing about the columns of the
 * text around it. comments.pdf sets Figure 13, a benchmark and nine figures a
 * row, twenty-six rows deep, over two columns of text: its nine gaps were
 * taken for the page's gutters, the text's own was crossed by its rows and
 * not found, and the whole page came back as one table, each line of the text
 * a row of it.
 */
function withoutTables(runs: ReadonlyArray<TextRun>): ReadonlyArray<TextRun> {
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize).filter((row) => row.some((r) => r.text.trim() !== ''));
  const tabled = new Set(blocksOf(rows, isTableRow, LEAST_TABLE_ROWS).flat(2));
  return tabled.size > 0 ? runs.filter((r) => !tabled.has(r)) : runs;
}

/**
 * The rows of a page that stand in runs of at least `least` for which `holds`
 * is true, each run of them a block.
 */
function blocksOf(
  rows: ReadonlyArray<ReadonlyArray<TextRun>>,
  holds: (row: ReadonlyArray<TextRun>) => boolean,
  least: number,
): Array<Array<ReadonlyArray<TextRun>>> {
  const out: Array<Array<ReadonlyArray<TextRun>>> = [];
  for (let from = 0; from < rows.length; ) {
    let to = from;
    while (to < rows.length && holds(rows[to]!)) to++;
    if (to - from >= least) out.push(rows.slice(from, to));
    from = Math.max(to, from + 1);
  }
  return out;
}

/**
 * Whether a row is a TABLE's: four cells or more — its words as far as each
 * gap wider than an em — and none of them as long as a line of prose. A row
 * of a page in columns is two long pieces; comments.pdf's Figure 13 is ten
 * short ones a row, a benchmark and nine figures, twenty-six rows deep.
 */
function isTableRow(runs: ReadonlyArray<TextRun>): boolean {
  const cells = cellsOf(runs);
  if (cells.length < TABLE_CELLS) return false;
  const size = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  return cells.every((c) => c.to - c.from <= size * WIDEST_CELL_EM);
}

/** A row's ink as its cells: its words, as far as each gap wider than an em. */
function cellsOf(runs: ReadonlyArray<TextRun>): Array<Extent> {
  const inks = runs
    .map((run) => ({ run, ink: runInk(run) }))
    .filter((r): r is { run: TextRun; ink: [number, number] } => r.ink !== undefined)
    .sort((a, b) => a.ink[0] - b.ink[0]);
  const cells: Array<{ from: number; to: number }> = [];
  for (const { run, ink } of inks) {
    const last = cells[cells.length - 1];
    if (last && ink[0] - last.to < (run.fontSizePt || 10) * CELL_GAP_EM) {
      last.to = Math.max(last.to, ink[1]);
    } else cells.push({ from: ink[0], to: ink[1] });
  }
  return cells;
}

/** How many cells make a row a table's. */
const TABLE_CELLS = 4;

/** How many such rows make a table. */
const LEAST_TABLE_ROWS = 3;

/**
 * …and how many make one across the page's columns with no row of it
 * reaching over the gutter on its own: more than two columns' rows fall
 * together by chance.
 */
const SPANNING_TABLE_ROWS = 6;

/** The white between two cells of a table row, in ems: wider than a word space ever is. */
const CELL_GAP_EM = 1;

/** The longest a table's cell runs, in ems; a line of a column of prose runs further. */
const WIDEST_CELL_EM = 12;

/** One run of runs, left to right on a shared baseline, as a {@link Line}. */
function lineOf(
  runs: ReadonlyArray<TextRun>,
  y: number,
  fontSize: number,
  stepped: boolean,
  shared: ReadonlyArray<number> = [],
  spaces?: FaceSpaces,
): Line {
  // A page may do both. A LaTeX document writes its prose with spaces in it and
  // sets its mathematics by stepping — TeX's thin space is a sixth of an em and
  // its medium one two ninths, both under the quarter a drawn-space page needs
  // — so bug1997343.pdf came back with "f(x) = sinx+cosx" where the file sets
  // "f(x) = sin x + cos x". A line with no space in it anywhere was stepped
  // across, whatever the rest of the page does.
  const steppedLine = stepped || (runs.length > 1 && !runs.some((r) => SPACE.test(r.text)));
  const { spans: ordered, stops, pieces } = lineSpans(runs, fontSize, steppedLine, shared, spaces);
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
    ...(tabbed(runs, fontSize) || stops.length > 0 ? { tabbed: true as const } : {}),
    ...(stops.length > 0 ? { stops, pieces } : {}),
    ...(isCode(runs) ? { code: true as const } : {}),
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

/**
 * Whether `run` carries on the leader `prev` is a character of: the same
 * character, a step no wider than a word space on from it. A dash a column
 * further on is a cell's: comments.pdf's Figure 13 marks three empty figures
 * "-", a column apart, and joined as one leader they came back "---", one
 * cell standing where three had.
 */
function continuesLeader(prev: TextRun, run: TextRun, fontSize: number): boolean {
  return (
    isLeader(prev.text) &&
    prev.text === run.text &&
    run.x - prev.endX < (run.fontSizePt || fontSize) * CELL_GAP_EM
  );
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
export function stepsBetweenWords(runs: ReadonlyArray<TextRun>): boolean {
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
  shared: ReadonlyArray<number> = [],
  spaces?: FaceSpaces,
): { spans: Array<TextSpan>; stops: Array<number>; pieces: Array<Extent> } {
  const spans: Array<TextSpan> = [];
  const stops: Array<number> = [];
  const pieces: Array<{ from: number; to: number }> = [{ from: Infinity, to: -Infinity }];
  // A table's row (see `isTableRow`) is its cells, each on the stop the page
  // set it at: a gap wider than an em between two of them is a tab, where one
  // narrower than a column's usual white became a space, and the rows of one
  // table came to stand on different numbers of stops.
  const table = isTableRow(runs);
  // A listing's blanks are its own, and stand as they were set.
  const code = isCode(runs);
  // §17.3.2.42 — the line's OWN baseline, which a script stands off. Taken from
  // the runs set at the line's size: the marks are the ones that moved.
  const body = runs.filter((r) => (r.fontSizePt || fontSize) > fontSize * SCRIPT_SIZE);
  const baseline = median((body.length > 0 ? body : runs).map((r) => r.y));
  let prev: TextRun | undefined;
  // Where the last run that MARKS anything ends: a stop is reached across the
  // blanks before it, not from them.
  let inked: number | undefined;
  for (const [i, run] of runs.entries()) {
    // A space that steps nowhere, with the ink on either side of it closed up,
    // is no word space: the page shows no gap for it to stand in. The file
    // named a glyph a space that is not one — bug1046314.pdf maps a Thai
    // mark of no width to U+0020, and "(คำแปล)" came back "(คำ แปล)", a word
    // broken in two. A space of no width that DOES stand in a gap is a word
    // space set by its gap, and stays.
    //
    // …and so is a space the next letter is set down INSIDE, whatever it
    // steps: the page drew it over the ink rather than stepping across it.
    // canvas.pdf writes the space that opens an empty cell at the column the
    // cell stands in, and where the name in the cell before it runs on into
    // that column the space lands between two of its letters — "p" set down
    // a point into it, the ink on either side closed up — and
    // "globalCompositeOperation" came back broken in two. Where the next word
    // starts where the space ends, the space was stepped, however far the run
    // before it claims to reach: freeculture.pdf's runs claim a quarter of an
    // em more than they ink, and its every word space overlaps the word
    // before it by as much.
    if (run.text.trim() === '' && inked !== undefined) {
      const next = runs.slice(i + 1).find((r) => r.text.trim() !== '');
      const over = run.endX - run.x <= 0 || (next !== undefined && next.x < (run.x + run.endX) / 2);
      if (over && next !== undefined && next.x - inked < spaceGap(next, fontSize, true)) continue;
    }
    // §9.10.2 — a glyph the file names no character for is a character this
    // reader cannot write, and dropped it takes its place on the line with it:
    // "6VOBWUGP-0010" came back "6VOBWUGP0010" and "Aug 11 – Sep 11" came back
    // "Aug 11Sep 11". What it stood in is still a gap between two words, and
    // the next run is measured from the last one that says something.
    if (run.text.replaceAll(UNMAPPED, '') === '' && run.text !== '') continue;
    // …and a run the page reached across blanks of its own, landing on a stop
    // the lines around it share, stands on that stop as a run reached across
    // white does: canvas.pdf spaces its way from a method's return type to
    // its name, "CanvasGradient" and two spaces to the column of names every
    // other row reaches with white, and the row came back one run of words,
    // the lines under it run on into it. The blanks were the page's way to
    // the stop, and so is the tab; set before it, they would push it on.
    if (
      !code &&
      prev !== undefined &&
      prev.text.trim() === '' &&
      run.text.trim() !== '' &&
      inked !== undefined &&
      run.x - inked >= (run.fontSizePt || fontSize) * STOP_GAP_EM &&
      shared.some((x) => Math.abs(x - run.x) <= STOP_SLACK_PT)
    ) {
      while (spans.length > 0 && spans[spans.length - 1]!.text.trim() === '') {
        if (spans[spans.length - 1]!.text === '\t') break;
        spans.pop();
      }
      spans.push({ text: '\t' });
      stops.push(run.x);
      pieces.push({ from: Infinity, to: -Infinity });
    } else if (
      prev !== undefined &&
      run.x - prev.endX > spaceGap(prev, fontSize, stepped) &&
      !continuesLeader(prev, run, fontSize)
    ) {
      // A gap no word space could be is a TAB, and the piece after it starts
      // where the page starts it. Written as a space the two pieces close up.
      // So is a narrower one that lands on a stop the lines around it share.
      const size = run.fontSizePt || fontSize;
      const onStop =
        inked !== undefined &&
        run.x - inked >= size * STOP_GAP_EM &&
        shared.some((x) => Math.abs(x - run.x) <= STOP_SLACK_PT);
      if (
        run.x - prev.endX >= size * TAB_GAP_EM ||
        onStop ||
        aligned(prev, run) ||
        (table && run.x - prev.endX >= size * CELL_GAP_EM)
      ) {
        spans.push({ text: '\t' });
        stops.push(run.x);
        pieces.push({ from: Infinity, to: -Infinity });
      } else {
        spans.push(fittedSpace(spaceAfter(spans[spans.length - 1], run), prev, run, spaces));
      }
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
    if (run.text.trim() !== '') {
      inked = Math.max(inked ?? run.endX, run.endX);
      const piece = pieces[pieces.length - 1]!;
      piece.from = Math.min(piece.from, run.x);
      piece.to = Math.max(piece.to, run.endX);
    }
  }
  return { spans, stops, pieces };
}

/** Face name → the advance of the space the document sets it with, in thousandths of an em. */
type FaceSpaces = ReadonlyMap<string, number>;

/**
 * §17.3.2.43 — a word space set as narrow as the page set it.
 *
 * TeX justifies a line by stretching its word spaces or SHRINKING them, and a
 * line it shrank holds more than the same words with the face's own spaces:
 * re-set with those, its last word no longer fits and goes to the next line,
 * and the paragraph grows by a line. comments.pdf sets its 9pt Times with
 * word spaces of 1.70 to 3.44 points against the face's 2.24, and every column
 * of it ran one to four lines long. A space the page set narrower than the
 * face's is set at that share of the face's width (`w:w`, whole percent
 * rounded down, so the line still fits). A wider one is left alone: a line TeX
 * stretched fits with the face's spaces too.
 *
 * The share of the glyph and not a spacing after it: LibreOffice lays no
 * character spacing (§17.3.2.35) after the last character of a run, and a
 * run that is one space has nothing else — Word and LibreOffice both narrow
 * the glyph itself.
 *
 * Only in a face the document carries, whose space is known: a substitute's
 * could be anything.
 *
 * @param space  The space as the line would have it.
 * @param before The run the space follows.
 * @param after  The run it precedes.
 * @param spaces The space each face the document carries is set with.
 * @returns The space, narrowed where the page set it narrower.
 */
function fittedSpace(
  space: TextSpan,
  before: TextRun,
  after: TextRun,
  spaces: FaceSpaces | undefined,
): TextSpan {
  // A space a run brings itself is the one written, and one more beside it
  // collapses into it — which a space narrowed on its own would not.
  if (spaces === undefined || /\s$/u.test(before.text) || /^\s/u.test(after.text)) return space;
  const advance = before.fontName !== undefined ? spaces.get(before.fontName) : undefined;
  const size = space.sizePt ?? before.fontSizePt;
  if (advance === undefined || !(size > 0)) return space;
  const own = (advance / 1000) * size;
  const gap = after.x - before.endX;
  if (gap > own - FIT_SLACK_PT) return space;
  const share = Math.floor((gap / own) * 100) / 100;
  return { ...space, widthScale: Math.max(share, 1 - MOST_SHRINK) };
}

/** A space this much narrower than the face's is the page's rounding, not a shrink. */
const FIT_SLACK_PT = 0.02;

/**
 * The most of a space a line may take back. TeX shrinks a Times space by a
 * quarter of itself and a Computer Modern one by a third — but the space after
 * a word in type is written in the type's face, and comments.pdf sets its
 * `primes` in 9pt Courier-wide cmtt9 and the space after it in Times: 2.25
 * points of a 4.72-point space. A gap narrower still is not a word space.
 */
const MOST_SHRINK = 0.6;

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

/**
 * Whether the line is a RULE drawn out of characters — a row of hyphens, dots
 * or underscores and nothing else.
 *
 * A page with no rule to draw types one. It is not a sentence, it does not run
 * on into the line under it and the line over it does not run on into it.
 *
 * @param line The line.
 * @returns Whether it is a rule rather than words.
 */
function ruleOfCharacters(line: Line): boolean {
  const text = line.text.trim();
  if (text.length < RULE_CHARS) return false;
  for (const ch of text) if (!LEADER_CHARS.has(ch)) return false;
  return true;
}

/** How many of them it takes before a row of marks is a rule. */
const RULE_CHARS = 3;

// Group consecutive lines into paragraphs: a vertical gap well over a single
// line's leading starts a new paragraph. `top` is the paragraph's first (highest) line.
function groupIntoParagraphs(
  lines: ReadonlyArray<Line>,
  column?: { left: number; right: number },
  pageHeight = 0,
  before?: LineBox,
  eachLine = false,
  reach?: number,
): Array<{
  spans: Array<TextSpan>;
  fontSize: number;
  top: number;
  /** The baseline of the paragraph's last line. */
  bottom: number;
  alignment?: 'center' | 'right';
  indentLeft?: number;
  indentFirstLine?: number;
  spacingBefore?: number;
  lineHeight: number;
  stops?: Array<number>;
  pieces?: Array<Extent>;
  /** How many lines the page set the paragraph in, and where the farthest of them ends. */
  lineCount: number;
  right: number;
  /** Its lines, where the white above it was more than the spacing may say (see `setOff`). */
  far?: ReadonlyArray<Line>;
}> {
  const groups: Array<Array<Line>> = [];
  const gaps: Array<number> = [];
  // A column set against both edges runs its every line to the measure but a
  // paragraph's last: there, a line short of it ENDED, where a ragged column's
  // quarter-measure rule (see `endedParagraph`) let a nearly full last line
  // run on into the next paragraph. comments.pdf's "…break even after running
  // a trace 270 times." stops twenty-seven points short, and "The other VMs we
  // compared…" came back joined to it.
  const flushEnds = column
    ? lines.filter((l) => Math.abs(column.right - (l.x + l.width)) <= JUSTIFIED_SLACK_PT).length
    : 0;
  const setJustified =
    column !== undefined &&
    lines.length >= JUSTIFIED_COLUMN_LINES &&
    flushEnds >= lines.length * JUSTIFIED_COLUMN_SHARE;
  let prev: Line | undefined;
  // …and the line before that, whose end says where the block's edge is.
  let beforePrev: Line | undefined;
  for (const line of lines) {
    const gap = prev !== undefined ? prev.y - line.y : 0;
    const opened = prev !== undefined && gap > line.fontSize * 1.5;
    if (
      groups.length === 0 ||
      eachLine ||
      opened ||
      // A line of nothing but rule characters is a RULE, and a rule is its own
      // line: an invoice sets one between the address it asks for cheques at
      // and the sentence above it, and joined to that sentence it came back at
      // the end of "…NOT to our San Francisco office. ----------------------".
      // It takes no line with it either, so what follows opens its own.
      ruleOfCharacters(line) ||
      line.code === true ||
      // …and nothing runs on into a line the page set out on stops, as it
      // runs on into nothing: canvas.pdf sets each method a row of its table,
      // its arguments on the lines under its name, and the next row came back
      // run on from the last argument of the row before, "[Variadic] any
      // args) Object getContext(", the return type torn off its own row.
      line.tabbed === true ||
      (prev !== undefined && opensItem(prev, line)) ||
      (prev !== undefined &&
        (ruleOfCharacters(prev) ||
          prev.code === true ||
          endedParagraph(prev, line, column) ||
          (reach !== undefined && roomForWord(prev, line, reach)) ||
          (setJustified && endedJustified(beforePrev, prev, line, column)) ||
          (column !== undefined && endedCentred(prev, line, column)) ||
          carriesLeader(prev) ||
          prev.tabbed === true))
    ) {
      groups.push([]);
      // The first paragraph stands off whatever the column set above it, where
      // that is known: the region before this one.
      gaps.push(prev === undefined ? (before !== undefined ? before.y - line.y : 0) : gap);
    }
    groups[groups.length - 1]!.push(line);
    beforePrev = prev;
    prev = line;
  }
  // §17.3.1.12 — where the column begins, which is what an indented paragraph
  // is indented from: the edge the .docx sets the column's text against, and
  // so the column's own edge on the page wherever it is known. Not where this
  // region's lines begin: a region is a stretch of a column, and its lines
  // stand in from the column's edge as far as the page stands them —
  // canvas.pdf sets its tables' rows seven points in from the headings of its
  // columns, and ZapfDingbats.pdf its entries thirteen in from its running
  // head, and indented from the rows and entries themselves both came back
  // against the edge of the column. Where no column is given, the low end of
  // the run of line starts: not the leftmost line, which may be a note set in
  // the margin, nor the median, since a page can be half list.
  const starts = lines.map((l) => l.x).sort((a, b) => a - b);
  const columnLeft = column?.left ?? starts[Math.floor(starts.length * COLUMN_LEFT_QUANTILE)] ?? 0;
  const heights: Array<number> = [];
  return groups.map((g, i) => {
    const first = g[0]!;
    const fontSize = Math.max(...g.map((l) => l.fontSize));
    // §17.3.1.33 — the lines stand EXACTLY as far apart as the page stood them.
    // Left to the reader's single spacing they stand as far apart as the face
    // it substitutes says: an invoice sets its 9pt lines 13.5 apart, and in
    // LibreOffice's sans they closed up to 10.4 — every block of the page rose
    // a little more than the one above it, and the rules drawn between them
    // were left behind on the page.
    const inner = g.slice(1).map((l, k) => g[k]!.y - l.y);
    let lineHeight = Math.min(
      Math.max(inner.length > 0 ? median(inner) : fontSize * NATURAL_LINE_EM, fontSize),
      fontSize * TALLEST_LINE_EM,
    );
    const above = i > 0 ? heights[i - 1] : before?.lineHeight;
    const gap = gaps[i] ?? 0;
    // A line alone has no pitch of its own, and takes no taller a box than the
    // gap above it leaves: a box the gap cannot hold pushes it down.
    if (inner.length === 0 && above !== undefined) {
      const room = (gap - (1 - BASELINE_AT) * above) / BASELINE_AT;
      if (room < lineHeight) lineHeight = Math.max(room, fontSize);
    }
    heights.push(lineHeight);
    // The gap that OPENED this paragraph, less the boxes the two lines stand
    // in, is the space its author put before it.
    //
    // Bounded by a third of the sheet, not by three lines: the gap is MEASURED,
    // and three lines is a guess overriding a measurement.
    // annotation-square-circle-without-appearance.pdf sets its two labels two
    // hundred points apart, each over its own pair of drawings, and capped at
    // thirty the second label came back inside the first drawing.
    const opened =
      above !== undefined ? gap - (1 - BASELINE_AT) * above - BASELINE_AT * lineHeight : 0;
    const most = pageHeight > 0 ? pageHeight / 3 : fontSize * 3;
    const spacingBefore = opened > SPACING_NOISE_PT ? Math.min(opened, most) : undefined;
    // §17.3.1.38 — the stops the page set the line out on, measured from where
    // the column begins. A tabbed line ends its paragraph, so these are one
    // line's stops and not a merge of several.
    const stops = g.length === 1 ? (first.stops ?? []) : [];
    // A line that starts where another line of the column starts, or on a stop
    // one of them is set out on, is set FROM there: "walonade@icloud.com"
    // closes the "Bill to" address at the x every line of it starts at, and
    // being short and near the middle of the sheet it read as centred.
    const flush = g.every((l) =>
      lines.some(
        (o) =>
          !g.includes(o) &&
          (Math.abs(o.x - l.x) <= STOP_SLACK_PT ||
            (o.stops ?? []).some((x) => Math.abs(x - l.x) <= STOP_SLACK_PT)),
      ),
    );
    const aligned = flush ? {} : alignmentOf(g, column);
    return {
      spans: joinLines(g),
      fontSize,
      top: first.y,
      bottom: g[g.length - 1]!.y,
      // §17.3.1.38 — a stop is measured from the text area's own left edge, not
      // from where this column's lines happen to start. Measured from the
      // latter, every stop stood a little right of where the page set it, and
      // the last column of a table came out a point wide.
      ...(stops.length > 0
        ? {
            stops: stops.map((x) => x - columnLeft),
            pieces: (first.pieces ?? []).map((p) => ({
              from: p.from - columnLeft,
              to: p.to - columnLeft,
            })),
          }
        : {}),
      ...(spacingBefore !== undefined ? { spacingBefore } : {}),
      lineHeight,
      ...aligned,
      ...indentOf(g, columnLeft, aligned.alignment),
      lineCount: g.length,
      right: Math.max(...g.map((l) => l.x + l.width)),
      ...(opened > most ? { far: g } : {}),
    };
  });
}

/**
 * Whether the first word of `next` would have fit on `prev` short of `reach` —
 * so the page did not break `prev` for want of room, and ended it.
 *
 * The width of the word is its share of its line's: a guess good to a letter,
 * and the margin it is compared with is the rest of the sheet.
 */
function roomForWord(prev: Line, next: Line, reach: number): boolean {
  const text = next.text.trimStart();
  const word = text.split(/\s/u)[0] ?? '';
  if (word.length === 0) return false;
  const width = (next.width * word.length) / text.length;
  return reach - (prev.x + prev.width) > width + prev.fontSize * WORD_SPACE_EM;
}

/**
 * Whether a centred line ended its paragraph: the line under it is centred too,
 * and its first word would have fit on this one. A centred paragraph that wraps
 * fills its lines to the measure but the last, as any paragraph does; lines a
 * page centres one by one do not. comments.pdf centres its authors'
 * affiliations a line apiece — "{gal,brendan,…}@mozilla.com", and under it
 * "Adobe Corporation" — and run together they came back re-wrapped a line
 * shorter, the columns under them risen into the white.
 */
function endedCentred(prev: Line, next: Line, column: { left: number; right: number }): boolean {
  const width = column.right - column.left;
  if (!(width > 0)) return false;
  const centred = (l: Line): boolean => {
    const lead = l.x - column.left;
    const trail = column.right - (l.x + l.width);
    return Math.abs(lead - trail) <= width * CENTRED_SLACK && Math.min(lead, trail) > 0;
  };
  if (!centred(prev) || !centred(next)) return false;
  // …and one of the two set in from both edges, as a full line is not. A
  // list's lines stand in from a column whose left edge is its bullets', and
  // where an overfull line pushed the right edge out they stood in by as much
  // there: comments.pdf's "…We expect to improve performance" and "on this
  // programs by improving…", two full lines of one item, read as centred, and
  // the item came back two paragraphs, its page a line over its sheet.
  const inset = (l: Line): number => Math.min(l.x - column.left, column.right - (l.x + l.width));
  if (Math.max(inset(prev), inset(next)) < width * CENTRED_INSET) return false;
  const text = next.text.trimStart();
  const word = text.split(/\s/u)[0] ?? '';
  if (word.length === 0) return false;
  const wordWidth = (next.width * word.length) / text.length;
  return width - prev.width > wordWidth + prev.fontSize * WORD_SPACE_EM;
}

/** How far off the middle of the measure, as a share of it, a centred line may stand. */
const CENTRED_SLACK = 0.06;

/**
 * How far in from both edges, as a share of the measure, a centred line
 * stands: further than a full line stands in from a column whose edges are
 * its bullets' and an overfull line's (comments.pdf's, a thirtieth at most),
 * and no further than a block of names needs — the paper's second line of
 * authors stands in by a thirteenth, and asked for a tenth, the line under it
 * came back run into it.
 */
const CENTRED_INSET = 0.05;

/**
 * Whether a line of a column set against both edges ended its paragraph: it
 * stops short of the measure, and the line after it is indented or its first
 * word would have fit on it.
 *
 * Not where a line beside it stops at the same place: that is the edge of a
 * narrower block, not the end of the text. A quotation set in from both sides
 * stops short of the measure on every line, and freeculture.pdf's came back a
 * paragraph a line.
 */
function endedJustified(
  before: Line | undefined,
  prev: Line,
  next: Line,
  column: { left: number; right: number },
): boolean {
  const end = prev.x + prev.width;
  if (column.right - end <= prev.fontSize * WORD_SPACE_EM) return false;
  const edge = (l: Line | undefined): boolean =>
    l !== undefined && Math.abs(l.x + l.width - end) <= JUSTIFIED_SLACK_PT;
  if (edge(before) || edge(next)) return false;
  const indented = next.x - prev.x >= prev.fontSize * INDENT_EM;
  return indented || roomForWord(prev, next, column.right);
}

/** How many lines a column needs to show it is set against both edges. */
const JUSTIFIED_COLUMN_LINES = 8;

/** …and how many of them run to the measure, at the least. */
const JUSTIFIED_COLUMN_SHARE = 0.6;

/** How far short of the measure a line of it may end and still run to it: the page's rounding. */
const JUSTIFIED_SLACK_PT = 1;

/** A word space, in ems: what stands between a line's end and the word put after it. */
const WORD_SPACE_EM = 0.3;

/**
 * The exact box a paragraph's first line stands in, where the paragraph is a
 * line of text set to one (see {@link groupIntoParagraphs}).
 */
function leadingLine(el: BodyElement): number | undefined {
  // A figure in the flow (see `figureIn`) is where the page's text begins as
  // much as a line is: its top is where the page set it.
  if (el.kind === 'shape' && el.shape.float === undefined) return 0;
  if (el.kind !== 'paragraph' || el.paragraph.runs.length === 0) return undefined;
  const { spacingLine, spacingLineRule } = el.paragraph.properties;
  return spacingLineRule === 'exact' && spacingLine !== undefined && spacingLine > 0
    ? spacingLine
    : undefined;
}

/** The tallest box a line is given, in ems — past it the gap is paragraph spacing. */
const TALLEST_LINE_EM = 3;

/** Spacing under this is the rounding of the page's own numbers, not white put there. */
const SPACING_NOISE_PT = 0.25;

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
      // A line the page ended with a space has its break written already:
      // joined with another, bug1057544.pdf's column came back "marks the  end
      // of a year's work", a gap twice as wide at every line it had broken.
      else if (!hard && !/\s$/u.test(ends) && !/^\s/u.test(line.spans[0]?.text ?? ''))
        out.push(spaceAfter(prev, line.spans[0]));
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
 * Whether a line opens a list item: it begins with a bullet set apart from the
 * words it marks, or with a label — "[19]", "3." — standing out left of the
 * line over it. Where the item before it ends on a line that runs to the
 * measure, or nearly, nothing else says it ended: comments.pdf's "• We explain
 * how to speculatively generate…" came back run into the item over it, its
 * bullet in the middle of a line, and its reference "[19] M. Zaleski…" into
 * "[18]". A label that does not stand out is a line of prose that begins
 * with one — a citation, a figure.
 *
 * @param prev The line over it.
 * @param line A line of the column.
 */
function opensItem(prev: Line, line: Line): boolean {
  if (BULLET_LEAD.test(line.text)) return true;
  return LABEL_LEAD.test(line.text) && prev.x - line.x >= line.fontSize * HANGING_EM;
}

/** A bullet and the white after it, at the head of a line. */
const BULLET_LEAD = /^[\u2022\u2023\u2043\u2219\u25aa\u25ab\u25cb\u25cf\u25e6\u25a0\u25a1]\s/u;

/** A reference's or a step's label, "[19]" or "3.", and the white after it. */
const LABEL_LEAD = /^(?:\[\d{1,4}\]|\d{1,3}\.)\s/u;

/** How far, in ems, a label stands out left of the line over it to hang. */
const HANGING_EM = 1;

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
 * §17.3.1.12 — how far a line the page set in ONE piece may run on into the
 * right margin, rather than wrap.
 *
 * The line is re-set in a face this reader has, and a face wider than the
 * page's pushes the last word of a line that ran nearly to the measure onto a
 * line of its own: bug1252420.pdf's one italic line came back with
 * "Centuries" under it, and on bug1108301.pdf's sheet, fifty points tall, the
 * word it pushed down fell off the paper altogether. A line the page did not
 * break is not broken: it may run on — so far as a sixth of its measure, and
 * never past the edge of the sheet. Only a line set flush, with nothing on
 * stops: a centred one would move its centre, and a tabbed one its stops.
 *
 * @param para      The paragraph.
 * @param measure   The column it was set in.
 * @param pageWidth The sheet's width.
 * @returns How far it may run past the measure, in points; 0 where it keeps to it.
 */
function runOn(
  para: {
    lineCount: number;
    right: number;
    alignment?: 'center' | 'right';
    stops?: ReadonlyArray<number>;
  },
  measure: { left: number; right: number } | undefined,
  pageWidth: number,
): number {
  if (measure === undefined || para.lineCount !== 1 || para.alignment !== undefined) return 0;
  if ((para.stops?.length ?? 0) > 0) return 0;
  const width = measure.right - measure.left;
  if (!(width > 0) || para.right - measure.left < width * FULL_LINE_SHARE) return 0;
  return Math.max(0, Math.min(width * RUN_ON_SHARE, pageWidth - measure.right));
}

/**
 * §17.3.1.33 — the last text of a sheet, set farther below everything above it
 * than a paragraph's spacing may say, stands where the page set it.
 *
 * The spacing before a paragraph is held to a third of the sheet, so that a
 * gap misread cannot throw everything after it off the page. The last text on
 * the sheet has nothing after it to throw, and held to a third it came back
 * half way up the paper: bug1989304.pdf signs its sheet "World" at the foot
 * and we set it in the middle. Its lines are placed instead, each where the
 * page drew it.
 *
 * @param blocks The sheet's blocks, in the order they are read.
 * @param far    The paragraphs whose spacing was held, with their lines.
 * @param place  A line as the box that places it, `k` counting them.
 */
function setOff(
  blocks: Array<{ band: number; col: number; top: number; el: BodyElement }>,
  far: WeakMap<BodyElement, ReadonlyArray<Line>>,
  place: (line: Line, k: number) => BodyElement,
): void {
  // The last of them as they will be read — the order the page's blocks are
  // sorted into, which they are not in yet.
  const after = (a: (typeof blocks)[number], b: (typeof blocks)[number]): boolean =>
    (a.band - b.band || columnOrder(a.col) - columnOrder(b.col) || b.top - a.top) > 0;
  let at = -1;
  blocks.forEach((b, k) => {
    if (b.el.kind !== 'paragraph' && b.el.kind !== 'table') return;
    if (at < 0 || after(b, blocks[at]!)) at = k;
  });
  const block = at >= 0 ? blocks[at] : undefined;
  const lines = block ? far.get(block.el) : undefined;
  if (!block || !lines) return;
  blocks.splice(
    at,
    1,
    ...lines.map((line, k) => ({
      band: block.band,
      col: block.col,
      top: line.y,
      el: place(line, k),
    })),
  );
}

/** Where the placed last text stands in the page's painting order: over it all. */
const FAR_Z = 1_000_000;

/** How far across its measure a line runs before it is at risk of wrapping when re-set. */
const FULL_LINE_SHARE = 0.85;

/** How far past its measure, as a share of it, a one-piece line may run on. */
const RUN_ON_SHARE = 1 / 6;

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
 * A line of a right-to-left script starts at the RIGHT and ends at the left,
 * and is read the other way round: ArabicCIDTrueType.pdf sets four lines, each
 * flush right and shorter than the one before, and read as a left-to-right
 * setting every pair of them ran together as one paragraph.
 *
 * @param prev   The line before: where it starts, how wide it is, its face,
 *               and its text when it has any.
 * @param next   The line after — where it starts, and how wide it is.
 * @param column The measure both were set in, when it is known.
 * @returns Whether the first line ended a paragraph.
 */
export function endedParagraph(
  prev: { x: number; width: number; fontSize: number; text?: string },
  next: { x: number; width?: number },
  column: { left: number; right: number } | undefined,
): boolean {
  if (!column) return false;
  const width = column.right - column.left;
  if (!(width > 0)) return false;
  const step = Math.max(prev.fontSize, 4);
  if (prev.text !== undefined && next.width !== undefined && isRightToLeft(prev.text)) {
    const indent = prev.x + prev.width - (next.x + next.width);
    if (indent < -step || indent > step * INDENT_LINES) return false;
    return prev.x - column.left > width * 0.25;
  }
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
export function runningFoot(
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
  const textSize = textSizeOf(pageRuns);
  const feet = pageRuns.map((runs, i) =>
    edgeLine(
      runs.filter((r) => !tooSmallToRead(r, textSize)),
      shown[i]?.height ?? 0,
      where,
    ),
  );
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
/**
 * §17.6.12 — how a document's pages are numbered, read off the band that
 * prints their numbers (see ./page-numbers).
 *
 * @param band The running head or foot whose numbers change from page to page.
 * @returns The numbering, or undefined where no band numbers the pages.
 */
export function numberingOf(
  band: { readonly lift: ReadonlyArray<ReadonlySet<TextRun>> } | undefined,
): PageNumbering | undefined {
  if (!band) return undefined;
  return pageNumberingOf(band.lift.map((set) => (set.size > 0 ? bandText([...set]) : undefined)));
}

/** A band's text on one page, with a word space wherever the page shows one. */
function bandText(runs: ReadonlyArray<TextRun>): string {
  return groupIntoLines(runs, false, stepsBetweenWords(runs))
    .map((line) => line.text)
    .join(' ');
}

/**
 * The numbering a section opening at `page` states: the numerals of the
 * sequence the page is counted in, and the number the sequence starts at where
 * it starts there.
 *
 * @param numbering The document's numbering.
 * @param page      The section's first page.
 * @returns The section's `pageNumberFormat` and `pageNumberStart`, as needed.
 */
export function numberedFrom(
  numbering: PageNumbering | undefined,
  page: number,
): Pick<SectionProperties, 'pageNumberFormat' | 'pageNumberStart'> {
  const run = numbering ? runOf(numbering.runs, page) : undefined;
  if (!run) return {};
  return {
    ...(run.format !== 'decimal' ? { pageNumberFormat: run.format } : {}),
    ...(run.from === page && (page > 0 || run.start !== 1) ? { pageNumberStart: run.start } : {}),
  };
}

/** A document of one section, numbered from its first page. */
function numberedAs(
  section: SectionProperties | undefined,
  numbering: PageNumbering | undefined,
): SectionProperties | undefined {
  return section ? { ...section, ...numberedFrom(numbering, 0) } : section;
}

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
  // The head is asked for a page of text under it, as it always was: the
  // lines that OPEN a page are far more often the text itself (see below).
  const head = where === 'head';
  if (runs.length < (head ? HEAD_RUNS : 1) || pageHeight <= 0) return undefined;
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize);
  if (rows.length === 0) return undefined;
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
  // A sheet that holds nothing BUT its foot has no text block to measure a gap
  // against: a receipt's second page is blank but for "Page 2 of 2", and asked
  // for four rows of it the reader found none, could not confirm the foot
  // repeated, and flowed page one's "Page 1 of 2" into the body — where it came
  // back at the TOP of the second sheet.
  // Blanks are not ink here either: both invoices set a space in the top
  // corner of every sheet, and counted as a row it stood outside the band and
  // said the page held something other than its foot.
  const inked = rows.filter((row) => row.some((r) => r.text.trim() !== ''));
  const inBand =
    inked.length > 0 &&
    inked.every((row) =>
      where === 'foot'
        ? edgeY(row) <= pageHeight * FOOT_BAND
        : edgeY(row) >= pageHeight * (1 - FOOT_BAND),
    );
  if (rows.length < 4 && (inBand || head)) {
    const all = inked.flat();
    const text = all
      .map((r) => r.text)
      .join('')
      .trim();
    return inBand && text.length > 0 && text.length <= FOOT_CHARS * inked.length
      ? { y: edgeY(inked[inked.length - 1] ?? []), runs: all }
      : undefined;
  }
  let found: ReadonlyArray<TextRun> | undefined;
  let group: Array<TextRun> = [];
  // Only the foot grows: the lines that OPEN a page are far more often the
  // text itself, and asked for a head of up to three ZapfDingbats.pdf gave up
  // its red note, its running head and the first line of its title.
  const most = head ? 1 : FOOT_ROWS;
  // A foot needs one line of text above it to stand clear of, and no more: a
  // chapter opens on a sheet of its own, its heading with the page number
  // under it, and asked for more basicapi.pdf's "page 2 / 3" stayed in the
  // body and came back half way up the sheet. The head keeps asking for more —
  // a chapter's heading, alone at the top of every sheet, is not the
  // document's running head.
  const above = head ? 2 : 1;
  for (let i = 0; i < Math.min(most, rows.length - above); i++) {
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

/** How many runs a page must hold before its first line is asked whether it is a running head. */
const HEAD_RUNS = 4;

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
 * @param run     The footer's run.
 * @param numeral The page's number as this band prints it, where it is known.
 * @returns The run, or the two or three it is cut into.
 */
function pageNumbered(run: Run, numeral?: string): Array<Run> {
  // The numeral the page's number is printed as — the one that changes from
  // page to page, which may be roman (see ./page-numbers) — or the first run
  // of digits where the bands said nothing more.
  const pattern =
    numeral !== undefined
      ? new RegExp(`(^|\\s)(${numeral})(\\s|$)`, 'u')
      : /(^|\s)(\d{1,4})(\s|$)/u;
  const found = pattern.exec(run.text);
  if (!found || run.field !== undefined) return [run];
  const at = found.index + found[1]!.length;
  const number = found[2]!;
  const before = run.text.slice(0, at);
  const after = run.text.slice(at + number.length);
  // §17.16.5.33 — and the number after it is how many sheets there are:
  // "Page 1 of 2" is two fields with a word between them, not one field and a
  // literal 2 that stays 2 on every page of a longer document.
  const total = /(^|\s)(\d{1,4})(\s|$)/u.exec(after);
  const tail =
    total === null
      ? after === ''
        ? []
        : [{ ...run, text: after }]
      : [
          ...(after.slice(0, total.index + total[1]!.length) === ''
            ? []
            : [{ ...run, text: after.slice(0, total.index + total[1]!.length) }]),
          { ...run, text: total[2]!, field: 'NUMPAGES' as const },
          ...(after.slice(total.index + total[1]!.length + total[2]!.length) === ''
            ? []
            : [
                {
                  ...run,
                  text: after.slice(total.index + total[1]!.length + total[2]!.length),
                },
              ]),
        ];
  return [
    ...(before === '' ? [] : [{ ...run, text: before }]),
    { ...run, text: number, field: 'PAGE' as const },
    ...tail,
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
 * @param numeral  That number as this band prints it, where it is known.
 */
export function footerBand(
  runs: ReadonlyArray<TextRun>,
  stepped: boolean,
  measure: { left: number; right: number } | undefined,
  numbered: boolean,
  numeral?: string,
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
        runs: numbered
          ? el.paragraph.runs.flatMap((run) => pageNumbered(run, numeral))
          : el.paragraph.runs,
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
    // A single line is set against the far edge when that is where it ends
    // and it does not start at the near one: "Page 1 of 2" at the foot of a
    // receipt. A paragraph's test wants a RAGGED left, which one line has not,
    // and the foot came back at the left margin.
    const only = lines.length === 1 ? lines[0] : undefined;
    const width = measure ? measure.right - measure.left : 0;
    const flushRight =
      only !== undefined &&
      measure !== undefined &&
      measure.right - (only.x + only.width) <= width * BAND_RIGHT_SHARE &&
      only.x - measure.left >= width * BAND_RIGHT_SHARE;
    const alignment = flushRight ? 'right' : alignmentOf(lines.slice(0, 1), measure).alignment;
    const spans = lines.flatMap((line, i) =>
      i === 0 ? line.spans : [spaceAfter(lines[i - 1]!.spans.at(-1), line.spans[0]), ...line.spans],
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
 * Whether the regions the gutters cut the page into are COLUMNS OF PROSE — the
 * only thing worth re-setting a document in columns for.
 *
 * A gutter is evidence about the page's white space, and white space is not
 * enough: an invoice sets its labels along the left and its amounts against the
 * right margin, and a dozen lines then break at the same x. Read as two columns
 * the amounts became a column of their own and every label was indented past
 * the strip it was given — "Total excluding tax" came back one letter per line.
 *
 * What tells them apart is the same thing that tells a cell from a line of
 * prose: a column of prose FILLS its measure, over and over, because that is
 * what wrapping does. A column of values never fills anything — its lines are
 * short and it is their right edges that line up, not their left. So every
 * region must hold lines of its own and they must reach across it.
 *
 * @param runs    The page's runs.
 * @param gutters The page's gutters.
 * @param edges   Where the page's text starts and ends.
 * @returns Whether the page is set in columns of prose.
 */
function proseColumns(
  runs: ReadonlyArray<TextRun>,
  gutters: ReadonlyArray<Gutter>,
  edges: { left: number; right: number } | undefined,
): boolean {
  if (gutters.length === 0) return false;
  if (!edges) return true; // Nothing to measure against: leave the reading alone.
  const fontSize = median(runs.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const rows = rowsOf(runs, fontSize).filter((row) => row.length > 0);
  if (rows.length === 0) return true;
  const bounds = columnBounds(runs, gutters, edges);
  const regions = bounds.slice(0, -1).map((lo, i) => [lo, bounds[i + 1]!] as const);
  const regionOf = (x: number): number => {
    for (let i = regions.length - 1; i >= 0; i--) if (x >= regions[i]![0]) return i;
    return 0;
  };
  const edgesIn = regions.map((): Array<{ left: number; right: number }> => []);
  for (const row of rows) {
    const byRegion = regions.map((): Array<TextRun> => []);
    for (const run of row) byRegion[regionOf(run.x)]!.push(run);
    byRegion.forEach((inRegion, i) => {
      if (inRegion.length === 0) return;
      edgesIn[i]!.push({
        left: Math.min(...inRegion.map((r) => r.x)),
        right: Math.max(...inRegion.map((r) => r.endX)),
      });
    });
  }
  // How many lines share an edge, which is what "flush" means.
  const agreeing = (xs: ReadonlyArray<number>): number => {
    let most = 0;
    for (const x of xs) most = Math.max(most, xs.filter((y) => Math.abs(y - x) <= FLUSH_PT).length);
    return most;
  };
  return edgesIn.every((lines, i) => {
    if (lines.length < MIN_PROSE_LINES) return false;
    // Prose is set flush LEFT and comes out ragged right: line after line
    // starts in the same place and ends wherever its last word ends. A column
    // of figures is the other way round — an invoice's amounts agree on their
    // right edge and on nothing else — and that is not a column to re-set a
    // document in.
    // …or it is JUSTIFIED, and flush on both sides: every full line ends
    // where the column does, and only a paragraph's first line starts
    // anywhere else. Its lines fill the column, which a column of amounts
    // never does. comments.pdf's pages agree on more right edges than left
    // ones, and five of its fourteen were read straight across both columns.
    const [lo, hi] = regions[i]!;
    const full = lines.filter((l) => l.right - l.left >= (hi - lo) * FULL_LINE_SHARE).length;
    return (
      agreeing(lines.map((l) => l.right)) <= agreeing(lines.map((l) => l.left)) ||
      full >= lines.length * FILLED_SHARE
    );
  });
}

/**
 * The drawn glyphs that are STRAY marks in a line of ordinary text.
 *
 * §9.6.6 — a glyph the file states no character for is drawn, because nothing
 * downstream can write it. That is the whole content of some pages: a subset
 * that names its glyphs `g18`, a program with no `cmap`, a font of pictures.
 * On such a page the drawings ARE the words and they stay.
 *
 * A single mark among readable words is the other case, and Stripe's invoices
 * are full of them — a `/ToUnicode` stating U+0000 for every piece of
 * punctuation it sets. The flowing reading has nowhere to put a shape inside a
 * line, so the colon of "Kazakhstan VAT: 86-1696045" was placed as a floating
 * drawing and landed a word away, on a line of its own, pushing the text
 * around it aside. Dropped, the line keeps its words and the loss report still
 * says a character was unrecoverable.
 *
 * The line decides: where nearly everything on it is drawn, the drawings are
 * the text — bug1151216.pdf sets three lines of prices that way. Where one or
 * two marks stand among several readable runs, they are punctuation.
 *
 * @param vectors The page's painted paths.
 * @param runs    The page's runs.
 * @returns The drawn glyphs to leave out of a flowing reading.
 */
function strayMarks(
  vectors: ReadonlyArray<PdfVector>,
  runs: ReadonlyArray<TextRun>,
): ReadonlySet<PdfVector> {
  const glyphs = vectors.filter((v) => v.glyph === true);
  const out = new Set<PdfVector>();
  if (glyphs.length === 0) return out;
  const readable = runs.filter((r) => r.text.trim() !== '' && !r.text.includes(UNMAPPED));
  for (const v of glyphs) {
    const mid = (v.minY + v.maxY) / 2;
    // The line's size, not the mark's: a hyphen is a tenth of an em tall, and
    // measured by its own height no word stood near enough to its line to
    // make it a mark in one — every hyphen of an invoice stayed a drawing.
    const words = readable.filter((r) => Math.abs(r.y - mid) <= (r.fontSizePt || 10));
    const size =
      words.length > 0 ? median(words.map((r) => r.fontSizePt || 10)) : v.maxY - v.minY || 10;
    // Ink beside it on the line is what makes it a mark IN that line rather
    // than a drawing of its own. On both sides where the mark is inside a word
    // — the hyphen of a postcode — and on one where the page ends its line with
    // it: "PAYMENT ADDRESS:" states nothing for its colon either, and kept, the
    // colon came back a word to the right of the line below.
    const before = words.some((r) => r.endX <= v.minX + size);
    const after = words.some((r) => r.x >= v.maxX - size);
    if (!before && !after) continue;
    const drawnHere = glyphs.filter((g) => Math.abs((g.minY + g.maxY) / 2 - mid) <= size).length;
    if (drawnHere <= MOST_STRAY_MARKS && words.length >= LEAST_READABLE_RUNS) out.add(v);
  }
  return out;
}

/**
 * §9.10.2 — the runs of a line with each glyph the file names no character for
 * read off the shape it draws (see {@link punctuationOf}), where it is one of
 * the stray marks inside a readable line.
 *
 * @param runs   The page's runs.
 * @param strays The drawn glyphs that stand inside lines of readable text.
 * @returns The runs, with every mark that could be read written as text.
 */
function readDrawnMarks(
  runs: ReadonlyArray<TextRun>,
  strays: ReadonlySet<PdfVector>,
): Array<TextRun> {
  if (strays.size === 0) return [...runs];
  const marks = [...strays];
  return runs.map((run) => {
    if (!run.text.includes(UNMAPPED) || run.text.replaceAll(UNMAPPED, '').trim() !== '') {
      return run;
    }
    const size = run.fontSizePt || 10;
    const mark = marks.find((v) => {
      const x = (v.minX + v.maxX) / 2;
      return (
        x >= run.x - 0.5 && x <= run.endX + 0.5 && Math.abs((v.minY + v.maxY) / 2 - run.y) <= size
      );
    });
    const read = mark ? punctuationOf(mark, run.y, size) : undefined;
    return read === undefined ? run : { ...run, text: run.text.replaceAll(UNMAPPED, read) };
  });
}

/**
 * §9.6.6 — the glyphs a page draws for want of characters, gathered into the
 * WORDS they stand in: a traced glyph that stands against the one before it,
 * on its line and in its colour, is drawn as one shape with it.
 *
 * Every mark kept apart is a shape of its own in the document, anchored to
 * the page — and a paper set in a subset whose glyphs name nothing is nothing
 * but such marks: TAMReview.pdf came back as forty-two thousand shapes, a
 * package no reader would open in under three minutes. The picture is the
 * same either way.
 *
 * @param vectors The page's painted paths, in painting order.
 * @returns The same paths, each run of traced glyphs made one.
 */
export function drawnWords(vectors: ReadonlyArray<PdfVector>): Array<PdfVector> {
  const out: Array<PdfVector> = [];
  for (const v of vectors) {
    const last = out[out.length - 1];
    const size = v.maxY - v.minY;
    if (
      v.glyph === true &&
      last?.glyph === true &&
      last.fillHex === v.fillHex &&
      last.strokeHex === v.strokeHex &&
      last.alpha === v.alpha &&
      // On one line: the two overlap across most of the shorter one's height.
      Math.min(last.maxY, v.maxY) - Math.max(last.minY, v.minY) >
        Math.min(size, last.maxY - last.minY) * 0.3 &&
      // …and next to each other, a word space at the most apart.
      v.minX >= last.minX &&
      v.minX - last.maxX <= Math.max(size, last.maxY - last.minY) * WORD_GAP_EM
    ) {
      out[out.length - 1] = {
        ...last,
        segs: [...last.segs, ...v.segs],
        minX: Math.min(last.minX, v.minX),
        minY: Math.min(last.minY, v.minY),
        maxX: Math.max(last.maxX, v.maxX),
        maxY: Math.max(last.maxY, v.maxY),
      };
      continue;
    }
    out.push(v);
  }
  return out;
}

/** How far apart, in heights of the taller glyph, two drawn glyphs may stand and be one word. */
const WORD_GAP_EM = 0.6;

/** How few drawn glyphs a line may hold before they are its text, not marks in it. */
const MOST_STRAY_MARKS = 2;

/** …and how many readable runs must stand on that line for them to be strays. */
const LEAST_READABLE_RUNS = 4;

/**
 * §17.4.38 — the consecutive lines a page set out on the SAME stops, as the
 * TABLE they are.
 *
 * A line broken at a gap no word space explains is a line the page laid out
 * (see {@link tabbed}); several of them one under another, broken in the same
 * places, is a table's head and its rows. An invoice's item table is exactly
 * that, and so is the payment history under it.
 *
 * Two stops at least, which is three columns: ONE stop is a contents entry with
 * its page number at the measure, and a list of those is a list.
 *
 * The stops do not have to agree to the point — a heading sits over its column
 * and a figure is set against the far side of it — so they are matched within
 * an em, and the column is put where the rows agree it is.
 *
 * @param paras   The column's paragraphs, in order.
 * @param measure The measure they are set across.
 * @returns For each paragraph that belongs to a table: the table at its first
 *          row, and `null` at the rows after it (their content is inside it).
 */
function tabbedRows(
  paras: ReadonlyArray<{
    spans: Array<TextSpan>;
    stops?: Array<number>;
    fontSize: number;
    top: number;
    spacingBefore?: number;
    lineHeight: number;
    pieces?: Array<Extent>;
  }>,
  measure: { left: number; right: number } | undefined,
): Map<number, BodyElement | null> {
  const out = new Map<number, BodyElement | null>();
  if (!measure) return out;
  for (let i = 0; i < paras.length; ) {
    const stops = paras[i]?.stops ?? [];
    if (stops.length < LEAST_TABLE_STOPS) {
      i++;
      continue;
    }
    const near = (a: Array<number>, b: Array<number>): boolean =>
      a.length === b.length &&
      a.every((x, k) => Math.abs(x - b[k]!) <= Math.max(paras[i]!.fontSize, TABLE_STOP_SLACK_PT));
    let to = i + 1;
    while (to < paras.length && near(stops, paras[to]?.stops ?? [])) to++;
    if (to - i < 2) {
      i++;
      continue;
    }
    const rows = paras.slice(i, to);
    // Where the EARLIEST row begins each column, and the last one runs to the
    // measure. Not the middle of them: a column is as wide as its widest cell
    // needs, and a heading set over a column of figures starts further left
    // than the figures do — measured to the middle, "Receipt number" came back
    // one letter per line down a thirty-point strip.
    const bounds = stops.map((_, k) => Math.min(...rows.map((r) => r.stops![k]!)));
    const width = measure.right - measure.left;
    // The ink each column covers, over every row; `undefined` where a row
    // does not say (a cell with nothing in it).
    const inks = [0, ...bounds].map((_, k): Extent | undefined => {
      const cells = rows.map((r) => r.pieces?.[k]);
      if (cells.some((c) => c === undefined || !Number.isFinite(c.from))) return undefined;
      return {
        from: Math.min(...cells.map((c) => c!.from)),
        to: Math.max(...cells.map((c) => c!.to)),
      };
    });
    // …and the HEADINGS set over the columns, where the line above the rows
    // stands over them: comments.pdf's Figure 13 heads its figures and leaves
    // the column of names bare, one stop short of every row under it, and the
    // headings came back as a line of their own above the table.
    // Close over the rows, and over half their columns at the least: a line
    // of an address with two pieces over a table of five heads nothing.
    const above = i > 0 && !out.has(i - 1) ? paras[i - 1] : undefined;
    const close =
      above !== undefined && above.top - rows[0]!.top <= rows[0]!.lineHeight * HEADING_LINES;
    const columnsFor = (heading: Heading | undefined) => {
      // §17.3.1.13 — a column whose cells END together and start apart is set
      // against its right edge: the "Qty", "Tax" and "Amount" of an item table,
      // figures and headings alike. Set from the left the "1" stood under the Q
      // of "Qty", eight points from the figure the page puts under its y.
      const flush = inks.map((ink, k) => {
        if (k === 0 || ink === undefined) return false;
        const tos = rows.map((r) => r.pieces![k]!.to);
        const froms = rows.map((r) => r.pieces![k]!.from);
        const even = Math.max(...tos) - Math.min(...tos);
        if (even > FLUSH_SLACK_PT) return false;
        if (Math.max(...froms) - Math.min(...froms) > even + FLUSH_SLACK_PT) return true;
        // Cells all as wide as each other say nothing of the side they are set
        // against, and the heading over them does: Figure 13's "Flushes" ends
        // where its column of noughts ends, and begins a word further left.
        const head = heading?.over[k];
        return (
          head !== undefined &&
          Math.abs(head.to - Math.max(...tos)) <= FLUSH_SLACK_PT &&
          Math.min(...froms) - head.from > FLUSH_SLACK_PT
        );
      });
      // What each column has to hold: its figures, and the heading over them.
      // Cut to the figures, "Traces/Tree" came back "Traces/" over "Tree".
      const holds = inks.map((ink, k): Extent | undefined => {
        const head = heading?.over[k];
        return ink !== undefined && head !== undefined
          ? { from: Math.min(ink.from, head.from), to: Math.max(ink.to, head.to) }
          : ink;
      });
      // Where each column begins. One set from its left begins at its stop; one
      // set against its right has no stop to begin at, and begins halfway across
      // the white before it — a cell only as wide as the page's figure wraps the
      // figure the moment the face it is re-set in runs a little wider.
      const edges = [0, ...bounds, width];
      for (let k = 1; k < bounds.length + 1; k++) {
        const before = holds[k - 1];
        const own = holds[k];
        if (!flush[k] || before === undefined || own === undefined) continue;
        const mid = (before.to + own.from) / 2;
        if (mid > edges[k - 1]! && mid < own.from) edges[k] = mid;
      }
      // §17.4.65 — and the first begins where its words do: comments.pdf centres
      // Figure 13 on the page, its names seventeen points in from the margin.
      const lead = holds[0]?.from ?? 0;
      if (lead > FLUSH_SLACK_PT && lead < edges[1]!) edges[0] = lead;
      const flushRight = flush.map((right, k) =>
        right ? Math.max(0, edges[k + 1]! - inks[k]!.to) : undefined,
      );
      // …and a heading the columns cannot hold is no heading of theirs: a
      // form's "Bankverbindung:" heads a label and its value together, and
      // set over the labels alone it broke in two.
      const held =
        heading?.over.every(
          (head, k) =>
            head === undefined ||
            edges[k + 1]! - edges[k]! - (flushRight[k] ?? 0) >=
              head.to - head.from - FLUSH_SLACK_PT,
        ) ?? true;
      return { holds, edges, flushRight, held };
    };
    const offered = above && close ? headingOver(above, inks) : undefined;
    const withHeading = offered ? columnsFor(offered) : undefined;
    const heading = withHeading?.held === true ? offered : undefined;
    const { holds, edges, flushRight } =
      heading && withHeading ? withHeading : columnsFor(undefined);
    const grid = edges.slice(0, -1).map((from, k) => pt(Math.max(edges[k + 1]! - from, 1)));
    const tableRows = above && heading ? [{ ...above, spans: heading.spans }, ...rows] : rows;
    const table: Table = {
      // §17.4.63/§17.4.72 — as wide as its columns, each cell as wide as its
      // column: the widths the grid gives, stated where a reader looks first.
      properties: {
        defaultCellMargins: { left: pt(0), right: pt(0) },
        layout: 'fixed',
        widthType: 'dxa',
        widthPt: pt(grid.reduce((sum, w) => sum + w, 0)),
        ...(edges[0]! > 0 ? { indentPt: pt(edges[0]!) } : {}),
      },
      grid,
      rows: tableRows.map((row, r) => {
        // The row stands as far from the next as the page stood it, and the
        // white BEFORE the table is the first row's own: a table has no
        // spacing of its own to carry it, and glued to the block above it the
        // invoice's item table came up against the address over it.
        const prev = tableRows[r - 1];
        // The white a row keeps from the one above it, less the boxes the two
        // lines stand in — the same measure a paragraph's spacing is read by.
        // Stated as the ROW's height instead, LibreOffice set the rows solid
        // and an invoice's heading sat on the item under it.
        const pitch = prev
          ? prev.top - row.top - (1 - BASELINE_AT) * prev.lineHeight - BASELINE_AT * row.lineHeight
          : 0;
        const opening = r === 0 ? row.spacingBefore : pitch > 0 ? pitch : undefined;
        return {
          properties: {},
          cells: splitAtTabs(row.spans).map((cell, k) => {
            const inset = flushRight[k];
            const own = grid[k];
            return {
              properties: own !== undefined ? { width: own } : {},
              content: [
                paragraphFromRuns(cell, undefined, {
                  ...(opening !== undefined ? { spacingBefore: pt(opening) } : {}),
                  spacingLine: pt(row.lineHeight),
                  spacingLineRule: 'exact',
                  ...(inset !== undefined
                    ? { alignment: 'right' as const, indentRight: pt(inset) }
                    : {}),
                }),
              ],
            };
          }),
        };
      }),
    };
    const ink = holds.filter((h): h is Extent => h !== undefined);
    tableReads.set(table, {
      rows: tableRows.map((row) => ({ top: row.top, lineHeight: row.lineHeight })),
      from: measure.left + Math.min(...ink.map((h) => h.from)),
      to: measure.left + Math.max(...ink.map((h) => h.to)),
      left: measure.left,
    });
    out.set(heading ? i - 1 : i, { kind: 'table', table });
    if (heading) out.set(i, null);
    for (let k = i + 1; k < to; k++) out.set(k, null);
    i = to;
  }
  return out;
}

/**
 * A line's cells as a heading row over a table's columns: each cell over the
 * column whose figures it overlaps most, left to right, and an empty cell over
 * a column it leaves bare. `undefined` where the line is no such row.
 *
 * @param line The line above the table, with the cells it was set out in.
 * @param inks The ink of each of the table's columns.
 * @returns The line's spans with a tab between every two columns, and the ink
 *          of the heading over each column.
 */
function headingOver(
  line: { spans: Array<TextSpan>; pieces?: Array<Extent> },
  inks: ReadonlyArray<Extent | undefined>,
): Heading | undefined {
  const cells = splitAtTabs(line.spans);
  const pieces = line.pieces ?? [];
  if (
    cells.length < 2 ||
    pieces.length !== cells.length ||
    cells.length > inks.length ||
    cells.length * 2 < inks.length
  ) {
    return undefined;
  }
  const columns: Array<number> = [];
  for (const piece of pieces) {
    let best = -1;
    let most = 0;
    inks.forEach((ink, k) => {
      if (ink === undefined) return;
      const overlap = Math.min(piece.to, ink.to) - Math.max(piece.from, ink.from);
      if (overlap > most) {
        most = overlap;
        best = k;
      }
    });
    if (best <= (columns[columns.length - 1] ?? -1)) return undefined;
    columns.push(best);
  }
  return {
    spans: inks.flatMap((_, k): Array<TextSpan> => {
      const own = columns.indexOf(k);
      return [...(k > 0 ? [{ text: '\t' }] : []), ...(own >= 0 ? cells[own]! : [])];
    }),
    over: inks.map((_, k) => {
      const own = columns.indexOf(k);
      return own >= 0 ? pieces[own] : undefined;
    }),
  };
}

/** A line set over a table's columns as its heading row (see `headingOver`). */
type Heading = { spans: Array<TextSpan>; over: Array<Extent | undefined> };

/**
 * Where a table read off the page stood on it (see `tabbedRows`): each row's
 * baseline and the box its line stands in, top to bottom, and how far across
 * the page its words reach — which is what a rule drawn between two of its
 * rows is measured against (see `ruleBorders`).
 */
const tableReads = new WeakMap<Table, TableRead>();

/** Where a table read off the page stood on it (see {@link tableReads}). */
type TableRead = {
  rows: ReadonlyArray<{ top: number; lineHeight: number }>;
  /** How far across the page its words reach. */
  from: number;
  to: number;
  /** Where on the page the measure its indent is stated from begins. */
  left: number;
};

/** How many of its rows' lines a heading may stand over a table and head it. */
const HEADING_LINES = 3;

/** How far apart the ends of a column's cells may stand and still be flush. */
const FLUSH_SLACK_PT = 2;

/** How many stops a line must stand on before a run of them is a table. */
const LEAST_TABLE_STOPS = 2;

/**
 * How far two rows' stops may stand apart and still be one column.
 *
 * A heading is set over its column and a figure against the far side of it: an
 * invoice's "Qty" stands eight points left of the "1" under it, and the heading
 * is set two sizes smaller, so its own em is not a wide enough tolerance.
 */
const TABLE_STOP_SLACK_PT = 10;

/** One line's spans cut into its cells, at the tabs the page set out. */
function splitAtTabs(spans: ReadonlyArray<TextSpan>): Array<Array<TextSpan>> {
  const cells: Array<Array<TextSpan>> = [[]];
  for (const span of spans) {
    if (span.text === '\t') cells.push([]);
    else cells[cells.length - 1]!.push(span);
  }
  return cells;
}

/**
 * Give each separator rule to the paragraph it separates, as that paragraph's
 * own border (§17.3.1.24), and say which rules were given away.
 *
 * A rule sits in the white between two blocks: under a table's headings, over a
 * total. The block BELOW it takes it as a top border, because that is the block
 * the rule introduces; where nothing follows closely enough, the block above
 * takes it as a bottom one.
 *
 * A rule drawn down one column is a border of a block in that column, and one
 * drawn across the columns of a block across them: comments.pdf rules off
 * Figure 9's caption in the right column, and taken by the nearest line under
 * it on the page, the rule came back over "Every time the trace recorder
 * emits…" in the left column, and none over the caption.
 *
 * @param vectors The page's painted paths.
 * @param blocks  The blocks read off the page so far, which the rule joins.
 * @param width   The page's width, which a rule is long relative to.
 * @param colOf   The column an x across the page stands in.
 * @returns The rules that became borders, and so must not be drawn again.
 */
function ruleBorders(
  vectors: ReadonlyArray<PdfVector>,
  blocks: Array<Block>,
  width: number,
  colOf: (x: number) => number = () => 0,
): ReadonlySet<PdfVector> {
  const given = new Set<PdfVector>();
  const paragraphs = blocks.filter((b) => b.el.kind === 'paragraph');
  const tables = blocks.filter((b) => b.el.kind === 'table');
  if (paragraphs.length === 0 && tables.length === 0) return given;
  // A rule drawn in PIECES is one rule. An invoice draws the rule under its
  // headings cell by cell — five bars on one baseline, one under each column —
  // and measured apart only the widest was long enough to be a rule: it became
  // the row's border and moved with the words, and the other four stayed
  // where the page drew them, struck through the figures when the row moved.
  for (const band of collinear(vectors.filter((v) => flatBar(v)))) {
    const v = band.pieces[0]!;
    if (band.to - band.from < width * RULE_SHARE) continue;
    const y = (v.minY + v.maxY) / 2;
    const border = {
      style: 'single' as const,
      width: pt(Math.max(v.lineWidth ?? v.maxY - v.minY, RULE_MIN_PT)),
      colorHex: v.strokeHex ?? v.fillHex ?? '000000',
    };
    // A rule between two rows of a table, across it, is the top edge of the
    // lower row's cells (§17.4.39): Figure 13's rule under its headings stayed
    // where the page drew it, and the table, set a little lower, ran its
    // headings through it. Only where it stands ON the edge: an invoice's
    // rule under its headings stands in the white between them and the item,
    // and moved onto the item's edge it rose nine points.
    const seat = seatOf(tables, band, y);
    if (seat !== undefined && seat.block.el.kind === 'table') {
      seat.block.el = {
        kind: 'table',
        table: ruledRow(seat.block.el.table, seat.row, border, band),
      };
      for (const piece of band.pieces) given.add(piece);
      continue;
    }
    if (paragraphs.length === 0) continue;
    // The column the rule is drawn in, or across them.
    const first = colOf(band.from + RULE_INSET_PT);
    const col = first === colOf(band.to - RULE_INSET_PT) ? first : SPANNING_COLUMN;
    const own = paragraphs.filter((b) => b.col === col);
    // The block the rule introduces: the nearest one under it. Failing that,
    // the one it closes off above.
    const below = own.filter((b) => b.top < y).sort((a, b) => b.top - a.top)[0];
    const above = own.filter((b) => b.top >= y).sort((a, b) => a.top - b.top)[0];
    const side =
      below !== undefined && y - below.top <= RULE_REACH_PT
        ? ({ block: below, edge: 'top' } as const)
        : above !== undefined && above.top - y <= RULE_REACH_PT
          ? ({ block: above, edge: 'bottom' } as const)
          : undefined;
    if (!side || side.block.el.kind !== 'paragraph') continue;
    const { paragraph } = side.block.el;
    // A rule takes the room it is drawn in (§17.3.1.24): stood over a line,
    // it pushes the line down by its own width, and five rules over five
    // totals put the last one four points below where the page has it.
    const before = paragraph.properties.spacingBefore;
    const room =
      side.edge === 'top' && before !== undefined
        ? { spacingBefore: pt(Math.max(0, before - border.width)) }
        : {};
    side.block.el = {
      kind: 'paragraph',
      paragraph: {
        ...paragraph,
        properties: {
          ...paragraph.properties,
          ...room,
          borders: {
            ...paragraph.properties.borders,
            [side.edge]: border,
            // §17.3.1.5 — paragraphs with the same borders are ONE bordered
            // set, and a set is ruled on its outside only: an invoice's totals
            // rule every line, and read as five tops Word and LibreOffice drew
            // one rule over "Subtotal" and none under it. The rule between two
            // members is the set's own edge, and it is this same rule.
            ...(side.edge === 'top' ? { insideH: border } : {}),
          },
        },
      },
    };
    for (const piece of band.pieces) given.add(piece);
  }
  return given;
}

/**
 * §17.3.1.33 — the white a block across the page leaves over the columns
 * under it, as the space after its last paragraph. Each column reads its own
 * lines from nothing (see `addColumn`), so the white went nowhere, and
 * comments.pdf's Figure 13 caption came down onto the text under it. As the
 * space before each column's first paragraph Word and LibreOffice drop it at
 * the head of the second column, which then stands higher than the first.
 *
 * @param blocks A page's blocks, as read; each block across the page with
 *               columns under it is given its white.
 */
function spaceUnderSpans(blocks: Array<Block>): void {
  for (const span of blocks) {
    if (span.col !== SPANNING_COLUMN || span.foot === undefined || span.el.kind !== 'paragraph') {
      continue;
    }
    // The columns under it, as far as the next block across the page — every
    // line across the page is a band of its own, so they are found by where
    // they stand, not by band.
    const next = blocks
      .filter((b) => b.col === SPANNING_COLUMN && b.top < span.top)
      .reduce((high, b) => Math.max(high, b.top), -Infinity);
    // The head of each: the box of its first line.
    const heads = blocks
      .filter((b) => b.col !== SPANNING_COLUMN && b.top < span.top && b.top > next)
      .map((b) => b.top + BASELINE_AT * (leadingLine(b.el) ?? 0));
    if (heads.length === 0) continue;
    const white = span.foot - Math.max(...heads);
    if (white <= SPACING_NOISE_PT) continue;
    const { paragraph } = span.el;
    span.el = {
      kind: 'paragraph',
      paragraph: {
        ...paragraph,
        properties: { ...paragraph.properties, spacingAfter: pt(white) },
      },
    };
  }
}

/**
 * Where each column of a page's last stretch of columns — under the last block
 * across the page — ends: the foot of its lowest paragraph, up the page.
 *
 * @param blocks The page's blocks.
 * @returns The feet, one a column, or nothing where fewer than two columns hold text.
 */
function lastColumnFeet(blocks: ReadonlyArray<Block>): Array<number> | undefined {
  const across = blocks
    .filter((b) => b.col === SPANNING_COLUMN)
    .reduce((low, b) => Math.min(low, b.top), Infinity);
  const feet = new Map<number, number>();
  for (const b of blocks) {
    if (b.col === SPANNING_COLUMN || b.foot === undefined || b.top >= across) continue;
    feet.set(b.col, Math.min(feet.get(b.col) ?? Infinity, b.foot));
  }
  return feet.size >= 2 ? [...feet.values()] : undefined;
}

/**
 * Whether a page's last columns are set BALANCED: they end within a couple of
 * lines of each other, and well above the foot of the text.
 *
 * @param feet  Where each column ends, up the page.
 * @param floor Where the text ends at its lowest, up the page.
 * @param size  The body's type size.
 */
function balancedEnd(feet: ReadonlyArray<number>, floor: number, size: number): boolean {
  const line = size * NATURAL_LINE_EM;
  const deepest = Math.min(...feet);
  return (
    Math.max(...feet) - deepest <= line * BALANCED_LINES && deepest - floor >= line * SHORT_LINES
  );
}

/** How many lines apart the ends of balanced columns may stand. */
const BALANCED_LINES = 4;

/** How many lines short of the foot a page's columns stop to be set short. */
const SHORT_LINES = 3;

/** Whether an element is a drawing anchored to the page, which takes no room in the flow. */
function floating(el: BodyElement | undefined): boolean {
  return (
    (el?.kind === 'shape' && el.shape.float !== undefined) ||
    (el?.kind === 'image' && el.image.float !== undefined)
  );
}

/** A block of a page's reading, as `ruleBorders` is handed it. */
type Block = {
  band: number;
  col: number;
  top: number;
  el: BodyElement;
  /** Where the box of a paragraph's last line ends, down the page. */
  foot?: number;
};

/**
 * A table with a rule over one of its rows: the top edge of each of the row's
 * cells (§17.4.39), the rule taking its room out of the white the row keeps
 * over it as a paragraph's does (see `ruleBorders`).
 */
function ruledRow(
  table: Table,
  at: number,
  border: { style: 'single'; width: Pt; colorHex: string },
  rule: { from: number; to: number },
): Table {
  const ruled: Table = {
    ...table,
    rows: table.rows.map((row, r) =>
      r !== at
        ? row
        : {
            ...row,
            cells: row.cells.map((cell) => ({
              properties: {
                ...cell.properties,
                borders: { ...cell.properties.borders, top: border },
              },
              content: cell.content.map((block, k) => {
                const before =
                  block.kind === 'paragraph' && k === 0
                    ? block.paragraph.properties.spacingBefore
                    : undefined;
                return before === undefined || block.kind !== 'paragraph'
                  ? block
                  : {
                      ...block,
                      paragraph: {
                        ...block.paragraph,
                        properties: {
                          ...block.paragraph.properties,
                          spacingBefore: pt(Math.max(0, before - border.width)),
                        },
                      },
                    };
              }),
            })),
          },
    ),
  };
  const read = tableReads.get(table);
  if (!read) return ruled;
  const fitted = fittedToRule(ruled, read, rule);
  tableReads.set(fitted, read);
  return fitted;
}

/**
 * §17.4.63 — a table as wide as the rule drawn across it.
 *
 * A table read off a page begins where its words do, and its last column runs
 * to the measure; a rule drawn across it is its own edge, standing a little
 * past its words on either side. comments.pdf rules Figure 13 from six points
 * left of its names to six right of its last figures, and written as the
 * edge of the row under its headings, the rule ran on twenty points past the
 * table, to the margin. The table now spans what the rule does: its first
 * column reaches out to where the rule begins, its words held where they
 * stood, and its last column ends where the rule ends.
 *
 * Only a rule that reaches past the words on both sides is the table's edge.
 *
 * @param table The table, its rule written as a row's border.
 * @param read  Where it stood on the page.
 * @param rule  The rule, across the page.
 */
function fittedToRule(table: Table, read: TableRead, rule: { from: number; to: number }): Table {
  const last = table.grid.length - 1;
  if (last < 0 || rule.from > read.from || rule.to < read.to) return table;
  const left = read.left + (table.properties.indentPt ?? 0);
  const right = left + table.grid.reduce((sum, w) => sum + w, 0);
  // How far the table's edges move to the rule's: out on the left, in on the
  // right, and the other way where the rule stands the other side.
  const out = left - rule.from;
  const inward = right - rule.to;
  if (Math.abs(out) < FIT_NOISE_PT && Math.abs(inward) < FIT_NOISE_PT) return table;
  const grid = table.grid.map((w, k) => {
    const wider = k === 0 ? w + out : w;
    return pt(Math.max(1, k === last ? wider - inward : wider));
  });
  /** A cell's paragraphs, its words held where the page set them. */
  const held = (content: ReadonlyArray<BodyElement>, k: number): Array<BodyElement> =>
    content.map((block) => {
      if (block.kind !== 'paragraph') return block;
      const { indentLeft, indentRight } = block.paragraph.properties;
      const properties = {
        ...block.paragraph.properties,
        ...(k === 0 && Math.abs(out) >= FIT_NOISE_PT
          ? { indentLeft: pt(Math.max(0, (indentLeft ?? 0) + out)) }
          : {}),
        ...(k === last && Math.abs(inward) >= FIT_NOISE_PT && indentRight !== undefined
          ? { indentRight: pt(Math.max(0, indentRight - inward)) }
          : {}),
      };
      return { ...block, paragraph: { ...block.paragraph, properties } };
    });
  return {
    ...table,
    properties: {
      ...table.properties,
      widthPt: pt(grid.reduce((sum, w) => sum + w, 0)),
      indentPt: pt((table.properties.indentPt ?? 0) - out),
    },
    grid,
    rows: table.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell, k) =>
        k !== 0 && k !== last
          ? cell
          : {
              properties: { ...cell.properties, width: grid[k]! },
              content: held(cell.content, k),
            },
      ),
    })),
  };
}

/** A fit smaller than this moves nothing a reader sees. */
const FIT_NOISE_PT = 0.5;

/**
 * The table row a rule is drawn on the top edge of: a row after the first,
 * the rule standing where its line's box meets the box of the line above it,
 * and reaching across the table's words. `undefined` where the rule is no
 * such row's.
 *
 * @param tables The page's tables, as blocks.
 * @param band   The rule, across the page.
 * @param y      Its height on the page.
 */
function seatOf(
  tables: ReadonlyArray<Block>,
  band: { from: number; to: number },
  y: number,
): { block: Block; row: number } | undefined {
  for (const block of tables) {
    const read = block.el.kind === 'table' ? tableReads.get(block.el.table) : undefined;
    if (!read) continue;
    if (band.from > read.from + RULE_SEAT_PT || band.to < read.to - RULE_SEAT_PT) continue;
    for (let r = 1; r < read.rows.length; r++) {
      const row = read.rows[r]!;
      const prev = read.rows[r - 1]!;
      const edge = row.top + BASELINE_AT * row.lineHeight;
      const over = prev.top - (1 - BASELINE_AT) * prev.lineHeight;
      if (y > row.top && y < prev.top && y - edge <= RULE_SEAT_PT && over - y <= RULE_SEAT_PT) {
        return { block, row: r };
      }
    }
  }
  return undefined;
}

/**
 * How far off the edge between two rows a rule may stand and still be that
 * edge — a point or two either way, which the white a row keeps does not
 * notice.
 */
const RULE_SEAT_PT = 3;

/** How far from a paragraph a rule may stand and still belong to it. */
const RULE_REACH_PT = 14;

/** How far in from its ends a rule is looked at to say which column it stands in. */
const RULE_INSET_PT = 2;

/** The thinnest a border may be drawn and still be seen. */
const RULE_MIN_PT = 0.5;

/**
 * Whether a painted path is flat and thin enough to be a piece of a RULE — a
 * line drawn to separate one block from the next, rather than a piece of the
 * page's artwork. How LONG the rule is is a question for the whole of it (see
 * {@link collinear}): a hairline the width of a column under a table's
 * headings, or the line a total is written over. Everything else — boxes,
 * panels, drawings — keeps the place on the page it was drawn at.
 *
 * @param v The painted path.
 * @returns Whether it could be a rule, or a piece of one.
 */
function flatBar(v: PdfVector): boolean {
  const w = v.maxX - v.minX;
  const h = v.maxY - v.minY;
  return h <= RULE_THICK_PT && w > h * RULE_RATIO;
}

/**
 * The flat bars of a page gathered into the rules they draw: pieces on one
 * baseline, end to end, are one rule from the first's left edge to the last's
 * right.
 *
 * @param bars The page's flat bars.
 * @returns One band per rule, with the pieces it was drawn in.
 */
function collinear(
  bars: ReadonlyArray<PdfVector>,
): Array<{ from: number; to: number; pieces: Array<PdfVector> }> {
  const out: Array<{ from: number; to: number; pieces: Array<PdfVector> }> = [];
  for (const bar of [...bars].sort((a, b) => b.maxY - a.maxY || a.minX - b.minX)) {
    const y = (bar.minY + bar.maxY) / 2;
    const band = out.find(
      (b) =>
        Math.abs((b.pieces[0]!.minY + b.pieces[0]!.maxY) / 2 - y) <= RULE_SAME_LINE_PT &&
        bar.minX <= b.to + RULE_JOIN_PT &&
        (bar.strokeHex ?? bar.fillHex) === (b.pieces[0]!.strokeHex ?? b.pieces[0]!.fillHex),
    );
    if (band) {
      band.from = Math.min(band.from, bar.minX);
      band.to = Math.max(band.to, bar.maxX);
      band.pieces.push(bar);
    } else out.push({ from: bar.minX, to: bar.maxX, pieces: [bar] });
  }
  return out;
}

/** How far apart two bars' baselines may be and still be one rule. */
const RULE_SAME_LINE_PT = 1.5;

/**
 * …and how wide a gap between two pieces of it the page may leave: an invoice
 * draws the rule under its item row cell by cell and skips the thirteen points
 * of gutter between two of them.
 */
const RULE_JOIN_PT = 18;

/** How thick a mark may be and still be a rule. */
const RULE_THICK_PT = 2;

/** How much of the sheet it must run across. */
const RULE_SHARE = 0.2;

/** …and how much longer than it is thick. */
const RULE_RATIO = 20;

/** How close two edges stand before they count as the same one. */
const FLUSH_PT = 1;

/** And how many lines it takes before a region is a column at all. */
const MIN_PROSE_LINES = 6;

/** How many of a column's lines fill it when it is justified prose. */
const FILLED_SHARE = 0.5;

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
