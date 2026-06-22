// Typed in-memory model produced by OOXML parsing and consumed by layout.
//
// All lengths are canonical points (Pt, 1/72"); readers convert format-native
// units (twips, half-points, EMU) at the parse boundary (ir-design.md §5).
//
// All "properties" objects use optional fields to allow inheritance / cascade:
// undefined = "inherit from parent in the style chain". The cascade resolver
// (style-cascade) merges these into ResolvedRunProperties / ResolvedParagraphProperties
// that the renderer consumes.

import type { NativeBag, Pt, ResourceId } from '@/core/ir';
import type { ShapeGradient } from '@/core/vector';

/** Paragraph horizontal alignment (`w:jc`); `both` = justified. */
export type Alignment = 'left' | 'right' | 'center' | 'both' | 'distribute';

/** Run underline style (ECMA-376 §17.18.99 ST_Underline, the subset Ream renders). */
export type UnderlineStyle =
  | 'none'
  | 'single'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'wave';

/** Run vertical alignment (`w:vertAlign`): baseline or super/subscript. */
export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

/**
 * The four script slots of `w:rFonts` (§17.3.2.26). A character picks its font
 * from the slot its Unicode range maps to (ASCII, high-ANSI, complex-script,
 * East-Asian).
 */
export interface FontFamilyMap {
  readonly ascii?: string;
  readonly hAnsi?: string;
  readonly cs?: string;
  readonly eastAsia?: string;
}

/**
 * ECMA-376 Part 1 §17.3.2 — Run Properties (`rPr`). All fields optional so the
 * cascade can inherit them (undefined = "inherit from parent in the style
 * chain"). Boolean toggle properties follow §17.17.4: an absent `<w:b/>` means
 * "inherit"; a present `<w:b/>` with no val (or val=true/1/on) means true;
 * val=false/0/off means explicit false.
 */
export interface RunProperties {
  readonly styleId?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: UnderlineStyle;
  readonly strike?: boolean;
  readonly fontSizePt?: Pt;
  readonly colorHex?: string;
  readonly fontFamily?: FontFamilyMap;
  readonly verticalAlign?: VerticalAlign;
  /**
   * ECMA-376 §17.3.2.30 — `w:rtl`. Marks the run as right-to-left, which seeds
   * the BiDi algorithm with an R/AL bias even for neutral characters.
   */
  readonly rtl?: boolean;
  /**
   * ECMA-376 §17.3.2.20 — `w:lang @w:val` (e.g. `"en-US"`, `"ru-RU"`). Used for
   * the tagged-PDF per-element `/Lang`; does not affect visual layout.
   */
  readonly lang?: string;
}

/**
 * ECMA-376 Part 1 §17.3.1 — Paragraph Properties (`pPr`). All lengths in
 * canonical Pt (converted from twips by the parser); all fields optional for
 * cascade inheritance.
 *
 * `indentFirstLine` encodes both `<w:ind w:firstLine="…"/>` (positive value)
 * and `<w:ind w:hanging="…"/>` (stored as the negative twips, so a hanging
 * indent of 360 twips becomes `indentFirstLine = -18pt` — the first line ends up
 * 360 twips to the left of `indentLeft`).
 */
export interface ParagraphProperties {
  readonly styleId?: string;
  readonly alignment?: Alignment;
  readonly spacingBefore?: Pt;
  readonly spacingAfter?: Pt;
  readonly spacingLine?: Pt;
  readonly spacingLineRule?: 'auto' | 'exact' | 'atLeast';
  readonly indentLeft?: Pt;
  readonly indentRight?: Pt;
  readonly indentFirstLine?: Pt;
  /** The implicit default run properties for runs in this paragraph (`w:pPr/w:rPr`). */
  readonly runProperties?: RunProperties;
  readonly numbering?: NumberingReference;
  /**
   * ECMA-376 Part 1 §17.3.1.21 — `w:pageBreakBefore`. When true, the paragraph
   * starts on a fresh page even if there is room on the current one.
   */
  readonly pageBreakBefore?: boolean;
  /**
   * ECMA-376 §17.3.1.6 — `w:bidi`. Sets the paragraph's base direction to RTL,
   * so the BiDi paragraph embedding level is 1 and default alignment is right.
   */
  readonly bidi?: boolean;
  /**
   * ECMA-376 §17.3.1.20 — `w:outlineLvl`. Outline level 0–8 maps to Heading 1–9;
   * 9 (or absent) is body text. Drives tagged-PDF heading structure (H1–H6).
   */
  readonly outlineLevel?: number;
}

/**
 * ECMA-376 §20.4.2.8 — `wp:inline` picture extent inside a `w:r`. Stored on
 * {@link Run} alongside (or instead of) `text` so layout can position the image
 * as if it were a glyph in the line box.
 */
export interface InlineImage {
  /**
   * Content-addressed bytes in the document's `ResourceStore`; absent when the
   * source relationship did not resolve (the layout box still reserves space).
   */
  readonly resource?: ResourceId;
  readonly width: Pt;
  readonly height: Pt;
}

/**
 * ECMA-376 Part 1 §22 — OfficeMathML (OMML). A recursive math element tree.
 * The union grows by milestone; the layout engine renders structural elements
 * (fraction bars, radicals, big operators, stretchy delimiters) as vector
 * paths and ordinary symbols as font glyphs.
 */
export type MathNode =
  | MathRow
  | MathRun
  | MathFraction
  | MathScript
  | MathRadical
  | MathNary
  | MathFunc
  | MathLimit
  | MathDelimiter
  | MathMatrix
  | MathAccent
  | MathBar
  | MathGroupChr
  | MathEqArray;

/** A sequence of sibling nodes (`m:oMath`, `m:e`, `m:num`, `m:den`, `m:fName` …). */
export interface MathRow {
  readonly type: 'row';
  readonly children: ReadonlyArray<MathNode>;
}

/**
 * `m:r` / `m:t` — literal symbols. `italic`/`bold` come from `m:rPr`/`m:sty`;
 * `nor` (normal text) forces upright. With none set, letters auto-italicise.
 */
export interface MathRun {
  readonly type: 'run';
  readonly text: string;
  readonly italic?: boolean;
  readonly bold?: boolean;
  readonly nor?: boolean;
}

/** `m:f` — fraction. `barless` = `m:fPr/m:type val="noBar"`. */
export interface MathFraction {
  readonly type: 'fraction';
  readonly num: MathNode;
  readonly den: MathNode;
  readonly barless?: boolean;
}

/** `m:sSup` / `m:sSub` / `m:sSubSup` / `m:sPre` (pre-scripts). */
export interface MathScript {
  readonly type: 'script';
  readonly base: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly pre?: boolean; // scripts sit before the base (m:sPre)
}

/** `m:rad` — radical with optional degree (`m:deg`). */
export interface MathRadical {
  readonly type: 'radical';
  readonly radicand: MathNode;
  readonly degree?: MathNode;
}

/** `m:nary` — n-ary operator (∑ ∏ ∫ …) with optional sub/sup limits. */
export interface MathNary {
  readonly type: 'nary';
  readonly op: string;
  readonly body: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly limLoc?: 'undOvr' | 'subSup'; // limits below/above vs at sub/sup
}

/** `m:func` — function application (sin x, …): a name applied to an argument. */
export interface MathFunc {
  readonly type: 'func';
  readonly name: MathNode;
  readonly body: MathNode;
}

/** `m:limLow` / `m:limUpp` — a limit below/above the base. */
export interface MathLimit {
  readonly type: 'limit';
  readonly base: MathNode;
  readonly lim: MathNode;
  readonly pos: 'low' | 'upp';
}

/** `m:d` — delimiters around one or more elements. */
export interface MathDelimiter {
  readonly type: 'delimiter';
  readonly begChr: string;
  readonly endChr: string;
  readonly sepChr?: string;
  readonly children: ReadonlyArray<MathNode>;
}

/** `m:m` — matrix of cells (rows × columns). */
export interface MathMatrix {
  readonly type: 'matrix';
  readonly rows: ReadonlyArray<ReadonlyArray<MathNode>>;
}

/** `m:acc` — accent character over a base (hat, bar, vec, dot, tilde …). */
export interface MathAccent {
  readonly type: 'accent';
  readonly char: string;
  readonly base: MathNode;
}

/** `m:bar` — a bar above or below the base. */
export interface MathBar {
  readonly type: 'bar';
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

/** `m:groupChr` — a grouping character (brace …) above or below the base. */
export interface MathGroupChr {
  readonly type: 'groupChr';
  readonly char: string;
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

/**
 * `m:eqArr` — an equation array: a stack of equations, each on its own line,
 * left-aligned, the block vertically centred on the math axis.
 */
export interface MathEqArray {
  readonly type: 'eqArr';
  readonly rows: ReadonlyArray<MathNode>;
}

/**
 * ECMA-376 Part 1 §17.3.2 — a run: a span of `text` (or an inline image / math
 * object) sharing one set of {@link RunProperties}, plus the reference markers
 * (hyperlink, note, comment, page-number field) the layout acts on.
 */
export interface Run {
  readonly native?: NativeBag;
  readonly text: string;
  readonly properties: RunProperties;
  /**
   * ECMA-376 §17.16.22 — the run sits inside a `w:hyperlink` whose `r:id`
   * resolved to an external target. The URL is stored as written in the rels
   * part; writers MUST pass it through the scheme allowlist (`core/links`)
   * before emitting anything clickable.
   */
  readonly href?: string;
  /**
   * §17.16.5.35 PAGE / §17.16.5.33 NUMPAGES — the run is a page-number field;
   * `text` holds the source's cached result. Header/footer rendering
   * substitutes the real number per page; body rendering keeps the cache.
   */
  readonly field?: 'PAGE' | 'NUMPAGES';
  /**
   * §17.16.22 `w:hyperlink @w:anchor` — internal link target: a bookmark name
   * in this document (never a URL — bypasses the scheme allowlist).
   */
  // §17.11.14 w:footnoteReference / §17.11.6 w:endnoteReference — the run
  // marks a note reference; the layout assigns sequential numbers in reading
  // order and renders them superscript.
  readonly anchor?: string;
  readonly footnoteRef?: string;
  readonly endnoteRef?: string;
  /**
   * §17.13.4.1 `w:commentReference` — the run anchors a review comment by id;
   * the comment's content/author live in `FlowDoc.comments`.
   */
  readonly commentRef?: string;
  /**
   * §17.13.4.3/4 `w:commentRangeStart`/`End` — the ids of the comment ranges
   * this run falls inside (a run may be covered by several). Renderers highlight
   * the covered span; the marker run (`commentRef`) sits at the range's end.
   */
  readonly commentRangeRefs?: ReadonlyArray<string>;
  /**
   * §17.11.13 `w:footnoteRef` / §17.11.5 `w:endnoteRef` — inside note content:
   * render the OWNING note's number here.
   */
  readonly noteNumber?: true;
  /**
   * The run is a list-item marker materialized by `applyNumbering` (`"1."`,
   * `"•"`). Tagged PDF wraps its glyphs in a Lbl structure element (§14.8.4.3.3).
   */
  readonly listMarker?: true;
  /** When set, the run renders this image inline in the line; `text` is ignored. */
  readonly inlineImage?: InlineImage;
  /** When set, the run is an inline OfficeMath object; `text` is ignored. */
  readonly math?: MathNode;
  /**
   * ECMA-376 Part 1 §17.3.3.1 — `w:br w:type="page"`. A forced page break: the
   * paragraph's following content starts on a new page.
   */
  readonly pageBreak?: boolean;
}

/** ECMA-376 Part 1 §17.3.1 — a paragraph: its {@link ParagraphProperties} and runs. */
export interface Paragraph {
  readonly native?: NativeBag;
  readonly properties: ParagraphProperties;
  readonly runs: ReadonlyArray<Run>;
  /**
   * §17.13.6.2 `w:bookmarkStart` — names of bookmarks opening in (or
   * immediately before) this paragraph. Paragraph-level v1: the destination
   * is the paragraph's first line.
   */
  readonly bookmarks?: ReadonlyArray<string>;
}

/**
 * ECMA-376 Part 1 §17.13.4.2 `w:comment` — a review comment: its block content
 * plus the author/date attribution. Anchored from a run's `commentRef`.
 */
export interface Comment {
  readonly content: ReadonlyArray<BodyElement>;
  readonly author?: string;
  readonly initials?: string;
  /** Raw w:date timestamp (ISO 8601), as authored — not reformatted. */
  readonly date?: string;
  /**
   * The author's resolved identity (usually an email) from word/people.xml
   * (`w15:person` → `w15:presenceInfo/@w15:userId`), matched on the author name.
   * Absent when the file ships no people part or the author is not listed.
   */
  readonly authorId?: string;
  /**
   * Microsoft commentsExtended (w15) — the id of the comment this one replies
   * to, forming a thread. Set only when word/commentsExtended.xml links this
   * comment's paragraph to a parent (w15:paraIdParent). Top-level comments omit it.
   */
  readonly parentId?: string;
  /** w15:done — the comment thread was marked resolved. */
  readonly done?: boolean;
}

/** ECMA-376 Part 1 §17.9 — `w:numFmt` list marker format. */
export type NumberingFormat =
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'bullet'
  | 'none';

/** A paragraph's list reference (`w:numPr`): the numbering instance id + level. */
export interface NumberingReference {
  readonly numId: string;
  readonly ilvl: number;
}

/** One level of an abstract numbering definition (`w:lvl`): its format and chrome. */
export interface NumberingLevel {
  readonly ilvl: number;
  readonly start: number;
  readonly format: NumberingFormat;
  /** §17.9.11 `w:lvlText` — the marker template (e.g. `"%1."`). */
  readonly lvlText: string;
  readonly paragraphProperties: ParagraphProperties;
  readonly runProperties: RunProperties;
}

/** §17.9.1 `w:abstractNum` — a reusable list definition keyed by level. */
export interface AbstractNumbering {
  readonly id: string;
  readonly levels: ReadonlyMap<number, NumberingLevel>;
}

/** §17.9.18 `w:num` — a concrete list instance bound to an {@link AbstractNumbering}. */
export interface NumberingInstance {
  readonly numId: string;
  readonly abstractNumId: string;
}

/** The parsed `word/numbering.xml`: abstract definitions + their instances. */
export interface Numbering {
  readonly abstractNums: ReadonlyMap<string, AbstractNumbering>;
  readonly numInstances: ReadonlyMap<string, NumberingInstance>;
}

/** ECMA-376 Part 1 §17.7 — `w:style @w:type` style category. */
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/**
 * §17.7.6 — one table-style formatting layer: the style's own base layer
 * (`w:style/tblPr` + `tcPr` + `rPr` + `pPr`) or one conditional override
 * (`w:tblStylePr`). Borders come from `tblBorders` (table layer) or `tcBorders`
 * (region layer) — whichever the layer carries.
 */
export interface TableStyleLayer {
  readonly borders?: CellBorders;
  readonly cellMargins?: CellMargins;
  readonly shading?: CellShading;
  readonly runProperties?: RunProperties;
  readonly paragraphProperties?: ParagraphProperties;
}

/** §17.7.6.3 `w:tblStylePr @w:type` — the table regions a conditional layer targets. */
export type TableStyleConditionType =
  | 'wholeTable'
  | 'band1Vert'
  | 'band2Vert'
  | 'band1Horz'
  | 'band2Horz'
  | 'firstCol'
  | 'lastCol'
  | 'firstRow'
  | 'lastRow'
  | 'nwCell'
  | 'neCell'
  | 'swCell'
  | 'seCell';

/** A conditional table-style override: the region `type` and the layer it applies. */
export interface TableStyleCondition {
  readonly type: TableStyleConditionType;
  readonly layer: TableStyleLayer;
}

/**
 * ECMA-376 §17.7 — one style definition: its id, category, `basedOn` parent,
 * and the run/paragraph properties it contributes (plus the table-style layers
 * for `type === 'table'`).
 */
export interface Style {
  readonly id: string;
  readonly type: StyleType;
  readonly basedOn?: string;
  readonly isDefault: boolean;
  readonly runProperties: RunProperties;
  readonly paragraphProperties: ParagraphProperties;
  /** Table styles only (§17.7.6): the base layer and conditional overrides. */
  readonly tableLayer?: TableStyleLayer;
  readonly tableConditions?: ReadonlyArray<TableStyleCondition>;
  /** `w:tblPr/w:tblStyleRowBandSize` / `ColBandSize` (default 1). */
  readonly rowBandSize?: number;
  readonly colBandSize?: number;
}

/**
 * The parsed `word/styles.xml`: document defaults (`docDefaults`) plus the
 * styles map keyed by id. The cascade resolver merges these into the resolved
 * property objects the renderer consumes.
 */
export interface StyleSheet {
  readonly defaultRunProperties: RunProperties;
  readonly defaultParagraphProperties: ParagraphProperties;
  readonly styles: ReadonlyMap<string, Style>;
}

// ECMA-376 Part 1 §17.4 — Tables.

/** §17.18.2 ST_Border — a cell-border line style (the subset Ream renders). */
export type BorderStyle = 'none' | 'single' | 'double' | 'thick' | 'dotted' | 'dashed';

/** One border edge: its line style plus optional width and colour. */
export interface Border {
  readonly style: BorderStyle;
  readonly width?: Pt;
  readonly colorHex?: string;
}

/** §17.4.39 `w:tcBorders` — the per-edge borders of a cell (or table). */
export interface CellBorders {
  readonly top?: Border;
  readonly right?: Border;
  readonly bottom?: Border;
  readonly left?: Border;
  readonly insideH?: Border;
  readonly insideV?: Border;
  /**
   * Diagonal strokes across the cell box: `diagonalDown` runs top-left →
   * bottom-right, `diagonalUp` bottom-left → top-right (Excel diagonal borders).
   */
  readonly diagonalDown?: Border;
  readonly diagonalUp?: Border;
}

/** §17.4.42 `w:tcMar` — a cell's inner padding per side. */
export interface CellMargins {
  readonly top?: Pt;
  readonly right?: Pt;
  readonly bottom?: Pt;
  readonly left?: Pt;
}

/** §17.4.33 `w:shd` — a cell's solid background fill (6-hex). */
export interface CellShading {
  readonly colorHex: string;
}

/**
 * A conditional-format data bar: a horizontal bar of width `fraction` (`0..1` of
 * the cell), painted over the shading and under the text (E-SHEET SC1c). It
 * starts at `startFraction` from the cell's left (default 0); a mixed-sign range
 * puts the axis inside the cell so negative bars run left of it (tail TC4).
 */
export interface CellDataBar {
  readonly fraction: number;
  readonly colorHex: string;
  readonly startFraction?: number;
}

/**
 * A conditional-format icon: a small glyph at the cell's left, chosen by the
 * value's bucket (E-SHEET SC1c). Format-neutral — the xlsx layer maps Excel's
 * named icon families (3TrafficLights, 3Arrows, …) onto these shapes + colours.
 */
export type CellIconShape =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangleUp'
  | 'triangleDown'
  | 'triangleRight'
  // Symbols families (3Symbols / 3Symbols2): a check / exclamation / cross mark.
  | 'check'
  | 'cross'
  | 'exclamation'
  // Meter families: ratings (a signal-strength bar histogram) and quarters (a
  // clock-fill pie). Both read `fill` for how many units are coloured in.
  | 'bars'
  | 'pie';
/** One resolved conditional-format icon: its {@link CellIconShape}, colour and fill level. */
export interface CellIcon {
  readonly shape: CellIconShape;
  readonly colorHex: string;
  /**
   * Meter glyphs (`bars` / `pie`): how many of `levels` units are filled with
   * `colorHex`; the rest are drawn in a neutral grey. Absent for single glyphs.
   */
  readonly fill?: { readonly filled: number; readonly levels: number };
}

/**
 * A sparkline: a mini chart filling the cell, plotting `values` (E-SHEET SC2).
 * Format-neutral — the xlsx layer resolves the data range to a value sequence;
 * the layout renders line / column / win-loss geometry sized to the cell.
 */
export interface CellSparkline {
  readonly kind: 'line' | 'column' | 'winLoss';
  /**
   * A blank/non-numeric cell in the range is a gap (`null`) so x-positions stay
   * aligned: a line breaks across it, a column/win-loss skips its slot.
   */
  readonly values: ReadonlyArray<number | null>;
  readonly colorHex?: string;
}

/**
 * Resolved position of a cell in a vertical merge group (ECMA-376 §17.4.85
 * `vMerge` markers are resolved by the reader): `start` opens a group that at
 * least one cell continues, `middle` / `end` are continuations; undefined =
 * not merged. Continuation cells stay in their rows (they hold the column
 * slot); writers that need an HTML-style rowSpan can derive it by counting a
 * start's middle/end run downwards.
 */
export type CellMerge = 'start' | 'middle' | 'end';

/** §17.4.30 `w:tcPr` — a cell's properties: span/merge, chrome, and CF overlays. */
export interface CellProperties {
  readonly width?: Pt;
  readonly colSpan?: number;
  readonly merge?: CellMerge;
  readonly borders?: CellBorders;
  readonly margins?: CellMargins;
  readonly shading?: CellShading;
  readonly dataBar?: CellDataBar;
  readonly icon?: CellIcon;
  readonly sparkline?: CellSparkline;
  /**
   * A data-validation `list` cell (E-SHEET SV1): the renderer paints an in-cell
   * dropdown affordance at the cell's right edge (a small button + ▾ glyph).
   */
  readonly dropdown?: boolean;
}

/** §17.4.81 `w:trPr` — a table row's properties: height, split/header flags. */
export interface RowProperties {
  readonly height?: Pt;
  readonly heightRule?: 'auto' | 'atLeast' | 'exact';
  readonly cantSplit?: boolean;
  readonly isHeader?: boolean;
  /**
   * Force this row to begin a new page (xlsx manual `<rowBreaks>`). The renderer
   * flushes the page before the row, then repeats any leading header rows.
   */
  readonly pageBreakBefore?: boolean;
}

/**
 * §17.4.62 `w:tblLook` — which of the table style's conditional formats apply.
 * Modern files carry explicit attributes; legacy files a hex bitmask (both
 * parsed). Band flags are negative ("no band") per the spec.
 */
export interface TableLook {
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
  readonly noHBand?: boolean;
  readonly noVBand?: boolean;
}

/** §17.4.60 `w:tblPr` — a table's properties: style ref, width/layout, chrome. */
export interface TableProperties {
  /**
   * §17.7.6 — raw reference to a table style (resolved by the reader's
   * `resolveTableStyles` transform; round-trip material afterwards).
   */
  readonly styleId?: string;
  readonly look?: TableLook;
  readonly widthPt?: Pt;
  readonly widthFraction?: number; // tblW type=pct: w/5000 (1.0 = full content width)
  readonly widthType?: 'auto' | 'dxa' | 'pct' | 'nil';
  readonly layout?: 'auto' | 'fixed';
  readonly defaultCellMargins?: CellMargins;
  readonly borders?: CellBorders;
  /**
   * ECMA-376 §17.4.27 (`w:jc`) / xlsx `<printOptions horizontalCentered>`.
   * Centers or right-aligns a table narrower than the content width; absent ⇒ left.
   */
  readonly alignment?: 'left' | 'center' | 'right';
  /**
   * A sticky-pane hint from a frozen worksheet view (E-SHEET SE3): the first
   * `rows` rows / `cols` columns stay pinned while the rest scrolls. Consumed
   * only by the HTML writer (an interactive target); PDF/SVG ignore it.
   */
  readonly frozen?: { readonly rows: number; readonly cols: number };
}

/** §17.4.66 `w:tc` — a table cell: its {@link CellProperties} and block content. */
export interface TableCell {
  readonly properties: CellProperties;
  readonly content: ReadonlyArray<BodyElement>;
}

/** §17.4.79 `w:tr` — a table row: its {@link RowProperties} and cells. */
export interface TableRow {
  readonly properties: RowProperties;
  readonly cells: ReadonlyArray<TableCell>;
}

/**
 * §17.4.38 `w:tbl` — a table: its {@link TableProperties}, the `w:tblGrid`
 * column widths (`grid`), and its rows.
 */
export interface Table {
  readonly native?: NativeBag;
  readonly properties: TableProperties;
  readonly grid: ReadonlyArray<Pt>;
  readonly rows: ReadonlyArray<TableRow>;
}

/**
 * ECMA-376 Part 1 §20.4.2.8 — a block-level image (`wp:inline` picture extent).
 * EMU = English Metric Units: 914400 per inch (1 pt = 12700 EMU).
 */
export interface ImageBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly resource?: ResourceId;
  readonly width: Pt;
  readonly height: Pt;
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

// ECMA-376 Part 1 §20 — DrawingML shapes (wps:wsp inside a w:drawing).
// A standalone shape (a paragraph whose only content is a shape drawing)
// collapses to a ShapeBlock, mirroring ImageBlock; it carries the paragraph's
// properties for block spacing / alignment.

/**
 * §20.1.10.55 prstGeom path command (custGeom `<a:pathLst>`), in path-space
 * units (`a:path @w/@h`). `quad` is elevated to cubic at layout; `arc` uses the
 * DrawingML angle convention (1/60000°, clockwise, y-down) and is converted
 * to beziers by the geometry layer.
 */
export type CustomPathCmd =
  | { readonly cmd: 'move'; readonly x: number; readonly y: number }
  | { readonly cmd: 'line'; readonly x: number; readonly y: number }
  | {
      readonly cmd: 'cubic';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly cmd: 'quad';
      readonly x1: number;
      readonly y1: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly cmd: 'arc';
      readonly wR: number;
      readonly hR: number;
      readonly stAng: number;
      readonly swAng: number;
    }
  | { readonly cmd: 'close' };

/**
 * §20.1.9.11 `custGeom` — a custom geometry: its path-space extent
 * (`pathWidth`/`pathHeight`) and the {@link CustomPathCmd} list that draws it.
 */
export interface CustomGeometry {
  readonly pathWidth: number;
  readonly pathHeight: number;
  readonly commands: ReadonlyArray<CustomPathCmd>;
}

/** §20.1.10.55 `prstGeom` / §20.1.9.11 `custGeom` — a shape's geometry source. */
export interface ShapeGeometry {
  readonly kind: 'preset' | 'custom';
  readonly preset?: string; // kind==='preset', e.g. 'roundRect'
  readonly adjust?: ReadonlyMap<string, number>; // a:gd name → raw val (usually 0..100000)
  readonly custom?: CustomGeometry; // kind==='custom'
}

/** A shape's fill mode (`a:noFill`/`a:solidFill`/`a:gradFill`). */
export type ShapeFillKind = 'none' | 'solid' | 'gradient';

/** A shape's fill: none, a solid colour, or a {@link ShapeGradient}. */
export interface ShapeFill {
  readonly kind: ShapeFillKind;
  readonly colorHex?: string; // resolved 6-hex (kind==='solid')
  readonly gradient?: ShapeGradient; // kind==='gradient' (a:gradFill, EP16)
}

/** §20.1.10.49 ST_PresetLineDashVal — a shape outline's preset dash pattern. */
export type ShapeDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'sysDash'
  | 'sysDot';

/** §20.1.2.2.24 `a:ln` — a shape's outline (stroke). */
export interface ShapeLine {
  readonly width?: Pt; // a:ln @w; default 0.75pt
  readonly colorHex?: string; // resolved 6-hex
  readonly dash?: ShapeDash; // a:prstDash @val
  readonly cap?: 'flat' | 'round' | 'square'; // a:ln @cap (flat=butt)
  readonly fill?: 'solid' | 'none'; // a:ln/a:noFill ⇒ no visible stroke
}

/** §20.1.7.6 `a:xfrm` — a shape's rotation (1/60000°, clockwise) + flips. */
export interface ShapeTransform {
  readonly rotation60k?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/** `wps:txbx` (`w:txbxContent`) + `wps:bodyPr` — text inside a shape. */
export interface ShapeTextBody {
  readonly content: ReadonlyArray<BodyElement>;
  readonly insetLeft?: Pt;
  readonly insetTop?: Pt;
  readonly insetRight?: Pt;
  readonly insetBottom?: Pt;
  readonly anchor?: 't' | 'ctr' | 'b'; // vertical anchor
}

/**
 * A block-level DrawingML shape (§20.1): its size, {@link ShapeGeometry}, fill,
 * outline, transform and optional text body. A standalone shape carries the
 * paragraph's properties for block spacing / alignment, mirroring
 * {@link ImageBlock}.
 */
export interface ShapeBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly width: Pt; // wp:extent cx (fallback a:ext cx)
  readonly height: Pt; // wp:extent cy
  readonly geometry: ShapeGeometry;
  readonly fill: ShapeFill;
  readonly line?: ShapeLine;
  readonly transform?: ShapeTransform;
  readonly text?: ShapeTextBody;
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

/**
 * ECMA-376 Part 1 §21.2 — the kind of a DrawingML chart. A chart is referenced
 * from a `w:drawing` (`a:graphicData uri=…/chart` → `c:chart r:id`) and its data
 * lives in a separate chart part (e.g. `word/charts/chart1.xml`). The parsed
 * {@link Chart} is keyed by that relationship id and supplied to the renderer
 * alongside the body (mirroring image bytes), so {@link ChartBlock} only carries
 * the ref.
 */
export type ChartType = 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'unknown';

/** A per-point colour override (`c:dPt`) at series index `idx`. */
export interface ChartDataPoint {
  readonly idx: number;
  readonly colorHex: string; // c:dPt per-point colour override
}

/** One chart data series (`c:ser`): its name, values and colour overrides. */
export interface ChartSeries {
  readonly name?: string;
  readonly values: ReadonlyArray<number>; // c:val/c:yVal numCache (idx-ordered, gaps → 0)
  readonly xValues?: ReadonlyArray<number>; // c:xVal numCache (scatter — paired with values)
  readonly colorHex?: string; // c:spPr solidFill
  readonly pointColors?: ReadonlyArray<ChartDataPoint>; // c:dPt overrides (pie slices)
}

/** A parsed chart (§21.2): its type, title, categories, series and rendering options. */
export interface Chart {
  readonly type: ChartType;
  readonly title?: string;
  readonly categories: ReadonlyArray<string>; // c:cat (shared across series)
  readonly series: ReadonlyArray<ChartSeries>;
  readonly hasLegend: boolean;
  readonly legendPos?: 'r' | 'l' | 't' | 'b';
  readonly barDir?: 'col' | 'bar'; // c:barDir (bar charts)
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  readonly doughnut?: boolean; // c:doughnutChart (a pie with a central hole)
  readonly showValues?: boolean; // c:dLbls/c:showVal — print each datum's value
  readonly catAxisTitle?: string; // c:catAx/c:title
  readonly valAxisTitle?: string; // c:valAx/c:title
  /**
   * MS-ODRAWXML `chartColorStyle` (`colorsN.xml`): the cycle of series colours;
   * overrides the built-in Office accent cycle when present.
   */
  readonly seriesColorCycle?: ReadonlyArray<string>;
}

/** A block-level chart reference (`c:chart`): its size and relationship id. */
export interface ChartBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly chartRelId: string; // c:chart @r:id (resolve against the document's rels)
  readonly width: Pt;
  readonly height: Pt;
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

/** One top-level body item: a discriminated union over the block kinds. */
export type BodyElement =
  | { readonly kind: 'paragraph'; readonly paragraph: Paragraph }
  | { readonly kind: 'table'; readonly table: Table }
  | { readonly kind: 'image'; readonly image: ImageBlock }
  | { readonly kind: 'shape'; readonly shape: ShapeBlock }
  | { readonly kind: 'chart'; readonly chart: ChartBlock };

/** ECMA-376 Part 1 §17.6.13 `w:pgSz` — the section's page dimensions + orientation. */
export interface PageSize {
  readonly width: Pt;
  readonly height: Pt;
  readonly orientation?: 'portrait' | 'landscape';
}

/** §17.6.11 `w:pgMar` — the section's page margins (and header/footer offsets). */
export interface PageMargins {
  readonly top: Pt;
  readonly right: Pt;
  readonly bottom: Pt;
  readonly left: Pt;
  readonly header?: Pt;
  readonly footer?: Pt;
}

/** §17.10 — which page class a header/footer reference applies to. */
export type HeaderFooterType = 'default' | 'first' | 'even';

/** §17.10.5/§17.10.2 `w:headerReference`/`w:footerReference` — a typed rels pointer. */
export interface HeaderFooterReference {
  readonly type: HeaderFooterType;
  readonly relationshipId: string;
}

/** §17.6.17 `w:sectPr` — a section's page setup, header/footer refs and columns. */
export interface SectionProperties {
  readonly pageSize?: PageSize;
  readonly margins?: PageMargins;
  readonly headers: ReadonlyArray<HeaderFooterReference>;
  readonly footers: ReadonlyArray<HeaderFooterReference>;
  /**
   * ECMA-376 §17.10.6 — `w:titlePg` toggle in `sectPr`. When true the first page
   * of the section uses the `first` header/footer references.
   */
  readonly titlePg?: boolean;
  /**
   * ECMA-376 §17.15.1.36 — `w:evenAndOddHeaders` toggle in `word/settings.xml`
   * (document-wide, not per-section). When true even-numbered pages use the
   * `even` header/footer references.
   */
  readonly evenAndOddHeaders?: boolean;
  /** §17.6.4 `w:cols` — multi-column section layout. */
  readonly columns?: SectionColumns;
}

/**
 * §20.4.2.3 `wp:anchor` — a floating drawing's placement. v1 honours
 * out-of-flow placement for wrap `'none'` (incl. `behindDoc`); the side-wrapping
 * modes (square/tight/through) and `topAndBottom` stay in flow as blocks.
 */
export interface FloatAnchor {
  readonly wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  readonly behind?: boolean; // wp:anchor @behindDoc
  readonly posH?: {
    readonly relativeFrom: 'margin' | 'page' | 'column';
    readonly offsetPt?: Pt; // wp:posOffset
    readonly align?: 'left' | 'center' | 'right'; // wp:align
  };
  readonly posV?: {
    readonly relativeFrom: 'margin' | 'page' | 'paragraph' | 'line';
    readonly offsetPt?: Pt;
  };
}

/**
 * §17.6.4 — column definitions: an equal-width `count` + gutter, or explicit
 * per-column widths/gutters (`w:col` children).
 */
export interface SectionColumns {
  readonly count: number;
  readonly spacePt: number;
  readonly explicit?: ReadonlyArray<{ readonly widthPt: number; readonly spacePt: number }>;
}

/**
 * One section descriptor for a multi-section document (ECMA-376 §17.6.17). Each
 * section's `sectPr` lives either inside a paragraph's `pPr` (mid-document
 * break) or as the final child of `w:body` (final section). Records the
 * properties and the upper-exclusive bound in `BodyElement[]` the section
 * covers — section i applies to
 * `body[sections[i-1].endIndex..sections[i].endIndex)`.
 */
export interface Section {
  readonly properties: SectionProperties;
  readonly endIndex: number;
}

/** The parsed WordprocessingML document: body, stylesheet, numbering and section setup. */
export interface DocumentModel {
  readonly body: ReadonlyArray<BodyElement>;
  readonly styleSheet: StyleSheet;
  readonly numbering?: Numbering;
  readonly section?: SectionProperties;
}

/**
 * Document metadata (PDF `/Info`-shaped, sourced from `docProps/core.xml` and/or
 * caller options). Lives in the model so FlowDoc can carry it format-neutrally.
 */
export interface DocumentInfo {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  /** Producer; defaults to "Ream" if not provided. */
  readonly producer?: string;
  /** ISO 8601 date; converted to PDF date format (`D:YYYYMMDDHHmmSS`). */
  readonly creationDate?: Date;
  readonly modificationDate?: Date;
}
