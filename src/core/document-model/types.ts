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

export type Alignment = 'left' | 'right' | 'center' | 'both' | 'distribute';

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

export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

export interface FontFamilyMap {
  readonly ascii?: string;
  readonly hAnsi?: string;
  readonly cs?: string;
  readonly eastAsia?: string;
}

// ECMA-376 Part 1 §17.3.2 — Run Properties (rPr).
// Boolean toggle properties follow §17.17.4: an absent <w:b/> means "inherit";
// a present <w:b/> with no val (or val=true/1/on) means true; val=false/0/off
// means explicit false.
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
  // ECMA-376 §17.3.2.30 — w:rtl. Marks the run as right-to-left, which seeds
  // the BiDi algorithm with an R/AL bias even for neutral characters.
  readonly rtl?: boolean;
  // ECMA-376 §17.3.2.20 — w:lang @w:val (e.g. "en-US", "ru-RU"). Used for the
  // tagged-PDF per-element /Lang; does not affect visual layout.
  readonly lang?: string;
}

// ECMA-376 Part 1 §17.3.1 — Paragraph Properties (pPr).
// All lengths in canonical Pt (converted from twips by the parser).
//
// indentFirstLine encodes both <w:ind w:firstLine="…"/> (positive value)
// and <w:ind w:hanging="…"/> (stored as the negative twips, so a hanging
// indent of 360 twips becomes indentFirstLine = -18pt — the first line ends up
// 360 twips to the left of indentLeftTwips).
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
  readonly runProperties?: RunProperties;
  readonly numbering?: NumberingReference;
  // ECMA-376 Part 1 §17.3.1.21 — w:pageBreakBefore. When true, the paragraph
  // starts on a fresh page even if there is room on the current one.
  readonly pageBreakBefore?: boolean;
  // ECMA-376 §17.3.1.6 — w:bidi. Sets the paragraph's base direction to RTL,
  // so the BiDi paragraph embedding level is 1 and default alignment is right.
  readonly bidi?: boolean;
  // ECMA-376 §17.3.1.20 — w:outlineLvl. Outline level 0–8 maps to Heading 1–9;
  // 9 (or absent) is body text. Drives tagged-PDF heading structure (H1–H6).
  readonly outlineLevel?: number;
}

// ECMA-376 §20.4.2.8 — wp:inline picture extent inside a w:r. Stored on Run
// alongside (or instead of) `text` so layout can position the image as if it
// were a glyph in the line box.
export interface InlineImage {
  // Content-addressed bytes in the document's ResourceStore; absent when the
  // source relationship did not resolve (the layout box still reserves space).
  readonly resource?: ResourceId;
  readonly width: Pt;
  readonly height: Pt;
}

// ECMA-376 Part 1 §22 — OfficeMathML (OMML). A recursive math element tree.
// The union grows by milestone; the layout engine renders structural elements
// (fraction bars, radicals, big operators, stretchy delimiters) as vector
// paths and ordinary symbols as font glyphs.
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

// A sequence of sibling nodes (m:oMath, m:e, m:num, m:den, m:fName …).
export interface MathRow {
  readonly type: 'row';
  readonly children: ReadonlyArray<MathNode>;
}

// m:r / m:t — literal symbols. `italic`/`bold` come from m:rPr/m:sty; `nor`
// (normal text) forces upright. With none set, letters auto-italicise.
export interface MathRun {
  readonly type: 'run';
  readonly text: string;
  readonly italic?: boolean;
  readonly bold?: boolean;
  readonly nor?: boolean;
}

// m:f — fraction. barless = m:fPr/m:type val="noBar".
export interface MathFraction {
  readonly type: 'fraction';
  readonly num: MathNode;
  readonly den: MathNode;
  readonly barless?: boolean;
}

// m:sSup / m:sSub / m:sSubSup / m:sPre (pre-scripts).
export interface MathScript {
  readonly type: 'script';
  readonly base: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly pre?: boolean; // scripts sit before the base (m:sPre)
}

// m:rad — radical with optional degree (m:deg).
export interface MathRadical {
  readonly type: 'radical';
  readonly radicand: MathNode;
  readonly degree?: MathNode;
}

// m:nary — n-ary operator (∑ ∏ ∫ …) with optional sub/sup limits.
export interface MathNary {
  readonly type: 'nary';
  readonly op: string;
  readonly body: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly limLoc?: 'undOvr' | 'subSup'; // limits below/above vs at sub/sup
}

// m:func — function application (sin x, …): a name applied to an argument.
export interface MathFunc {
  readonly type: 'func';
  readonly name: MathNode;
  readonly body: MathNode;
}

// m:limLow / m:limUpp — a limit below/above the base.
export interface MathLimit {
  readonly type: 'limit';
  readonly base: MathNode;
  readonly lim: MathNode;
  readonly pos: 'low' | 'upp';
}

// m:d — delimiters around one or more elements.
export interface MathDelimiter {
  readonly type: 'delimiter';
  readonly begChr: string;
  readonly endChr: string;
  readonly sepChr?: string;
  readonly children: ReadonlyArray<MathNode>;
}

// m:m — matrix of cells (rows × columns).
export interface MathMatrix {
  readonly type: 'matrix';
  readonly rows: ReadonlyArray<ReadonlyArray<MathNode>>;
}

// m:acc — accent character over a base (hat, bar, vec, dot, tilde …).
export interface MathAccent {
  readonly type: 'accent';
  readonly char: string;
  readonly base: MathNode;
}

// m:bar — a bar above or below the base.
export interface MathBar {
  readonly type: 'bar';
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

// m:groupChr — a grouping character (brace …) above or below the base.
export interface MathGroupChr {
  readonly type: 'groupChr';
  readonly char: string;
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

// m:eqArr — an equation array: a stack of equations, each on its own line,
// left-aligned, the block vertically centred on the math axis.
export interface MathEqArray {
  readonly type: 'eqArr';
  readonly rows: ReadonlyArray<MathNode>;
}

export interface Run {
  readonly native?: NativeBag;
  readonly text: string;
  readonly properties: RunProperties;
  // ECMA-376 §17.16.22 — the run sits inside a w:hyperlink whose r:id resolved
  // to an external target. The URL is stored as written in the rels part;
  // writers MUST pass it through the scheme allowlist (core/links) before
  // emitting anything clickable.
  readonly href?: string;
  // §17.16.5.35 PAGE / §17.16.5.33 NUMPAGES — the run is a page-number field;
  // `text` holds the source's cached result. Header/footer rendering
  // substitutes the real number per page; body rendering keeps the cache.
  readonly field?: 'PAGE' | 'NUMPAGES';
  // §17.11.14 w:footnoteReference / §17.11.6 w:endnoteReference — the run
  // marks a note reference; the layout assigns sequential numbers in reading
  // order and renders them superscript.
  // §17.16.22 w:hyperlink @w:anchor — internal link target: a bookmark name
  // in this document (never a URL — bypasses the scheme allowlist).
  readonly anchor?: string;
  readonly footnoteRef?: string;
  readonly endnoteRef?: string;
  // §17.11.13 w:footnoteRef / §17.11.5 w:endnoteRef — inside note content:
  // render the OWNING note's number here.
  readonly noteNumber?: true;
  // When set, the run renders this image inline in the line; `text` is ignored.
  readonly inlineImage?: InlineImage;
  // When set, the run is an inline OfficeMath object; `text` is ignored.
  readonly math?: MathNode;
  // ECMA-376 Part 1 §17.3.3.1 — w:br w:type="page". A forced page break: the
  // paragraph's following content starts on a new page.
  readonly pageBreak?: boolean;
}

export interface Paragraph {
  readonly native?: NativeBag;
  readonly properties: ParagraphProperties;
  readonly runs: ReadonlyArray<Run>;
  // §17.13.6.2 w:bookmarkStart — names of bookmarks opening in (or
  // immediately before) this paragraph. Paragraph-level v1: the destination
  // is the paragraph's first line.
  readonly bookmarks?: ReadonlyArray<string>;
}

// ECMA-376 Part 1 §17.9 — Numbering.
export type NumberingFormat =
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'bullet'
  | 'none';

export interface NumberingReference {
  readonly numId: string;
  readonly ilvl: number;
}

export interface NumberingLevel {
  readonly ilvl: number;
  readonly start: number;
  readonly format: NumberingFormat;
  readonly lvlText: string;
  readonly paragraphProperties: ParagraphProperties;
  readonly runProperties: RunProperties;
}

export interface AbstractNumbering {
  readonly id: string;
  readonly levels: ReadonlyMap<number, NumberingLevel>;
}

export interface NumberingInstance {
  readonly numId: string;
  readonly abstractNumId: string;
}

export interface Numbering {
  readonly abstractNums: ReadonlyMap<string, AbstractNumbering>;
  readonly numInstances: ReadonlyMap<string, NumberingInstance>;
}

// ECMA-376 Part 1 §17.7 — Styles.
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

// §17.7.6 — one table-style formatting layer: the style's own base layer
// (w:style/tblPr + tcPr + rPr + pPr) or one conditional override
// (w:tblStylePr). Borders come from tblBorders (table layer) or tcBorders
// (region layer) — whichever the layer carries.
export interface TableStyleLayer {
  readonly borders?: CellBorders;
  readonly cellMargins?: CellMargins;
  readonly shading?: CellShading;
  readonly runProperties?: RunProperties;
  readonly paragraphProperties?: ParagraphProperties;
}

// §17.7.6.3 w:tblStylePr @w:type — the table regions a conditional layer
// targets.
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

export interface TableStyleCondition {
  readonly type: TableStyleConditionType;
  readonly layer: TableStyleLayer;
}

export interface Style {
  readonly id: string;
  readonly type: StyleType;
  readonly basedOn?: string;
  readonly isDefault: boolean;
  readonly runProperties: RunProperties;
  readonly paragraphProperties: ParagraphProperties;
  // Table styles only (§17.7.6): the base layer and conditional overrides.
  readonly tableLayer?: TableStyleLayer;
  readonly tableConditions?: ReadonlyArray<TableStyleCondition>;
  // w:tblPr/w:tblStyleRowBandSize / ColBandSize (default 1).
  readonly rowBandSize?: number;
  readonly colBandSize?: number;
}

export interface StyleSheet {
  readonly defaultRunProperties: RunProperties;
  readonly defaultParagraphProperties: ParagraphProperties;
  readonly styles: ReadonlyMap<string, Style>;
}

// ECMA-376 Part 1 §17.4 — Tables.

export type BorderStyle = 'none' | 'single' | 'double' | 'thick' | 'dotted' | 'dashed';

export interface Border {
  readonly style: BorderStyle;
  readonly width?: Pt;
  readonly colorHex?: string;
}

export interface CellBorders {
  readonly top?: Border;
  readonly right?: Border;
  readonly bottom?: Border;
  readonly left?: Border;
  readonly insideH?: Border;
  readonly insideV?: Border;
}

export interface CellMargins {
  readonly top?: Pt;
  readonly right?: Pt;
  readonly bottom?: Pt;
  readonly left?: Pt;
}

export interface CellShading {
  readonly colorHex: string;
}

// Resolved position of a cell in a vertical merge group (ECMA-376 §17.4.85
// vMerge markers are resolved by the reader): 'start' opens a group that at
// least one cell continues, 'middle' / 'end' are continuations; undefined =
// not merged. Continuation cells stay in their rows (they hold the column
// slot); writers that need an HTML-style rowSpan can derive it by counting a
// start's middle/end run downwards.
export type CellMerge = 'start' | 'middle' | 'end';

export interface CellProperties {
  readonly width?: Pt;
  readonly colSpan?: number;
  readonly merge?: CellMerge;
  readonly borders?: CellBorders;
  readonly margins?: CellMargins;
  readonly shading?: CellShading;
}

export interface RowProperties {
  readonly height?: Pt;
  readonly heightRule?: 'auto' | 'atLeast' | 'exact';
  readonly cantSplit?: boolean;
  readonly isHeader?: boolean;
  // Force this row to begin a new page (xlsx manual <rowBreaks>). The renderer
  // flushes the page before the row, then repeats any leading header rows.
  readonly pageBreakBefore?: boolean;
}

// §17.4.62 w:tblLook — which of the table style's conditional formats apply.
// Modern files carry explicit attributes; legacy files a hex bitmask (both
// parsed). Band flags are negative ("no band") per the spec.
export interface TableLook {
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
  readonly noHBand?: boolean;
  readonly noVBand?: boolean;
}

export interface TableProperties {
  // §17.7.6 — raw reference to a table style (resolved by the reader's
  // resolveTableStyles transform; round-trip material afterwards).
  readonly styleId?: string;
  readonly look?: TableLook;
  readonly widthPt?: Pt;
  readonly widthFraction?: number; // tblW type=pct: w/5000 (1.0 = full content width)
  readonly widthType?: 'auto' | 'dxa' | 'pct' | 'nil';
  readonly layout?: 'auto' | 'fixed';
  readonly defaultCellMargins?: CellMargins;
  readonly borders?: CellBorders;
  // ECMA-376 §17.4.27 (w:jc) / xlsx <printOptions horizontalCentered>. Centers
  // or right-aligns a table narrower than the content width; absent ⇒ left.
  readonly alignment?: 'left' | 'center' | 'right';
}

export interface TableCell {
  readonly properties: CellProperties;
  readonly content: ReadonlyArray<BodyElement>;
}

export interface TableRow {
  readonly properties: RowProperties;
  readonly cells: ReadonlyArray<TableCell>;
}

export interface Table {
  readonly native?: NativeBag;
  readonly properties: TableProperties;
  readonly grid: ReadonlyArray<Pt>;
  readonly rows: ReadonlyArray<TableRow>;
}

// ECMA-376 Part 1 §20.4.2.8 — wp:inline picture extent. EMU = English
// Metric Units: 914400 per inch (1 pt = 12700 EMU).
export interface ImageBlock {
  readonly resource?: ResourceId;
  readonly width: Pt;
  readonly height: Pt;
  readonly paragraphProperties: ParagraphProperties;
  // wp:docPr @descr/@title — alternate text for the tagged-PDF Figure (/Alt).
  readonly altText?: string;
}

// ECMA-376 Part 1 §20 — DrawingML shapes (wps:wsp inside a w:drawing).
// A standalone shape (a paragraph whose only content is a shape drawing)
// collapses to a ShapeBlock, mirroring ImageBlock; it carries the paragraph's
// properties for block spacing / alignment.

// §20.1.10.55 prstGeom path command (custGeom <a:pathLst>), in path-space
// units (a:path @w/@h). quad is elevated to cubic at layout; arc uses
// DrawingML angle convention (1/60000°, clockwise, y-down) and is converted
// to beziers by the geometry layer.
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

export interface CustomGeometry {
  readonly pathWidth: number;
  readonly pathHeight: number;
  readonly commands: ReadonlyArray<CustomPathCmd>;
}

// §20.1.10.55 prstGeom / §20.1.9.11 custGeom.
export interface ShapeGeometry {
  readonly kind: 'preset' | 'custom';
  readonly preset?: string; // kind==='preset', e.g. 'roundRect'
  readonly adjust?: ReadonlyMap<string, number>; // a:gd name → raw val (usually 0..100000)
  readonly custom?: CustomGeometry; // kind==='custom'
}

export type ShapeFillKind = 'none' | 'solid';

export interface ShapeFill {
  readonly kind: ShapeFillKind;
  readonly colorHex?: string; // resolved 6-hex (kind==='solid')
}

export type ShapeDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'sysDash'
  | 'sysDot';

// §20.1.2.2.24 a:ln — outline.
export interface ShapeLine {
  readonly width?: Pt; // a:ln @w; default 0.75pt
  readonly colorHex?: string; // resolved 6-hex
  readonly dash?: ShapeDash; // a:prstDash @val
  readonly cap?: 'flat' | 'round' | 'square'; // a:ln @cap (flat=butt)
  readonly fill?: 'solid' | 'none'; // a:ln/a:noFill ⇒ no visible stroke
}

// §20.1.7.6 a:xfrm — rotation (1/60000°, clockwise) + flips.
export interface ShapeTransform {
  readonly rotation60k?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

// wps:txbx (w:txbxContent) + wps:bodyPr — text inside a shape.
export interface ShapeTextBody {
  readonly content: ReadonlyArray<BodyElement>;
  readonly insetLeft?: Pt;
  readonly insetTop?: Pt;
  readonly insetRight?: Pt;
  readonly insetBottom?: Pt;
  readonly anchor?: 't' | 'ctr' | 'b'; // vertical anchor
}

export interface ShapeBlock {
  readonly width: Pt; // wp:extent cx (fallback a:ext cx)
  readonly height: Pt; // wp:extent cy
  readonly geometry: ShapeGeometry;
  readonly fill: ShapeFill;
  readonly line?: ShapeLine;
  readonly transform?: ShapeTransform;
  readonly text?: ShapeTextBody;
  readonly paragraphProperties: ParagraphProperties;
  // wp:docPr @descr/@title — alternate text for the tagged-PDF Figure (/Alt).
  readonly altText?: string;
}

// ECMA-376 Part 1 §21.2 — DrawingML charts. A chart is referenced from a
// w:drawing (a:graphicData uri=…/chart → c:chart r:id) and its data lives in a
// separate chart part (e.g. word/charts/chart1.xml). The parsed Chart is keyed
// by that relationship id and supplied to the renderer alongside the body
// (mirroring how image bytes are supplied), so ChartBlock only carries the ref.
export type ChartType = 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'unknown';

export interface ChartDataPoint {
  readonly idx: number;
  readonly colorHex: string; // c:dPt per-point colour override
}

export interface ChartSeries {
  readonly name?: string;
  readonly values: ReadonlyArray<number>; // c:val/c:yVal numCache (idx-ordered, gaps → 0)
  readonly xValues?: ReadonlyArray<number>; // c:xVal numCache (scatter — paired with values)
  readonly colorHex?: string; // c:spPr solidFill
  readonly pointColors?: ReadonlyArray<ChartDataPoint>; // c:dPt overrides (pie slices)
}

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
}

export interface ChartBlock {
  readonly chartRelId: string; // c:chart @r:id (resolve against the document's rels)
  readonly width: Pt;
  readonly height: Pt;
  readonly paragraphProperties: ParagraphProperties;
  // wp:docPr @descr/@title — alternate text for the tagged-PDF Figure (/Alt).
  readonly altText?: string;
}

export type BodyElement =
  | { readonly kind: 'paragraph'; readonly paragraph: Paragraph }
  | { readonly kind: 'table'; readonly table: Table }
  | { readonly kind: 'image'; readonly image: ImageBlock }
  | { readonly kind: 'shape'; readonly shape: ShapeBlock }
  | { readonly kind: 'chart'; readonly chart: ChartBlock };

// ECMA-376 Part 1 §17.6 — Sections.
export interface PageSize {
  readonly width: Pt;
  readonly height: Pt;
  readonly orientation?: 'portrait' | 'landscape';
}

export interface PageMargins {
  readonly top: Pt;
  readonly right: Pt;
  readonly bottom: Pt;
  readonly left: Pt;
  readonly header?: Pt;
  readonly footer?: Pt;
}

export type HeaderFooterType = 'default' | 'first' | 'even';

export interface HeaderFooterReference {
  readonly type: HeaderFooterType;
  readonly relationshipId: string;
}

export interface SectionProperties {
  readonly pageSize?: PageSize;
  readonly margins?: PageMargins;
  readonly headers: ReadonlyArray<HeaderFooterReference>;
  readonly footers: ReadonlyArray<HeaderFooterReference>;
  // ECMA-376 §17.10.6 — w:titlePg toggle in sectPr. When true the first page
  // of the section uses the `first` header/footer references.
  readonly titlePg?: boolean;
  // ECMA-376 §17.15.1.36 — w:evenAndOddHeaders toggle in word/settings.xml
  // (document-wide, not per-section). When true even-numbered pages use the
  // `even` header/footer references.
  readonly evenAndOddHeaders?: boolean;
  // §17.6.4 w:cols — multi-column section layout.
  readonly columns?: SectionColumns;
}

// §17.6.4 — column definitions: equal-width count + gutter, or explicit
// per-column widths/gutters (w:col children).
export interface SectionColumns {
  readonly count: number;
  readonly spacePt: number;
  readonly explicit?: ReadonlyArray<{ readonly widthPt: number; readonly spacePt: number }>;
}

// A document can contain multiple sections (ECMA-376 §17.6.17). Each section's
// sectPr lives either inside a paragraph's pPr (mid-document break) or as the
// final child of w:body (final section). The Section descriptor below records
// the properties and the upper-exclusive bound in BodyElement[] that the
// section covers — section i applies to body[sections[i-1].endIndex..sections[i].endIndex).
export interface Section {
  readonly properties: SectionProperties;
  readonly endIndex: number;
}

export interface DocumentModel {
  readonly body: ReadonlyArray<BodyElement>;
  readonly styleSheet: StyleSheet;
  readonly numbering?: Numbering;
  readonly section?: SectionProperties;
}

// Document metadata (PDF /Info-shaped, sourced from docProps/core.xml and/or
// caller options). Lives in the model so FlowDoc can carry it format-neutrally.
export interface DocumentInfo {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  // Producer defaults to "Ream" if not provided.
  readonly producer?: string;
  // ISO 8601 dates; converted to PDF date format (D:YYYYMMDDHHmmSS).
  readonly creationDate?: Date;
  readonly modificationDate?: Date;
}
