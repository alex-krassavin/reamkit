// Styled page renderer for BodyElement[] (paragraphs + tables).
//
// Pipeline:
//   1. Pre-scan body recursively → collect every (variant, gid) actually used,
//      embed each font variant once (subsetted).
//   2. For each body block, lay out:
//        paragraph → list of styled lines (existing path)
//        table     → column widths + per-cell laid-out content + row heights
//   3. Paginate: stack laid-out blocks; advance cursorY; break to new page
//      when a block (or table row) does not fit.
//   4. Emit:
//        borders pass (graphics ops outside BT/ET) draws cell rectangles
//        text pass    (BT…ET) shifts Tm and emits Tj per token
//
// Justify alignment is M4. Table row split across pages currently stays on
// one page if it fits; otherwise the whole row is moved to the next page.
// vMerge / gridSpan / nested-tables inside cells parse but render as plain
// independent cells.

import type {
  BodyElement,
  Border,
  BorderStyle,
  CellBorders,
  Chart,
  ChartBlock,
  DocumentInfo,
  HeaderFooterReference,
  HeaderFooterType,
  ImageBlock,
  MathNode,
  Numbering,
  Paragraph,
  ParagraphProperties,
  Run,
  Section,
  SectionProperties,
  ShapeBlock,
  ShapeDash,
  ShapeGeometry,
  ShapeLine,
  StyleSheet,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';
import type { FontRegistry, ParsedTtf } from '@/core/font';
import type { FamilyKey } from '@/core/fonts';
import type { Hyphenator } from '@/core/hyphenation';
import type { Item } from '@/core/line-breaker';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type {
  ChartLabel,
  ChartPolygon,
  ChartPolyline,
  ChartRect,
  ChartWedge,
} from '@/pdf/chart-geometry';
import type { EmbeddedImage, PreparedImage } from '@/pdf/image-xobject';
import type { MathDrawItem, MathVariant, MeasureMath } from '@/pdf/math-layout';
import type { PdfDict, PdfRef, PdfValue } from '@/pdf/objects';
import type { PathSegment, StrokeStyle, VectorPath, VectorShape } from '@/pdf/vector-graphics';
import type { AttachedFile } from '@/pdf/embedded-file';
import type { StructNode, StructType } from '@/pdf/struct-tree';
import type { SignaturePlaceholder } from '@/pdf/signature';
import type { ResourceId } from '@/core/ir';
import type { EmbeddedFont, FontMeasure } from '@/pdf/cid-font';
import { ResourceStore, halfPtToPt } from '@/core/ir';
import { shapeText } from '@/core/font';
import { resolveFamilyKey } from '@/core/fonts';
import {
  analyzeString,
  computeBidi,
  hasBidiCharacters,
  reorderVisual,
  reverseByCodePoint,
} from '@/core/bidi';
import { FORCED_BREAK, breakLines } from '@/core/line-breaker';
import { NumberingState } from '@/core/numbering';
import {
  DEFAULT_RESOLVED_PARAGRAPH,
  DEFAULT_RESOLVED_RUN,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';

import { arcPoint, arcToBeziers } from '@/pdf/arc-to-bezier';
import { buildChartScene } from '@/pdf/chart-geometry';
import { createFontMeasure, embedTtfFont } from '@/pdf/cid-font';
import { buildSrgbIccProfile } from '@/pdf/icc-profile';
import { addImage, prepareImage } from '@/pdf/image-xobject';
import { emitStyledPdf } from '@/pdf/styled-page-emitter';
import { layoutMath, mathGlyphSegments, variantStyle } from '@/pdf/math-layout';
import { dict, name, ref, stream, unicodeString } from '@/pdf/objects';
import { customPaths, presetPaths, rectPath } from '@/pdf/preset-geometry';
import { PathBuilder, emitVectorShape } from '@/pdf/vector-graphics';
import { embedAssociatedFile } from '@/pdf/embedded-file';
import { StructTreeBuilder } from '@/pdf/struct-tree';
import { addSignaturePlaceholder } from '@/pdf/signature';
import { PdfDocument } from '@/pdf/writer';
import { buildXmpPacket } from '@/pdf/xmp';

// PDF/A profiles: part 1 (ISO 19005-1, PDF 1.4) / 2 (ISO 19005-2) / 3
// (ISO 19005-3, both PDF 1.7); conformance level a (tagged) / b (visual) /
// u (Unicode — only 2/3).
export type PdfALevel =
  | 'PDF/A-1b'
  | 'PDF/A-1a'
  | 'PDF/A-2b'
  | 'PDF/A-2u'
  | 'PDF/A-2a'
  | 'PDF/A-3b'
  | 'PDF/A-3u'
  | 'PDF/A-3a';

export interface PdfAProfile {
  readonly part: 1 | 2 | 3;
  readonly level: 'a' | 'b' | 'u';
  readonly tagged: boolean; // level a
  readonly version: '1.4' | '1.7'; // part 1 → 1.4, else 1.7
}

// Decompose a "PDF/A-<part><level>" string. The 7th/8th chars are the part and
// level (e.g. "PDF/A-2u" → part 2, level u).
function parsePdfAProfile(pdfA: PdfALevel): PdfAProfile {
  const part = Number(pdfA.charAt(6)) as 1 | 2 | 3;
  const level = pdfA.charAt(7) as 'a' | 'b' | 'u';
  return { part, level, tagged: level === 'a', version: part === 1 ? '1.4' : '1.7' };
}

export interface StyledRenderOptions {
  readonly registry: FontRegistry;
  // Per-run font resolution: when supplied, each text run picks the registry of
  // its declared family (sans→roboto / serif→tinos / mono→cousine via the run's
  // w:ascii) instead of always using `registry`. Absent ⇒ single-family (every
  // run uses `registry`), byte-identical to before. `registry` remains the
  // guaranteed fallback for math/chart/default glyphs and any missing family.
  readonly registriesByFamily?: ReadonlyMap<FamilyKey, FontRegistry>;
  // The document's OWN embedded fonts (word/fonts/*.odttf, de-obfuscated), keyed
  // by normalized font name. A run whose w:ascii matches one renders with the
  // real font — glyph-exact, no substitution. Highest priority.
  readonly embeddedFonts?: ReadonlyMap<string, FontRegistry>;
  readonly styles: StyleSheet;
  readonly numbering?: Numbering;
  // Single-section legacy entry-point. If `sections` is set it takes
  // precedence and `section` is ignored.
  readonly section?: SectionProperties;
  // ECMA-376 §17.6 — ordered list of sections. Each section's endIndex is the
  // exclusive bound into the body array (section N covers
  // body[sections[N-1].endIndex..sections[N].endIndex)).
  readonly sections?: ReadonlyArray<Section>;
  readonly headersFooters?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  // Content-addressed binary store; image nodes reference it by ResourceId.
  readonly resources?: ResourceStore;
  // Parsed charts keyed by relationship id (ChartBlock.chartRelId). Supplied by
  // the converter, which resolves the chart parts from the package.
  readonly charts?: ReadonlyMap<string, Chart>;
  readonly pageWidth?: number;
  readonly pageHeight?: number;
  readonly marginLeft?: number;
  readonly marginRight?: number;
  readonly marginTop?: number;
  readonly marginBottom?: number;
  // Optional Liang hyphenator. When set, each word token is split at allowed
  // hyphenation positions and offered to Knuth-Plass as potential break
  // points (with a small disincentive). Improves justified paragraph rags.
  readonly hyphenator?: Hyphenator;
  // Optional /Info dictionary metadata (ISO 32000-1 §14.3.3). Unset fields
  // are omitted; if any field is set a PDF /Info entry is emitted.
  readonly info?: DocumentInfo;
  // When set, emit a PDF/A-conformant file: an OutputIntent with an embedded
  // sRGB ICC profile, document XMP /Metadata (the pdfaid identifier), /ID, and
  // subset-tagged fonts with a /CIDSet. The profile picks the rest:
  //   part 1 → PDF 1.4 + flattened image alpha (no transparency);
  //   part 2/3 → PDF 1.7 + preserved transparency (image /SMask + page group);
  //   part 3 → may carry embedded associated files (see `attachments`);
  //   level a → tagged (logical structure); b → visual; u → b + Unicode mapping.
  readonly pdfA?: PdfALevel;
  // Emit a tagged PDF (ISO 32000-1 §14.8): a /StructTreeRoot describing reading
  // order, marked content (BDC/EMC + MCID) on body text, and /Artifact marking
  // of page decoration. Implied by pdfA: 'PDF/A-1a'. Independent of PDF/A
  // otherwise (a plain tagged PDF is useful on its own).
  readonly tagged?: boolean;
  // Document natural language (BCP 47, e.g. "en-US", "ru-RU") for the tagged-PDF
  // catalog /Lang (§14.9.2). Defaults to "en-US". The docx converter fills this
  // from the document's default w:lang.
  readonly language?: string;
  // Files to embed as associated files (catalog /AF + /Names /EmbeddedFiles).
  // Only emitted for plain PDF and PDF/A-3 (PDF/A-1/2 forbid arbitrary embedded
  // files); ignored for PDF/A-1/2. The docx/xlsx converters can embed the source
  // document automatically via `embedSource`.
  readonly attachments?: ReadonlyArray<AttachedFile>;
  // Emit an (invisible) signature field + signature dictionary with placeholder
  // /ByteRange and /Contents (ISO 32000 §12.8). The result is an UNSIGNED PDF;
  // pass it to signPdf() to fill the placeholder with a real PKCS#7 signature.
  readonly signaturePlaceholder?: SignaturePlaceholder;
}

// Re-exported from the document model (moved there so FlowDoc can carry it).
export type { DocumentInfo } from '@/core/document-model';

export const A4_WIDTH = 595;
export const A4_HEIGHT = 842;
const TWIP_TO_PT = 1 / 20;
const EIGHTH_PT = 1 / 8;
const DEFAULT_CELL_PADDING_TWIPS = 108;
const DEFAULT_BORDER_SIZE_EIGHTH = 4;
void DEFAULT_BORDER_SIZE_EIGHTH;

const encoder = new TextEncoder();

export interface FontResource {
  readonly resourceName: string;
  readonly parsed: ParsedTtf;
  // Pure measurement/encoding (no PdfDocument): layout measures with this and
  // emit encodes with it; the PDF font objects are created in the emit phase.
  readonly measure: FontMeasure;
  // Glyphs collected by the usage walk — the emit phase subsets to these.
  readonly gids: ReadonlySet<number>;
}

export interface TextToken {
  readonly kind: 'text';
  readonly text: string;
  readonly isSpace: boolean;
  readonly resolvedRun: ResolvedRunProperties;
  readonly font: FontResource;
  readonly fontSizePt: number;
  readonly widthPt: number;
  // UAX #9 embedding level of this token's characters (0 for pure-LTR docs).
  // Even = LTR, odd = RTL. Used for visual reordering at emit time.
  readonly bidiLevel: number;
}

export interface ImageToken {
  readonly kind: 'image';
  readonly imageResourceName: string;
  readonly widthPt: number;
  readonly heightPt: number;
  // Constants kept to satisfy Token consumers — they never read these for images.
  readonly isSpace: false;
  readonly bidiLevel: number;
}

// A math draw item with its font already resolved (math-layout emits a variant;
// the tokenizer maps it to a FontResource so the emit phase needs no fonts).
type ResolvedMathItem =
  | {
      readonly kind: 'glyph';
      readonly x: number;
      readonly y: number;
      readonly text: string;
      readonly font: FontResource;
      readonly sizePt: number;
    }
  | {
      readonly kind: 'rule';
      readonly x: number;
      readonly y: number;
      readonly w: number;
      readonly h: number;
    }
  | {
      readonly kind: 'path';
      readonly segments: ReadonlyArray<PathSegment>;
      readonly strokeWidthPt?: number;
      readonly fill?: boolean;
    };

// An inline OfficeMath object — an atomic box straddling the baseline (its own
// ascent/descent extend the line height).
export interface MathToken {
  readonly kind: 'math';
  readonly items: ReadonlyArray<ResolvedMathItem>;
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly isSpace: false;
  readonly bidiLevel: number;
}

export type Token = TextToken | ImageToken | MathToken;

export interface Line {
  readonly tokens: ReadonlyArray<Token>;
  readonly contentWidthPt: number;
  readonly maxFontSizePt: number;
  readonly availableWidthPt: number;
  readonly firstLine: boolean;
  readonly resolved: ResolvedParagraphProperties;
  isLastInParagraph: boolean;
  // Max ascent/descent contributed by math tokens (0 when none) — they straddle
  // the baseline, so the line height/descent must grow to fit them.
  readonly mathAscentPt?: number;
  readonly mathDescentPt?: number;
}

type LaidOutBlock =
  | ParagraphBlock
  | TableBlock
  | ImageBlockLaidOut
  | ShapeBlockLaidOut
  | ChartBlockLaidOut;

// A chart laid out into draw primitives in a LOCAL frame (origin = chart box
// bottom-left, y-up). Pagination translates these to the page position. Vector
// primitives reuse the shape pass; text primitives reuse the text (line) pass.
interface ChartShapePrim {
  readonly paths: ReadonlyArray<VectorPath>;
  readonly fillColorHex?: string;
  readonly stroke?: StrokeStyle;
}
interface ChartTextPrim {
  readonly line: Line;
  readonly x: number; // local baseline origin
  readonly y: number;
}
interface ChartLayout {
  readonly shapes: ReadonlyArray<ChartShapePrim>;
  readonly texts: ReadonlyArray<ChartTextPrim>;
}
interface ChartBlockLaidOut {
  readonly kind: 'chart';
  readonly widthPt: number;
  readonly heightPt: number;
  readonly layout: ChartLayout;
  readonly resolvedAlignment: 'left' | 'center' | 'right' | 'both' | 'distribute';
  readonly spacingBeforePt: number;
  readonly spacingAfterPt: number;
  readonly altText?: string;
}

interface ImageBlockLaidOut {
  readonly kind: 'image';
  readonly widthPt: number;
  readonly heightPt: number;
  readonly resolvedAlignment: 'left' | 'center' | 'right' | 'both' | 'distribute';
  readonly resourceName: string;
  readonly spacingBeforePt: number;
  readonly spacingAfterPt: number;
  readonly altText?: string;
}

// A DrawingML shape laid out for placement. Geometry is built in the local
// y-up frame at (widthPt × heightPt); the page-placement transform (rotation
// about the centre + flips + translate) is finalised at pagination, once the
// shape's page position is known.
interface ShapeBlockLaidOut {
  readonly kind: 'shape';
  readonly widthPt: number;
  readonly heightPt: number;
  readonly paths: ReadonlyArray<VectorPath>;
  readonly fillColorHex?: string;
  readonly stroke?: StrokeStyle;
  readonly rotation60k: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly resolvedAlignment: 'left' | 'center' | 'right' | 'both' | 'distribute';
  readonly spacingBeforePt: number;
  readonly spacingAfterPt: number;
  // Text box (wps:txbx) laid out within the inset rect, anchored vertically.
  readonly textLines: ReadonlyArray<Line>;
  readonly textHeightPt: number;
  readonly insetLeftPt: number;
  readonly insetRightPt: number;
  readonly insetTopPt: number;
  readonly insetBottomPt: number;
  readonly anchor: 't' | 'ctr' | 'b';
  readonly altText?: string;
}

interface ParagraphBlock {
  readonly kind: 'paragraph';
  readonly resolved: ResolvedParagraphProperties;
  readonly lines: ReadonlyArray<Line>;
  readonly heightPt: number;
  readonly spacingBeforePt: number;
  readonly spacingAfterPt: number;
  // ECMA-376 §17.3.3.1 — the paragraph carries a forced page break (w:br
  // w:type="page"); subsequent blocks start on a new page.
  readonly pageBreakAfter?: boolean;
  // Tagged PDF: when this paragraph is a list item (w:numPr), its list id and
  // nesting level (w:ilvl) so pagination can build the L/LI/LBody structure.
  readonly list?: { readonly numId: string; readonly level: number };
}

type MergeRole = 'standalone' | 'start' | 'middle' | 'end';

interface CellLayout {
  readonly widthPt: number;
  readonly padTopPt: number;
  readonly padRightPt: number;
  readonly padBottomPt: number;
  readonly padLeftPt: number;
  readonly borders: CellBorders;
  readonly shadingColorHex?: string;
  readonly lines: ReadonlyArray<Line>;
  // Nested tables (a w:tbl inside this cell) rendered below the lines.
  readonly nestedTables?: ReadonlyArray<TableBlock>;
  readonly contentHeightPt: number;
  readonly totalHeightPt: number;
  readonly colStart: number;
  readonly colSpan: number;
  readonly mergeRole: MergeRole;
}

interface RowLayout {
  readonly heightPt: number;
  readonly cells: ReadonlyArray<CellLayout>;
  readonly columnXOffsets: ReadonlyArray<number>;
  readonly rowIdx: number;
  readonly rowCount: number;
  // ECMA-376 §17.4.49 (w:tblHeader) / xlsx _xlnm.Print_Titles — a header row that
  // repeats at the top of each page the table spills onto.
  readonly isHeader?: boolean;
  // xlsx manual <rowBreaks> — force a page break before this row.
  readonly breakBefore?: boolean;
}

interface TableBlock {
  readonly kind: 'table';
  readonly rows: ReadonlyArray<RowLayout>;
  readonly heightPt: number;
  readonly totalWidthPt: number;
  // Grid column count — needed when a nested table is re-emitted inside a cell.
  readonly colCount: number;
  // Horizontal offset from the left margin for a center/right-aligned table
  // narrower than the content width (0 for the default left alignment).
  readonly xOffsetPt: number;
}

// PageDoc draft items (ir-design §6, stage 3c): the positioned, layout-output
// vocabulary a page is made of. Coordinates are PDF points in the PDF page
// frame (y-up, bottom-left origin) for now — the top-left flip and branded Pt
// coordinates land when the types stabilize against the svg-writer (stage 6).
export interface PageItemBase {
  // Tagged PDF (§14.8): the logical structure node this item's content belongs
  // to. Set only on body content in tagged mode; undefined text in the line
  // pass is treated as an artifact. Ignored when not tagging.
  readonly structId?: number;
  // Tagged PDF: explicitly mark this item as a pagination artifact (running
  // header/footer, §14.8.2.2.2). Distinguishes header/footer text from
  // not-yet-tagged body content so it is typed /Artifact /Pagination, never a P.
  readonly artifact?: 'pagination';
}

// A laid-out line of text (tokens carry their fonts/sizes/positions).
export interface TextLineItem extends PageItemBase {
  readonly type: 'line';
  readonly line: Line;
  readonly originX: number;
  readonly baselineY: number;
}

// One edge of a table-cell frame.
export interface BorderItem extends PageItemBase {
  readonly type: 'border';
  readonly side: 'top' | 'right' | 'bottom' | 'left';
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly borderSizePt: number;
  readonly borderColorHex: string;
}

// A filled rectangle (cell shading).
export interface FillItem extends PageItemBase {
  readonly type: 'fill';
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly fillColorHex: string;
}

// A placed raster image (resource name binds to the page XObject dict).
export interface ImageItem extends PageItemBase {
  readonly type: 'image';
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly imageResourceName: string;
}

// DrawingML vector geometry.
export interface ShapeItem extends PageItemBase {
  readonly type: 'shape';
  readonly shape: VectorShape;
}

export type PageItem = TextLineItem | BorderItem | FillItem | ImageItem | ShapeItem;
// Canonical paint order of a laid-out page (oop-design §3.2): fills under
// everything, then images, borders, vector shapes, finally text. ONE owner
// for every writer (pdf emit, svg, future canvas) — and the switch is
// exhaustive, so a new PageItem kind refuses to compile until each group
// has a home. Order within a group is the layout's emission order.
export interface PagePaintPlan {
  readonly fills: ReadonlyArray<FillItem>;
  readonly images: ReadonlyArray<ImageItem>;
  readonly borders: ReadonlyArray<BorderItem>;
  readonly shapes: ReadonlyArray<ShapeItem>;
  readonly lines: ReadonlyArray<TextLineItem>;
}

export function paintPlan(commands: ReadonlyArray<PageItem>): PagePaintPlan {
  const fills: Array<FillItem> = [];
  const images: Array<ImageItem> = [];
  const borders: Array<BorderItem> = [];
  const shapes: Array<ShapeItem> = [];
  const lines: Array<TextLineItem> = [];
  for (const c of commands) {
    switch (c.type) {
      case 'fill':
        fills.push(c);
        break;
      case 'image':
        images.push(c);
        break;
      case 'border':
        borders.push(c);
        break;
      case 'shape':
        shapes.push(c);
        break;
      case 'line':
        lines.push(c);
        break;
      default:
        assertNeverPageItem(c);
    }
  }
  return { fills, images, borders, shapes, lines };
}

function assertNeverPageItem(item: never): never {
  throw new Error(`Unhandled PageItem kind: ${String((item as PageItem).type)}`);
}

export type DrawCommand = PageItem; // internal alias during the stage-3 migration

// The seam between the two halves of the pipeline (ir-design §6 / stage 3):
// everything the emit phase needs from layout. `pages` is the PageDoc draft —
// positioned DrawCommands per page; the resource maps and struct tree ride
// along until stage 3b/3c make layout PdfDocument-free and the types public.
export interface LaidOutDocument {
  readonly pages: ReadonlyArray<LaidOutPage>;
  // Content-addressed binaries the items reference (images) — ir-design §6.
  readonly resources: ResourceStore;
  readonly fontResources: Map<string, FontResource>;
  readonly imageResources: Map<ResourceId, ImageResource>;
  readonly structBuilder: StructTreeBuilder | undefined;
  readonly sectionCtxs: ReadonlyArray<SectionRenderCtx>;
  readonly pdfaProfile: PdfAProfile | undefined;
  readonly tagged: boolean;
}

export function renderStyledPdf(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): Uint8Array {
  const laid = layoutStyledDocument(body, options);
  const doc = new PdfDocument();
  return emitStyledPdf(laid, options, doc);
}

// Layout phase (@experimental — the FlowDoc→PageDoc transform of ir-design §7):
// body → positioned pages (PageItems), font/image resources,
// logical structure. Still takes `doc` because font/image embedding currently
// happens up-front (the layout measures with the embedded fonts) — stage 3b
// splits collect/measure from embed and drops the parameter.
export function layoutStyledDocument(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): LaidOutDocument {
  const sectionList = resolveSectionList(body, options);

  const numberedBody = applyNumbering(body, options.numbering);
  const numberedHeadersFooters = applyNumberingToHeadersFooters(
    options.headersFooters,
    options.numbering,
  );

  // Tagged PDF (ISO 32000-1 §14.8) — implied by PDF/A-1a. When on, paginate
  // builds a logical structure tree and emit marks body text / artifacts.
  const pdfaProfile = options.pdfA ? parsePdfAProfile(options.pdfA) : undefined;
  const tagged = options.tagged === true || (pdfaProfile?.tagged ?? false);
  const structBuilder = tagged ? new StructTreeBuilder() : undefined;
  const fontResources = collectFontResources(numberedBody, numberedHeadersFooters, options);
  const imageResources = collectImageResources(numberedBody, numberedHeadersFooters, options);

  // Pre-compute per-section render context (geometry + header/footer bands).
  const sectionCtxs: Array<SectionRenderCtx> = sectionList.map((s) =>
    buildSectionContext(s, options, numberedHeadersFooters, fontResources),
  );

  // Layout each body block within its owning section's content width.
  let sectionIdx = 0;
  const blocks: Array<LaidOutBlock> = numberedBody.map((el, idx) => {
    while (sectionIdx < sectionCtxs.length - 1 && idx >= sectionCtxs[sectionIdx]!.endIndex) {
      sectionIdx++;
    }
    const ctx = sectionCtxs[sectionIdx]!;
    return layoutBodyElement(
      el,
      options,
      fontResources,
      imageResources,
      ctx.contentWidth,
      ctx.pageContentHeight,
    );
  });

  const pages = paginateSections(blocks, sectionCtxs, structBuilder, options.language ?? 'en-US');

  return {
    pages,
    resources: options.resources ?? new ResourceStore(),
    fontResources,
    imageResources,
    structBuilder,
    sectionCtxs,
    pdfaProfile,
    tagged,
  };
}

// Emit phase: PageDoc draft → PDF objects (content streams, page dicts,
// catalog, PDF/A apparatus, structure tree, signature placeholder) → bytes.

// Priority for page geometry:
//   1. explicit value in StyledRenderOptions (test/library caller override)
//   2. value from section properties (sectPr/pgSz/pgMar from the docx)
//   3. A4 + 1-inch margins fallback
export interface PageDimensions {
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly headerOffsetPt: number;
  readonly footerOffsetPt: number;
}

function resolvePageDimensions(
  options: StyledRenderOptions,
  section: SectionProperties | undefined,
): PageDimensions {
  const sectionPageWidth = section?.pageSize !== undefined ? section.pageSize.width : undefined;
  const sectionPageHeight = section?.pageSize !== undefined ? section.pageSize.height : undefined;
  const sectionLeft = section?.margins?.left !== undefined ? section.margins.left : undefined;
  const sectionRight = section?.margins?.right !== undefined ? section.margins.right : undefined;
  const sectionTop = section?.margins?.top !== undefined ? section.margins.top : undefined;
  const sectionBottom = section?.margins?.bottom !== undefined ? section.margins.bottom : undefined;
  const headerOffsetPt = section?.margins?.header ?? 720 * TWIP_TO_PT;
  const footerOffsetPt = section?.margins?.footer ?? 720 * TWIP_TO_PT;

  return {
    pageWidth: options.pageWidth ?? sectionPageWidth ?? A4_WIDTH,
    pageHeight: options.pageHeight ?? sectionPageHeight ?? A4_HEIGHT,
    marginLeft: options.marginLeft ?? sectionLeft ?? 72,
    marginRight: options.marginRight ?? sectionRight ?? 72,
    marginTop: options.marginTop ?? sectionTop ?? 72,
    marginBottom: options.marginBottom ?? sectionBottom ?? 72,
    headerOffsetPt,
    footerOffsetPt,
  };
}

export interface SectionRenderCtx {
  readonly endIndex: number;
  readonly properties: SectionProperties;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly marginLeft: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly contentWidth: number;
  readonly pageContentHeight: number;
  readonly headerSet: HeaderFooterSet;
  readonly footerSet: HeaderFooterSet;
  readonly titlePg: boolean;
  readonly evenAndOddHeaders: boolean;
}

// Pick the final list of sections to render. Precedence:
//   1. options.sections (the typical docx path).
//   2. options.section (single-section legacy / xlsx path).
//   3. A single fallback section covering the whole body.
function resolveSectionList(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): ReadonlyArray<Section> {
  if (options.sections && options.sections.length > 0) {
    return options.sections;
  }
  if (options.section) {
    return [{ properties: options.section, endIndex: body.length }];
  }
  return [{ properties: { headers: [], footers: [] }, endIndex: body.length }];
}

function buildSectionContext(
  section: Section,
  options: StyledRenderOptions,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  fontResources: ReadonlyMap<string, FontResource>,
): SectionRenderCtx {
  const dims = resolvePageDimensions(options, section.properties);
  const contentWidth = dims.pageWidth - dims.marginLeft - dims.marginRight;
  const headerSet = layoutHeaderSet(
    section.properties,
    headersFooters,
    options,
    fontResources,
    contentWidth,
    dims.marginLeft,
    dims.pageHeight,
    dims.headerOffsetPt,
  );
  const footerSet = layoutFooterSet(
    section.properties,
    headersFooters,
    options,
    fontResources,
    contentWidth,
    dims.marginLeft,
    dims.footerOffsetPt,
  );
  return {
    endIndex: section.endIndex,
    properties: section.properties,
    pageWidth: dims.pageWidth,
    pageHeight: dims.pageHeight,
    marginLeft: dims.marginLeft,
    marginTop: dims.marginTop,
    marginBottom: dims.marginBottom,
    contentWidth,
    pageContentHeight: dims.pageHeight - dims.marginTop - dims.marginBottom,
    headerSet,
    footerSet,
    titlePg: section.properties.titlePg === true,
    evenAndOddHeaders: section.properties.evenAndOddHeaders === true,
  };
}

function applyNumberingToHeadersFooters(
  hf: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
  numbering: Numbering | undefined,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  if (!hf || hf.size === 0) return new Map();
  if (!numbering || numbering.abstractNums.size === 0) return hf;
  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const [key, value] of hf) {
    out.set(key, applyNumbering(value, numbering));
  }
  return out;
}

function refByType(
  refs: ReadonlyArray<HeaderFooterReference> | undefined,
  type: HeaderFooterType,
): HeaderFooterReference | undefined {
  if (!refs || refs.length === 0) return undefined;
  return refs.find((r) => r.type === type);
}

type HfBand = 'default' | 'first' | 'even';

interface HeaderFooterSet {
  readonly default: Array<DrawCommand>;
  readonly first: Array<DrawCommand>;
  readonly even: Array<DrawCommand>;
}

// Tag every command in a header/footer band as a pagination artifact so the
// tagged-PDF emit keeps it out of the structure tree (§14.8.2.2.2). A no-op for
// non-tagged output — the field is simply ignored at emit.
function markPagination(cmds: Array<DrawCommand>): Array<DrawCommand> {
  return cmds.map((c) => ({ ...c, artifact: 'pagination' as const }));
}

// Pre-layout the three possible header/footer bands. Each band is an array of
// draw commands ready to be merged onto a body page. Missing bands fall back
// to the default during page assembly, so producing an empty array here is
// fine (we always check before falling back).
function layoutHeaderSet(
  section: SectionProperties,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  contentWidth: number,
  marginLeft: number,
  pageHeight: number,
  headerOffsetPt: number,
): HeaderFooterSet {
  const band = (type: HeaderFooterType): Array<DrawCommand> => {
    const ref = refByType(section.headers, type);
    if (!ref) return [];
    const content = headersFooters.get(ref.relationshipId);
    if (!content) return [];
    const blocks = laidOutBlocksFor(content, options, fontResources, contentWidth);
    return markPagination(drawBlocksSequentially(blocks, marginLeft, pageHeight - headerOffsetPt));
  };
  return { default: band('default'), first: band('first'), even: band('even') };
}

function layoutFooterSet(
  section: SectionProperties,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  contentWidth: number,
  marginLeft: number,
  footerOffsetPt: number,
): HeaderFooterSet {
  const band = (type: HeaderFooterType): Array<DrawCommand> => {
    const ref = refByType(section.footers, type);
    if (!ref) return [];
    const content = headersFooters.get(ref.relationshipId);
    if (!content) return [];
    const blocks = laidOutBlocksFor(content, options, fontResources, contentWidth);
    const totalHeight = blocks.reduce(
      (sum, b) =>
        sum +
        (b.kind === 'paragraph' ? b.spacingBeforePt + b.heightPt + b.spacingAfterPt : b.heightPt),
      0,
    );
    return markPagination(drawBlocksSequentially(blocks, marginLeft, footerOffsetPt + totalHeight));
  };
  return { default: band('default'), first: band('first'), even: band('even') };
}

function pickBand(set: HeaderFooterSet, band: HfBand): Array<DrawCommand> {
  if (band === 'first') return set.first.length > 0 ? set.first : set.default;
  if (band === 'even') return set.even.length > 0 ? set.even : set.default;
  return set.default;
}

function bandForPage(
  pageInSection: number,
  globalPageIdx: number,
  titlePg: boolean,
  evenAndOddHeaders: boolean,
): HfBand {
  // titlePg is per-section: the first page of each section uses 'first' when
  // the section's titlePg toggle is on.
  if (pageInSection === 0 && titlePg) return 'first';
  // evenAndOddHeaders is document-wide and keyed off the human-visible page
  // number (page 2, 4, ... → even).
  if (evenAndOddHeaders && (globalPageIdx + 1) % 2 === 0) return 'even';
  return 'default';
}

function laidOutBlocksFor(
  elements: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  contentWidth: number,
): Array<LaidOutBlock> {
  return elements.map((el) =>
    layoutBodyElement(el, options, fontResources, undefined, contentWidth),
  );
}

function layoutBodyElement(
  el: BodyElement,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
  maxHeight?: number,
): LaidOutBlock {
  if (el.kind === 'paragraph') {
    return layoutParagraphBlock(el.paragraph, options, fontResources, imageResources, contentWidth);
  }
  if (el.kind === 'table') {
    return layoutTableBlock(el.table, options, fontResources, imageResources, contentWidth);
  }
  if (el.kind === 'image') {
    return layoutImageBlock(el.image, imageResources, contentWidth);
  }
  if (el.kind === 'chart') {
    return layoutChartBlock(el.chart, options, fontResources, contentWidth, maxHeight);
  }
  return layoutShapeBlock(
    el.shape,
    options,
    fontResources,
    imageResources,
    contentWidth,
    maxHeight,
  );
}

export interface ImageResource {
  readonly resourceName: string;
  // Decoded/validated at layout time (the probe); the emit phase replays it
  // without touching the source bytes again.
  readonly prepared: PreparedImage;
}

// 1 inch = 914400 EMU = 72 pt, so 1 pt = 12700 EMU.
const EMU_PER_PT = 12700;

function layoutImageBlock(
  image: ImageBlock,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
): ImageBlockLaidOut {
  let widthPt: number = image.width;
  let heightPt: number = image.height;
  if (widthPt > contentWidth) {
    const scale = contentWidth / widthPt;
    widthPt = contentWidth;
    heightPt = heightPt * scale;
  }
  const res = image.resource ? imageResources?.get(image.resource) : undefined;
  const resolvedAlignment = image.paragraphProperties.alignment ?? 'left';
  return {
    kind: 'image',
    widthPt,
    heightPt,
    resolvedAlignment,
    resourceName: res?.resourceName ?? '',
    spacingBeforePt: image.paragraphProperties.spacingBefore ?? 0,
    spacingAfterPt: image.paragraphProperties.spacingAfter ?? 0,
    ...(image.altText ? { altText: image.altText } : {}),
  };
}

// a:ln default width when @w is absent (9525 EMU = 0.75pt).
const DEFAULT_LINE_WIDTH_EMU = 9525;

// Word's default text-box insets (§20.1.2.1) — 0.1" L/R, 0.05" T/B in EMU.
const DEFAULT_INSET_LR_EMU = 91440;
const DEFAULT_INSET_TB_EMU = 45720;

function layoutShapeBlock(
  shape: ShapeBlock,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
  maxHeight?: number,
): ShapeBlockLaidOut {
  let widthPt: number = shape.width;
  let heightPt: number = shape.height;
  // Clamp width to the content area like images, scaling height to keep aspect.
  if (widthPt > contentWidth && widthPt > 0) {
    const scale = contentWidth / widthPt;
    widthPt = contentWidth;
    heightPt *= scale;
  }
  // Clamp height to the page content area so an oversized shape stays on one
  // page (shapes are atomic). Scale width with it to preserve aspect.
  if (maxHeight !== undefined && heightPt > maxHeight && heightPt > 0) {
    const scale = maxHeight / heightPt;
    widthPt *= scale;
    heightPt = maxHeight;
  }
  const paths = buildShapePaths(shape.geometry, widthPt, heightPt);
  const fillColorHex = shape.fill.kind === 'solid' ? shape.fill.colorHex : undefined;
  const stroke = buildStroke(shape.line);
  const t = shape.transform;
  const pp = shape.paragraphProperties;

  const text = shape.text;
  const insetLeftPt = text?.insetLeft ?? DEFAULT_INSET_LR_EMU / EMU_PER_PT;
  const insetRightPt = text?.insetRight ?? DEFAULT_INSET_LR_EMU / EMU_PER_PT;
  const insetTopPt = text?.insetTop ?? DEFAULT_INSET_TB_EMU / EMU_PER_PT;
  const insetBottomPt = text?.insetBottom ?? DEFAULT_INSET_TB_EMU / EMU_PER_PT;
  const textLines: Array<Line> = [];
  let textHeightPt = 0;
  if (text && text.content.length > 0) {
    const innerWidth = Math.max(1, widthPt - insetLeftPt - insetRightPt);
    for (const el of text.content) {
      if (el.kind !== 'paragraph') continue; // tables/nested shapes in a text box: out of scope
      const blk = layoutParagraphBlock(
        el.paragraph,
        options,
        fontResources,
        imageResources,
        innerWidth,
      );
      for (const line of blk.lines) {
        textLines.push(line);
        textHeightPt += computeLineHeight(line, blk.resolved);
      }
      textHeightPt += blk.spacingAfterPt;
    }
  }

  return {
    kind: 'shape',
    widthPt,
    heightPt,
    paths,
    ...(fillColorHex ? { fillColorHex } : {}),
    ...(stroke ? { stroke } : {}),
    rotation60k: t?.rotation60k ?? 0,
    flipH: t?.flipH ?? false,
    flipV: t?.flipV ?? false,
    resolvedAlignment: pp.alignment ?? 'left',
    spacingBeforePt: pp.spacingBefore ?? 0,
    spacingAfterPt: pp.spacingAfter ?? 0,
    textLines,
    textHeightPt,
    insetLeftPt,
    insetRightPt,
    insetTopPt,
    insetBottomPt,
    anchor: text?.anchor ?? 't',
    ...(shape.altText ? { altText: shape.altText } : {}),
  };
}

function buildShapePaths(
  geometry: ShapeGeometry,
  widthPt: number,
  heightPt: number,
): Array<VectorPath> {
  if (geometry.kind === 'preset') {
    const paths = presetPaths(
      geometry.preset ?? 'rect',
      widthPt,
      heightPt,
      geometry.adjust ?? new Map(),
    );
    return paths ?? [rectPath(widthPt, heightPt)];
  }
  if (geometry.custom) return customPaths(geometry.custom, widthPt, heightPt);
  return [rectPath(widthPt, heightPt)];
}

function buildStroke(line: ShapeLine | undefined): StrokeStyle | undefined {
  if (!line || line.fill === 'none') return undefined;
  const widthPt = line.width ?? DEFAULT_LINE_WIDTH_EMU / EMU_PER_PT;
  const dash = line.dash && line.dash !== 'solid' ? dashPattern(line.dash, widthPt) : undefined;
  // DrawingML 'flat' cap is PDF butt; round/square map straight through.
  const cap: StrokeStyle['cap'] | undefined = line.cap === 'flat' ? 'butt' : line.cap;
  return {
    colorHex: line.colorHex ?? '000000',
    widthPt,
    ...(dash ? { dash } : {}),
    ...(cap ? { cap } : {}),
  };
}

// Dash patterns expressed in multiples of the line width (a common rendering
// convention), in points. 'solid' has no pattern.
function dashPattern(dash: ShapeDash, w: number): Array<number> | undefined {
  const u = Math.max(w, 0.1);
  switch (dash) {
    case 'solid':
      return undefined;
    case 'dot':
      return [u, 2 * u];
    case 'dash':
      return [4 * u, 3 * u];
    case 'dashDot':
      return [4 * u, 3 * u, u, 3 * u];
    case 'lgDash':
      return [8 * u, 3 * u];
    case 'lgDashDot':
      return [8 * u, 3 * u, u, 3 * u];
    case 'sysDash':
      return [3 * u, u];
    case 'sysDot':
      return [u, u];
  }
}

// Build the `cm` matrix that places a shape's local y-up frame on the page at
// bottom-left (pageX, pageY), rotated about its centre and optionally flipped.
// DrawingML rot is clockwise in y-down space ⇒ a negative angle in PDF y-up.
function buildShapeTransform(
  pageX: number,
  pageY: number,
  widthPt: number,
  heightPt: number,
  rotation60k: number,
  flipH: boolean,
  flipV: boolean,
): [number, number, number, number, number, number] {
  const theta = (-rotation60k / 60000) * (Math.PI / 180);
  const sx = flipH ? -1 : 1;
  const sy = flipV ? -1 : 1;
  const cos = Math.cos(theta);
  const sin = Math.sin(theta);
  const a = sx * cos;
  const b = sx * sin;
  const c = -sy * sin;
  const d = sy * cos;
  const cxL = widthPt / 2;
  const cyL = heightPt / 2;
  const centerX = pageX + cxL;
  const centerY = pageY + cyL;
  const e = centerX - (a * cxL + c * cyL);
  const f = centerY - (b * cxL + d * cyL);
  return [a, b, c, d, e, f];
}

function layoutChartBlock(
  block: ChartBlock,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  contentWidth: number,
  maxHeight?: number,
): ChartBlockLaidOut {
  let widthPt: number = block.width;
  let heightPt: number = block.height;
  if (widthPt > contentWidth && widthPt > 0) {
    const scale = contentWidth / widthPt;
    widthPt = contentWidth;
    heightPt *= scale;
  }
  if (maxHeight !== undefined && heightPt > maxHeight && heightPt > 0) {
    const scale = maxHeight / heightPt;
    widthPt *= scale;
    heightPt = maxHeight;
  }
  const chart = options.charts?.get(block.chartRelId);
  const { variant } = options.registry.resolveByStyle(false, false);
  const font = fontResources.get(variant);
  const layout =
    chart && font ? buildChartLayout(chart, widthPt, heightPt, font) : { shapes: [], texts: [] };
  const pp = block.paragraphProperties;
  // Figure alt text: the drawing's docPr description, else the chart's own title.
  const altText = block.altText ?? chart?.title;
  return {
    kind: 'chart',
    widthPt,
    heightPt,
    layout,
    resolvedAlignment: pp.alignment ?? 'left',
    spacingBeforePt: pp.spacingBefore ?? 0,
    spacingAfterPt: pp.spacingAfter ?? 0,
    ...(altText ? { altText } : {}),
  };
}

// Build a chart's draw primitives (local y-up frame, origin bottom-left) from
// the pure geometry scene. Chart text uses the regular font. Unsupported chart
// types (no scene) fall back to a light bounding box reserving the space.
function buildChartLayout(
  chart: Chart,
  widthPt: number,
  heightPt: number,
  font: FontResource,
): ChartLayout {
  const measure = (text: string, sizePt: number): number => font.measure.textWidthPt(text, sizePt);
  const scene = buildChartScene(chart, widthPt, heightPt, measure);
  if (!scene) {
    return {
      shapes: [
        { paths: [rectPath(widthPt, heightPt)], stroke: { colorHex: 'D9D9D9', widthPt: 1 } },
      ],
      texts: [],
    };
  }
  const shapes: Array<ChartShapePrim> = [];
  // Area-fill polygons sit at the bottom of the z-order (below gridlines/labels).
  for (const pg of scene.polygons ?? []) shapes.push(polygonPrim(pg));
  for (const r of scene.rects) shapes.push(rectPrim(r));
  for (const p of scene.polylines) shapes.push(polylinePrim(p));
  for (const w of scene.wedges) shapes.push(wedgePrim(w));
  const texts: Array<ChartTextPrim> = scene.labels.map((l) => labelPrim(l, font));
  return { shapes, texts };
}

function rectAtPath(x: number, y: number, w: number, h: number): VectorPath {
  return new PathBuilder()
    .moveTo(x, y)
    .lineTo(x + w, y)
    .lineTo(x + w, y + h)
    .lineTo(x, y + h)
    .close()
    .build();
}

function rectPrim(r: ChartRect): ChartShapePrim {
  return {
    paths: [rectAtPath(r.x, r.y, r.w, r.h)],
    ...(r.fillHex ? { fillColorHex: r.fillHex } : {}),
    ...(r.strokeHex ? { stroke: { colorHex: r.strokeHex, widthPt: r.strokeWidthPt ?? 1 } } : {}),
  };
}

function polylinePrim(p: ChartPolyline): ChartShapePrim {
  const b = new PathBuilder();
  p.points.forEach(([x, y], i) => (i === 0 ? b.moveTo(x, y) : b.lineTo(x, y)));
  return { paths: [b.build()], stroke: { colorHex: p.strokeHex, widthPt: p.widthPt } };
}

function polygonPrim(p: ChartPolygon): ChartShapePrim {
  const b = new PathBuilder();
  p.points.forEach(([x, y], i) => (i === 0 ? b.moveTo(x, y) : b.lineTo(x, y)));
  b.close();
  return {
    paths: [b.build()],
    fillColorHex: p.fillHex,
    ...(p.strokeHex ? { stroke: { colorHex: p.strokeHex, widthPt: p.widthPt ?? 1 } } : {}),
  };
}

function wedgePrim(w: ChartWedge): ChartShapePrim {
  const start = arcPoint(w.cx, w.cy, w.r, w.r, w.startRad);
  const b = new PathBuilder()
    .moveTo(w.cx, w.cy)
    .lineTo(start[0], start[1])
    .append(arcToBeziers(w.cx, w.cy, w.r, w.r, w.startRad, w.sweepRad))
    .close();
  return {
    paths: [b.build()],
    fillColorHex: w.fillHex,
    ...(w.strokeHex ? { stroke: { colorHex: w.strokeHex, widthPt: 1 } } : {}),
  };
}

function labelPrim(l: ChartLabel, font: FontResource): ChartTextPrim {
  const line = makeChartLabelLine(l.text, font, l.sizePt, l.colorHex);
  const w = line.contentWidthPt;
  const x = l.align === 'center' ? l.x - w / 2 : l.align === 'right' ? l.x - w : l.x;
  return { line, x, y: l.y };
}

// A minimal single-token Line for a positioned chart label.
function makeChartLabelLine(
  text: string,
  font: FontResource,
  sizePt: number,
  colorHex: string,
): Line {
  const widthPt = font.measure.textWidthPt(text, sizePt);
  const token: TextToken = {
    kind: 'text',
    text,
    isSpace: false,
    resolvedRun: {
      ...DEFAULT_RESOLVED_RUN,
      colorHex,
      fontSizePt: halfPtToPt(Math.round(sizePt * 2)),
    },
    font,
    fontSizePt: sizePt,
    widthPt,
    bidiLevel: 0,
  };
  return {
    tokens: [token],
    contentWidthPt: widthPt,
    maxFontSizePt: sizePt,
    availableWidthPt: widthPt,
    firstLine: true,
    resolved: DEFAULT_RESOLVED_PARAGRAPH,
    isLastInParagraph: true,
  };
}

function collectImageResources(
  body: ReadonlyArray<BodyElement>,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
): Map<ResourceId, ImageResource> {
  const out = new Map<ResourceId, ImageResource>();
  if (!options.resources || options.resources.size === 0) return out;

  const seen = new Set<ResourceId>();
  const visit = (elements: ReadonlyArray<BodyElement>) => {
    for (const el of elements) {
      if (el.kind === 'image') {
        if (el.image.resource) seen.add(el.image.resource);
      } else if (el.kind === 'paragraph') {
        for (const run of el.paragraph.runs) {
          if (run.inlineImage?.resource) seen.add(run.inlineImage.resource);
        }
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) {
          for (const cell of row.cells) visit(cell.content);
        }
      } else if (el.kind === 'shape') {
        if (el.shape.text) visit(el.shape.text.content);
      }
    }
  };
  visit(body);
  for (const hf of headersFooters.values()) visit(hf);

  // Only PDF/A-1 forbids transparency; PDF/A-2/3 keep the image soft mask.
  const flattenAlpha = options.pdfA ? parsePdfAProfile(options.pdfA).part === 1 : false;
  let counter = 0;
  for (const resourceId of seen) {
    const bytes = options.resources.get(resourceId);
    if (!bytes) continue;
    // An unsupported or corrupt image must not abort the whole document — skip
    // it. It then has no resource name, so nothing is drawn for it (its layout
    // box still reserves space). prepareImage is the pure decode/validate
    // expert; the emit phase replays its result, so skip semantics match by
    // construction.
    let prepared: PreparedImage;
    try {
      prepared = prepareImage(bytes, { flattenAlpha });
    } catch {
      continue;
    }
    counter++;
    out.set(resourceId, { resourceName: `Im${counter}`, prepared });
  }
  return out;
}

// Sequential, non-paginated draw used for header/footer bands. Tables in
// headers/footers are uncommon in practice and skipped here for simplicity.
function drawBlocksSequentially(
  blocks: ReadonlyArray<LaidOutBlock>,
  startX: number,
  startY: number,
): Array<DrawCommand> {
  const out: Array<DrawCommand> = [];
  let cursorY = startY;
  for (const block of blocks) {
    if (block.kind !== 'paragraph') continue;
    cursorY -= block.spacingBeforePt;
    for (const line of block.lines) {
      const h = computeLineHeight(line, block.resolved);
      cursorY -= h;
      const indentLeft =
        block.resolved.indentLeft + (line.firstLine ? block.resolved.indentFirstLine : 0);
      const offset = alignmentOffset(
        block.resolved.alignment,
        line.contentWidthPt,
        line.availableWidthPt,
      );
      out.push({
        type: 'line',
        line,
        originX: startX + indentLeft + offset,
        baselineY: cursorY + lineDescent(line),
      });
    }
    cursorY -= block.spacingAfterPt;
  }
  return out;
}

// Walk the body in document order, advancing list counters and prepending a
// marker Run (e.g. "1.", "•") plus a tab to every paragraph that references
// a list level. The level's pPr indent is applied if the paragraph has no
// indent of its own — that pushes the body text right of the marker, while a
// hanging indent (negative indentFirstLine) places the marker itself
// to the left of the body indent.
function applyNumbering(
  body: ReadonlyArray<BodyElement>,
  numbering: Numbering | undefined,
): Array<BodyElement> {
  if (!numbering || numbering.abstractNums.size === 0) return body.map((b) => b);
  const state = new NumberingState();

  const transformParagraph = (p: Paragraph): Paragraph => {
    const ref = p.properties.numbering;
    if (!ref) return p;
    const marker = state.resolveMarker(numbering, ref);
    if (marker === null) return p;
    const instance = numbering.numInstances.get(ref.numId);
    const abstractNum = instance ? numbering.abstractNums.get(instance.abstractNumId) : undefined;
    const level = abstractNum?.levels.get(ref.ilvl);

    const markerRun: Run = {
      text: `${marker}\t`,
      properties: level?.runProperties ?? {},
    };
    const newProps: ParagraphProperties = mergeIndentFromLevel(
      p.properties,
      level?.paragraphProperties,
    );
    return { properties: newProps, runs: [markerRun, ...p.runs] };
  };

  const visit = (el: BodyElement): BodyElement => {
    if (el.kind === 'paragraph') {
      return { kind: 'paragraph', paragraph: transformParagraph(el.paragraph) };
    }
    if (el.kind === 'image') return el;
    // Shapes don't advance list counters (a floating/inline shape isn't a list
    // item). Numbered lists inside a shape's text box are out of scope for M5.
    if (el.kind === 'shape') return el;
    if (el.kind === 'chart') return el;
    return {
      kind: 'table',
      table: {
        ...el.table,
        rows: el.table.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({ ...cell, content: cell.content.map(visit) })),
        })),
      },
    };
  };

  return body.map(visit);
}

function mergeIndentFromLevel(
  paragraphProps: ParagraphProperties,
  levelProps: ParagraphProperties | undefined,
): ParagraphProperties {
  if (!levelProps) return paragraphProps;
  const out: { -readonly [K in keyof ParagraphProperties]: ParagraphProperties[K] } = {
    ...paragraphProps,
  };
  if (out.indentLeft === undefined && levelProps.indentLeft !== undefined) {
    out.indentLeft = levelProps.indentLeft;
  }
  if (out.indentRight === undefined && levelProps.indentRight !== undefined) {
    out.indentRight = levelProps.indentRight;
  }
  if (out.indentFirstLine === undefined && levelProps.indentFirstLine !== undefined) {
    out.indentFirstLine = levelProps.indentFirstLine;
  }
  return out;
}

// Resolve a run's font to a (fontKey, parsed) pair. In multi-family mode
// (options.registriesByFamily set) the key is `${familyKey}:${variant}` chosen
// from the run's declared family; otherwise it is just the variant — which keeps
// the single-family output byte-identical to before.
function runFontKeyAndParsed(
  options: StyledRenderOptions,
  ascii: string | undefined,
  bold: boolean,
  italic: boolean,
): { fontKey: string; parsed: ParsedTtf } {
  // The document's own embedded font (word/fonts/*.odttf) — glyph-exact, takes
  // priority over any substitution.
  if (ascii && options.embeddedFonts) {
    const name = ascii.trim().toLowerCase();
    const emb = options.embeddedFonts.get(name);
    if (emb) {
      const { variant, parsed } = emb.resolveByStyle(bold, italic);
      return { fontKey: `embed:${name}:${variant}`, parsed };
    }
  }
  const byFamily = options.registriesByFamily;
  if (byFamily && byFamily.size > 0) {
    let key = resolveFamilyKey(ascii);
    let reg = byFamily.get(key);
    if (!reg) {
      key = byFamily.keys().next().value as FamilyKey;
      reg = byFamily.get(key)!;
    }
    const { variant, parsed } = reg.resolveByStyle(bold, italic);
    return { fontKey: `${key}:${variant}`, parsed };
  }
  const { variant, parsed } = options.registry.resolveByStyle(bold, italic);
  return { fontKey: variant, parsed };
}

// Tolerant lookup for placeholder fonts (inline-image / math outer run) whose
// declared family may not have been embedded — falls back to any embedded font.
function lookupFont(resources: ReadonlyMap<string, FontResource>, fontKey: string): FontResource {
  return resources.get(fontKey) ?? resources.values().next().value!;
}

function collectFontResources(
  body: ReadonlyArray<BodyElement>,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
): Map<string, FontResource> {
  const used = new Map<string, { parsed: ParsedTtf; gids: Set<number> }>();
  const addRun = (
    run: { text: string; properties: { bold?: boolean; italic?: boolean; styleId?: string } },
    para: Paragraph,
  ) => {
    const resolved = resolveRunProperties(run.properties, para.properties, options.styles);
    const { fontKey, parsed } = runFontKeyAndParsed(
      options,
      resolved.fontFamily.ascii,
      resolved.bold,
      resolved.italic,
    );
    let bucket = used.get(fontKey);
    if (!bucket) {
      bucket = { parsed, gids: new Set<number>() };
      used.set(fontKey, bucket);
    }
    // Collect the SHAPED glyphs (ligatures applied), matching exactly what the
    // emit phase encodes — otherwise a ligature glyph (e.g. fi) would be
    // rendered but pruned from the subset / absent from the /CIDSet and
    // /ToUnicode (PDF/A §6.3.5 / §6.3.8).
    const shaped = shapeText(
      run.text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      parsed.kerning,
      parsed.joiningForms,
    );
    for (const g of shaped.gids) bucket.gids.add(g);
  };

  // Inline math glyphs go to the variant the layout engine will use (italic for
  // letters, etc.) so the correct subset is embedded.
  const addMath = (node: MathNode) => {
    for (const seg of mathGlyphSegments(node)) {
      const { bold, italic } = variantStyle(seg.variant);
      const { variant, parsed } = options.registry.resolveByStyle(bold, italic);
      let bucket = used.get(variant);
      if (!bucket) {
        bucket = { parsed, gids: new Set<number>() };
        used.set(variant, bucket);
      }
      for (const ch of seg.text) bucket.gids.add(parsed.glyphForCodepoint(ch.codePointAt(0)!));
    }
  };

  const visit = (elements: ReadonlyArray<BodyElement>) => {
    for (const el of elements) {
      if (el.kind === 'paragraph') {
        for (const run of el.paragraph.runs) {
          if (run.math) {
            addMath(run.math);
            continue;
          }
          // Skip runs that carry an inline image only (no text glyphs needed).
          if (run.inlineImage && !run.text) continue;
          addRun(run, el.paragraph);
        }
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) {
          for (const cell of row.cells) visit(cell.content);
        }
      } else if (el.kind === 'shape') {
        if (el.shape.text) visit(el.shape.text.content);
      } else if (el.kind === 'chart') {
        const chart = options.charts?.get(el.chart.chartRelId);
        if (chart) {
          const reg = options.registry.resolveByStyle(false, false);
          let bucket = used.get(reg.variant);
          if (!bucket) {
            bucket = { parsed: reg.parsed, gids: new Set<number>() };
            used.set(reg.variant, bucket);
          }
          const add = (s: string): void => {
            for (const ch of s) bucket.gids.add(reg.parsed.glyphForCodepoint(ch.codePointAt(0)!));
          };
          if (chart.title) add(chart.title);
          for (const c of chart.categories) add(c);
          for (const sr of chart.series) if (sr.name) add(sr.name);
          add('0123456789.,-%() '); // value-axis tick labels
        }
      }
      // image-block elements use no fonts
    }
  };
  visit(body);
  for (const hf of headersFooters.values()) visit(hf);

  if (used.size === 0) {
    const regular = options.registry.resolveByStyle(false, false);
    used.set(regular.variant, { parsed: regular.parsed, gids: new Set([0]) });
  }

  const out = new Map<string, FontResource>();
  let counter = 0;
  for (const [variant, info] of used) {
    counter++;
    const resourceName = `F${counter}`;
    out.set(variant, {
      resourceName,
      parsed: info.parsed,
      measure: createFontMeasure(info.parsed),
      gids: info.gids,
    });
  }
  return out;
}

function layoutParagraphBlock(
  paragraph: Paragraph,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
): ParagraphBlock {
  const baseResolved = resolveParagraphProperties(paragraph.properties, options.styles);
  const baseDir = paragraphBaseDirection(paragraph, baseResolved);
  // RTL paragraphs default to right alignment. We only override the cascade's
  // 'left' default (which is also what an absent jc collapses to) — explicit
  // center/right/justify are preserved.
  const resolved: ResolvedParagraphProperties =
    baseDir === 'rtl' && baseResolved.alignment === 'left'
      ? { ...baseResolved, alignment: 'right' }
      : baseResolved;
  const tokens = tokenizeParagraph(
    paragraph,
    options,
    fontResources,
    imageResources,
    contentWidth,
    baseDir,
  );
  const firstLineWidth = paragraphMaxWidth(resolved, contentWidth, true);
  const otherWidth = paragraphMaxWidth(resolved, contentWidth, false);
  const lines = wrap(tokens, firstLineWidth, otherWidth, resolved, options.hyphenator);

  let heightPt = 0;
  for (const line of lines) heightPt += computeLineHeight(line, resolved);
  const numbering = paragraph.properties.numbering;
  return {
    kind: 'paragraph',
    resolved,
    lines,
    heightPt,
    spacingBeforePt: resolved.spacingBefore,
    spacingAfterPt: resolved.spacingAfter,
    ...(numbering ? { list: { numId: numbering.numId, level: numbering.ilvl } } : {}),
    ...(paragraph.runs.some((r) => r.pageBreak) ? { pageBreakAfter: true } : {}),
  };
}

// Object Replacement Character — stands in for an inline image when computing
// BiDi levels (treated as a neutral object).
const OBJECT_REPLACEMENT = 0xfffc;
// Explicit directional formatting codes used to honour run-level w:rtl.
const RLE = 0x202b;
const PDF_FMT = 0x202c;
const LRE = 0x202a;

interface RunPlan {
  readonly run: Paragraph['runs'][number];
  readonly resolvedRun: ResolvedRunProperties;
  readonly font: FontResource;
  readonly fontSizePt: number;
  readonly isImage: boolean;
  readonly imageWidthPt: number;
  readonly imageHeightPt: number;
  readonly imageResourceName: string;
  readonly math?: {
    readonly items: ReadonlyArray<ResolvedMathItem>;
    readonly widthPt: number;
    readonly ascentPt: number;
    readonly descentPt: number;
  };
}

// Resolve a math box's variant-tagged glyph items to concrete FontResources so
// the emit phase needs no font lookup.
function resolveMathItems(
  items: ReadonlyArray<MathDrawItem>,
  fontFor: (v: MathVariant) => FontResource,
): Array<ResolvedMathItem> {
  return items.map((it): ResolvedMathItem => {
    if (it.kind === 'glyph') {
      return {
        kind: 'glyph',
        x: it.x,
        y: it.y,
        text: it.text,
        font: fontFor(it.variant),
        sizePt: it.sizePt,
      };
    }
    if (it.kind === 'rule') {
      return { kind: 'rule', x: it.x, y: it.y, w: it.w, h: it.h };
    }
    return {
      kind: 'path',
      segments: it.segments,
      ...(it.strokeWidthPt !== undefined ? { strokeWidthPt: it.strokeWidthPt } : {}),
      ...(it.fill ? { fill: true } : {}),
    };
  });
}

function tokenizeParagraph(
  paragraph: Paragraph,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
  baseDir: 'ltr' | 'rtl',
): Array<Token> {
  // First pass — resolve each run's style and decide image vs text.
  const plans: Array<RunPlan> = paragraph.runs.map((run) => {
    if (run.inlineImage) {
      const naturalW = run.inlineImage.width;
      const widthPt = Math.min(naturalW, contentWidth);
      const scale = naturalW > 0 ? widthPt / naturalW : 1;
      const heightPt = run.inlineImage.height * scale;
      const res = run.inlineImage.resource
        ? imageResources?.get(run.inlineImage.resource)
        : undefined;
      const resolvedRun = resolveRunProperties(
        run.properties,
        paragraph.properties,
        options.styles,
      );
      const { fontKey } = runFontKeyAndParsed(
        options,
        resolvedRun.fontFamily.ascii,
        resolvedRun.bold,
        resolvedRun.italic,
      );
      return {
        run,
        resolvedRun,
        font: lookupFont(fontResources, fontKey),
        fontSizePt: resolvedRun.fontSizePt,
        isImage: true,
        imageWidthPt: widthPt,
        imageHeightPt: heightPt,
        imageResourceName: res?.resourceName ?? '',
      };
    }
    if (run.math) {
      const resolvedRun = resolveRunProperties(
        run.properties,
        paragraph.properties,
        options.styles,
      );
      const sizePt = resolvedRun.fontSizePt;
      const fontFor = (v: MathVariant): FontResource => {
        const { bold, italic } = variantStyle(v);
        const r = options.registry.resolveByStyle(bold, italic);
        return (
          fontResources.get(r.variant) ??
          fontResources.get(options.registry.resolveByStyle(false, false).variant)!
        );
      };
      const measure: MeasureMath = (text, sz, v) => fontFor(v).measure.textWidthPt(text, sz);
      const box = layoutMath(run.math, { sizePt }, measure);
      const { variant } = options.registry.resolveByStyle(resolvedRun.bold, resolvedRun.italic);
      return {
        run,
        resolvedRun,
        font: lookupFont(fontResources, variant),
        fontSizePt: sizePt,
        isImage: false,
        imageWidthPt: 0,
        imageHeightPt: 0,
        imageResourceName: '',
        math: {
          items: resolveMathItems(box.items, fontFor),
          widthPt: box.width,
          ascentPt: box.ascent,
          descentPt: box.descent,
        },
      };
    }
    const resolvedRun = resolveRunProperties(run.properties, paragraph.properties, options.styles);
    const { fontKey } = runFontKeyAndParsed(
      options,
      resolvedRun.fontFamily.ascii,
      resolvedRun.bold,
      resolvedRun.italic,
    );
    return {
      run,
      resolvedRun,
      font: lookupFont(fontResources, fontKey),
      fontSizePt: resolvedRun.fontSizePt,
      isImage: false,
      imageWidthPt: 0,
      imageHeightPt: 0,
      imageResourceName: '',
    };
  });

  // Decide whether BiDi processing is needed at all. Pure-LTR paragraphs with
  // no RTL runs skip it entirely (the overwhelming common case).
  const anyRtlRun = plans.some((p) => p.resolvedRun.rtl);
  const anyBidiChars = plans.some((p) => !p.isImage && hasBidiCharacters(p.run.text));
  const needsBidi = baseDir === 'rtl' || anyRtlRun || anyBidiChars;

  if (!needsBidi) {
    // Fast path — everything at level 0, no reordering.
    return tokenizePlansLtr(plans);
  }

  // Build the BiDi input. RTL runs are wrapped in RLE…PDF (LTR runs in an RTL
  // paragraph in LRE…PDF) so run-level direction overrides neutral resolution.
  const bidiCps: Array<number> = [];
  // For each bidi code point, the index of the "real" code point it maps to,
  // or -1 for inserted control characters.
  const realIndexOfBidi: Array<number> = [];
  let realCount = 0;

  for (const plan of plans) {
    const isObject = plan.isImage || plan.math !== undefined;
    const wrap = plan.resolvedRun.rtl ? RLE : baseDir === 'rtl' && !isObject ? LRE : 0;
    if (wrap) {
      bidiCps.push(wrap);
      realIndexOfBidi.push(-1);
    }
    if (isObject) {
      bidiCps.push(OBJECT_REPLACEMENT);
      realIndexOfBidi.push(realCount);
      realCount++;
    } else {
      for (const ch of plan.run.text) {
        bidiCps.push(ch.codePointAt(0)!);
        realIndexOfBidi.push(realCount);
        realCount++;
      }
    }
    if (wrap) {
      bidiCps.push(PDF_FMT);
      realIndexOfBidi.push(-1);
    }
  }

  const { levels } = computeBidi(bidiCps, baseDir);
  // Project levels back onto real code points.
  const realLevels: Array<number> = new Array(realCount).fill(baseDir === 'rtl' ? 1 : 0);
  for (let i = 0; i < bidiCps.length; i++) {
    const ri = realIndexOfBidi[i]!;
    if (ri >= 0) realLevels[ri] = levels[i]!;
  }

  return tokenizePlansBidi(plans, realLevels);
}

// Fast LTR tokenization — splits each run on whitespace only.
function tokenizePlansLtr(plans: ReadonlyArray<RunPlan>): Array<Token> {
  const tokens: Array<Token> = [];
  for (const plan of plans) {
    if (plan.isImage) {
      tokens.push({
        kind: 'image',
        imageResourceName: plan.imageResourceName,
        widthPt: plan.imageWidthPt,
        heightPt: plan.imageHeightPt,
        isSpace: false,
        bidiLevel: 0,
      });
      continue;
    }
    if (plan.math) {
      tokens.push({
        kind: 'math',
        items: plan.math.items,
        widthPt: plan.math.widthPt,
        ascentPt: plan.math.ascentPt,
        descentPt: plan.math.descentPt,
        isSpace: false,
        bidiLevel: 0,
      });
      continue;
    }
    for (const t of tokenizeText(plan.run.text)) {
      tokens.push({
        kind: 'text',
        text: t.text,
        isSpace: t.isSpace,
        resolvedRun: plan.resolvedRun,
        font: plan.font,
        fontSizePt: plan.fontSizePt,
        widthPt: plan.font.measure.textWidthPt(t.text, plan.fontSizePt),
        bidiLevel: 0,
      });
    }
  }
  return tokens;
}

// BiDi-aware tokenization — splits each run on whitespace boundaries AND on
// embedding-level changes so every token carries a single uniform level.
function tokenizePlansBidi(
  plans: ReadonlyArray<RunPlan>,
  realLevels: ReadonlyArray<number>,
): Array<Token> {
  const tokens: Array<Token> = [];
  let realIdx = 0;
  for (let r = 0; r < plans.length; r++) {
    const plan = plans[r]!;
    if (plan.isImage) {
      tokens.push({
        kind: 'image',
        imageResourceName: plan.imageResourceName,
        widthPt: plan.imageWidthPt,
        heightPt: plan.imageHeightPt,
        isSpace: false,
        bidiLevel: realLevels[realIdx] ?? 0,
      });
      realIdx++;
      continue;
    }
    if (plan.math) {
      tokens.push({
        kind: 'math',
        items: plan.math.items,
        widthPt: plan.math.widthPt,
        ascentPt: plan.math.ascentPt,
        descentPt: plan.math.descentPt,
        isSpace: false,
        bidiLevel: realLevels[realIdx] ?? 0,
      });
      realIdx++;
      continue;
    }
    // Iterate code points, grouping by (isSpace, level).
    const chars = [...plan.run.text];
    let bufStart = 0;
    let curSpace = false;
    let curLevel = -1;
    const flush = (endExclusive: number) => {
      if (endExclusive <= bufStart) return;
      const text = chars.slice(bufStart, endExclusive).join('');
      tokens.push({
        kind: 'text',
        text,
        isSpace: curSpace,
        resolvedRun: plan.resolvedRun,
        font: plan.font,
        fontSizePt: plan.fontSizePt,
        widthPt: plan.font.measure.textWidthPt(text, plan.fontSizePt),
        bidiLevel: curLevel,
      });
    };
    for (let c = 0; c < chars.length; c++) {
      const ch = chars[c]!;
      const isSpace = /\s/.test(ch);
      const level = realLevels[realIdx] ?? 0;
      if (c === 0) {
        curSpace = isSpace;
        curLevel = level;
      } else if (isSpace !== curSpace || level !== curLevel) {
        flush(c);
        bufStart = c;
        curSpace = isSpace;
        curLevel = level;
      }
      realIdx++;
    }
    flush(chars.length);
  }
  return tokens;
}

// Resolve the paragraph's BiDi base direction. An explicit w:bidi sets RTL;
// otherwise we auto-detect from the first strong character so that an
// untagged Hebrew/Arabic paragraph still gets an RTL base (matching how
// viewers render such content).
function paragraphBaseDirection(
  paragraph: Paragraph,
  resolved: ResolvedParagraphProperties,
): 'ltr' | 'rtl' {
  if (resolved.bidi) return 'rtl';
  let text = '';
  for (const run of paragraph.runs) {
    if (!run.inlineImage) text += run.text;
    if (text.length > 64) break; // first strong char is near the start
  }
  if (!hasBidiCharacters(text)) return 'ltr';
  return analyzeString(text, 'auto').paragraphLevel === 1 ? 'rtl' : 'ltr';
}

function paragraphMaxWidth(
  p: ResolvedParagraphProperties,
  contentWidth: number,
  firstLine: boolean,
): number {
  const indentLeft = p.indentLeft;
  const indentRight = p.indentRight;
  const firstLineExtra = firstLine ? p.indentFirstLine : 0;
  return Math.max(1, contentWidth - indentLeft - indentRight - firstLineExtra);
}

function tokenizeText(text: string): Array<{ text: string; isSpace: boolean }> {
  const out: Array<{ text: string; isSpace: boolean }> = [];
  if (text.length === 0) return out;
  const re = /(\s+)|(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1] !== undefined) out.push({ text: m[1], isSpace: true });
    else if (m[2] !== undefined) out.push({ text: m[2], isSpace: false });
  }
  return out;
}

// Penalty value for hyphenation breaks — small enough that KP prefers a
// hyphen break over a badly stretched line, but large enough to avoid
// hyphenating when a clean glue break is available.
const HYPHENATION_PENALTY = 50;

// Line-wrap a token stream using Knuth-Plass. With a hyphenator supplied,
// word tokens are split at allowed positions and each split becomes a
// flagged penalty (KP may "buy" a line break here at a small cost). When
// the chosen break lands on a hyphenation penalty, the last token of the
// line gets a trailing "-" so the reader sees the word was hyphenated.
// One entry per Knuth–Plass item: the displayable token it came from (null
// for glue/penalties/sentinels) and, on hyphenation penalties, the hyphen
// width to fold onto the previous fragment when a line breaks here. Replaces
// four parallel arrays (oop-design §4.3, A6).
interface StreamEntry {
  readonly item: Item;
  readonly token: Token | null;
  readonly hyphenWidthPt?: number;
}

// Tokens → the Knuth–Plass item stream: spaces become glue, images/math are
// atomic boxes, text optionally splits at hyphenation points with flagged
// penalties, and the paragraph closes with infinite glue + a forced break.
function paragraphItemStream(
  tokens: ReadonlyArray<Token>,
  hyphenator: Hyphenator | undefined,
): Array<StreamEntry> {
  const entries: Array<StreamEntry> = [];
  for (const tok of tokens) {
    if (tok.isSpace || tok.kind === 'image' || tok.kind === 'math') {
      // Spaces are glue; images and math boxes are atomic (un-hyphenatable) boxes.
      entries.push({
        item: tok.isSpace
          ? {
              type: 'glue',
              width: tok.widthPt,
              stretch: tok.widthPt * 0.6,
              shrink: tok.widthPt * 0.3,
            }
          : { type: 'box', width: tok.widthPt },
        token: tok,
      });
      continue;
    }
    // Text non-space token. Try hyphenation; if no breaks (or too short), one
    // box covers the whole word.
    const positions = hyphenator ? hyphenator.hyphenate(tok.text) : [];
    if (positions.length === 0) {
      entries.push({ item: { type: 'box', width: tok.widthPt }, token: tok });
      continue;
    }
    const hyphenWidth = tok.font.measure.textWidthPt('-', tok.fontSizePt);
    let prev = 0;
    for (let pi = 0; pi <= positions.length; pi++) {
      const end = pi < positions.length ? positions[pi]! : tok.text.length;
      const fragText = tok.text.substring(prev, end);
      const fragWidth = tok.font.measure.textWidthPt(fragText, tok.fontSizePt);
      const fragTok: Token = {
        kind: 'text',
        text: fragText,
        isSpace: false,
        resolvedRun: tok.resolvedRun,
        font: tok.font,
        fontSizePt: tok.fontSizePt,
        widthPt: fragWidth,
        bidiLevel: tok.bidiLevel,
      };
      entries.push({ item: { type: 'box', width: fragWidth }, token: fragTok });
      if (pi < positions.length) {
        entries.push({
          item: {
            type: 'penalty',
            width: hyphenWidth,
            penalty: HYPHENATION_PENALTY,
            flagged: true,
          },
          token: null,
          hyphenWidthPt: hyphenWidth,
        });
      }
      prev = end;
    }
  }

  // Final glue + forced penalty — same convention as before.
  entries.push({ item: { type: 'glue', width: 0, stretch: 1e6, shrink: 0 }, token: null });
  entries.push({
    item: { type: 'penalty', width: 0, penalty: FORCED_BREAK, flagged: false },
    token: null,
  });
  return entries;
}

// Rebuild one Line from the chosen break range [start, breakIdx): trim
// edge spaces/sentinels, fold the hyphen glyph when the break sits on a
// hyphenation penalty, and aggregate the line metrics.
function lineFromRange(
  entries: ReadonlyArray<StreamEntry>,
  start: number,
  breakIdx: number,
  availableWidthPt: number,
  isFirst: boolean,
  resolved: ResolvedParagraphProperties,
): Line | null {
  let st = start;
  let et = breakIdx;
  // Skip leading nulls / spaces.
  while (st < et && (entries[st]!.token === null || entries[st]!.token?.isSpace)) st++;
  // Trim trailing nulls / spaces.
  while (et > st && (entries[et - 1]!.token === null || entries[et - 1]!.token?.isSpace)) et--;
  if (st >= et) return null;

  const lineTokens: Array<Token> = [];
  for (let i = st; i < et; i++) {
    const ft = entries[i]!.token;
    if (ft) lineTokens.push(ft);
  }

  // If the chosen break is at a hyphenation penalty, fold the hyphen glyph
  // onto the last text token of the line.
  const hyphenWidth = entries[breakIdx]?.hyphenWidthPt;
  if (hyphenWidth !== undefined) {
    const lastIdx = lineTokens.length - 1;
    const last = lineTokens[lastIdx];
    if (last && last.kind === 'text') {
      lineTokens[lastIdx] = {
        ...last,
        text: last.text + '-',
        widthPt: last.widthPt + hyphenWidth,
      };
    }
  }

  const tokenLineSize = (t: Token): number =>
    t.kind === 'text' ? t.fontSizePt : t.kind === 'image' ? t.heightPt : 0;
  let contentWidth = 0;
  let maxSize = 0;
  let mathAscent = 0;
  let mathDescent = 0;
  for (const t of lineTokens) {
    contentWidth += t.widthPt;
    const sz = tokenLineSize(t);
    if (sz > maxSize) maxSize = sz;
    if (t.kind === 'math') {
      mathAscent = Math.max(mathAscent, t.ascentPt);
      mathDescent = Math.max(mathDescent, t.descentPt);
    }
  }
  return {
    tokens: lineTokens,
    contentWidthPt: contentWidth,
    maxFontSizePt: maxSize,
    availableWidthPt,
    firstLine: isFirst,
    resolved,
    isLastInParagraph: false,
    mathAscentPt: mathAscent,
    mathDescentPt: mathDescent,
  };
}

function wrap(
  tokens: ReadonlyArray<Token>,
  firstLineWidth: number,
  otherWidth: number,
  resolved: ResolvedParagraphProperties,
  hyphenator: Hyphenator | undefined,
): Array<Line> {
  if (tokens.length === 0) return [];

  const entries = paragraphItemStream(tokens, hyphenator);
  const { breaks } = breakLines(
    entries.map((e) => e.item),
    [firstLineWidth, otherWidth],
  );

  const lines: Array<Line> = [];
  let start = 0;
  let isFirst = true;
  for (const breakIdx of breaks) {
    const line = lineFromRange(
      entries,
      start,
      breakIdx,
      isFirst ? firstLineWidth : otherWidth,
      isFirst,
      resolved,
    );
    if (line) lines.push(line);
    start = breakIdx + 1;
    isFirst = false;
  }

  if (lines.length > 0) lines[lines.length - 1]!.isLastInParagraph = true;
  return lines;
}

function computeLineHeight(line: Line, p: ResolvedParagraphProperties): number {
  const fontSize = line.maxFontSizePt || 12;
  // A math box straddles the baseline; the line must be at least tall enough to
  // hold its full ascent+descent (plus a little leading).
  const mathH = (line.mathAscentPt ?? 0) + (line.mathDescentPt ?? 0);
  const mathNeed = mathH > 0 ? mathH * 1.05 : 0;
  if (p.spacingLineRule === 'exact' && p.spacingLine > 0) {
    return Math.max(p.spacingLine, mathNeed);
  }
  if (p.spacingLineRule === 'atLeast' && p.spacingLine > 0) {
    return Math.max(fontSize * 1.2, p.spacingLine, mathNeed);
  }
  // "Multiple" spacing is defined in 240ths (240 twips = single). Recover the
  // integer twips before dividing — historically this divided the raw int, and
  // (twips*(1/20))/12 differs from twips/240 in the last ulp.
  const lineTwips = Math.round(p.spacingLine * 20);
  const multiple = lineTwips > 0 ? lineTwips / 240 : 1;
  return Math.max(fontSize * 1.2 * multiple, mathNeed);
}

function lineDescent(line: Line): number {
  const fs = line.maxFontSizePt || 12;
  return Math.max(fs * 0.2, line.mathDescentPt ?? 0);
}

function alignmentOffset(
  alignment: ResolvedParagraphProperties['alignment'],
  lineWidth: number,
  available: number,
): number {
  if (alignment === 'right') return Math.max(0, available - lineWidth);
  if (alignment === 'center') return Math.max(0, (available - lineWidth) / 2);
  return 0;
}

function layoutTableBlock(
  table: Table,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
): TableBlock {
  const columnWidthsPt = computeColumnWidths(table, options, fontResources, contentWidth);
  // Vertical-merge roles are resolved by the readers (CellProperties.merge);
  // standalone cells carry no marker.
  const mergeRoles: Array<Array<MergeRole>> = table.rows.map((r) =>
    r.cells.map((c) => c.properties.merge ?? 'standalone'),
  );

  const rows: Array<RowLayout> = [];
  const colCount = columnWidthsPt.length;
  // Per-column borders of the row above (for a cell's top-border fallback to the
  // cell above's bottom). Rebuilt after each row.
  let aboveBordersByCol: Array<CellBorders | undefined> = [];
  for (let r = 0; r < table.rows.length; r++) {
    const rl = layoutTableRow(
      table.rows[r]!,
      r,
      table.rows.length,
      colCount,
      table.properties,
      columnWidthsPt,
      options,
      fontResources,
      imageResources,
      mergeRoles[r] ?? [],
      aboveBordersByCol,
    );
    const nextAbove: Array<CellBorders | undefined> = [];
    let ci = 0;
    for (const cell of table.rows[r]!.cells) {
      const span = Math.max(1, cell.properties.colSpan ?? 1);
      for (let k = 0; k < span; k++) nextAbove[ci + k] = cell.properties.borders;
      ci += span;
    }
    aboveBordersByCol = nextAbove;
    const rp = table.rows[r]!.properties;
    rows.push(
      rp.isHeader || rp.pageBreakBefore
        ? {
            ...rl,
            ...(rp.isHeader ? { isHeader: true } : {}),
            ...(rp.pageBreakBefore ? { breakBefore: true } : {}),
          }
        : rl,
    );
  }
  const heightPt = rows.reduce((s, r) => s + r.heightPt, 0);
  const totalWidthPt = columnWidthsPt.reduce((s, w) => s + w, 0);
  const xOffsetPt = tableXOffset(table.properties.alignment, contentWidth, totalWidthPt);
  return { kind: 'table', rows, heightPt, totalWidthPt, colCount, xOffsetPt };
}

// ECMA-376 §17.4.27 (w:jc) — horizontal placement of a table narrower than the
// content width. Left (default) ⇒ 0; center ⇒ half the slack; right ⇒ all of it.
function tableXOffset(
  alignment: TableProperties['alignment'],
  contentWidth: number,
  totalWidthPt: number,
): number {
  const slack = contentWidth - totalWidthPt;
  if (slack <= 0 || !alignment || alignment === 'left') return 0;
  return alignment === 'center' ? slack / 2 : slack;
}

// ECMA-376 Part 1 §17.4.20 — tblLayout.
//   "fixed"     → use tblGrid widths verbatim (scaled to tblW if set)
//   "auto"|absent → auto-fit: column widths derived from cell content widths;
//                   tblGrid is treated as a hint only.
function computeColumnWidths(
  table: Table,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  contentWidth: number,
): Array<number> {
  let colCount = table.grid.length;
  for (const row of table.rows) {
    let rowCols = 0;
    for (const cell of row.cells) rowCols += Math.max(1, cell.properties.colSpan ?? 1);
    if (rowCols > colCount) colCount = rowCols;
  }
  if (colCount === 0) return [];

  if (table.properties.layout === 'fixed') {
    return gridWidthsScaled(table, contentWidth, colCount);
  }

  const colNaturalWidths = new Array<number>(colCount).fill(0);
  for (const row of table.rows) {
    let colIdx = 0;
    for (let i = 0; i < row.cells.length; i++) {
      const cell = row.cells[i]!;
      const span = Math.max(1, cell.properties.colSpan ?? 1);
      const padLeft =
        cell.properties.margins?.left ??
        table.properties.defaultCellMargins?.left ??
        DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
      const padRight =
        cell.properties.margins?.right ??
        table.properties.defaultCellMargins?.right ??
        DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
      let maxContent = 0;
      for (const el of cell.content) {
        if (el.kind !== 'paragraph') continue;
        const w = measureSingleLine(el.paragraph, options, fontResources);
        if (w > maxContent) maxContent = w;
      }
      const need = maxContent + padLeft + padRight;
      const perColumn = need / span;
      for (let k = 0; k < span && colIdx + k < colCount; k++) {
        if (perColumn > colNaturalWidths[colIdx + k]!) colNaturalWidths[colIdx + k] = perColumn;
      }
      colIdx += span;
    }
  }

  const naturalTotal = colNaturalWidths.reduce((s, w) => s + w, 0);
  const explicitTarget = explicitTableTargetWidth(table, contentWidth);

  if (naturalTotal === 0) return gridWidthsScaled(table, contentWidth, colCount);

  if (explicitTarget !== undefined && explicitTarget > 0) {
    const scale = explicitTarget / naturalTotal;
    return colNaturalWidths.map((w) => w * scale);
  }

  if (naturalTotal > contentWidth) {
    const scale = contentWidth / naturalTotal;
    return colNaturalWidths.map((w) => w * scale);
  }

  return colNaturalWidths;
}

// Recover the integer twips behind a Pt grid column. Grid ratios and sums were
// historically computed on the raw ints; float-summing the Pt values instead
// would drift in the last ulp and break the byte-identical gate.
function gridColTwips(table: Table): Array<number> {
  return table.grid.map((w) => Math.round(w * 20));
}

function gridWidthsScaled(table: Table, contentWidth: number, colCount: number): Array<number> {
  const gridTwips = gridColTwips(table);
  const totalGridTwips = gridTwips.reduce((s, w) => s + w, 0);
  if (totalGridTwips > 0) {
    const target = totalTableTarget(table, contentWidth);
    return gridTwips.map((w) => (w / totalGridTwips) * target);
  }
  const each = contentWidth / colCount;
  return new Array<number>(colCount).fill(each);
}

function totalTableTarget(table: Table, contentWidth: number): number {
  const explicit = explicitTableTargetWidth(table, contentWidth);
  if (explicit !== undefined && explicit > 0) return explicit;
  const sum = gridColTwips(table).reduce((s, g) => s + g, 0);
  if (sum > 0) return sum * TWIP_TO_PT;
  return contentWidth;
}

function explicitTableTargetWidth(table: Table, contentWidth: number): number | undefined {
  const w = table.properties.widthPt;
  const type = table.properties.widthType;
  if (w !== undefined && w > 0 && type === 'dxa') return w;
  const f = table.properties.widthFraction;
  if (f !== undefined && type === 'pct') return f * contentWidth;
  return undefined;
}

function measureSingleLine(
  paragraph: Paragraph,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
): number {
  let total = 0;
  for (const run of paragraph.runs) {
    const resolved = resolveRunProperties(run.properties, paragraph.properties, options.styles);
    // The same resolution every other phase uses (embedded → per-family →
    // base registry). The old direct registry lookup keyed fontResources by
    // bare variant — with registriesByFamily/embeddedFonts those keys do not
    // exist, so table auto-layout measured with the wrong font or crashed.
    const { fontKey } = runFontKeyAndParsed(
      options,
      resolved.fontFamily.ascii,
      resolved.bold,
      resolved.italic,
    );
    const font = lookupFont(fontResources, fontKey);
    const fontSizePt = resolved.fontSizePt;
    total += font.measure.textWidthPt(run.text, fontSizePt);
  }
  return total;
}

function layoutTableRow(
  row: TableRow,
  rowIdx: number,
  rowCount: number,
  colCount: number,
  tableProps: TableProperties,
  columnWidthsPt: ReadonlyArray<number>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  rowMergeRoles: ReadonlyArray<MergeRole>,
  aboveBordersByCol: ReadonlyArray<CellBorders | undefined>,
): RowLayout {
  const cells: Array<CellLayout> = [];
  const columnXOffsets: Array<number> = [];
  let cursorX = 0;
  let colIdx = 0;
  for (let cellIdx = 0; cellIdx < row.cells.length; cellIdx++) {
    const cell = row.cells[cellIdx]!;
    const span = Math.max(1, cell.properties.colSpan ?? 1);
    let widthPt = 0;
    for (let k = 0; k < span && colIdx + k < columnWidthsPt.length; k++) {
      widthPt += columnWidthsPt[colIdx + k]!;
    }
    columnXOffsets.push(cursorX);
    const mergeRole = rowMergeRoles[cellIdx] ?? 'standalone';
    const leftNeighbor = cellIdx > 0 ? row.cells[cellIdx - 1]!.properties.borders : undefined;
    const cellLayout = layoutTableCell(
      cell,
      tableProps,
      widthPt,
      options,
      fontResources,
      imageResources,
      rowIdx,
      colIdx,
      span,
      rowCount,
      colCount,
      mergeRole,
      leftNeighbor,
      aboveBordersByCol[colIdx],
    );
    cells.push(cellLayout);
    cursorX += widthPt;
    colIdx += span;
  }
  let heightPt = 0;
  for (const c of cells) if (c.totalHeightPt > heightPt) heightPt = c.totalHeightPt;
  if (row.properties.height && row.properties.heightRule === 'exact') {
    heightPt = row.properties.height;
  } else if (row.properties.height && row.properties.heightRule === 'atLeast') {
    heightPt = Math.max(heightPt, row.properties.height);
  }
  return { heightPt, cells, columnXOffsets, rowIdx, rowCount };
}

function layoutTableCell(
  cell: TableCell,
  tableProps: TableProperties,
  widthPt: number,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  rowIdx: number,
  colStart: number,
  colSpan: number,
  rowCount: number,
  colCount: number,
  mergeRole: MergeRole,
  leftNeighborBorders?: CellBorders,
  aboveNeighborBorders?: CellBorders,
): CellLayout {
  const cellMar = cell.properties.margins;
  const tableMar = tableProps.defaultCellMargins;
  const padTopPt = cellMar?.top ?? tableMar?.top ?? 0;
  const padBottomPt = cellMar?.bottom ?? tableMar?.bottom ?? 0;
  const padLeftPt = cellMar?.left ?? tableMar?.left ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
  const padRightPt = cellMar?.right ?? tableMar?.right ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;

  const innerWidth = Math.max(1, widthPt - padLeftPt - padRightPt);
  const lines: Array<Line> = [];
  const nestedTables: Array<TableBlock> = [];
  let contentHeightPt = 0;
  // Continuation cells (vMerge=continue) render no content — their text lives
  // in the 'start' cell above.
  if (mergeRole !== 'middle' && mergeRole !== 'end') {
    for (const el of cell.content) {
      if (el.kind === 'paragraph') {
        const block = layoutParagraphBlock(
          el.paragraph,
          options,
          fontResources,
          imageResources,
          innerWidth,
        );
        for (const line of block.lines) {
          lines.push(line);
          contentHeightPt += computeLineHeight(line, block.resolved);
        }
        contentHeightPt += block.spacingAfterPt;
      } else if (el.kind === 'table') {
        // Nested table (a w:tbl inside this w:tc): lay it out within the cell's
        // inner width; it renders below the cell's paragraph lines.
        const nested = layoutTableBlock(
          el.table,
          options,
          fontResources,
          imageResources,
          innerWidth,
        );
        nestedTables.push(nested);
        contentHeightPt += nested.heightPt;
      }
      // image/shape/chart inside a cell are not yet rendered (skipped).
    }
  }
  const colEnd = colStart + colSpan - 1;
  const borders = resolveCellBorders(
    cell.properties.borders,
    tableProps.borders,
    rowIdx,
    colStart,
    colEnd,
    rowCount,
    colCount,
    mergeRole,
    leftNeighborBorders,
    aboveNeighborBorders,
  );
  const totalHeightPt =
    mergeRole === 'middle' || mergeRole === 'end' ? 0 : padTopPt + contentHeightPt + padBottomPt;
  return {
    widthPt,
    padTopPt,
    padRightPt,
    padBottomPt,
    padLeftPt,
    borders,
    ...(cell.properties.shading?.colorHex
      ? { shadingColorHex: cell.properties.shading.colorHex }
      : {}),
    lines,
    ...(nestedTables.length > 0 ? { nestedTables } : {}),
    contentHeightPt,
    totalHeightPt,
    colStart,
    colSpan,
    mergeRole,
  };
}

// ECMA-376 Part 1 §17.4.39 — cell border resolution.
// Outer cell edges inherit from the table's top/bottom/left/right; internal
// edges inherit from insideH (horizontal) and insideV (vertical). No default
// fallback: if neither cell nor table specifies a border, none is drawn.
//
// For vMerge cells (mergeRole != 'standalone') we suppress the borders that
// would split the merge group — middle/end cells lose their top, start/middle
// cells lose their bottom. Span cells use colStart/colEnd to decide first/last
// column status so a right border still snaps to the last spanned column.
// Style precedence for breaking a same-width border conflict (CSS border-collapse
// / §17.4 ordering): a heavier line style wins.
const BORDER_STYLE_RANK: Readonly<Record<BorderStyle, number>> = {
  double: 5,
  thick: 4,
  single: 3,
  dashed: 2,
  dotted: 1,
  none: 0,
};

// Effective weight of a border for conflict resolution: its width in points
// point (default 4 = ½pt), or 0 when absent / explicitly 'none'.
function borderWeight(b: Border | undefined): number {
  if (!b || b.style === 'none') return 0;
  return b.width ?? 0.5;
}

// §17.4 — the winner of a shared cell edge: the heavier border, ties broken by
// style precedence. A side with weight 0 (absent or 'none') loses to any real
// border; if both are empty the edge stays unbordered.
function heavierBorder(a: Border | undefined, b: Border | undefined): Border | undefined {
  const wa = borderWeight(a);
  const wb = borderWeight(b);
  if (wa !== wb) return wa > wb ? a : b;
  if (wa === 0) return undefined;
  return BORDER_STYLE_RANK[a!.style] >= BORDER_STYLE_RANK[b!.style] ? a : b;
}

function resolveCellBorders(
  cellBorders: CellBorders | undefined,
  tableBorders: CellBorders | undefined,
  rowIdx: number,
  colStart: number,
  colEnd: number,
  rowCount: number,
  colCount: number,
  mergeRole: MergeRole,
  leftNeighbor?: CellBorders,
  aboveNeighbor?: CellBorders,
): CellBorders {
  const isFirstRow = rowIdx === 0;
  const isLastRow = rowIdx === rowCount - 1;
  const isFirstCol = colStart === 0;
  const isLastCol = colEnd === colCount - 1;
  const out: { -readonly [K in keyof CellBorders]: CellBorders[K] } = {};
  // Each cell edge takes the cell's own specified border, inheriting the table
  // default (insideH/insideV for internal edges, top/left/bottom/right for the
  // table's outer edges) when the cell omits it.
  const ownTop = cellBorders?.top ?? (isFirstRow ? tableBorders?.top : tableBorders?.insideH);
  const ownBottom =
    cellBorders?.bottom ?? (isLastRow ? tableBorders?.bottom : tableBorders?.insideH);
  const ownLeft = cellBorders?.left ?? (isFirstCol ? tableBorders?.left : tableBorders?.insideV);
  const ownRight = cellBorders?.right ?? (isLastCol ? tableBorders?.right : tableBorders?.insideV);
  // §17.4 border-conflict resolution: an INTERNAL edge is shared by two cells and
  // drawn once (we draw it on the lower/right cell's top/left side). The two
  // facing borders compete — the heavier wins (§17.4.* table border conflict).
  // The neighbour's facing border also inherits the table's insideH/insideV.
  const aboveBottom = aboveNeighbor ? (aboveNeighbor.bottom ?? tableBorders?.insideH) : undefined;
  const leftRight = leftNeighbor ? (leftNeighbor.right ?? tableBorders?.insideV) : undefined;
  const top = isFirstRow ? ownTop : heavierBorder(ownTop, aboveBottom);
  const left = isFirstCol ? ownLeft : heavierBorder(ownLeft, leftRight);
  const bottom = ownBottom;
  const right = ownRight;
  const suppressTop = mergeRole === 'middle' || mergeRole === 'end';
  const suppressBottom = mergeRole === 'start' || mergeRole === 'middle';
  if (top && !suppressTop) out.top = top;
  if (bottom && !suppressBottom) out.bottom = bottom;
  if (left) out.left = left;
  if (right) out.right = right;
  return out;
}

export interface LaidOutPage {
  readonly commands: Array<DrawCommand>;
  readonly width: number;
  readonly height: number;
}

// Map a paragraph's resolved properties to a tagged-PDF structure type
// (ISO 32000-1 §14.8.4). A heading comes from the resolved outline level
// (w:outlineLvl, 0–8 → H1–H9 clamped to the H1–H6 range), or — when a heading
// style carried no outline level — from a "Heading N" / "Title" style id.
// Everything else is body text (P).
function paragraphStructType(resolved: ResolvedParagraphProperties): StructType {
  const lvl = resolved.outlineLevel;
  if (lvl !== undefined && lvl >= 0 && lvl <= 8) {
    return `H${Math.min(lvl, 5) + 1}` as StructType;
  }
  return headingFromStyleId(resolved.styleId) ?? 'P';
}

function headingFromStyleId(styleId: string | undefined): StructType | null {
  if (!styleId) return null;
  const m = /^Heading\s*([1-9])$/i.exec(styleId);
  if (m) return `H${Math.min(Number(m[1]), 6)}` as StructType;
  if (/^(Title|Subtitle)$/i.test(styleId)) return 'H1';
  return null;
}

// The dominant natural language of a paragraph's text tokens (weighted by
// character count), or undefined when no run carries a w:lang. Drives the
// tagged-PDF per-element /Lang.
function dominantParagraphLang(lines: ReadonlyArray<Line>): string | undefined {
  const counts = new Map<string, number>();
  for (const line of lines) {
    for (const tok of line.tokens) {
      if (tok.kind !== 'text') continue;
      const l = tok.resolvedRun.lang;
      if (l) counts.set(l, (counts.get(l) ?? 0) + tok.text.length);
    }
  }
  let best: string | undefined;
  let bestN = 0;
  for (const [l, n] of counts) {
    if (n > bestN) {
      best = l;
      bestN = n;
    }
  }
  return best;
}

// Create a Figure structure node under the document root with non-empty /Alt.
// PDF/A-1a requires every Figure to carry alternate text; fall back to a generic
// label when the drawing has no docPr description. Returns the node id.
function createFigure(
  builder: StructTreeBuilder,
  alt: string | undefined,
  fallback: string,
): number {
  const node = builder.create('Figure', builder.root);
  node.alt = alt && alt.trim().length > 0 ? alt : fallback;
  return node.id;
}

// One open list level on the nesting stack: its L element and the most recent
// item's LBody (where a deeper nested list attaches).
interface ListFrame {
  readonly level: number;
  readonly numId: string;
  readonly listNode: StructNode;
  lbody: StructNode | null;
}

// Resolve the structure node (a P) for a list-item paragraph, growing/shrinking
// the open-list stack by nesting level (w:ilvl). Each item is L → LI → LBody →
// P; a deeper level opens a nested L inside the parent item's LBody. The list
// marker stays inside the P (no separate Lbl in this milestone). Returns the P
// node id to stamp onto the paragraph's lines.
function listItemParagraphNode(
  builder: StructTreeBuilder,
  stack: Array<ListFrame>,
  list: { numId: string; level: number },
): number {
  const lvl = list.level;
  // Close any deeper levels, or a same-level list with a different numId.
  while (stack.length > 0) {
    const t = stack[stack.length - 1]!;
    if (t.level > lvl || (t.level === lvl && t.numId !== list.numId)) stack.pop();
    else break;
  }
  let frame = stack[stack.length - 1];
  if (!frame || frame.level < lvl) {
    // Open a new L: nested in the parent item's LBody, or at the document root.
    const parent = frame?.lbody ?? builder.root;
    const listNode = builder.create('L', parent);
    frame = { level: lvl, numId: list.numId, listNode, lbody: null };
    stack.push(frame);
  }
  const li = builder.create('LI', frame.listNode);
  const lbody = builder.create('LBody', li);
  frame.lbody = lbody;
  return builder.create('P', lbody).id;
}

function paginateSections(
  blocks: ReadonlyArray<LaidOutBlock>,
  sectionCtxs: ReadonlyArray<SectionRenderCtx>,
  builder?: StructTreeBuilder,
  defaultLang = 'en-US',
): Array<LaidOutPage> {
  if (sectionCtxs.length === 0) return [];
  const pages: Array<LaidOutPage> = [];

  let ctx = sectionCtxs[0]!;
  let secIdx = 0;
  let pageInSection = 0;
  let globalPageIdx = 0;
  let current: Array<DrawCommand> = [];
  let pendingPageBreak = false;
  let cursorY = ctx.pageHeight - ctx.marginTop;
  // Tagged PDF: the stack of open list levels (for L/LI/LBody nesting). Cleared
  // whenever a non-list-item block interrupts the run of list paragraphs.
  const listStack: Array<ListFrame> = [];

  const flushPage = (force = false) => {
    if (current.length === 0 && !force) return;
    const band = bandForPage(pageInSection, globalPageIdx, ctx.titlePg, ctx.evenAndOddHeaders);
    const header = pickBand(ctx.headerSet, band);
    const footer = pickBand(ctx.footerSet, band);
    pages.push({
      commands: [...header, ...current, ...footer],
      width: ctx.pageWidth,
      height: ctx.pageHeight,
    });
    current = [];
    pageInSection++;
    globalPageIdx++;
    cursorY = ctx.pageHeight - ctx.marginTop;
  };

  for (let blockIdx = 0; blockIdx < blocks.length; blockIdx++) {
    // Advance to the section that owns this block. A section boundary forces
    // a page break before the next section's first block.
    while (secIdx < sectionCtxs.length - 1 && blockIdx >= ctx.endIndex) {
      flushPage();
      secIdx++;
      ctx = sectionCtxs[secIdx]!;
      pageInSection = 0;
      cursorY = ctx.pageHeight - ctx.marginTop;
    }

    const block = blocks[blockIdx]!;
    // A forced page break (w:br w:type="page") carried by the previous block:
    // start this block on a fresh page.
    if (pendingPageBreak) {
      pendingPageBreak = false;
      if (current.length > 0) flushPage();
    }
    // A non-list-item block ends any open list run (tagged PDF).
    if (builder && !(block.kind === 'paragraph' && block.list)) listStack.length = 0;
    if (block.kind === 'paragraph') {
      if (block.resolved.pageBreakBefore && current.length > 0) flushPage();
      cursorY -= block.spacingBeforePt;
      // Tagged PDF: a plain paragraph → one P (or heading) element; a list item
      // → an L/LI/LBody/P built on the nesting stack. Its lines all reference
      // the resulting leaf by MCID.
      let structId: number | undefined;
      if (builder) {
        structId = block.list
          ? listItemParagraphNode(builder, listStack, block.list)
          : builder.create(paragraphStructType(block.resolved), builder.root).id;
        // §14.9.2 per-element /Lang: tag a paragraph whose dominant run language
        // differs from the document default so AT switches pronunciation.
        const lang = dominantParagraphLang(block.lines);
        if (lang && lang !== defaultLang) builder.node(structId).lang = lang;
      }
      for (const line of block.lines) {
        const h = computeLineHeight(line, block.resolved);
        if (cursorY - h < ctx.marginBottom && current.length > 0) flushPage();
        cursorY -= h;
        const indentLeft =
          block.resolved.indentLeft + (line.firstLine ? block.resolved.indentFirstLine : 0);
        const offset = alignmentOffset(
          block.resolved.alignment,
          line.contentWidthPt,
          line.availableWidthPt,
        );
        current.push({
          type: 'line',
          line,
          originX: ctx.marginLeft + indentLeft + offset,
          baselineY: cursorY + lineDescent(line),
          ...(structId !== undefined ? { structId } : {}),
        });
      }
      cursorY -= block.spacingAfterPt;
      if (block.pageBreakAfter) pendingPageBreak = true;
    } else if (block.kind === 'image') {
      const figId = builder ? createFigure(builder, block.altText, 'Image') : undefined;
      cursorY -= block.spacingBeforePt;
      if (cursorY - block.heightPt < ctx.marginBottom && current.length > 0) flushPage();
      cursorY -= block.heightPt;
      current.push({
        type: 'image',
        x: ctx.marginLeft,
        y: cursorY,
        width: block.widthPt,
        height: block.heightPt,
        imageResourceName: block.resourceName,
        ...(figId !== undefined ? { structId: figId } : {}),
      });
      cursorY -= block.spacingAfterPt;
    } else if (block.kind === 'shape') {
      // Shapes are atomic — never split across pages.
      const figId = builder ? createFigure(builder, block.altText, 'Shape') : undefined;
      cursorY -= block.spacingBeforePt;
      if (cursorY - block.heightPt < ctx.marginBottom && current.length > 0) flushPage();
      cursorY -= block.heightPt;
      const offset = alignmentOffset(block.resolvedAlignment, block.widthPt, ctx.contentWidth);
      const x = ctx.marginLeft + offset;
      const transform = buildShapeTransform(
        x,
        cursorY,
        block.widthPt,
        block.heightPt,
        block.rotation60k,
        block.flipH,
        block.flipV,
      );
      current.push({
        type: 'shape',
        shape: {
          paths: block.paths,
          ...(block.fillColorHex ? { fillColorHex: block.fillColorHex } : {}),
          ...(block.stroke ? { stroke: block.stroke } : {}),
          transform,
        },
        ...(figId !== undefined ? { structId: figId } : {}),
      });
      // Shape text: laid out axis-aligned, anchored vertically within the inset
      // rect, emitted as ordinary line commands so it rides the text pass on
      // top of the fill. (Rotated text boxes keep upright text in M5.)
      if (block.textLines.length > 0) {
        const shapeBottom = cursorY;
        const shapeTop = cursorY + block.heightPt;
        const innerWidth = Math.max(1, block.widthPt - block.insetLeftPt - block.insetRightPt);
        let textY: number;
        if (block.anchor === 'b') {
          textY = shapeBottom + block.insetBottomPt + block.textHeightPt;
        } else if (block.anchor === 'ctr') {
          textY = shapeBottom + (block.heightPt + block.textHeightPt) / 2;
        } else {
          textY = shapeTop - block.insetTopPt;
        }
        for (const line of block.textLines) {
          const h = computeLineHeight(line, line.resolved);
          textY -= h;
          const lineOffset = alignmentOffset(
            line.resolved.alignment,
            line.contentWidthPt,
            innerWidth,
          );
          current.push({
            type: 'line',
            line,
            originX: x + block.insetLeftPt + lineOffset,
            baselineY: textY + lineDescent(line),
            ...(figId !== undefined ? { structId: figId } : {}),
          });
        }
      }
      cursorY -= block.spacingAfterPt;
    } else if (block.kind === 'chart') {
      // Charts are atomic. Their primitives are in a local y-up frame; place by
      // translating to the chart box's page-space bottom-left (x, y). The whole
      // chart is one Figure (alt = its title); its shapes + labels carry that id.
      const figId = builder ? createFigure(builder, block.altText, 'Chart') : undefined;
      const fig = figId !== undefined ? { structId: figId } : {};
      cursorY -= block.spacingBeforePt;
      if (cursorY - block.heightPt < ctx.marginBottom && current.length > 0) flushPage();
      cursorY -= block.heightPt;
      const offset = alignmentOffset(block.resolvedAlignment, block.widthPt, ctx.contentWidth);
      const x = ctx.marginLeft + offset;
      const y = cursorY;
      for (const s of block.layout.shapes) {
        current.push({
          type: 'shape',
          shape: {
            paths: s.paths,
            ...(s.fillColorHex ? { fillColorHex: s.fillColorHex } : {}),
            ...(s.stroke ? { stroke: s.stroke } : {}),
            transform: [1, 0, 0, 1, x, y],
          },
          ...fig,
        });
      }
      for (const t of block.layout.texts) {
        current.push({ type: 'line', line: t.line, originX: x + t.x, baselineY: y + t.y, ...fig });
      }
      cursorY -= block.spacingAfterPt;
    } else {
      const colCount = block.rows.reduce(
        (max, r) =>
          Math.max(
            max,
            r.cells.reduce((s, c) => Math.max(s, c.colStart + c.colSpan), 0),
          ),
        0,
      );
      // Leading header rows (w:tblHeader / _xlnm.Print_Titles) repeat at the top
      // of every continuation page. Only the maximal leading prefix repeats —
      // a header flagged mid-table is not a repeating title.
      const headerRows: Array<RowLayout> = [];
      for (const r of block.rows) {
        if (r.isHeader) headerRows.push(r);
        else break;
      }
      const headerHeightPt = headerRows.reduce((s, r) => s + r.heightPt, 0);

      // Tagged PDF: Table → one TR per row → one TD per logical cell (skipping
      // vMerge continuation cells, which are covered by the origin's span) → P
      // holding the cell's content. The same TD/P node is reused across row
      // chunks (page splits), so its MCRs accumulate like a split paragraph.
      const tableNode = builder ? builder.create('Table', builder.root) : undefined;
      const tableX = ctx.marginLeft + block.xOffsetPt;
      for (let ri = 0; ri < block.rows.length; ri++) {
        const row = block.rows[ri]!;
        const isLeadingHeader = ri < headerRows.length;
        let cellStructIds: Array<number | undefined> | undefined;
        if (builder && tableNode) {
          const b = builder;
          const tr = b.create('TR', tableNode);
          // Header-row cells are TH with /Scope /Column (§14.8.5.2) so AT binds
          // them to the data cells beneath; all other cells are TD. Spanning
          // cells carry /ColSpan (gridSpan) and /RowSpan (vertical merge).
          cellStructIds = row.cells.map((cell) => {
            if (cell.mergeRole === 'middle' || cell.mergeRole === 'end') return undefined;
            const cellNode = b.create(row.isHeader ? 'TH' : 'TD', tr);
            if (row.isHeader) cellNode.scope = 'Column';
            if (cell.colSpan > 1) cellNode.colSpan = cell.colSpan;
            const rowSpan = tableCellRowSpan(block.rows, ri, cell.colStart);
            if (rowSpan > 1) cellNode.rowSpan = rowSpan;
            return b.create('P', cellNode).id;
          });
        }
        const chunks =
          row.heightPt > ctx.pageContentHeight
            ? splitRowIntoChunks(row, ctx.pageContentHeight)
            : [row];

        for (let ci = 0; ci < chunks.length; ci++) {
          const chunk = chunks[ci]!;
          // A manual <rowBreaks> break (first chunk only) forces a new page even
          // when the row would fit; an overflow break starts one when it won't.
          const forcedBreak = ci === 0 && row.breakBefore && !isLeadingHeader && current.length > 0;
          const overflow = cursorY - chunk.heightPt < ctx.marginBottom && current.length > 0;
          if (forcedBreak || overflow) {
            flushPage();
            // Re-emit the header rows on the fresh page (visual repetition →
            // artifacts, no structIds), but only when the breaking row is not
            // itself a header and the row still fits beneath the repeated band.
            if (
              !isLeadingHeader &&
              headerRows.length > 0 &&
              cursorY - headerHeightPt - chunk.heightPt >= ctx.marginBottom
            ) {
              for (const hr of headerRows) {
                emitRowChunk(current, hr, tableX, cursorY, colCount, undefined);
                cursorY -= hr.heightPt;
              }
            }
          }
          emitRowChunk(current, chunk, tableX, cursorY, colCount, cellStructIds);
          cursorY -= chunk.heightPt;
        }
      }
    }
  }

  // Trailing content on the last in-progress page.
  flushPage();
  // A body that produced no flushable content (e.g. text lives only in the
  // header/footer bands — a header/footer-only document) must still emit one
  // page so those bands render, instead of falling back to a blank page.
  if (pages.length === 0) flushPage(true);
  return pages;
}

// §14.8.5.2 /RowSpan — how many rows a vertical-merge origin spans. Walk down
// from the origin row at the same column, counting vMerge continuation cells
// (mergeRole middle/end) until the group ends. A standalone cell spans 1.
function tableCellRowSpan(
  rows: ReadonlyArray<RowLayout>,
  startRowIdx: number,
  colStart: number,
): number {
  let span = 1;
  for (let r = startRowIdx + 1; r < rows.length; r++) {
    const cont = rows[r]!.cells.find(
      (c) => c.colStart === colStart && (c.mergeRole === 'middle' || c.mergeRole === 'end'),
    );
    if (!cont) break;
    span++;
    if (cont.mergeRole === 'end') break;
  }
  return span;
}

function emitRowChunk(
  out: Array<DrawCommand>,
  row: RowLayout,
  marginLeft: number,
  cursorY: number,
  colCount: number,
  cellStructIds?: ReadonlyArray<number | undefined>,
): void {
  const rowTop = cursorY;
  const rowBottom = cursorY - row.heightPt;
  for (let i = 0; i < row.cells.length; i++) {
    const cell = row.cells[i]!;
    const cellX = marginLeft + (row.columnXOffsets[i] ?? 0);
    const structId = cellStructIds?.[i];
    if (cell.shadingColorHex && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      out.push({
        type: 'fill',
        x: cellX,
        y: rowBottom,
        width: cell.widthPt,
        height: row.heightPt,
        fillColorHex: cell.shadingColorHex,
      });
    }
    emitCellBorders(out, cell, cellX, rowBottom, row.heightPt, row.rowIdx, row.rowCount, colCount);
    if (cell.mergeRole === 'middle' || cell.mergeRole === 'end') continue;
    let textY = rowTop - cell.padTopPt;
    for (const line of cell.lines) {
      const h = computeLineHeight(line, line.resolved);
      textY -= h;
      const offset = alignmentOffset(
        line.resolved.alignment,
        line.contentWidthPt,
        cell.widthPt - cell.padLeftPt - cell.padRightPt,
      );
      out.push({
        type: 'line',
        line,
        originX: cellX + cell.padLeftPt + offset,
        baselineY: textY + lineDescent(line),
        ...(structId !== undefined ? { structId } : {}),
      });
    }
    // Nested tables render below the cell's paragraph lines, inset to the
    // cell's content box. Each nested row reuses emitRowChunk; when tagged, its
    // content is marked under the parent cell's structId.
    if (cell.nestedTables) {
      const nestedX = cellX + cell.padLeftPt;
      for (const nt of cell.nestedTables) {
        for (const nrow of nt.rows) {
          const nestedIds = structId !== undefined ? nrow.cells.map(() => structId) : undefined;
          emitRowChunk(out, nrow, nestedX, textY, nt.colCount, nestedIds);
          textY -= nrow.heightPt;
        }
      }
    }
  }
}

// Slice an oversized row into vertical chunks, each ≤ capacity. Each chunk is
// shaped like a normal RowLayout so the emit code doesn't have to special-case
// it. Convention:
//   - first chunk keeps padTop and the top border
//   - last chunk keeps padBottom and the bottom border
//   - middle chunks have no top/bottom padding or borders
// Left/right borders and shading are kept on every chunk.
function splitRowIntoChunks(row: RowLayout, capacity: number): Array<RowLayout> {
  // A cell containing a nested table is not split across pages — the nested
  // layout would be duplicated into every chunk. Keep such a row whole.
  if (row.cells.some((c) => c.nestedTables && c.nestedTables.length > 0)) return [row];
  type Queue = { remaining: Array<Line>; template: CellLayout };
  const queues: Array<Queue> = row.cells.map((c) => ({ remaining: [...c.lines], template: c }));

  const anyHasLines = () =>
    queues.some((q) =>
      q.template.mergeRole === 'middle' || q.template.mergeRole === 'end'
        ? false
        : q.remaining.length > 0,
    );

  const out: Array<RowLayout> = [];
  let isFirst = true;
  // Cap iterations as a safety net — a malformed input shouldn't hang the renderer.
  let safety = 1000;
  while (anyHasLines() && safety-- > 0) {
    const chunkCells: Array<CellLayout> = [];
    let chunkHeight = 0;

    for (let i = 0; i < queues.length; i++) {
      const q = queues[i]!;
      const tpl = q.template;
      if (tpl.mergeRole === 'middle' || tpl.mergeRole === 'end') {
        chunkCells.push(tpl);
        continue;
      }
      const padTop = isFirst ? tpl.padTopPt : 0;
      const capacityForLines = Math.max(0, capacity - padTop);
      const taken: Array<Line> = [];
      let takenHeight = 0;
      while (q.remaining.length > 0) {
        const next = q.remaining[0]!;
        const lh = computeLineHeight(next, next.resolved);
        // Always take at least one line per chunk to guarantee forward progress
        // even when a single line is taller than capacity.
        if (taken.length > 0 && takenHeight + lh > capacityForLines) break;
        taken.push(next);
        takenHeight += lh;
        q.remaining.shift();
      }
      const cellHeight = padTop + takenHeight;
      const chunkCell: CellLayout = {
        ...tpl,
        padTopPt: padTop,
        padBottomPt: 0,
        lines: taken,
        contentHeightPt: takenHeight,
        totalHeightPt: cellHeight,
      };
      chunkCells.push(chunkCell);
      if (cellHeight > chunkHeight) chunkHeight = cellHeight;
    }

    const isLast = !anyHasLines();
    if (isLast) {
      // Apply padBottom to all real cells; recompute chunkHeight.
      for (let i = 0; i < chunkCells.length; i++) {
        const tpl = queues[i]!.template;
        if (tpl.mergeRole === 'middle' || tpl.mergeRole === 'end') continue;
        const cell = chunkCells[i]!;
        const total = cell.padTopPt + cell.contentHeightPt + tpl.padBottomPt;
        chunkCells[i] = {
          ...cell,
          padBottomPt: tpl.padBottomPt,
          totalHeightPt: total,
        };
        if (total > chunkHeight) chunkHeight = total;
      }
    }

    // Borders: keep top only on first chunk, bottom only on last; left/right always.
    for (let i = 0; i < chunkCells.length; i++) {
      const cell = chunkCells[i]!;
      const src = cell.borders;
      const borders: CellBorders = {
        ...(isFirst && src.top ? { top: src.top } : {}),
        ...(isLast && src.bottom ? { bottom: src.bottom } : {}),
        ...(src.left ? { left: src.left } : {}),
        ...(src.right ? { right: src.right } : {}),
      };
      chunkCells[i] = { ...cell, borders };
    }

    out.push({
      heightPt: chunkHeight,
      cells: chunkCells,
      columnXOffsets: row.columnXOffsets,
      rowIdx: row.rowIdx,
      rowCount: row.rowCount,
    });
    isFirst = false;
  }

  return out.length > 0 ? out : [row];
}

// Each shared edge between adjacent cells is the same physical line, so we
// render it exactly once. Convention: every cell paints its top + left; the
// last row paints its bottom and the last spanned column paints its right.
function emitCellBorders(
  out: Array<DrawCommand>,
  cell: CellLayout,
  cellX: number,
  cellY: number,
  rowHeight: number,
  rowIdx: number,
  rowCount: number,
  colCount: number,
): void {
  const pushSide = (
    side: 'top' | 'right' | 'bottom' | 'left',
    border: CellBorders[keyof CellBorders],
  ) => {
    if (!border || border.style === 'none') return;
    const sz = border.width ?? DEFAULT_BORDER_SIZE_EIGHTH * EIGHTH_PT;
    out.push({
      type: 'border',
      side,
      x: cellX,
      y: cellY,
      width: cell.widthPt,
      height: rowHeight,
      borderSizePt: sz,
      borderColorHex: border.colorHex ?? '000000',
    });
  };
  pushSide('top', cell.borders.top);
  pushSide('left', cell.borders.left);
  if (rowIdx === rowCount - 1) pushSide('bottom', cell.borders.bottom);
  if (cell.colStart + cell.colSpan - 1 === colCount - 1) pushSide('right', cell.borders.right);
}

// Per-page tagging state threaded into emitPageContent. `next` is the running
// MCID counter (reset per page); `assigned` records whether any tagged marked
// content was emitted (so the page gets /StructParents); `record` ties an
// assigned MCID back to its structure node.
export interface PageTagging {
  next: number;
  assigned: boolean;
  record: (structId: number, mcid: number) => void;
  // The marked-content tag for a structure node — its structure type, so the
  // BDC tag matches the StructElem /S (§14.7.2: a heading is /H1, a cell's
  // paragraph /P, …), not a hardcoded /P.
  tagFor: (structId: number) => string;
}
