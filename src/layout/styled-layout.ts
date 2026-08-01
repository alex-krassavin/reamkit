// The layout engine (ir-design §7, moved to src/layout/ at stage 6.4):
// FlowDoc body → PageDoc (positioned PageItems per page) + the PDF companion.
//
// Pipeline:
//   1. Pre-scan body recursively → collect every (variant, gid) actually used,
//      and probe/decode every referenced image (the resource collections).
//   2. For each body block, lay out:
//        paragraph → list of styled lines (Knuth-Plass + hyphenation + BiDi)
//        table     → column widths + per-cell laid-out content + row heights
//   3. Paginate: stack laid-out blocks; advance cursorY; break to new page
//      when a block (or table row) does not fit. Emit PageItems in the
//      frozen top-left page frame (see layout/page-doc).
//
// Dependency direction: layout → core only — with three deliberate,
// type-shaped exceptions into pdf/ that the options/companion carry:
// StructTreeBuilder (the tagged logical structure the PDF emitter consumes),
// and the AttachedFile/SignaturePlaceholder option types. Moving those value
// objects out of pdf/ is mechanical follow-up work, not part of the freeze.

import type {
  BodyElement,
  Border,
  BorderStyle,
  CellBorders,
  CellDataBar,
  CellIcon,
  CellSparkline,
  Chart,
  ChartBlock,
  Comment,
  DocumentInfo,
  FloatAnchor,
  HeaderFooterReference,
  HeaderFooterType,
  ImageBlock,
  ImageCrop,
  MathNode,
  Numbering,
  Paragraph,
  Run,
  Section,
  SectionColumns,
  SectionProperties,
  ShapeBlock,
  ShapeShadow,
  StyleSheet,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';
import type { FontRegistry, ParsedTtf } from '@/core/font';
import type { FamilyKey } from '@/core/fonts';
import type { Hyphenator } from '@/core/hyphenation';
import type { PreparedImage } from '@/core/images';
import type { Item } from '@/core/line-breaker';
import type { Pt, ResourceId } from '@/core/ir';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { ShapeGradient, StrokeStyle, VectorPath } from '@/core/vector';
import type {
  ChartLabel,
  ChartPolygon,
  ChartPolyline,
  ChartRect,
  ChartWedge,
} from '@/core/drawingml/chart-geometry';
import type { MathDrawItem, MathVariant, MeasureMath } from '@/layout/math-layout';
import type {
  FontResource,
  ImageResource,
  LaidOutDocument,
  LaidOutPage,
  Line,
  PageItem,
  ResolvedMathItem,
  TextToken,
  Token,
} from '@/layout/page-doc';
// Deliberate layout→pdf residue (see the header note).
import type { AttachedFile } from '@/pdf/embedded-file';
import type { SignaturePlaceholder } from '@/pdf/signature';
import type { PdfEncryptOptions } from '@/pdf/encryption';
import type { StructNode, StructType } from '@/pdf/struct-tree';

import { ResourceStore, halfPtToPt, pt } from '@/core/ir';
import { createFontMeasure, shapeText } from '@/core/font';
import { resolveFamilyKey } from '@/core/fonts';
import { prepareImage } from '@/core/images';
import { analyzeString, hasBidiCharacters, segmentLevels } from '@/core/bidi';
import {
  FORCED_BREAK,
  breakLines,
  cjkBreakBetween,
  greedyBreakLines,
  splitCjkSegment,
} from '@/core/line-breaker';
import { applyNumbering, applyNumberingToHeadersFooters } from '@/core/numbering';
import {
  DEFAULT_RESOLVED_PARAGRAPH,
  DEFAULT_RESOLVED_RUN,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';
import { PathBuilder, flipTransform } from '@/core/vector';
import { buildSparkline } from '@/core/drawingml/sparkline-geometry';
import { arcPoint, arcToBeziers } from '@/core/arc-to-bezier';
import { buildChartScene, legendSeriesName } from '@/core/drawingml/chart-geometry';
import { layoutMath, mathGlyphSegments, variantStyle } from '@/layout/math-layout';
import { rectPath } from '@/core/drawingml/preset-geometry';
import {
  DEFAULT_INSET_LR_PT,
  DEFAULT_INSET_TB_PT,
  buildShapePaths,
  buildShapeTransform,
  buildStroke,
  gradientToSolid,
} from '@/core/drawingml/shape-render';
import { StructTreeBuilder } from '@/pdf/struct-tree';

/**
 * PDF/A conformance string: part 1 (ISO 19005-1, PDF 1.4) / 2 (ISO 19005-2) / 3
 * (ISO 19005-3, both PDF 1.7); conformance level a (tagged) / b (visual) /
 * u (Unicode — only 2/3).
 */
export type PdfALevel =
  | 'PDF/A-1b'
  | 'PDF/A-1a'
  | 'PDF/A-2b'
  | 'PDF/A-2u'
  | 'PDF/A-2a'
  | 'PDF/A-3b'
  | 'PDF/A-3u'
  | 'PDF/A-3a';

/** A {@link PdfALevel} decomposed into the part, level, and the apparatus they imply. */
export interface PdfAProfile {
  readonly part: 1 | 2 | 3;
  readonly level: 'a' | 'b' | 'u';
  /** Whether the level mandates a tagged PDF (level `a`). */
  readonly tagged: boolean; // level a
  /** PDF version the part pins: part 1 → 1.4, else 1.7. */
  readonly version: '1.4' | '1.7'; // part 1 → 1.4, else 1.7
}

// Decompose a "PDF/A-<part><level>" string. The 7th/8th chars are the part and
// level (e.g. "PDF/A-2u" → part 2, level u).
function parsePdfAProfile(pdfA: PdfALevel): PdfAProfile {
  const part = Number(pdfA.charAt(6)) as 1 | 2 | 3;
  const level = pdfA.charAt(7) as 'a' | 'b' | 'u';
  return { part, level, tagged: level === 'a', version: part === 1 ? '1.4' : '1.7' };
}

/**
 * E-PARITY: renderer-compatibility profile for the line-height model.
 *   `'ream'` (default) — Ream's flat 1.2× leading; byte-identical to before.
 *   `'word'` — leading from the font's OS/2 usWin metrics (the GDI cell box).
 *   `'libreoffice'` — leading from hhea (or OS/2 typo when USE_TYPO_METRICS is set).
 *
 * Opt-in: a profile emulates that renderer's vertical rhythm for closer visual
 * parity. It never changes default (`'ream'`) output.
 */
export type LayoutProfile = 'ream' | 'word' | 'libreoffice';

/**
 * The full option set the layout engine and PDF emitter consume: the resolved
 * font registry, the style/numbering tables, the section model and page
 * geometry, the after-body apparatus (footnotes, endnotes, comments), and every
 * PDF-output toggle (PDF/A, tagged, PDF/UA, encryption, attachments,
 * signatures). Only `registry` and `styles` are required; everything else falls
 * back to a sensible default.
 */
export interface StyledRenderOptions {
  /** The resolved font set every run draws with (the guaranteed fallback registry). */
  readonly registry: FontRegistry;
  /** Renderer-compatibility profile for the line-height model (default `'ream'`). */
  readonly layoutProfile?: LayoutProfile;
  /**
   * Per-run font resolution: when supplied, each text run picks the registry of
   * its declared family (sans→arimo / serif→tinos / mono→cousine via the run's
   * `w:ascii`) instead of always using `registry`. Absent ⇒ single-family (every
   * run uses `registry`), byte-identical to before. `registry` remains the
   * guaranteed fallback for math/chart/default glyphs and any missing family.
   */
  readonly registriesByFamily?: ReadonlyMap<FamilyKey, FontRegistry>;
  /**
   * The document's OWN embedded fonts (`word/fonts/*.odttf`, de-obfuscated),
   * keyed by normalized font name. A run whose `w:ascii` matches one renders with
   * the real font — glyph-exact, no substitution. Highest priority.
   */
  readonly embeddedFonts?: ReadonlyMap<string, FontRegistry>;
  /** §18.8 / ECMA-376 style table the cascade resolves against. */
  readonly styles: StyleSheet;
  /** §17.9 numbering definitions; applied to list paragraphs before layout. */
  readonly numbering?: Numbering;
  /**
   * Single-section legacy entry-point. If `sections` is set it takes precedence
   * and `section` is ignored.
   */
  readonly section?: SectionProperties;
  /**
   * ECMA-376 §17.6 — ordered list of sections. Each section's `endIndex` is the
   * exclusive bound into the body array (section N covers
   * `body[sections[N-1].endIndex..sections[N].endIndex)`).
   */
  readonly sections?: ReadonlyArray<Section>;
  /** Header/footer body content keyed by relationship id. */
  readonly headersFooters?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  /**
   * §17.11 notes content by id. Footnotes render in a reserved band at the
   * bottom of the referencing page; endnotes flow after the body.
   */
  readonly footnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  readonly endnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  /**
   * §17.13.4 review comments by id; rendered as superscript markers in text and
   * a list after the body (after endnotes), each with its author/date.
   */
  readonly comments?: ReadonlyMap<string, Comment>;
  /**
   * CM2b — also emit each comment as a native PDF `/Text` (sticky-note)
   * annotation at its marker. Opt-in and interactive-only: suppressed under
   * PDF/A and tagged output (where it would need annotation/appearance
   * conformance), since the clickable marker + Comments section already carry
   * the content there.
   */
  readonly commentAnnotations?: boolean;
  /** Content-addressed binary store; image nodes reference it by {@link ResourceId}. */
  readonly resources?: ResourceStore;
  /**
   * Parsed charts keyed by relationship id (`ChartBlock.chartRelId`). Supplied by
   * the converter, which resolves the chart parts from the package.
   */
  readonly charts?: ReadonlyMap<string, Chart>;
  readonly pageWidth?: number;
  readonly pageHeight?: number;
  readonly marginLeft?: number;
  readonly marginRight?: number;
  readonly marginTop?: number;
  readonly marginBottom?: number;
  /**
   * Optional Liang hyphenator. When set, each word token is split at allowed
   * hyphenation positions and offered to Knuth-Plass as potential break points
   * (with a small disincentive). Improves justified paragraph rags.
   */
  readonly hyphenator?: Hyphenator;
  /**
   * Optional `/Info` dictionary metadata (ISO 32000-1 §14.3.3). Unset fields are
   * omitted; if any field is set a PDF `/Info` entry is emitted.
   */
  readonly info?: DocumentInfo;
  /**
   * When set, emit a PDF/A-conformant file: an OutputIntent with an embedded
   * sRGB ICC profile, document XMP `/Metadata` (the pdfaid identifier), `/ID`,
   * and subset-tagged fonts with a `/CIDSet`. The profile picks the rest:
   *   part 1 → PDF 1.4 + flattened image alpha (no transparency);
   *   part 2/3 → PDF 1.7 + preserved transparency (image `/SMask` + page group);
   *   part 3 → may carry embedded associated files (see `attachments`);
   *   level a → tagged (logical structure); b → visual; u → b + Unicode mapping.
   */
  readonly pdfA?: PdfALevel;
  /**
   * Emit a tagged PDF (ISO 32000-1 §14.8): a `/StructTreeRoot` describing reading
   * order, marked content (BDC/EMC + MCID) on body text, and `/Artifact` marking
   * of page decoration. Implied by `pdfA: 'PDF/A-1a'`. Independent of PDF/A
   * otherwise (a plain tagged PDF is useful on its own).
   */
  readonly tagged?: boolean;
  /**
   * PDF/UA-1 (ISO 14289-1): implies tagged; the XMP carries `pdfuaid:part=1` and
   * the document always gets a title (AT announces it). Combines freely with
   * `pdfA` level-a profiles.
   */
  readonly pdfUA?: boolean;
  /**
   * Document natural language (BCP 47, e.g. `"en-US"`, `"ru-RU"`) for the
   * tagged-PDF catalog `/Lang` (§14.9.2). Defaults to `"en-US"`. The docx
   * converter fills this from the document's default `w:lang`.
   */
  readonly language?: string;
  /**
   * §7.6 PDF encryption (AES-256, R6). Only honoured on the ASYNC conversion
   * path (WebCrypto); mutually exclusive with `pdfA` (ISO 19005 forbids
   * `/Encrypt`) and with signatures (v1).
   */
  readonly encrypt?: PdfEncryptOptions;
  /**
   * Files to embed as associated files (catalog `/AF` + `/Names`
   * `/EmbeddedFiles`). Only emitted for plain PDF and PDF/A-3 (PDF/A-1/2 forbid
   * arbitrary embedded files); ignored for PDF/A-1/2. The docx/xlsx converters
   * can embed the source document automatically via `embedSource`.
   */
  readonly attachments?: ReadonlyArray<AttachedFile>;
  /**
   * Emit an (invisible) signature field + signature dictionary with placeholder
   * `/ByteRange` and `/Contents` (ISO 32000 §12.8). The result is an UNSIGNED
   * PDF; pass it to `signPdf()` to fill the placeholder with a real PKCS#7
   * signature.
   */
  readonly signaturePlaceholder?: SignaturePlaceholder;
}

// Re-exported from the document model (moved there so FlowDoc can carry it).
export type { DocumentInfo } from '@/core/document-model';

/** A4 page width in points (the page-geometry fallback). */
export const A4_WIDTH = 595;
/** A4 page height in points (the page-geometry fallback). */
export const A4_HEIGHT = 842;
/**
 * How far a space may be squeezed below its natural width — the shrink the
 * line breaker offers when it weighs a line, and therefore the shrink the
 * emitter owes it back when it draws one.
 */
export const GLUE_SHRINK_RATIO = 0.3;
const TWIP_TO_PT = 1 / 20;
const EIGHTH_PT = 1 / 8;
const DEFAULT_CELL_PADDING_TWIPS = 108;
const DEFAULT_BORDER_SIZE_EIGHTH = 4;
void DEFAULT_BORDER_SIZE_EIGHTH;

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
  /** Counter-clockwise degrees about the origin (rotated axis titles). */
  readonly rotationDeg?: number;
}
interface ChartLayout {
  readonly shapes: ReadonlyArray<ChartShapePrim>;
  readonly texts: ReadonlyArray<ChartTextPrim>;
}
interface ChartBlockLaidOut {
  readonly kind: 'chart';
  readonly float?: FloatAnchor;
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
  readonly float?: FloatAnchor;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly resolvedAlignment: 'left' | 'center' | 'right' | 'both' | 'distribute';
  readonly resourceName: string;
  /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
  readonly crop?: ImageCrop;
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
  readonly float?: FloatAnchor;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly paths: ReadonlyArray<VectorPath>;
  readonly fillColorHex?: string;
  readonly fillGradient?: ShapeGradient;
  readonly stroke?: StrokeStyle;
  readonly shadow?: ShapeShadow;
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
  // §17.13.6.2 — bookmark names anchored to this paragraph; pagination
  // records the first line's page + y as their GoTo destination.
  readonly bookmarks?: ReadonlyArray<string>;
  // The source paragraph — kept so pagination can re-wrap the block with
  // per-line widths when it overlaps a float's exclusion area.
  readonly source?: Paragraph;
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
  readonly dataBar?: CellDataBar;
  readonly icon?: CellIcon;
  readonly sparkline?: CellSparkline;
  // A data-validation `list` cell paints a dropdown button at the right edge.
  readonly dropdown?: boolean;
  readonly lines: ReadonlyArray<Line>;
  // Nested tables (a w:tbl inside this cell) rendered below the lines.
  readonly nestedTables?: ReadonlyArray<TableBlock>;
  readonly contentHeightPt: number;
  readonly totalHeightPt: number;
  readonly verticalAlign?: 'top' | 'center' | 'bottom';
  readonly colStart: number;
  readonly colSpan: number;
  readonly mergeRole: MergeRole;
  readonly clipped?: boolean;
  /**
   * The width the cell's own fill and rules cover, when a text-overflow span
   * made {@link widthPt} wider than the cell itself.
   */
  readonly paintWidthPt?: number;
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

/**
 * §17.13.6.2 — a bookmark's GoTo destination: the page (0-based) and the y-up
 * top of the anchoring paragraph's first line.
 */
export interface BookmarkPosition {
  readonly pageIdx: number;
  readonly yTopPt: number;
}

/** A comment's `/Text` annotation payload: the author and the flattened body (CM2b). */
export interface CommentNote {
  readonly author?: string;
  readonly contents: string;
}

/**
 * PDF-only companion the same layout pass produces (oop-design A13): the
 * logical-structure tree, per-section geometry (the emit fallback page), and the
 * parsed PDF/A profile. Consumed only by `emitStyledPdf`; the SVG writer never
 * sees it.
 */
export interface PdfLayoutAux {
  readonly structBuilder: StructTreeBuilder | undefined;
  readonly sectionCtxs: ReadonlyArray<SectionRenderCtx>;
  readonly pdfaProfile: PdfAProfile | undefined;
  readonly tagged: boolean;
  readonly bookmarks: ReadonlyMap<string, BookmarkPosition>;
  /**
   * CM2b — comment marker anchor (`comment-${n}`) → the note the emitter attaches
   * as a `/Text` annotation. Present only when `commentAnnotations` was requested.
   */
  readonly commentNotes?: ReadonlyMap<string, CommentNote>;
}

/**
 * What {@link layoutStyledDocument} actually returns: the PageDoc with the PDF
 * companion riding on `pdf`. Assignable to the narrow `LaidOutDocument`, so
 * PageDoc-only consumers (`writeSvg`) take it as-is.
 */
export interface LaidOutPdfDocument extends LaidOutDocument {
  readonly pdf: PdfLayoutAux;
}

// §17.11 footnote machinery: the separator rule above the notes band and the
// space it occupies (rule + breathing room), in points.
const FOOTNOTE_RULE_PT = 0.75;
const FOOTNOTE_RULE_GAP_ABOVE = 4;
const FOOTNOTE_SEPARATOR_HEIGHT = 10;

// Float text wrapping: the gap between a side-wrapped float and the text
// flowing beside it, and the floor below which line narrowing stops.
const FLOAT_TEXT_GAP = 6;
const MIN_WRAP_WIDTH = 36;

// Out-of-flow floats: wrap none renders at its anchor with no text effect;
// the side-wrapping modes additionally claim an exclusion rectangle the
// paragraphs flow around. topAndBottom stays a plain in-flow block.
function isOutOfFlowFloat(f: FloatAnchor | undefined): f is FloatAnchor {
  return f !== undefined && f.wrap !== 'topAndBottom';
}
const FOOTNOTE_RULE_WIDTH = 144; // Word's ~2" short separator

// Assign sequential numbers to note references in reading order (§17.11:
// footnotes and endnotes each keep their own counter) and rewrite each
// reference run to render its number superscript. Returns copies — direct
// renderStyledPdf callers own their trees.
function assignNoteNumbers(
  body: ReadonlyArray<BodyElement>,
  // Comment ids whose content is actually rendered (the after-body entries):
  // their markers become clickable jumps; a dangling ref gets a marker only.
  commentIds?: ReadonlySet<string>,
): {
  body: ReadonlyArray<BodyElement>;
  footnotes: ReadonlyMap<string, number>;
  endnotes: ReadonlyMap<string, number>;
  comments: ReadonlyMap<string, number>;
  // Footnote ids whose references sit OUTSIDE top-level paragraphs (table
  // cells, shape text): greedy bottom-of-page placement only tracks paragraph
  // lines, so these notes flow after the body instead (documented v1).
  deferredFootnotes: ReadonlyArray<string>;
} {
  const footnotes = new Map<string, number>();
  const endnotes = new Map<string, number>();
  const comments = new Map<string, number>();
  const paragraphFootnotes = new Set<string>();

  const numberRun = (run: Run, n: number): Run => ({
    ...run,
    text: String(n),
    properties: { ...run.properties, verticalAlign: 'superscript' },
  });
  // A comment marker reads `[n]` (bracketed) to read distinctly from a footnote
  // number, and — when the comment's entry is rendered — becomes a clickable
  // jump to it (an internal GoTo to the `comment-${n}` bookmark, reusing the
  // link path; PDF/A- and tagged-safe). A dangling ref gets a plain marker.
  const commentMarkerRun = (run: Run, n: number): Run => ({
    ...run,
    text: `[${n}]`,
    ...(run.commentRef !== undefined && commentIds?.has(run.commentRef)
      ? { anchor: `comment-${n}` }
      : {}),
    properties: { ...run.properties, verticalAlign: 'superscript' },
  });

  const mapParagraph = (paragraph: Paragraph): Paragraph => {
    if (
      !paragraph.runs.some(
        (r) =>
          r.footnoteRef !== undefined || r.endnoteRef !== undefined || r.commentRef !== undefined,
      )
    ) {
      return paragraph;
    }
    return {
      ...paragraph,
      runs: paragraph.runs.map((run) => {
        if (run.footnoteRef !== undefined) {
          let n = footnotes.get(run.footnoteRef);
          if (n === undefined) {
            n = footnotes.size + 1;
            footnotes.set(run.footnoteRef, n);
          }
          return numberRun(run, n);
        }
        if (run.endnoteRef !== undefined) {
          let n = endnotes.get(run.endnoteRef);
          if (n === undefined) {
            n = endnotes.size + 1;
            endnotes.set(run.endnoteRef, n);
          }
          return numberRun(run, n);
        }
        if (run.commentRef !== undefined) {
          let n = comments.get(run.commentRef);
          if (n === undefined) {
            n = comments.size + 1;
            comments.set(run.commentRef, n);
          }
          return commentMarkerRun(run, n);
        }
        return run;
      }),
    };
  };

  const mapElement = (el: BodyElement, topLevel: boolean): BodyElement => {
    if (el.kind === 'paragraph') {
      if (topLevel) {
        for (const r of el.paragraph.runs) {
          if (r.footnoteRef !== undefined) paragraphFootnotes.add(r.footnoteRef);
        }
      }
      return { kind: 'paragraph', paragraph: mapParagraph(el.paragraph) };
    }
    if (el.kind === 'table') {
      return {
        kind: 'table',
        table: {
          ...el.table,
          rows: el.table.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: cell.content.map((c) => mapElement(c, false)),
            })),
          })),
        },
      };
    }
    if (el.kind === 'shape' && el.shape.text) {
      return {
        kind: 'shape',
        shape: {
          ...el.shape,
          text: {
            ...el.shape.text,
            content: el.shape.text.content.map((c) => mapElement(c, false)),
          },
        },
      };
    }
    return el;
  };

  // Cheap pre-check: most documents carry no notes at all.
  const hasRefs = (els: ReadonlyArray<BodyElement>): boolean =>
    els.some((el) => {
      if (el.kind === 'paragraph') {
        return el.paragraph.runs.some(
          (r) =>
            r.footnoteRef !== undefined || r.endnoteRef !== undefined || r.commentRef !== undefined,
        );
      }
      if (el.kind === 'table') {
        return el.table.rows.some((row) => row.cells.some((c) => hasRefs(c.content)));
      }
      if (el.kind === 'shape' && el.shape.text) return hasRefs(el.shape.text.content);
      return false;
    });
  if (!hasRefs(body)) {
    return { body, footnotes, endnotes, comments, deferredFootnotes: [] };
  }

  const mapped = body.map((el) => mapElement(el, true));
  const deferredFootnotes = [...footnotes.keys()].filter((id) => !paragraphFootnotes.has(id));
  return { body: mapped, footnotes, endnotes, comments, deferredFootnotes };
}

// Replace the note's own-number placeholder (w:footnoteRef) with the number,
// or prepend it when the source omitted the placeholder.
function substituteNoteNumber(
  content: ReadonlyArray<BodyElement>,
  n: number,
): ReadonlyArray<BodyElement> {
  const hasPlaceholder = content.some(
    (el) => el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.noteNumber),
  );
  const out = content.map((el) => {
    if (el.kind !== 'paragraph') return el;
    if (!el.paragraph.runs.some((r) => r.noteNumber)) return el;
    return {
      kind: 'paragraph' as const,
      paragraph: {
        ...el.paragraph,
        runs: el.paragraph.runs.map((r) =>
          r.noteNumber
            ? {
                ...r,
                text: String(n),
                properties: { ...r.properties, verticalAlign: 'superscript' as const },
              }
            : r,
        ),
      },
    };
  });
  if (hasPlaceholder || out.length === 0) return out;
  const first = out[0]!;
  if (first.kind !== 'paragraph') return out;
  return [
    {
      kind: 'paragraph',
      paragraph: {
        ...first.paragraph,
        runs: [
          {
            text: `${n} `,
            properties: { verticalAlign: 'superscript' as const },
          },
          ...first.paragraph.runs,
        ],
      },
    },
    ...out.slice(1),
  ];
}

// Flatten a comment's block content to one plain string (paragraphs joined by
// newlines) — the /Text annotation's pop-up body (CM2b).
function flattenCommentText(content: ReadonlyArray<BodyElement>): string {
  const lines: Array<string> = [];
  const visit = (els: ReadonlyArray<BodyElement>): void => {
    for (const el of els) {
      if (el.kind === 'paragraph') {
        lines.push(el.paragraph.runs.map((r) => r.text).join(''));
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(content);
  return lines.join('\n').trim();
}

// Map each numbered comment's marker anchor (`comment-${n}`) to its note payload
// for the emitter's /Text annotations (CM2b).
function buildCommentNotes(
  numbers: ReadonlyMap<string, number>,
  comments: ReadonlyMap<string, Comment>,
): Map<string, CommentNote> {
  const out = new Map<string, CommentNote>();
  for (const [id, n] of numbers) {
    const c = comments.get(id);
    if (!c) continue;
    out.set(`comment-${n}`, {
      ...(c.author !== undefined ? { author: c.author } : {}),
      contents: flattenCommentText(c.content),
    });
  }
  return out;
}

// A review comment's after-body entry (E-COMMENTS CM1): its content led by an
// `[n]` marker and the author/date, mirroring how endnotes prepend their number.
// A reply (CM4) is indented by its thread depth and notes the parent it answers;
// a resolved thread is tagged `(resolved)`. The cues stay ASCII so they render
// in any embedded font.
function commentTailBlocks(
  comment: Comment,
  n: number,
  opts?: { parentN?: number; depth?: number; done?: boolean },
): ReadonlyArray<BodyElement> {
  const who = [comment.author, comment.date]
    .filter((s): s is string => s !== undefined && s.length > 0)
    .join(', ');
  const inReplyTo = opts?.parentN !== undefined ? ` (in reply to [${opts.parentN}])` : '';
  const resolved = opts?.done ? ' (resolved)' : '';
  const labelRun: Run = {
    text: who ? `[${n}] ${who}${inReplyTo}${resolved}: ` : `[${n}]${inReplyTo}${resolved} `,
    properties: {},
  };
  // The entry is the destination the in-text marker jumps to (E-COMMENTS CM2).
  const bm = `comment-${n}`;
  const indent = opts?.depth && opts.depth > 0 ? { indentLeft: pt(18 * opts.depth) } : {};
  const first = comment.content[0];
  if (first && first.kind === 'paragraph') {
    return [
      {
        kind: 'paragraph',
        paragraph: {
          ...first.paragraph,
          properties: { ...first.paragraph.properties, ...indent },
          runs: [labelRun, ...first.paragraph.runs],
          bookmarks: [bm, ...(first.paragraph.bookmarks ?? [])],
        },
      },
      ...comment.content.slice(1),
    ];
  }
  return [
    {
      kind: 'paragraph',
      paragraph: { properties: { ...indent }, runs: [labelRun], bookmarks: [bm] },
    },
    ...comment.content,
  ];
}

// What pagination needs to place footnotes: per-id content, numbers, and a
// per-section lazily-cached layout of each note at that section's width.
interface NotePlan {
  readonly numbers: ReadonlyMap<string, number>;
  readonly layout: (
    ctx: SectionRenderCtx,
    id: string,
  ) => { blocks: ReadonlyArray<LaidOutBlock>; heightPt: number } | undefined;
}

/**
 * Layout phase (the FlowDoc→PageDoc transform of ir-design §7): body →
 * positioned pages (PageItems), font/image resources, logical structure. Drives
 * the whole pass — numbering and note numbering, per-section geometry, per-block
 * layout, the footnote/endnote/comment tails, then pagination — and returns the
 * {@link LaidOutPdfDocument} (the PageDoc plus its PDF companion).
 *
 * @param body    The document body the section model partitions.
 * @param options The resolved fonts, styles, section model, and PDF-output toggles.
 * @returns The positioned pages plus the PDF-only {@link PdfLayoutAux} companion.
 */
export function layoutStyledDocument(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): LaidOutPdfDocument {
  const sectionList = resolveSectionList(body, options);

  const noteAssigned = assignNoteNumbers(
    applyNumbering(body, options.numbering),
    options.comments ? new Set(options.comments.keys()) : undefined,
  );
  const numberedBody = noteAssigned.body;
  const numberedHeadersFooters = applyNumberingToHeadersFooters(
    options.headersFooters,
    options.numbering,
  );

  // Tagged PDF (ISO 32000-1 §14.8) — implied by PDF/A-1a. When on, paginate
  // builds a logical structure tree and emit marks body text / artifacts.
  const pdfaProfile = options.pdfA ? parsePdfAProfile(options.pdfA) : undefined;
  const tagged =
    options.tagged === true || options.pdfUA === true || (pdfaProfile?.tagged ?? false);
  const structBuilder = tagged ? new StructTreeBuilder() : undefined;
  const fontResources = collectFontResources(numberedBody, numberedHeadersFooters, options);
  const imageResources = collectImageResources(numberedBody, numberedHeadersFooters, options);

  // Pre-compute per-section render context (geometry + header/footer bands).
  const sectionCtxs: Array<SectionRenderCtx> = sectionList.map((s) =>
    buildSectionContext(s, options, numberedHeadersFooters, fontResources, imageResources),
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
      ctx.columns ? ctx.columns[0]!.widthPt : ctx.contentWidth,
      ctx.pageContentHeight,
    );
  });

  // §17.3.1.9 `w:contextualSpacing` — the space between two paragraphs of the
  // SAME style is dropped, which is how a list sits together while keeping its
  // distance from the text around it. Read nowhere, every item of
  // ComplexNumberedLists.docx stood ten points from the next.
  for (let i = 0; i < blocks.length; i++) {
    const here = blocks[i];
    const next = blocks[i + 1];
    if (here?.kind !== 'paragraph' || next?.kind !== 'paragraph') continue;
    if (here.source?.properties.styleId !== next.source?.properties.styleId) continue;
    if (here.resolved.contextualSpacing) blocks[i] = { ...here, spacingAfterPt: 0 };
    if (next.resolved.contextualSpacing) blocks[i + 1] = { ...next, spacingBeforePt: 0 };
  }

  // Footnote plan: per-section lazily-cached layout of each note's content at
  // that section's width (notes referenced only from tables/shape text flow
  // after the body instead — see assignNoteNumbers).
  const noteBlockCache = new Map<SectionRenderCtx, Map<string, ReturnType<NotePlan['layout']>>>();
  const notePlan: NotePlan | undefined =
    options.footnotes && noteAssigned.footnotes.size > 0
      ? {
          numbers: noteAssigned.footnotes,
          layout: (sectionCtx, id) => {
            let byId = noteBlockCache.get(sectionCtx);
            if (!byId) {
              byId = new Map();
              noteBlockCache.set(sectionCtx, byId);
            }
            if (byId.has(id)) return byId.get(id);
            const content = options.footnotes?.get(id);
            const n = noteAssigned.footnotes.get(id);
            let laid: ReturnType<NotePlan['layout']> = undefined;
            if (content && n !== undefined && !noteAssigned.deferredFootnotes.includes(id)) {
              const noteBlocks = substituteNoteNumber(content, n).map((el) =>
                layoutBodyElement(
                  el,
                  options,
                  fontResources,
                  imageResources,
                  sectionCtx.contentWidth,
                  sectionCtx.pageContentHeight,
                ),
              );
              const heightPt = noteBlocks.reduce(
                (sum, b) =>
                  sum +
                  (b.kind === 'paragraph' ? b.spacingBeforePt + b.heightPt + b.spacingAfterPt : 0),
                0,
              );
              laid = { blocks: noteBlocks, heightPt };
            }
            byId.set(id, laid);
            return laid;
          },
        }
      : undefined;

  // Endnotes (and footnotes whose references the greedy pass cannot track)
  // flow after the body at the LAST section's width.
  const lastCtx = sectionCtxs[sectionCtxs.length - 1];
  if (lastCtx) {
    const tailNotes: Array<{ content: ReadonlyArray<BodyElement>; n: number }> = [];
    for (const id of noteAssigned.deferredFootnotes) {
      const content = options.footnotes?.get(id);
      const n = noteAssigned.footnotes.get(id);
      if (content && n !== undefined) tailNotes.push({ content, n });
    }
    for (const [id, n] of noteAssigned.endnotes) {
      const content = options.endnotes?.get(id);
      if (content) tailNotes.push({ content, n });
    }
    for (const note of tailNotes.sort((a, b) => a.n - b.n)) {
      for (const el of substituteNoteNumber(note.content, note.n)) {
        blocks.push(
          layoutBodyElement(
            el,
            options,
            fontResources,
            imageResources,
            lastCtx.contentWidth,
            lastCtx.pageContentHeight,
          ),
        );
      }
    }
    // Review comments follow the notes (E-COMMENTS CM1): each entry led by its
    // [n] marker and author/date, anchoring the in-text marker. Replies indent
    // by thread depth and note their parent's number (CM4).
    const commentNums = noteAssigned.comments;
    const threadDepth = (id: string): number => {
      let depth = 0;
      const seen = new Set<string>([id]);
      let cur = options.comments?.get(id)?.parentId;
      while (cur !== undefined && commentNums.has(cur) && !seen.has(cur)) {
        depth++;
        seen.add(cur);
        cur = options.comments?.get(cur)?.parentId;
      }
      return depth;
    };
    const commentTail = [...commentNums]
      .map(([id, n]) => ({ id, comment: options.comments?.get(id), n }))
      .filter((e): e is { id: string; comment: Comment; n: number } => e.comment !== undefined)
      .sort((a, b) => a.n - b.n);
    for (const { id, comment, n } of commentTail) {
      const parentId = comment.parentId;
      const parentN =
        parentId !== undefined && commentNums.has(parentId) ? commentNums.get(parentId) : undefined;
      for (const el of commentTailBlocks(comment, n, {
        ...(parentN !== undefined ? { parentN } : {}),
        depth: threadDepth(id),
        ...(comment.done ? { done: true } : {}),
      })) {
        blocks.push(
          layoutBodyElement(
            el,
            options,
            fontResources,
            imageResources,
            lastCtx.contentWidth,
            lastCtx.pageContentHeight,
          ),
        );
      }
    }
  }

  const bookmarks = new Map<string, BookmarkPosition>();
  // CM2b — native /Text annotations for comments (opt-in). The emitter attaches
  // them at the marker, gated to interactive (non-PDF/A, non-tagged) output.
  const commentNotes =
    options.commentAnnotations && options.comments
      ? buildCommentNotes(noteAssigned.comments, options.comments)
      : undefined;
  // Float text wrapping: pagination re-wraps an overlapped paragraph with
  // per-line widths; the closure re-runs the paragraph layout at the given
  // column width with those widths.
  const reflowParagraph = (paragraph: Paragraph, width: number, widths: ReadonlyArray<number>) =>
    layoutParagraphBlock(paragraph, options, fontResources, imageResources, width, widths);
  const pages = paginateSections(
    blocks,
    sectionCtxs,
    structBuilder,
    options.language ?? 'en-US',
    notePlan,
    bookmarks,
    reflowParagraph,
  );

  return {
    pages,
    resources: options.resources ?? new ResourceStore(),
    fontResources,
    imageResources,
    pdf: {
      structBuilder,
      sectionCtxs,
      pdfaProfile,
      tagged,
      bookmarks,
      ...(commentNotes ? { commentNotes } : {}),
    },
  };
}

// Emit phase: PageDoc draft → PDF objects (content streams, page dicts,
// catalog, PDF/A apparatus, structure tree, signature placeholder) → bytes.

/**
 * The resolved page box + header/footer offsets, in points. Priority for page
 * geometry:
 *   1. explicit value in {@link StyledRenderOptions} (test/library caller override)
 *   2. value from section properties (`sectPr/pgSz/pgMar` from the docx)
 *   3. A4 + 1-inch margins fallback
 */
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

/**
 * Per-section render context: the section's exclusive body `endIndex`, its
 * resolved {@link PageDimensions} (flattened), the derived content box, the
 * column geometry, the pre-laid header/footer bands, and the title-page /
 * even-and-odd header toggles. Built once per section and threaded through both
 * layout and pagination.
 */
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
  /**
   * §17.6.4 multi-column sections: per-column x-offset (from `marginLeft`) and
   * width. Absent for single-column sections. Body blocks are laid out at the
   * FIRST column's width (explicit unequal widths degrade to flowing without
   * re-wrap); headers/footers and footnotes keep the full content width.
   */
  readonly columns?: ReadonlyArray<{ readonly xOffsetPt: number; readonly widthPt: number }>;
  readonly headerSet: HeaderFooterSet;
  readonly footerSet: HeaderFooterSet;
  readonly titlePg: boolean;
  readonly evenAndOddHeaders: boolean;
  /**
   * §17.6.22 — the section begins on the page already in hand, at the point the
   * one before it stopped, rather than on a fresh one.
   */
  readonly continuous: boolean;
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
  imageResources: ReadonlyMap<ResourceId, ImageResource>,
): SectionRenderCtx {
  const dims = resolvePageDimensions(options, section.properties);
  const contentWidth = dims.pageWidth - dims.marginLeft - dims.marginRight;
  const headerSet = layoutHeaderSet(
    section.properties,
    headersFooters,
    options,
    fontResources,
    imageResources,
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
    imageResources,
    contentWidth,
    dims.marginLeft,
    dims.pageHeight,
    dims.footerOffsetPt,
  );
  const columns = buildColumnGeometry(section.properties.columns, contentWidth);
  // A header band taller than the gap between its own offset and the top margin
  // would otherwise be drawn straight over the first rows of the body. Excel and
  // Word both push the body down instead; a multi-line header is ordinary (a
  // spreadsheet header region may carry a literal line break).
  const headerBottom =
    dims.headerOffsetPt +
    Math.max(
      headerSet.default.heightPt ?? 0,
      headerSet.first.heightPt ?? 0,
      headerSet.even.heightPt ?? 0,
    );
  const marginTop = Math.max(dims.marginTop, headerBottom);
  return {
    endIndex: section.endIndex,
    properties: section.properties,
    pageWidth: dims.pageWidth,
    pageHeight: dims.pageHeight,
    marginLeft: dims.marginLeft,
    marginTop,
    marginBottom: dims.marginBottom,
    contentWidth,
    pageContentHeight: dims.pageHeight - marginTop - dims.marginBottom,
    ...(columns ? { columns } : {}),
    headerSet,
    footerSet,
    titlePg: section.properties.titlePg === true,
    evenAndOddHeaders: section.properties.evenAndOddHeaders === true,
    continuous: section.properties.sectionStart === 'continuous',
  };
}

// §17.6.4: explicit w:col list as given; otherwise equal widths separated by
// the shared gutter.
function buildColumnGeometry(
  cols: SectionColumns | undefined,
  contentWidth: number,
): Array<{ xOffsetPt: number; widthPt: number }> | undefined {
  if (!cols || cols.count <= 1) return undefined;
  const out: Array<{ xOffsetPt: number; widthPt: number }> = [];
  if (cols.explicit && cols.explicit.length > 1) {
    let x = 0;
    for (const c of cols.explicit) {
      out.push({ xOffsetPt: x, widthPt: c.widthPt });
      x += c.widthPt + c.spacePt;
    }
    return out;
  }
  const width = (contentWidth - cols.spacePt * (cols.count - 1)) / cols.count;
  if (width <= 0) return undefined;
  for (let i = 0; i < cols.count; i++) {
    out.push({ xOffsetPt: i * (width + cols.spacePt), widthPt: width });
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

// One header/footer band. Static bands carry their pre-rendered commands
// (the byte-identical fast path). A band containing PAGE/NUMPAGES fields is
// DYNAMIC: it re-lays out per page once pagination knows both numbers
// (§17.16.5.33/.35) — substitution changes text widths, so this is an honest
// re-layout, not a glyph swap. w:pgNumType start offsets are not applied (v1).
interface HfBandEntry {
  readonly commands: Array<PageItem>;
  readonly renderDynamic?: (pageNumber: number, totalPages: number) => Array<PageItem>;
  /** Laid-out height of the band, so the body can be kept clear of it. */
  readonly heightPt?: number;
}

interface HeaderFooterSet {
  readonly default: HfBandEntry;
  readonly first: HfBandEntry;
  readonly even: HfBandEntry;
}

// Tag every command in a header/footer band as a pagination artifact so the
// tagged-PDF emit keeps it out of the structure tree (§14.8.2.2.2). A no-op for
// non-tagged output — the field is simply ignored at emit.
function markPagination(cmds: Array<PageItem>): Array<PageItem> {
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
  imageResources: ReadonlyMap<ResourceId, ImageResource> | undefined,
  contentWidth: number,
  marginLeft: number,
  pageHeight: number,
  headerOffsetPt: number,
): HeaderFooterSet {
  const band = (type: HeaderFooterType): HfBandEntry => {
    const ref = refByType(section.headers, type);
    if (!ref) return { commands: [] };
    const content = headersFooters.get(ref.relationshipId);
    if (!content) return { commands: [] };
    let measured = 0;
    const render = (c: ReadonlyArray<BodyElement>): Array<PageItem> => {
      const blocks = laidOutBlocksFor(c, options, fontResources, contentWidth, imageResources);
      measured = blocksHeight(blocks);
      return markPagination(
        drawBlocksSequentially(
          blocks,
          marginLeft,
          pageHeight - headerOffsetPt,
          pageHeight,
          contentWidth,
        ),
      );
    };
    if (contentHasPageFields(content)) {
      // Measure once with placeholder values so the height is known before the
      // first page is composed; the per-page render replaces the commands.
      render(substitutePageFields(content, 1, 1));
      return {
        commands: [],
        heightPt: measured,
        renderDynamic: (n, total) => render(substitutePageFields(content, n, total)),
      };
    }
    const commands = render(content);
    return { commands, heightPt: measured };
  };
  return { default: band('default'), first: band('first'), even: band('even') };
}

/** Total laid-out height of a run of blocks, paragraph spacing included. */
function blocksHeight(blocks: ReadonlyArray<LaidOutBlock>): number {
  return blocks.reduce(
    (sum, b) =>
      sum +
      (b.kind === 'paragraph' ? b.spacingBeforePt + b.heightPt + b.spacingAfterPt : b.heightPt),
    0,
  );
}

function layoutFooterSet(
  section: SectionProperties,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<ResourceId, ImageResource> | undefined,
  contentWidth: number,
  marginLeft: number,
  pageHeight: number,
  footerOffsetPt: number,
): HeaderFooterSet {
  const band = (type: HeaderFooterType): HfBandEntry => {
    const ref = refByType(section.footers, type);
    if (!ref) return { commands: [] };
    const content = headersFooters.get(ref.relationshipId);
    if (!content) return { commands: [] };
    const render = (c: ReadonlyArray<BodyElement>): Array<PageItem> => {
      const blocks = laidOutBlocksFor(c, options, fontResources, contentWidth, imageResources);
      const totalHeight = blocks.reduce(
        (sum, b) =>
          sum +
          (b.kind === 'paragraph' ? b.spacingBeforePt + b.heightPt + b.spacingAfterPt : b.heightPt),
        0,
      );
      return markPagination(
        drawBlocksSequentially(
          blocks,
          marginLeft,
          footerOffsetPt + totalHeight,
          pageHeight,
          contentWidth,
        ),
      );
    };
    if (contentHasPageFields(content)) {
      return {
        commands: [],
        renderDynamic: (n, total) => render(substitutePageFields(content, n, total)),
      };
    }
    return { commands: render(content) };
  };
  return { default: band('default'), first: band('first'), even: band('even') };
}

// A band is dynamic when any of its paragraphs carries a PAGE/NUMPAGES field
// run (bands render paragraphs only).
function contentHasPageFields(content: ReadonlyArray<BodyElement>): boolean {
  for (const el of content) {
    if (el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.field !== undefined)) {
      return true;
    }
  }
  return false;
}

// Clone the band content with field runs' cached text replaced by the real
// numbers for this page.
function substitutePageFields(
  content: ReadonlyArray<BodyElement>,
  pageNumber: number,
  totalPages: number,
): ReadonlyArray<BodyElement> {
  return content.map((el) => {
    if (el.kind !== 'paragraph') return el;
    if (!el.paragraph.runs.some((r) => r.field !== undefined)) return el;
    return {
      kind: 'paragraph',
      paragraph: {
        ...el.paragraph,
        runs: el.paragraph.runs.map((r) =>
          r.field === undefined
            ? r
            : { ...r, text: String(r.field === 'PAGE' ? pageNumber : totalPages) },
        ),
      },
    };
  });
}

function pickBand(set: HeaderFooterSet, band: HfBand): HfBandEntry {
  const has = (e: HfBandEntry) => e.commands.length > 0 || e.renderDynamic !== undefined;
  // §17.10.6 — `w:titlePg` says the first page's header is DIFFERENT. A section
  // that turns it on and declares no `first` part means different by being
  // empty: falling back to the default put a header on the title page that
  // neither Word nor LibreOffice draws (ImageCrop.docx, whose only page is one).
  // `w:evenAndOddHeaders` is not the same bargain — LibreOffice prints the
  // default on even pages when the document declares no even part.
  if (band === 'first') return set.first;
  if (band === 'even') return has(set.even) ? set.even : set.default;
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
  imageResources?: ReadonlyMap<ResourceId, ImageResource>,
): Array<LaidOutBlock> {
  return elements.map((el) =>
    layoutBodyElement(el, options, fontResources, imageResources, contentWidth),
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
    ...(image.crop ? { crop: image.crop } : {}),
    spacingBeforePt: image.paragraphProperties.spacingBefore ?? 0,
    spacingAfterPt: image.paragraphProperties.spacingAfter ?? 0,
    ...(image.altText ? { altText: image.altText } : {}),
    ...(image.float ? { float: image.float } : {}),
  };
}

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
  const fillGradient = shape.fill.kind === 'gradient' ? shape.fill.gradient : undefined;
  const fillColorHex =
    shape.fill.kind === 'solid'
      ? shape.fill.colorHex
      : fillGradient
        ? gradientToSolid(fillGradient)
        : undefined;
  const stroke = buildStroke(shape.line);
  const t = shape.transform;
  const pp = shape.paragraphProperties;

  const text = shape.text;
  const insetLeftPt = text?.insetLeft ?? DEFAULT_INSET_LR_PT;
  const insetRightPt = text?.insetRight ?? DEFAULT_INSET_LR_PT;
  const insetTopPt = text?.insetTop ?? DEFAULT_INSET_TB_PT;
  const insetBottomPt = text?.insetBottom ?? DEFAULT_INSET_TB_PT;
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
      // §21.1.2.2.3 — a paragraph with no runs is still a LINE (the blank line
      // the author typed); layoutParagraphBlock makes it, so a shape's text no
      // longer closes up over one (shape-macro-ext-ref.xlsx opens its button's
      // caption with a blank line and every paragraph after it drew 14pt high).
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
    ...(fillGradient ? { fillGradient } : {}),
    ...(stroke ? { stroke } : {}),
    ...(shape.shadow ? { shadow: shape.shadow } : {}),
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
    ...(shape.float ? { float: shape.float } : {}),
  };
}

// Build the `cm` matrix that places a shape's local y-up frame on the page at
// bottom-left (pageX, pageY), rotated about its centre and optionally flipped.

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
    ...(block.float ? { float: block.float } : {}),
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
  // Z-order: the chart-space frame under everything, then the gridlines, then
  // the plotted data over both. Gridlines drawn after the bars ruled white
  // lines straight across every one of them (123233_charts.xlsx).
  if (scene.background) shapes.push(rectPrim(scene.background));
  for (const g of scene.gridlines ?? []) shapes.push(polylinePrim(g));
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

// --- Conditional-format icons (E-SHEET SC1c) -------------------------------
// A glyph drawn in the cell's left gutter (CF_ICON_GUTTER_PT wide), sized to
// CF_ICON_SIZE_PT and vertically centred in the row. Built in a local y-up
// [0,size]² frame; emitRowChunk composes the page-flip transform. Single-glyph
// families (lights, arrows, signs, flags, symbols) draw one prim; the meter
// families — ratings (a bar histogram) and quarters (a clock pie) — draw several.
const CF_ICON_SIZE_PT = 9;
const CF_ICON_GUTTER_PT = 12;
// The unfilled portion of a meter glyph (ratings bars / quarter pie).
const CF_ICON_EMPTY_HEX = 'BFBFBF';

// A data-validation `list` cell paints a dropdown button at its right edge
// (E-SHEET SV1): a light-grey square (thin border) with a dark ▾. Sized to the
// row, capped at DROPDOWN_BUTTON_PT; the cell reserves DROPDOWN_GUTTER_PT of
// right padding so the value never runs under the button.
const DROPDOWN_BUTTON_PT = 11;
const DROPDOWN_GUTTER_PT = 13;
const DROPDOWN_FILL_HEX = 'F0F0F0';
const DROPDOWN_BORDER_HEX = 'B0B0B0';
const DROPDOWN_ARROW_HEX = '595959';

// The button (filled, thin-bordered square) + a centred ▾, in a local y-up
// [0,s]² frame; emitRowChunk composes the page-flip transform (mirrors the CF
// icon path). The triangle's apex points toward the bottom (smaller y).
function buildDropdownPrims(s: number): ReadonlyArray<ChartShapePrim> {
  return [
    {
      paths: [rectAtPath(0, 0, s, s)],
      fillColorHex: DROPDOWN_FILL_HEX,
      stroke: { colorHex: DROPDOWN_BORDER_HEX, widthPt: 0.5 },
    },
    {
      paths: [trianglePath([s / 2, s * 0.34], [s * 0.26, s * 0.64], [s * 0.74, s * 0.64])],
      fillColorHex: DROPDOWN_ARROW_HEX,
    },
  ];
}

type FilledIconShape =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangleUp'
  | 'triangleDown'
  | 'triangleRight';

function buildCellIconShape(icon: CellIcon, s: number): ReadonlyArray<ChartShapePrim> {
  const color = icon.colorHex;
  switch (icon.shape) {
    case 'bars':
      return barsIconPrims(s, color, icon.fill);
    case 'pie':
      return pieIconPrims(s, color, icon.fill);
    case 'check':
      return [{ paths: [checkPath(s)], stroke: { colorHex: color, widthPt: s * 0.16 } }];
    case 'cross':
      return [{ paths: crossPaths(s), stroke: { colorHex: color, widthPt: s * 0.16 } }];
    case 'exclamation':
      return [
        {
          paths: [
            rectAtPath(s * 0.41, s * 0.34, s * 0.18, s * 0.56),
            rectAtPath(s * 0.41, s * 0.1, s * 0.18, s * 0.16),
          ],
          fillColorHex: color,
        },
      ];
    default:
      return [{ paths: [cellIconPath(icon.shape, s)], fillColorHex: color }];
  }
}

function cellIconPath(shape: FilledIconShape, s: number): VectorPath {
  switch (shape) {
    case 'square':
      return rectAtPath(s * 0.12, s * 0.12, s * 0.76, s * 0.76);
    case 'diamond':
      return new PathBuilder()
        .moveTo(s / 2, s * 0.92)
        .lineTo(s * 0.92, s / 2)
        .lineTo(s / 2, s * 0.08)
        .lineTo(s * 0.08, s / 2)
        .close()
        .build();
    case 'triangleUp':
      return trianglePath([s / 2, s * 0.9], [s * 0.1, s * 0.12], [s * 0.9, s * 0.12]);
    case 'triangleDown':
      return trianglePath([s / 2, s * 0.1], [s * 0.1, s * 0.88], [s * 0.9, s * 0.88]);
    case 'triangleRight':
      return trianglePath([s * 0.9, s / 2], [s * 0.12, s * 0.1], [s * 0.12, s * 0.9]);
    case 'circle':
      return circlePath(s / 2, s / 2, s * 0.42);
  }
}

// Ratings (4/5): `levels` ascending bars left→right, the first `filled` coloured.
function barsIconPrims(
  s: number,
  color: string,
  fill: CellIcon['fill'],
): ReadonlyArray<ChartShapePrim> {
  const n = Math.max(1, fill?.levels ?? 4);
  const filled = fill?.filled ?? 0;
  const gap = s * 0.12;
  const bw = (s - gap * (n - 1)) / n;
  const prims: Array<ChartShapePrim> = [];
  for (let i = 0; i < n; i++) {
    const h = s * (0.32 + (0.68 * (i + 1)) / n);
    prims.push({
      paths: [rectAtPath(i * (bw + gap), 0, bw, h)],
      fillColorHex: i < filled ? color : CF_ICON_EMPTY_HEX,
    });
  }
  return prims;
}

// Quarters (5): a clock pie — `filled` of `levels` slices coloured clockwise from
// the top over a grey base circle (bucket 0 = empty, full = a solid disc).
function pieIconPrims(
  s: number,
  color: string,
  fill: CellIcon['fill'],
): ReadonlyArray<ChartShapePrim> {
  const cx = s / 2;
  const cy = s / 2;
  const r = s * 0.46;
  const levels = Math.max(1, fill?.levels ?? 4);
  const filled = Math.max(0, Math.min(levels, fill?.filled ?? 0));
  const prims: Array<ChartShapePrim> = [
    { paths: [circlePath(cx, cy, r)], fillColorHex: CF_ICON_EMPTY_HEX },
  ];
  if (filled >= levels) {
    prims.push({ paths: [circlePath(cx, cy, r)], fillColorHex: color });
  } else if (filled > 0) {
    const start = Math.PI / 2; // 12 o'clock in the y-up frame
    const sweep = -(filled / levels) * 2 * Math.PI; // clockwise
    const [sx, sy] = arcPoint(cx, cy, r, r, start);
    prims.push({
      paths: [
        new PathBuilder()
          .moveTo(cx, cy)
          .lineTo(sx, sy)
          .append(arcToBeziers(cx, cy, r, r, start, sweep))
          .close()
          .build(),
      ],
      fillColorHex: color,
    });
  }
  return prims;
}

// Symbols: a check mark (✓) and a cross (✗) as stroked polylines.
function checkPath(s: number): VectorPath {
  return new PathBuilder()
    .moveTo(s * 0.16, s * 0.5)
    .lineTo(s * 0.4, s * 0.26)
    .lineTo(s * 0.84, s * 0.78)
    .build();
}

function crossPaths(s: number): ReadonlyArray<VectorPath> {
  return [
    new PathBuilder()
      .moveTo(s * 0.24, s * 0.24)
      .lineTo(s * 0.76, s * 0.76)
      .build(),
    new PathBuilder()
      .moveTo(s * 0.24, s * 0.76)
      .lineTo(s * 0.76, s * 0.24)
      .build(),
  ];
}

function trianglePath(a: [number, number], b: [number, number], c: [number, number]): VectorPath {
  return new PathBuilder().moveTo(a[0], a[1]).lineTo(b[0], b[1]).lineTo(c[0], c[1]).close().build();
}

// Four cubic Béziers approximating a circle (kappa = 4/3·(√2 − 1)).
function circlePath(cx: number, cy: number, r: number): VectorPath {
  const k = 0.5522847498307936 * r;
  return new PathBuilder()
    .moveTo(cx + r, cy)
    .cubicTo(cx + r, cy + k, cx + k, cy + r, cx, cy + r)
    .cubicTo(cx - k, cy + r, cx - r, cy + k, cx - r, cy)
    .cubicTo(cx - r, cy - k, cx - k, cy - r, cx, cy - r)
    .cubicTo(cx + k, cy - r, cx + r, cy - k, cx + r, cy)
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
  const shift = l.align === 'center' ? -w / 2 : l.align === 'right' ? -w : 0;
  // A rotated label reads along the rotated axis, so its own alignment shifts
  // it there and not across the page.
  if (l.rotationDeg) {
    return { line, x: l.x, y: l.y + shift, rotationDeg: l.rotationDeg };
  }
  return { line, x: l.x + shift, y: l.y };
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
  for (const note of options.footnotes?.values() ?? []) visit(note);
  for (const note of options.endnotes?.values() ?? []) visit(note);
  for (const comment of options.comments?.values() ?? []) visit(comment.content);

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
// `startY` is in the internal y-up frame the band math works in; the emitted
// items carry top-left coordinates like everything else on a page.
function drawBlocksSequentially(
  blocks: ReadonlyArray<LaidOutBlock>,
  startX: number,
  startY: number,
  pageHeight: number,
  contentWidth: number,
  // Tagged PDF: stamp every emitted line with this structure node (used by
  // the footnote band; header/footer bands stay artifact-marked instead).
  structId?: number,
): Array<PageItem> {
  const out: Array<PageItem> = [];
  let cursorY = startY;
  for (const block of blocks) {
    // A header is not only paragraphs. A letterhead is very often ONE TABLE —
    // logo in the left cell, institute in the right — and the band drew
    // paragraphs alone, so such a document lost its header and its footer
    // whole: 090716_Studentische_Arbeit_VWS.docx prints a green crest and two
    // rules across every one of its six pages and we printed none of it.
    // No pagination here and no tagging: a band is one page's worth by
    // construction, and its contents are artifacts.
    if (block.kind === 'table') {
      const tableX = startX + block.xOffsetPt;
      for (const row of block.rows) {
        emitRowChunk(out, row, tableX, cursorY, pageHeight, block.colCount);
        cursorY -= row.heightPt;
      }
      continue;
    }
    // …and a first-page header is very often nothing BUT the crest.
    if (block.kind === 'image') {
      cursorY -= block.spacingBeforePt;
      const offset = alignmentOffset(block.resolvedAlignment, block.widthPt, contentWidth);
      out.push({
        type: 'image',
        x: pt(startX + offset),
        y: pt(pageHeight - cursorY),
        width: pt(block.widthPt),
        height: pt(block.heightPt),
        imageResourceName: block.resourceName,
        ...(block.crop ? { crop: block.crop } : {}),
      });
      cursorY -= block.heightPt + block.spacingAfterPt;
      continue;
    }
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
        originX: pt(startX + indentLeft + offset),
        baselineY: pt(pageHeight - (cursorY + lineDescent(line))),
        ...(structId !== undefined ? { structId } : {}),
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
          // A PAGE/NUMPAGES field renders substituted digits per page — make
          // sure every digit is in the subset, not just the cached result.
          if (run.field !== undefined) {
            addRun({ text: `${run.text}0123456789`, properties: run.properties }, el.paragraph);
            continue;
          }
          // Note references and the note's own-number placeholder render
          // substituted numbers — subset every digit for their fonts too.
          if (run.footnoteRef !== undefined || run.endnoteRef !== undefined || run.noteNumber) {
            addRun({ text: `${run.text}0123456789 `, properties: run.properties }, el.paragraph);
            continue;
          }
          addRun(run, el.paragraph);
        }
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) {
          for (const cell of row.cells) {
            // A number too wide for its column is drawn as a row of `#` in
            // place of its own digits (CellProperties.hashOnOverflow) — a
            // character that appears nowhere in the document's text. Left out
            // of the subset, the page encodes a glyph the embedded font no
            // longer has: bug69812.xlsx is one such cell, and its page came out
            // blank where both references print a number.
            if (cell.properties.hashOnOverflow === true) {
              for (const inner of cell.content) {
                if (inner.kind !== 'paragraph') continue;
                for (const run of inner.paragraph.runs) {
                  addRun({ text: '#', properties: run.properties }, inner.paragraph);
                }
              }
            }
            visit(cell.content);
          }
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
          // …and the legend name of a series that HAS no name is invented at
          // draw time (`Series1`), so ask for the same string the legend will
          // draw: 57362.xlsx's capital S is on the page nowhere else, and the
          // subset it was left out of drew "eries1".
          chart.series.forEach((sr, i) => add(legendSeriesName(sr, i)));
          // The axis titles are drawn too, and a character that appears ONLY
          // there was left out of the subset and drew blank — shape-macro-ext-
          // ref.xlsx prints "Translation X [mm]" as "Translation    mm", with
          // the text layer still claiming the missing glyphs.
          if (chart.catAxisTitle) add(chart.catAxisTitle);
          if (chart.valAxisTitle) add(chart.valAxisTitle);
          // §21.2.2.49 — a label the author typed is arbitrary text, not digits.
          for (const sr of chart.series) for (const label of sr.pointLabels ?? []) add(label.text);
          add('0123456789.,-%() '); // value-axis tick labels
          // …and whatever the axis's number format puts around them. A currency
          // code draws a `$` that appears nowhere else on the page, and left out
          // of the subset it drew blank: 123233_charts.xlsx labelled its axis
          // "2,000,000,000.00" where every reader writes "$2,000,000,000.00".
          if (chart.numberFormat) add(chart.numberFormat);
        }
      }
      // image-block elements use no fonts
    }
  };
  visit(body);
  for (const hf of headersFooters.values()) visit(hf);
  for (const note of options.footnotes?.values() ?? []) visit(note);
  for (const note of options.endnotes?.values() ?? []) visit(note);
  // A comment's content is only half of what gets drawn. `commentTailBlocks`
  // synthesises the rest — the author, the date, the `[n]` marker, the `re [n]`
  // reply cue, the `(resolved)` tag — and none of it lives in `comment.content`,
  // so a character appearing ONLY there never reached the subset. The review
  // fixture's author is "Alice Reviewer" and its capital A is nowhere else in
  // the document. Ask the builder for its blocks rather than restating its
  // wording here, and throw in every digit because the numbers are per-document.
  for (const comment of options.comments?.values() ?? []) {
    visit(comment.content);
    visit(commentTailBlocks(comment, 1, { parentN: 1, depth: 1, done: true }));
    visit([
      {
        kind: 'paragraph',
        paragraph: { properties: {}, runs: [{ text: '0123456789', properties: {} }] },
      },
    ]);
  }

  if (used.size === 0) {
    // A document with no text still needs a font RESOURCE — `lookupFont` falls
    // back to the first one and would otherwise hand back undefined — but it
    // needs no font PROGRAM. Seeding `.notdef` here made us embed a 13 KB
    // Roboto subset into 56017.xlsx, whose page draws nothing at all: 89% of
    // the file was a font no operator names. The emitter skips a resource with
    // no glyphs; ISO 32000-1 §7.8.3 wants the resources the stream needs.
    const regular = options.registry.resolveByStyle(false, false);
    used.set(regular.variant, { parsed: regular.parsed, gids: new Set<number>() });
  }

  const out = new Map<string, FontResource>();
  let counter = 0;
  for (const [variant, info] of used) {
    counter++;
    const resourceName = `F${counter}`;
    out.set(variant, {
      resourceName,
      parsed: info.parsed,
      // E-PARITY FP4: the 'word' profile measures kern-free (Word's default).
      measure: createFontMeasure(info.parsed, options.layoutProfile !== 'word'),
      gids: info.gids,
    });
  }
  return out;
}

/**
 * The one line an empty paragraph still occupies — no tokens, but as tall as
 * the paragraph mark's own font (§17.3.1.31 `w:rPr` on the `w:pPr`, which is
 * exactly the run properties Word keeps for a paragraph that holds no runs).
 */
function emptyLine(
  paragraph: Paragraph,
  options: StyledRenderOptions,
  resolved: ResolvedParagraphProperties,
  availableWidthPt: number,
  fontResources: ReadonlyMap<string, FontResource>,
): Line {
  const mark = resolveRunProperties(
    paragraph.properties.runProperties ?? {},
    paragraph.properties,
    options.styles,
  );
  const fontSizePt = mark.fontSizePt;
  const profile = options.layoutProfile ?? 'ream';
  const { variant } = options.registry.resolveByStyle(mark.bold, mark.italic);
  const parsed = fontResources.get(variant)?.parsed;
  const metric =
    profile !== 'ream' && parsed ? fontLeadingPt(parsed, fontSizePt, profile) : undefined;
  return {
    tokens: [],
    contentWidthPt: 0,
    maxFontSizePt: fontSizePt,
    availableWidthPt,
    firstLine: true,
    resolved,
    isLastInParagraph: true,
    ...(metric && metric.ascentPt + metric.descentPt > 0
      ? {
          metricHeightPt: metric.ascentPt + metric.descentPt + metric.lineGapPt,
          metricDescentPt: metric.descentPt,
        }
      : {}),
  };
}

function layoutParagraphBlock(
  paragraph: Paragraph,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
  imageResources: ReadonlyMap<string, ImageResource> | undefined,
  contentWidth: number,
  // Float wrapping: explicit per-line widths (overrides first/other widths).
  lineWidths?: ReadonlyArray<number>,
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
  const wrapped = wrap(
    tokens,
    firstLineWidth,
    otherWidth,
    resolved,
    options.hyphenator,
    options.layoutProfile ?? 'ream',
    lineWidths,
  );
  // §17.3.1 — a paragraph with nothing in it is still a LINE: the blank line
  // the author typed, as tall as the font it would have been typed in. Wrapping
  // zero tokens yields no line, so the paragraph took no room and everything
  // after it moved up — IllustrativeCases.docx opens each table's label column
  // with an empty paragraph, and every label ended up beside the value of the
  // row ABOVE it, "Gross Income" against "€" and the euro sign's own row gone.
  const lines =
    wrapped.length > 0
      ? wrapped
      : [emptyLine(paragraph, options, resolved, Math.max(firstLineWidth, 0), fontResources)];

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
    ...(paragraph.bookmarks && paragraph.bookmarks.length > 0
      ? { bookmarks: paragraph.bookmarks }
      : {}),
    source: paragraph,
  };
}

/**
 * Word and Excel both draw a super/subscript at about two thirds of the base
 * size, raised a third of it or dropped a sixth. The base line height does not
 * change: a superscript inside a line of body text must not open the leading.
 */
const SCRIPT_SCALE = 0.66;
const SCRIPT_OFFSET: Partial<Record<string, number>> = { superscript: 1 / 3, subscript: -1 / 6 };

interface RunPlan {
  readonly run: Paragraph['runs'][number];
  readonly resolvedRun: ResolvedRunProperties;
  readonly font: FontResource;
  readonly fontSizePt: number;
  /** How far off the baseline this run draws (super/subscript), in points. */
  readonly risePt?: number;
  readonly isImage: boolean;
  readonly imageWidthPt: number;
  readonly imageHeightPt: number;
  readonly imageResourceName: string;
  readonly imageCrop?: ImageCrop;
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
        ...(run.inlineImage.crop ? { imageCrop: run.inlineImage.crop } : {}),
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
    // §17.3.2.42 / §18.4.2 `vertAlign` — a super/subscript run draws SMALLER and
    // off the baseline. The model carried the flag and the HTML writer honoured
    // it; the PDF layout did neither, so a footnote marker and a cell's
    // "Salary⁽²⁾" came out full size, on the line (45540_classic_Header.xlsx).
    const script = SCRIPT_OFFSET[resolvedRun.verticalAlign] ?? 0;
    return {
      run,
      resolvedRun,
      font: lookupFont(fontResources, fontKey),
      fontSizePt: script === 0 ? resolvedRun.fontSizePt : resolvedRun.fontSizePt * SCRIPT_SCALE,
      ...(script === 0 ? {} : { risePt: resolvedRun.fontSizePt * script }),
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

  // Per-real-position embedding levels via the core/bidi segment facade —
  // the explicit-formatting protocol (RLE/LRE/PDF wrapping, U+FFFC objects)
  // lives there, not in the PDF layer (stage 6 / A5).
  const realLevels = segmentLevels(
    plans.map((plan) => ({
      text: plan.run.text,
      ...(plan.isImage || plan.math !== undefined ? { object: true } : {}),
      ...(plan.resolvedRun.rtl ? { rtl: true } : {}),
    })),
    baseDir,
  );

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
        ...(plan.imageCrop ? { crop: plan.imageCrop } : {}),
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
    const highlight = (plan.run.commentRangeRefs?.length ?? 0) > 0;
    for (const t of tokenizeText(plan.run.text)) {
      tokens.push({
        kind: 'text',
        text: t.text,
        isSpace: t.isSpace,
        ...(t.tab ? { tab: true as const } : {}),
        ...(plan.run.href !== undefined ? { href: plan.run.href } : {}),
        ...(plan.run.footnoteRef !== undefined ? { footnoteRef: plan.run.footnoteRef } : {}),
        ...(plan.run.anchor !== undefined ? { anchor: plan.run.anchor } : {}),
        ...(plan.run.listMarker ? { listMarker: true } : {}),
        ...(highlight ? { highlight: true } : {}),
        resolvedRun: plan.resolvedRun,
        font: plan.font,
        fontSizePt: plan.fontSizePt,
        ...(plan.risePt !== undefined ? { risePt: plan.risePt } : {}),
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
        ...(plan.imageCrop ? { crop: plan.imageCrop } : {}),
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
        ...(plan.run.href !== undefined ? { href: plan.run.href } : {}),
        ...(plan.run.footnoteRef !== undefined ? { footnoteRef: plan.run.footnoteRef } : {}),
        ...(plan.run.anchor !== undefined ? { anchor: plan.run.anchor } : {}),
        ...(plan.run.listMarker ? { listMarker: true } : {}),
        ...((plan.run.commentRangeRefs?.length ?? 0) > 0 ? { highlight: true } : {}),
        resolvedRun: plan.resolvedRun,
        font: plan.font,
        fontSizePt: plan.fontSizePt,
        ...(plan.risePt !== undefined ? { risePt: plan.risePt } : {}),
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

function tokenizeText(text: string): Array<{ text: string; isSpace: boolean; tab?: true }> {
  const out: Array<{ text: string; isSpace: boolean; tab?: true }> = [];
  if (text.length === 0) return out;
  // §17.3.1.38 — a tab is not whitespace that happens to be wide: it advances
  // to a POSITION, so it gets a token of its own to be measured against the
  // stops once the line is known.
  const re = /(\t)|(\s+)|(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1] !== undefined) out.push({ text: '', isSpace: true, tab: true });
    else if (m[2] !== undefined) out.push({ text: m[2], isSpace: true });
    else if (m[3] !== undefined) {
      // CJK runs carry no whitespace, so split them into per-ideograph pieces —
      // each becomes its own box and a wrap opportunity opens between them
      // (see paragraphItemStream). A segment with no CJK is returned unchanged,
      // so Latin tokenization stays byte-identical.
      for (const piece of splitCjkSegment(m[3])) out.push({ text: piece, isSpace: false });
    }
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
// The last Unicode code point of a string (surrogate-pair aware).
function lastCodePoint(s: string): number {
  const chars = [...s];
  return chars.length > 0 ? chars[chars.length - 1]!.codePointAt(0)! : 0;
}

function paragraphItemStream(
  tokens: ReadonlyArray<Token>,
  hyphenator: Hyphenator | undefined,
  // The widest line this paragraph will be broken to, when there is one. A word
  // wider than THAT can never be placed, and has to break inside itself.
  maxLineWidthPt?: number,
): Array<StreamEntry> {
  const entries: Array<StreamEntry> = [];
  // Last code point of the previous text box, for CJK wrap-opportunity detection;
  // reset at spaces / images / math (a break is already present or not wanted).
  let prevBoxEndCp: number | null = null;
  for (const tok of tokens) {
    // §17.3.3.1 `<w:br/>` — a soft line break, which the reader carries as a
    // newline inside the run's text. Left to the whitespace path it became
    // glue, and the line builder drew the glue's own characters: a newline has
    // no glyph, so each break printed as a replacement box. 60329.docx puts 85
    // of them inside its table cells and every paragraph in there ran on with
    // a pair of them mid-sentence. The pair below — stretchy glue, then a
    // forced penalty — is the same one that ends a paragraph.
    const breaks = tok.isSpace ? (tok.text.match(/\r\n|[\r\n]/gu)?.length ?? 0) : 0;
    if (breaks > 0) {
      // One break per newline, so the two the author typed leave the blank
      // line between paragraphs that they asked for.
      for (let i = 0; i < breaks; i++) {
        entries.push({ item: { type: 'glue', width: 0, stretch: 1e6, shrink: 0 }, token: null });
        entries.push({
          item: { type: 'penalty', width: 0, penalty: FORCED_BREAK, flagged: false },
          token: null,
        });
      }
      prevBoxEndCp = null;
      continue;
    }
    if (tok.isSpace || tok.kind === 'image' || tok.kind === 'math') {
      // Spaces are glue; images and math boxes are atomic (un-hyphenatable) boxes.
      entries.push({
        item: tok.isSpace
          ? {
              type: 'glue',
              width: tok.widthPt,
              stretch: tok.widthPt * 0.6,
              shrink: tok.widthPt * GLUE_SHRINK_RATIO,
            }
          : { type: 'box', width: tok.widthPt },
        token: tok,
      });
      prevBoxEndCp = null;
      continue;
    }
    // Text non-space token. A CJK wrap opportunity opens before it when a wide
    // ideograph sits on either side of the boundary (core/line-breaker/cjk) —
    // mark it with a zero-width, zero-cost penalty. `cjkBreakBetween` is false for
    // non-CJK boundaries, so Latin item streams stay byte-identical.
    const startCp = tok.text.codePointAt(0) ?? 0;
    if (prevBoxEndCp !== null && cjkBreakBetween(prevBoxEndCp, startCp)) {
      entries.push({
        item: { type: 'penalty', width: 0, penalty: 0, flagged: false },
        token: null,
      });
    }
    prevBoxEndCp = lastCodePoint(tok.text);
    // Try hyphenation; if no breaks (or too short), one box covers the whole word.
    const positions = hyphenator ? hyphenator.hyphenate(tok.text) : [];
    if (positions.length === 0) {
      // …unless the word is wider than the line itself, and so can never be
      // placed: a URL has no spaces and no hyphenation points, and one box for
      // it is a line that overflows by however long the URL is. The cell then
      // clips it and everything past the first box width is gone —
      // no_drawing_patriarch.xlsx keeps a catalogue link in every one of its
      // 7 465 rows and lost all of them. Excel, Calc and every browser break
      // such a word between characters; so do we, at the last one that fits.
      // A tenth of headroom, the same the projection gives its own estimates:
      // a word that overruns its box by a hair is drawn and clipped, as it
      // always was. Without it an 11-digit product code in a column measured
      // for the WORKBOOK's font — narrower than the face we draw with — broke
      // across two lines on every one of no_drawing_patriarch.xlsx's 7 465
      // rows.
      const chunks =
        maxLineWidthPt !== undefined && maxLineWidthPt > 0 && tok.widthPt > maxLineWidthPt * 1.1
          ? splitToWidth(tok, maxLineWidthPt)
          : undefined;
      if (chunks) {
        chunks.forEach((chunk, ci) => {
          if (ci > 0) {
            entries.push({
              item: { type: 'penalty', width: 0, penalty: 0, flagged: false },
              token: null,
            });
          }
          entries.push({ item: { type: 'box', width: chunk.widthPt }, token: chunk });
        });
        continue;
      }
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

/** Word's own default when a document names none (§17.15.1.25): half an inch. */
const DEFAULT_TAB_STOP_PT = 36;

const TAB_LEADER_CHARS: ReadonlyMap<string, string> = new Map([
  ['dot', '.'],
  ['hyphen', '-'],
  ['underscore', '_'],
  ['middleDot', '\u00b7'],
]);

/**
 * Give each tab on a line the width that carries it to its stop, and the
 * leader that fills the way there.
 *
 * §17.3.1.37 — a tab advances to a POSITION, not by a fixed amount, and where
 * that position is depends on how far the line has already run. Measured as
 * ordinary whitespace it collapsed to a space: the page numbers of
 * FDO77715.docx's index sat against their titles with no dot leader between,
 * and its header ran its left and right halves together on one line.
 *
 * A `right`, `center` or `decimal` stop positions the text AFTER the tab, so
 * each is resolved against the segment that follows it.
 *
 * @param tokens   The line's tokens, replaced in place.
 * @param resolved The paragraph's resolved properties: the stops and indents.
 * @param isFirst  Whether this is the paragraph's first line.
 */
function resolveTabs(
  tokens: Array<Token>,
  resolved: ResolvedParagraphProperties,
  isFirst: boolean,
): void {
  const stops = resolved.tabs;
  let x = resolved.indentLeft + (isFirst ? resolved.indentFirstLine : 0);
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) continue;
    if (!(tok.kind === 'text' && tok.tab === true)) {
      x += tok.widthPt;
      continue;
    }
    // The run of tokens up to the next tab: what this stop has to place.
    let segment = 0;
    let decimalAt = 0;
    let seenDecimal = false;
    for (let j = i + 1; j < tokens.length; j++) {
      const next = tokens[j]!;
      if (next.kind === 'text' && next.tab === true) break;
      if (!seenDecimal && next.kind === 'text') {
        const dot = next.text.indexOf('.');
        if (dot >= 0) {
          decimalAt =
            segment + next.font.measure.textWidthPt(next.text.slice(0, dot), next.fontSizePt);
          seenDecimal = true;
        }
      }
      segment += next.widthPt;
    }
    // Two tabs in a row and one stop between them: the first places nothing,
    // so it claims nothing and the stop falls to the second. FDO77715.docx's
    // header is written that way — two tabs against a single right stop at the
    // margin — and letting the first take it drove the title off the page.
    const after = tokens[i + 1];
    if (segment === 0 && after?.kind === 'text' && after.tab === true) {
      tokens[i] = { ...tok, widthPt: 0, text: '' };
      continue;
    }
    const stop = stops.find((sp) => sp.positionPt > x + 0.01);
    const position = stop
      ? stop.positionPt
      : (Math.floor(x / DEFAULT_TAB_STOP_PT) + 1) * DEFAULT_TAB_STOP_PT;
    const alignment = stop?.alignment ?? 'left';
    let target = position;
    if (alignment === 'right') target = position - segment;
    else if (alignment === 'center') target = position - segment / 2;
    else if (alignment === 'decimal') target = position - (seenDecimal ? decimalAt : segment);
    const width = Math.max(0, target - x);
    const leaderChar = stop?.leader === undefined ? undefined : TAB_LEADER_CHARS.get(stop.leader);
    const text = leaderChar === undefined ? '' : fillLeader(tok, leaderChar, width);
    tokens[i] = { ...tok, widthPt: width, text };
    x += width;
  }
}

/** As many leader characters as fit the gap, measured in the tab's own font. */
function fillLeader(tab: TextToken, leader: string, widthPt: number): string {
  const unit = tab.font.measure.textWidthPt(leader, tab.fontSizePt);
  if (!(unit > 0)) return '';
  return leader.repeat(Math.max(0, Math.floor(widthPt / unit)));
}

/**
 * Cut an unbreakable word into the widest pieces that fit `widthPt`, measured
 * with the token's own font. Each piece keeps the run's formatting, so the line
 * builder reassembles them exactly as it does hyphenation fragments.
 *
 * @param tok     The text token.
 * @param widthPt The line width to fit each piece into.
 * @returns The pieces in order, or undefined when even one character overruns
 *          (nothing can be gained by splitting then).
 */
function splitToWidth(tok: Token, widthPt: number): Array<Token> | undefined {
  if (tok.kind !== 'text') return undefined;
  const chars = [...tok.text];
  // A break that leaves two or three characters to the line is not a line
  // break, it is shredding — and a box that narrow is usually a width nobody
  // meant (a nested table's cell measured before its own grid is known). Below
  // that, the word keeps its shape and overflows, as it did before.
  if (
    tok.font.measure.textWidthPt(chars.slice(0, MIN_SPLIT_CHARS).join(''), tok.fontSizePt) > widthPt
  ) {
    return undefined;
  }
  const out: Array<Token> = [];
  let start = 0;
  while (start < chars.length) {
    let end = start + 1;
    let text = chars[start]!;
    let width = tok.font.measure.textWidthPt(text, tok.fontSizePt);
    if (width > widthPt && start === 0 && chars.length === 1) return undefined;
    while (end < chars.length) {
      const next = text + chars[end]!;
      const nextWidth = tok.font.measure.textWidthPt(next, tok.fontSizePt);
      if (nextWidth > widthPt) break;
      text = next;
      width = nextWidth;
      end++;
    }
    out.push({ ...tok, text, widthPt: width });
    start = end;
  }
  return out.length > 1 ? out : undefined;
}

/** The fewest characters a mid-word break may leave on a line. */
const MIN_SPLIT_CHARS = 6;

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
  profile: LayoutProfile,
  allowEmpty = false,
): Line | null {
  let st = start;
  let et = breakIdx;
  // Skip leading nulls, and leading spaces on a CONTINUATION line — where they
  // are the break's own whitespace. On the first line they are the author's:
  // 45540_classic_Header.xlsx indents a footnote's continuation by writing four
  // spaces into the cell (`<t xml:space="preserve">    in September…`), and
  // dropping them put the text flush against the column edge.
  while (
    st < et &&
    (entries[st]!.token === null || (!isFirst && entries[st]!.token?.isSpace === true))
  ) {
    st++;
  }
  // Trim trailing nulls / spaces — except a space that paints. Whitespace is
  // dropped at a line end because it is invisible, and an underlined or struck
  // space is not: Excel and LibreOffice both rule right across the run of them
  // that a header pads its region with (tdf171828.xlsx underlines its title and
  // the 130 spaces after it, which is where the rule across the page comes from).
  const paints = (t: Token | null | undefined): boolean =>
    t !== null &&
    t !== undefined &&
    t.kind === 'text' &&
    (t.resolvedRun.underline !== 'none' || t.resolvedRun.strike);
  while (
    et > st &&
    (entries[et - 1]!.token === null ||
      (entries[et - 1]!.token?.isSpace === true && !paints(entries[et - 1]!.token)))
  ) {
    et--;
  }
  // An empty range is a line with nothing on it. At the end of a paragraph
  // that is the final forced break and there is no line to draw; in the MIDDLE
  // it is the blank line two `<w:br/>` in a row ask for, and dropping it ran
  // the paragraphs of 60329.docx's table cells together.
  if (st >= et && !allowEmpty) return null;

  const lineTokens: Array<Token> = [];
  for (let i = st; i < et; i++) {
    const ft = entries[i]!.token;
    if (ft) lineTokens.push(ft);
  }

  resolveTabs(lineTokens, resolved, isFirst);

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
  // E-PARITY: under a non-default profile, derive the line's natural height and
  // descent from the max of its text-token fonts' vertical metrics.
  const useMetric = profile !== 'ream';
  let metricAscent = 0;
  let metricDescent = 0;
  let metricLineGap = 0;
  for (const t of lineTokens) {
    contentWidth += t.widthPt;
    const sz = tokenLineSize(t);
    if (sz > maxSize) maxSize = sz;
    if (t.kind === 'math') {
      mathAscent = Math.max(mathAscent, t.ascentPt);
      mathDescent = Math.max(mathDescent, t.descentPt);
    } else if (useMetric && t.kind === 'text') {
      const m = fontLeadingPt(t.font.parsed, t.fontSizePt, profile);
      if (m.ascentPt > metricAscent) metricAscent = m.ascentPt;
      if (m.descentPt > metricDescent) metricDescent = m.descentPt;
      if (m.lineGapPt > metricLineGap) metricLineGap = m.lineGapPt;
    }
  }
  const hasMetric = useMetric && metricAscent + metricDescent > 0;
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
    ...(hasMetric
      ? {
          metricHeightPt: metricAscent + metricDescent + metricLineGap,
          metricDescentPt: metricDescent,
        }
      : {}),
  };
}

function wrap(
  tokens: ReadonlyArray<Token>,
  firstLineWidth: number,
  otherWidth: number,
  resolved: ResolvedParagraphProperties,
  hyphenator: Hyphenator | undefined,
  profile: LayoutProfile,
  // Float text wrapping: explicit per-line widths (the last reuses for the
  // tail, the Knuth-Plass convention). Overrides first/other when given.
  lineWidths?: ReadonlyArray<number>,
): Array<Line> {
  if (tokens.length === 0) return [];

  const widths = lineWidths ?? [firstLineWidth, otherWidth];
  // The WIDEST line the paragraph will be broken to: a word that fits one of
  // them can be placed, and only a word too wide for all of them has to break
  // inside itself.
  const entries = paragraphItemStream(tokens, hyphenator, Math.max(0, ...widths));
  const items = entries.map((e) => e.item);
  // E-PARITY FP3: a renderer-compat profile breaks lines greedily (first-fit,
  // like Word/LibreOffice); the default 'ream' keeps Knuth-Plass total-fit.
  const breaks =
    profile === 'ream' ? breakLines(items, widths).breaks : greedyBreakLines(items, widths);

  const buildLines = (at: ReadonlyArray<number>): Array<Line> => {
    const out: Array<Line> = [];
    let start = 0;
    let isFirst = true;
    let lineIdx = 0;
    for (const [i, breakIdx] of at.entries()) {
      const width = lineIdx < widths.length ? widths[lineIdx]! : widths[widths.length - 1]!;
      const line = lineFromRange(
        entries,
        start,
        breakIdx,
        width,
        isFirst,
        resolved,
        profile,
        i < at.length - 1,
      );
      if (line) out.push(line);
      start = breakIdx + 1;
      isFirst = false;
      lineIdx++;
    }
    return out;
  };

  let lines = buildLines(breaks);
  // One word too wide for the measure must not cost the paragraph every OTHER
  // break. Knuth-Plass scores a line by how badly it fits, and a line that
  // cannot stretch — one word alone in a narrow cell — is infinitely loose,
  // while an overfull line clamps to a small fixed badness. So the total-fit
  // answer for a wrapping cell holding "SELF EMPLOYED" in a column that fits
  // neither word was ONE overfull line, spilling across the cells beside it,
  // where breaking after "SELF" costs nothing but white space. First-fit has no
  // such preference: it breaks wherever the next box will not fit.
  //
  // Only when the whole paragraph collapsed to that single line. A longer one
  // that overflows somewhere in the middle made its own trade — a tight line
  // against a gaping one — and keeps it; taking first-fit there would rewrite
  // every break in the paragraph to fix one.
  const spill = (ls: ReadonlyArray<Line>): number =>
    Math.max(0, ...ls.map((l) => l.contentWidthPt - l.availableWidthPt));
  if (profile === 'ream' && lines.length === 1 && spill(lines) > 0.01) {
    const greedy = buildLines(greedyBreakLines(items, widths));
    if (greedy.length > 1 && spill(greedy) < spill(lines) - 0.01) lines = greedy;
  }

  if (lines.length > 0) lines[lines.length - 1]!.isLastInParagraph = true;
  return lines;
}

// E-PARITY: a font's natural ascent/descent/line-gap (Pt) at the given size for
// the renderer-compat profile. 'word' uses the OS/2 win metrics (the GDI cell
// box, no external leading); 'libreoffice' uses hhea, or the OS/2 typo triple
// when the font requests USE_TYPO_METRICS. descender/typoDescent are stored
// negative, so the descent magnitude negates them.
function fontLeadingPt(
  parsed: ParsedTtf,
  sizePt: number,
  profile: LayoutProfile,
): { ascentPt: number; descentPt: number; lineGapPt: number } {
  const s = sizePt / parsed.unitsPerEm;
  const vm = parsed.vmetrics;
  if (profile === 'word') {
    return { ascentPt: vm.winAscent * s, descentPt: vm.winDescent * s, lineGapPt: 0 };
  }
  const asc = vm.useTypoMetrics ? vm.typoAscent : parsed.ascender;
  const desc = vm.useTypoMetrics ? vm.typoDescent : parsed.descender;
  const gap = vm.useTypoMetrics ? vm.typoLineGap : parsed.lineGap;
  return { ascentPt: asc * s, descentPt: -desc * s, lineGapPt: gap * s };
}

function computeLineHeight(line: Line, p: ResolvedParagraphProperties): number {
  const fontSize = line.maxFontSizePt || 12;
  // Natural single-line height: a font-metric value under a layoutProfile
  // (E-PARITY), else Ream's flat 1.2×. The spacing rules below scale it.
  const natural = line.metricHeightPt ?? fontSize * 1.2;
  // A math box straddles the baseline; the line must be at least tall enough to
  // hold its full ascent+descent (plus a little leading).
  const mathH = (line.mathAscentPt ?? 0) + (line.mathDescentPt ?? 0);
  const mathNeed = mathH > 0 ? mathH * 1.05 : 0;
  // §17.3.1.33 — "exact" means exactly that, zero included: a line told to be
  // no tall is no tall. (Absent spacing resolves to 'auto', so this reads only
  // a height the document actually declared.)
  if (p.spacingLineRule === 'exact' && p.spacingLine >= 0) {
    return Math.max(p.spacingLine, mathNeed);
  }
  if (p.spacingLineRule === 'atLeast' && p.spacingLine > 0) {
    return Math.max(natural, p.spacingLine, mathNeed);
  }
  // "Multiple" spacing is defined in 240ths (240 twips = single). Recover the
  // integer twips before dividing — historically this divided the raw int, and
  // (twips*(1/20))/12 differs from twips/240 in the last ulp.
  const lineTwips = Math.round(p.spacingLine * 20);
  const multiple = lineTwips > 0 ? lineTwips / 240 : 1;
  return Math.max(natural * multiple, mathNeed);
}

function lineDescent(line: Line): number {
  const fs = line.maxFontSizePt || 12;
  // Metric descent under a layoutProfile (E-PARITY), else the flat 0.2×.
  const base = line.metricDescentPt ?? fs * 0.2;
  return Math.max(base, line.mathDescentPt ?? 0);
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

  // §17.4.49 `w:tblGrid` is the document's own answer, not a hint. Word writes
  // the grid it last laid the table out at, and every reader — Word reopening
  // the file included — draws it at those widths; autofit only ever recomputes
  // them when the user edits. Refitting to the natural width of the text
  // instead shrank a table to its words: 2_table_doc.docx declares two 221pt
  // columns across the page and we drew a 100pt box in the corner.
  //
  // The measured autofit stays for a table that brings no grid at all.
  if (gridColTwips(table).some((w) => w > 0)) {
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

/**
 * The above-neighbour a cell resolves its top edge against.
 *
 * A spanning cell covers several columns and paints ONE top across all of them,
 * so taking the first column's neighbour throws away any rule the others
 * declare. tdf100034.xlsx rules the bottom of its two header cells, and a label
 * overflowing from two columns to their left spanned over both and wiped that
 * rule off the page. Take the heaviest bottom under the span instead — it is
 * the one edge the span can draw, so it should be the strongest claim on it.
 */
function aboveOfSpan(
  aboveBordersByCol: ReadonlyArray<CellBorders | undefined>,
  colIdx: number,
  span: number,
): CellBorders | undefined {
  const first = aboveBordersByCol[colIdx];
  if (span <= 1) return first;
  let bottom = first?.bottom;
  for (let k = 1; k < span; k++) {
    const b = aboveBordersByCol[colIdx + k]?.bottom;
    if (heavierBorder(bottom, b) === b) bottom = b;
  }
  if (bottom === first?.bottom) return first;
  return { ...first, ...(bottom ? { bottom } : {}) };
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
    // A text-overflow span is wider than the box the cell's own fill and rules
    // belong in (CellProperties.paintColumns).
    const paintSpan = Math.min(span, Math.max(1, cell.properties.paintColumns ?? span));
    let paintWidthPt = 0;
    for (let k = 0; k < paintSpan && colIdx + k < columnWidthsPt.length; k++) {
      paintWidthPt += columnWidthsPt[colIdx + k]!;
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
      aboveOfSpan(aboveBordersByCol, colIdx, span),
    );
    cells.push(paintWidthPt < widthPt ? { ...cellLayout, paintWidthPt } : cellLayout);
    cursorX += widthPt;
    colIdx += span;
  }
  let heightPt = 0;
  for (const c of cells) if (c.totalHeightPt > heightPt) heightPt = c.totalHeightPt;
  if (row.properties.height && row.properties.heightRule === 'exact') {
    heightPt = row.properties.height;
    // A row pinned to a height keeps it, and what does not fit is CUT — Excel
    // and Word both clip a cell to its row rather than run it over the row
    // below. Drawing the overflow instead put tdf118668.xlsx's second lines on
    // top of the next row's text: a form of 9.75pt rows whose labels wrap.
    return {
      heightPt,
      cells: cells.map((c) => clipCellToHeight(c, heightPt)),
      columnXOffsets,
      rowIdx,
      rowCount,
    };
  } else if (row.properties.height && row.properties.heightRule === 'atLeast') {
    heightPt = Math.max(heightPt, row.properties.height);
  }
  return { heightPt, cells, columnXOffsets, rowIdx, rowCount };
}

/**
 * A cell with only the lines that fit inside `heightPt`.
 *
 * Whole lines only: half a line of text reads as a rendering fault, where a
 * missing one reads as the clipping it is. A cell that already fits, or whose
 * first line does not, is returned untouched — the first line always draws, as
 * it does in Excel.
 *
 * @param cell     The laid-out cell.
 * @param heightPt The row's pinned height.
 * @returns The cell, clipped.
 */
function clipCellToHeight(cell: CellLayout, heightPt: number): CellLayout {
  const room = heightPt - cell.padTopPt - cell.padBottomPt;
  if (cell.lines.length <= 1 || cell.contentHeightPt <= room) return cell;
  const kept: Array<Line> = [];
  let used = 0;
  for (const line of cell.lines) {
    const h = computeLineHeight(line, line.resolved);
    if (kept.length > 0 && used + h > room) break;
    kept.push(line);
    used += h;
  }
  return { ...cell, lines: kept, contentHeightPt: used };
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
  const padLeftBase = cellMar?.left ?? tableMar?.left ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
  const padRightBase = cellMar?.right ?? tableMar?.right ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
  // A conditional-format icon (E-SHEET SC1c) reserves a left gutter; the cell's
  // text is inset past it so the glyph and the value never overlap.
  const padLeftPt = padLeftBase + (cell.properties.icon ? CF_ICON_GUTTER_PT : 0);
  // A data-validation dropdown is NOT laid out. The arrow is a selection
  // affordance: it appears when a cell is selected and neither Excel nor
  // LibreOffice ever prints it. Drawing it put a grey button inside printed
  // table cells, and its gutter and minimum height pushed tdf58243.xlsx onto a
  // third page the reference fits in two. CellProperties.dropdown survives for
  // the HTML writer, which renders an interactive view rather than a page.
  const padRightPt = padRightBase;

  const innerWidth = Math.max(1, widthPt - padLeftPt - padRightPt);
  const lines: Array<Line> = [];
  const nestedTables: Array<TableBlock> = [];
  let contentHeightPt = 0;
  let clipped = false;
  // Continuation cells (vMerge=continue) render no content — their text lives
  // in the 'start' cell above.
  if (mergeRole !== 'middle' && mergeRole !== 'end') {
    for (const el of cell.content) {
      if (el.kind === 'paragraph') {
        // A spreadsheet cell that does not ask to wrap never does: the text runs
        // as far as its box allows and the rest is cut. Only the layout knows
        // where that falls — it has the font metrics and the resolved width — so
        // the projection marks the cell and the layout does the cutting.
        //
        // It has to cut at the GLYPH, though. Taking the line breaker's first
        // line instead looks equivalent and is not: the breaker only breaks at
        // spaces, so a string a hair too wide loses its whole last word rather
        // than its last letter. tdf82984's "Carta geologica - litologica" came
        // out as "Carta" where every other reader shows "Carta geol" — 22pt of
        // white space in a column whose neighbours run full. So measure the
        // whole line, then trim it to the box.
        const noWrap = cell.properties.noWrap === true;
        const block = layoutParagraphBlock(
          el.paragraph,
          options,
          fontResources,
          imageResources,
          noWrap ? NO_WRAP_MEASURE_WIDTH : innerWidth,
        );
        // Up to the cell's right EDGE, not to its inner box: the right padding
        // keeps wrapped text off the border, but a clip is the border. Excel and
        // LibreOffice both run the text to it — measured against the inner box
        // instead, tdf167019's dates lost their last digit and printed
        // "25/10/201".
        const cutAt = widthPt - padLeftPt;
        if (noWrap && (block.lines[0]?.contentWidthPt ?? 0) > cutAt) clipped = true;
        const keep = noWrap
          ? clipLineToWidth(block.lines, cutAt, cell.properties.hashOnOverflow === true)
          : block.lines;
        for (const line of keep) {
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
    ...(cell.properties.dataBar ? { dataBar: cell.properties.dataBar } : {}),
    ...(cell.properties.icon ? { icon: cell.properties.icon } : {}),
    ...(cell.properties.sparkline ? { sparkline: cell.properties.sparkline } : {}),
    lines,
    ...(nestedTables.length > 0 ? { nestedTables } : {}),
    contentHeightPt,
    totalHeightPt,
    ...(cell.properties.verticalAlign ? { verticalAlign: cell.properties.verticalAlign } : {}),
    colStart,
    colSpan,
    mergeRole,
    ...(clipped ? { clipped: true } : {}),
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
  dashDot: 2,
  dashDotDot: 2,
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

// Resolve the structure nodes for a list-item paragraph, growing/shrinking
// the open-list stack by nesting level (w:ilvl). Each item is L → LI →
// [Lbl +] LBody → P; a deeper level opens a nested L inside the parent item's
// LBody. `wantLbl` creates the marker's Lbl element (§14.8.4.3.3) — requested
// only when the first line will actually split its marker tokens out.
function listItemParagraphNode(
  builder: StructTreeBuilder,
  stack: Array<ListFrame>,
  list: { numId: string; level: number },
  wantLbl: boolean,
): { pId: number; lblId?: number } {
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
  const lblId = wantLbl ? builder.create('Lbl', li).id : undefined;
  const lbody = builder.create('LBody', li);
  frame.lbody = lbody;
  return {
    pId: builder.create('P', lbody).id,
    ...(lblId !== undefined ? { lblId } : {}),
  };
}

/**
 * A9 (oop-design §4.3): the page-assembly state machine, extracted verbatim from
 * `paginateSections`. One instance assembles all pages of a document: fields are
 * the former local state, methods the former closures — the block loop drives it.
 * Push order into `current`/`pages` IS the byte order of the emitted PDF, so the
 * bodies move unchanged.
 */
class PageAssembler {
  /**
   * @param sectionCtxs      The per-section render contexts, in order; assembly
   *                         starts on the first section's first page.
   * @param builder          The tagged-PDF structure builder, or `undefined` for
   *                         untagged output.
   * @param notes            The footnote plan (lazy per-section layout), or
   *                         `undefined` when there are none.
   * @param bookmarkPositions Out-param: bookmark name → its GoTo destination.
   */
  constructor(
    readonly sectionCtxs: ReadonlyArray<SectionRenderCtx>,
    readonly builder: StructTreeBuilder | undefined,
    readonly notes: NotePlan | undefined,
    readonly bookmarkPositions: Map<string, BookmarkPosition> | undefined,
  ) {
    this.ctx = sectionCtxs[0]!;
    this.cursorY = this.ctx.pageHeight - this.ctx.marginTop;
    this.colStartY = this.cursorY;
    this.bandTopY = this.cursorY;
  }

  readonly pages: Array<LaidOutPage> = [];

  ctx: SectionRenderCtx;
  secIdx = 0;
  pageInSection = 0;
  globalPageIdx = 0;
  current: Array<PageItem> = [];
  pendingPageBreak = false;
  cursorY: number;

  /**
   * §17.6.4 multi-column flow: content fills column after column before the page
   * flushes. `colStartLen` marks where the current column's items begin in
   * `current` — the single-column "page has content" guard generalizes to "this
   * column has content" (identical when there is one column).
   */
  colIdx = 0;
  colStartLen = 0;
  /** The cursor's y at the top of the current column — see {@link colHasContent}. */
  colStartY: number;
  /**
   * Where this page's column band begins. The top margin on an ordinary page;
   * further down when a §17.6.22 `continuous` section started mid-page, whose
   * columns run beside each other from there, not from the paper's top.
   */
  bandTopY: number;
  /** The current column's left edge (`marginLeft` plus the column x-offset). */
  colLeft = (): number => this.ctx.marginLeft + (this.ctx.columns?.[this.colIdx]?.xOffsetPt ?? 0);
  /** The current column's width (the section content width when single-column). */
  colWidth = (): number => this.ctx.columns?.[this.colIdx]?.widthPt ?? this.ctx.contentWidth;
  /**
   * Whether the current column is already in use.
   *
   * This gates every overflow break, so that a block too tall for an empty page
   * is placed rather than looping forever. It used to ask whether the column had
   * received any drawable ITEM, which is not the same question: a run of empty
   * table rows consumes vertical space and draws nothing, so the guard stayed
   * false, no break ever fired, and the rows marched off the bottom of the page.
   * A spreadsheet is full of such rows — the sheet in tdf171828.xlsx has 162 of
   * them — and they were silently costing whole pages of pagination.
   *
   * Consumed space counts as use, whether or not any ink went with it.
   */
  colHasContent = (): boolean =>
    this.current.length > this.colStartLen || this.cursorY < this.colStartY;
  /**
   * Whether this page has ANYTHING on it — including drawings, which is the
   * whole point: a page carrying only floats has an empty `current`, so a page
   * break asked for after it was silently dropped and the next band's drawings
   * landed on top of the first's. tdf111980_radioButtons.xlsx has no cells at
   * all and eleven controls wider than its page.
   */
  pageHasContent = (): boolean =>
    this.current.length > 0 || this.floatsBehind.length > 0 || this.floatsFront.length > 0;
  /** Overflow step: next column on this page, or a fresh page after the last. */
  advanceColumn = (): void => {
    if (this.ctx.columns && this.colIdx + 1 < this.ctx.columns.length) {
      this.colIdx++;
      this.colStartLen = this.current.length;
      this.cursorY = this.bandTopY;
      this.colStartY = this.cursorY;
    } else {
      this.flushPage();
    }
  };

  /**
   * §17.6.22 — start a `continuous` section on the page in hand: its columns
   * begin where the section does, not at the top of the paper.
   *
   * @param next The section to continue into.
   */
  startBandSection = (next: SectionRenderCtx): void => {
    this.secIdx++;
    this.ctx = next;
    this.colIdx = 0;
    this.colStartLen = this.current.length;
    // The section's own top margin still governs an empty page — it is the
    // first thing on the paper — but past that the cursor is where it is.
    this.cursorY = Math.min(this.cursorY, next.pageHeight - next.marginTop);
    this.bandTopY = this.cursorY;
    this.colStartY = this.cursorY;
  };

  /**
   * §20.4.2.3 out-of-flow drawings (wrap `'none'`): they render at their anchored
   * position without moving the cursor. `behindDoc` sinks below the body text,
   * everything else above it; both flush with the page.
   */
  floatsBehind: Array<PageItem> = [];
  floatsFront: Array<PageItem> = [];
  /**
   * Side-wrapping floats (wrapSquare/tight/through): rectangles the body text
   * must flow around. Page-scoped, like the float graphics.
   */
  exclusions: Array<{ x0: number; x1: number; topYUp: number; bottomYUp: number }> = [];
  /** The float's left edge on the page, resolved from its horizontal anchor. */
  floatX = (f: FloatAnchor, widthPt: number): number => {
    const h = f.posH;
    if (!h) return this.colLeft();
    const base =
      h.relativeFrom === 'page'
        ? 0
        : h.relativeFrom === 'column'
          ? this.colLeft()
          : this.ctx.marginLeft;
    const span =
      h.relativeFrom === 'page'
        ? this.ctx.pageWidth
        : h.relativeFrom === 'column'
          ? this.colWidth()
          : this.ctx.contentWidth;
    if (h.align === 'center') return base + (span - widthPt) / 2;
    if (h.align === 'right') return base + span - widthPt;
    return base + (h.offsetPt ?? 0);
  };
  /**
   * The drawing's TOP in the y-up cursor frame. Paragraph/line-relative offsets
   * hang off the anchoring paragraph's current position.
   */
  floatTopYUp = (f: FloatAnchor): number => {
    const v = f.posV;
    if (!v) return this.cursorY;
    if (v.relativeFrom === 'page') return this.ctx.pageHeight - (v.offsetPt ?? 0);
    if (v.relativeFrom === 'margin')
      return this.ctx.pageHeight - this.ctx.marginTop - (v.offsetPt ?? 0);
    return this.cursorY - (v.offsetPt ?? 0);
  };
  /**
   * Tagged PDF: the stack of open list levels (for L/LI/LBody nesting). Cleared
   * whenever a non-list-item block interrupts the run of list paragraphs.
   */
  readonly listStack: Array<ListFrame> = [];

  /**
   * §17.11 footnotes: `pageNotes` reserved for the CURRENT page (greedy — a line
   * carrying a reference pulls its note's height out of the page bottom, so the
   * line and its note land together). `placedNotes` is global: a note renders
   * once, on its first reference's page.
   */
  pageNotes: Array<{ n: number; blocks: ReadonlyArray<LaidOutBlock>; heightPt: number }> = [];
  noteReserve = 0;
  readonly placedNotes = new Set<string>();
  /** The page's usable bottom: the margin plus whatever the notes band has claimed so far. */
  bottomLimit = (): number => this.ctx.marginBottom + this.noteReserve;

  /**
   * The first exclusion a line spanning `[yTop-h, yTop]` collides with in the
   * current column (x-overlap with the column required).
   */
  exclusionAt = (
    yTop: number,
    h: number,
  ): { x0: number; x1: number; topYUp: number; bottomYUp: number } | undefined => {
    const cl = this.colLeft();
    const cr = cl + this.colWidth();
    for (const e of this.exclusions) {
      if (yTop - h >= e.topYUp || yTop <= e.bottomYUp) continue; // no y overlap
      if (e.x1 <= cl || e.x0 >= cr) continue; // outside this column
      return e;
    }
    return undefined;
  };

  /**
   * Per-line geometry beside an exclusion: the narrowed width and the x shift.
   * Text goes on the WIDER side of the float (one side per line, v1).
   */
  lineGeometryAt = (yTop: number, h: number): { width: number; xOffset: number } => {
    const full = this.colWidth();
    const e = this.exclusionAt(yTop, h);
    if (!e) return { width: full, xOffset: 0 };
    const cl = this.colLeft();
    const leftRoom = e.x0 - FLOAT_TEXT_GAP - cl;
    const rightRoom = cl + full - (e.x1 + FLOAT_TEXT_GAP);
    if (rightRoom >= leftRoom) {
      const w = Math.max(MIN_WRAP_WIDTH, rightRoom);
      return { width: w, xOffset: full - w };
    }
    return { width: Math.max(MIN_WRAP_WIDTH, leftRoom), xOffset: 0 };
  };

  /**
   * Estimated per-line widths for a paragraph starting at `startY`: narrowed
   * while lines (estimated at the first line's height) overlap an exclusion, then
   * one full width Knuth-Plass reuses for the tail. `undefined` when nothing
   * overlaps — the caller keeps the original block.
   */
  lineWidthsFor = (
    block: { readonly lines: ReadonlyArray<Line>; readonly resolved: ResolvedParagraphProperties },
    startY: number,
  ): ReadonlyArray<number> | undefined => {
    if (this.exclusions.length === 0 || block.lines.length === 0) return undefined;
    const h0 = computeLineHeight(block.lines[0]!, block.resolved);
    if (h0 <= 0) return undefined;
    const widths: Array<number> = [];
    let narrowed = false;
    let y = startY;
    for (let i = 0; i < 200; i++) {
      const g = this.lineGeometryAt(y, h0);
      widths.push(g.width);
      if (g.width < this.colWidth()) narrowed = true;
      else if (i > 0) break; // first full-width line after the float ends the scan
      y -= h0;
      if (y < this.bottomLimit()) break;
    }
    return narrowed ? widths : undefined;
  };

  /**
   * New (unplaced) footnotes referenced by a line's tokens, with their layout at
   * the current section's width.
   */
  lineFootnotes = (
    line: Line,
  ): Array<{ id: string; n: number; blocks: ReadonlyArray<LaidOutBlock>; heightPt: number }> => {
    if (!this.notes) return [];
    const out: Array<{
      id: string;
      n: number;
      blocks: ReadonlyArray<LaidOutBlock>;
      heightPt: number;
    }> = [];
    for (const tok of line.tokens) {
      if (tok.kind !== 'text' || tok.footnoteRef === undefined) continue;
      const id = tok.footnoteRef;
      if (this.placedNotes.has(id) || out.some((o) => o.id === id)) continue;
      const laid = this.notes.layout(this.ctx, id);
      const n = this.notes.numbers.get(id);
      if (!laid || n === undefined) continue;
      out.push({ id, n, blocks: laid.blocks, heightPt: laid.heightPt });
    }
    return out;
  };

  /**
   * The notes band for the flushing page: separator rule + each note's blocks
   * stacked inside the reserved area. Tagged: each note is a Note→P element.
   */
  renderNotesBand = (): Array<PageItem> => {
    if (this.pageNotes.length === 0) return [];
    const out: Array<PageItem> = [];
    const top = this.ctx.marginBottom + this.noteReserve; // y-up top of the reserve
    out.push({
      type: 'fill',
      x: pt(this.ctx.marginLeft),
      y: pt(this.ctx.pageHeight - (top - FOOTNOTE_RULE_GAP_ABOVE)),
      width: pt(FOOTNOTE_RULE_WIDTH),
      height: pt(FOOTNOTE_RULE_PT),
      fillColorHex: '000000',
    });
    let cursor = top - FOOTNOTE_SEPARATOR_HEIGHT;
    for (const note of this.pageNotes.sort((a, b) => a.n - b.n)) {
      let structId: number | undefined;
      if (this.builder) {
        const noteNode = this.builder.create('Note', this.builder.root);
        structId = this.builder.create('P', noteNode).id;
      }
      out.push(
        ...drawBlocksSequentially(
          note.blocks,
          this.ctx.marginLeft,
          cursor,
          this.ctx.pageHeight,
          this.ctx.contentWidth,
          structId,
        ),
      );
      cursor -= note.heightPt;
    }
    return out;
  };

  /**
   * Dynamic PAGE/NUMPAGES bands re-render after pagination (both numbers are
   * known only then); each use records where its commands must be spliced.
   */
  readonly dynBands: Array<{
    pageIdx: number;
    pageNumber: number;
    position: 'header' | 'footer';
    render: (pageNumber: number, totalPages: number) => Array<PageItem>;
  }> = [];

  /**
   * Finish the in-progress page: pick the header/footer band, queue any dynamic
   * band, push the assembled commands as a page, then reset all per-page state
   * (cursor, floats, exclusions, notes, columns) for the next one.
   *
   * @param force Emit a page even when no body content was placed (used to
   *              guarantee one page for a header/footer-only document).
   */
  flushPage = (force = false): void => {
    // A page counts as used when it holds drawn items OR when something consumed
    // vertical space on it. Testing only for items discarded pages made of empty
    // rows — and, because the early return also skips the cursor reset below,
    // let the cursor keep descending so every later row piled onto the same
    // page. See colHasContent for the same conflation on the other side.
    // …and drawings count. A page whose only content floats — a sheet of form
    // controls with no cells at all — has an empty `current` and an untouched
    // cursor, so it was discarded and the next band's drawings piled onto the
    // page before it.
    if (!this.pageHasContent() && this.cursorY >= this.colStartY && !force) return;
    const band = bandForPage(
      this.pageInSection,
      this.globalPageIdx,
      this.ctx.titlePg,
      this.ctx.evenAndOddHeaders,
    );
    const header = pickBand(this.ctx.headerSet, band);
    const footer = pickBand(this.ctx.footerSet, band);
    if (header.renderDynamic) {
      this.dynBands.push({
        pageIdx: this.pages.length,
        pageNumber: this.globalPageIdx + 1,
        position: 'header',
        render: header.renderDynamic,
      });
    }
    if (footer.renderDynamic) {
      this.dynBands.push({
        pageIdx: this.pages.length,
        pageNumber: this.globalPageIdx + 1,
        position: 'footer',
        render: footer.renderDynamic,
      });
    }
    this.pages.push({
      commands: [
        ...header.commands,
        ...this.floatsBehind,
        ...this.current,
        ...this.floatsFront,
        ...this.renderNotesBand(),
        ...footer.commands,
      ],
      width: pt(this.ctx.pageWidth),
      height: pt(this.ctx.pageHeight),
    });
    this.current = [];
    this.floatsBehind = [];
    this.floatsFront = [];
    this.exclusions = [];
    this.pageNotes = [];
    this.noteReserve = 0;
    this.colIdx = 0;
    this.colStartLen = 0;
    this.pageInSection++;
    this.globalPageIdx++;
    this.cursorY = this.ctx.pageHeight - this.ctx.marginTop;
    this.colStartY = this.cursorY;
    this.bandTopY = this.cursorY;
  };

  /**
   * Throw the in-progress page away and reset for the next one.
   *
   * A section boundary reached with a page that draws nothing is the one place
   * this is wanted: {@link flushPage} keeps such a page whenever the cursor
   * moved at all, which is right in the middle of a flow and wrong at a
   * boundary, where the space consumed was the section break itself.
   */
  dropPage = (): void => {
    this.current = [];
    this.floatsBehind = [];
    this.floatsFront = [];
    this.exclusions = [];
    this.pageNotes = [];
    this.noteReserve = 0;
    this.colIdx = 0;
    this.colStartLen = 0;
    this.cursorY = this.ctx.pageHeight - this.ctx.marginTop;
    this.colStartY = this.cursorY;
    this.bandTopY = this.cursorY;
  };
}

function paginateSections(
  blocks: ReadonlyArray<LaidOutBlock>,
  sectionCtxs: ReadonlyArray<SectionRenderCtx>,
  builder?: StructTreeBuilder,
  defaultLang = 'en-US',
  notes?: NotePlan,
  // §17.13.6.2 — out-param: bookmark name → its destination (the page and
  // y-up top of the anchoring paragraph's first line).
  bookmarkPositions?: Map<string, BookmarkPosition>,
  // Float text wrapping: re-runs a paragraph's layout with per-line widths.
  reflowParagraph?: (
    paragraph: Paragraph,
    width: number,
    widths: ReadonlyArray<number>,
  ) => ParagraphBlock,
): Array<LaidOutPage> {
  if (sectionCtxs.length === 0) return [];
  const asm = new PageAssembler(sectionCtxs, builder, notes, bookmarkPositions);

  for (let blockIdx = 0; blockIdx < blocks.length; blockIdx++) {
    // Advance to the section that owns this block. A section boundary forces
    // a page break before the next section's first block.
    while (asm.secIdx < sectionCtxs.length - 1 && blockIdx >= asm.ctx.endIndex) {
      const next = sectionCtxs[asm.secIdx + 1]!;
      // §17.6.22 — a `continuous` section starts on the page in hand, where the
      // one before it stopped. Ending the page instead put every section on its
      // own: IndexFieldFlagF.docx writes its index as five continuous sections
      // of differing column counts and printed five pages of one line each,
      // where LibreOffice fits the lot on one. Word makes the break a page
      // break anyway when the paper changes, since a page has one size.
      const staysOnPage =
        next.continuous &&
        next.pageWidth === asm.ctx.pageWidth &&
        next.pageHeight === asm.ctx.pageHeight;
      if (staysOnPage) {
        asm.startBandSection(next);
        continue;
      }
      // Forced: a section owns a page even when the body it holds draws
      // nothing. A title page is written exactly that way — one empty
      // paragraph carrying the sectPr, everything visible living in the
      // header — and letting the page go took the header with it.
      // 090716_Studentische_Arbeit_VWS.docx lost the crest off its first page
      // while keeping it on the rest. Only when the section has produced no
      // page yet, though: past that an empty tail is an empty tail, and the
      // page it would make is thrown away rather than printed blank.
      if (asm.pageHasContent() || asm.pageInSection === 0) asm.flushPage(true);
      else asm.dropPage();
      asm.secIdx++;
      asm.ctx = sectionCtxs[asm.secIdx]!;
      asm.pageInSection = 0;
      asm.cursorY = asm.ctx.pageHeight - asm.ctx.marginTop;
    }

    const block = blocks[blockIdx]!;
    // A forced page break (w:br w:type="page") carried by the previous block:
    // start this block on a fresh page.
    if (asm.pendingPageBreak) {
      asm.pendingPageBreak = false;
      // Forced: a break on a page that holds nothing still ends it. Two breaks
      // in a row are how a document asks for a blank page — 60293.docx does
      // exactly that and prints three, of which we printed two, collapsing the
      // pair into one.
      asm.flushPage(true);
    }
    // A non-list-item block ends any open list run (tagged PDF).
    if (builder && !(block.kind === 'paragraph' && block.list)) asm.listStack.length = 0;
    if (block.kind === 'paragraph') {
      if (block.resolved.pageBreakBefore && asm.pageHasContent()) asm.flushPage();
      asm.cursorY -= block.spacingBeforePt;
      // Float text wrapping: when the paragraph overlaps an exclusion, re-wrap
      // it with per-line widths (the source paragraph re-lays at the column
      // width); the line loop below adds the matching x offsets.
      let pb = block;
      if (reflowParagraph && pb.source && asm.exclusions.length > 0) {
        const widths = asm.lineWidthsFor(block, asm.cursorY);
        if (widths) pb = reflowParagraph(pb.source, asm.colWidth(), widths);
      }
      // Tagged PDF: a plain paragraph → one P (or heading) element; a list item
      // → an L/LI/LBody/P built on the nesting stack. Its lines all reference
      // the resulting leaf by MCID.
      let structId: number | undefined;
      let markerLblId: number | undefined;
      // §14.8.4.3.3 Lbl: the first line's leading marker tokens split into
      // their own marked-content sequence when the geometry is simple — base
      // LTR, left-aligned (the dominant list case). Justified/centered/RTL
      // lines keep the marker inside the P, exactly as before.
      const firstLine = pb.lines[0];
      const splitMarker =
        builder !== undefined &&
        pb.list !== undefined &&
        firstLine !== undefined &&
        pb.resolved.alignment === 'left' &&
        pb.resolved.bidi !== true &&
        !firstLine.tokens.some((t) => t.bidiLevel % 2 === 1) &&
        firstLine.tokens.some((t) => t.kind === 'text' && t.listMarker) &&
        firstLine.tokens.some((t) => !(t.kind === 'text' && t.listMarker));
      if (builder) {
        if (pb.list) {
          const nodes = listItemParagraphNode(builder, asm.listStack, pb.list, splitMarker);
          structId = nodes.pId;
          markerLblId = nodes.lblId;
        } else {
          structId = builder.create(paragraphStructType(pb.resolved), builder.root).id;
        }
        // §14.9.2 per-element /Lang: tag a paragraph whose dominant run language
        // differs from the document default so AT switches pronunciation.
        const lang = dominantParagraphLang(pb.lines);
        if (lang && lang !== defaultLang) builder.node(structId).lang = lang;
      }
      let firstLineOfBlock = true;
      for (const line of pb.lines) {
        const h = computeLineHeight(line, pb.resolved);
        let newNotes = asm.lineFootnotes(line);
        const addedReserve = (sub: typeof newNotes) =>
          sub.reduce((sum, x) => sum + x.heightPt, 0) +
          (asm.pageNotes.length === 0 && sub.length > 0 ? FOOTNOTE_SEPARATOR_HEIGHT : 0);
        if (asm.cursorY - h < asm.bottomLimit() + addedReserve(newNotes) && asm.colHasContent()) {
          asm.advanceColumn();
          newNotes = asm.lineFootnotes(line); // reserve restarts on the fresh page
        }
        if (newNotes.length > 0) {
          asm.noteReserve += addedReserve(newNotes);
          for (const x of newNotes) {
            asm.placedNotes.add(x.id);
            asm.pageNotes.push({ n: x.n, blocks: x.blocks, heightPt: x.heightPt });
          }
        }
        if (firstLineOfBlock) {
          firstLineOfBlock = false;
          if (pb.bookmarks && asm.bookmarkPositions) {
            for (const bookmarkName of pb.bookmarks) {
              if (!asm.bookmarkPositions.has(bookmarkName)) {
                asm.bookmarkPositions.set(bookmarkName, {
                  pageIdx: asm.pages.length,
                  yTopPt: asm.cursorY,
                });
              }
            }
          }
        }
        asm.cursorY -= h;
        const indentLeft =
          pb.resolved.indentLeft + (line.firstLine ? pb.resolved.indentFirstLine : 0);
        const offset = alignmentOffset(
          pb.resolved.alignment,
          line.contentWidthPt,
          line.availableWidthPt,
        );
        const baselineY = pt(asm.ctx.pageHeight - (asm.cursorY + lineDescent(line)));
        const exclusionShift =
          asm.exclusions.length > 0 ? asm.lineGeometryAt(asm.cursorY + h, h).xOffset : 0;
        const originX = asm.colLeft() + indentLeft + offset + exclusionShift;
        if (markerLblId !== undefined && line === firstLine) {
          // Marker glyphs → Lbl, the rest of the line → P. Both segments keep
          // the original baseline; the body segment starts where the marker's
          // advance ends (left-aligned guard ⇒ offset is 0 for both).
          let k = 0;
          let markerWidth = 0;
          while (k < line.tokens.length) {
            const t = line.tokens[k]!;
            if (t.kind !== 'text' || !t.listMarker) break;
            markerWidth += t.widthPt;
            k++;
          }
          asm.current.push({
            type: 'line',
            line: { ...line, tokens: line.tokens.slice(0, k), isLastInParagraph: false },
            originX: pt(originX),
            baselineY,
            structId: markerLblId,
          });
          asm.current.push({
            type: 'line',
            line: { ...line, tokens: line.tokens.slice(k) },
            originX: pt(originX + markerWidth),
            baselineY,
            ...(structId !== undefined ? { structId } : {}),
          });
        } else {
          asm.current.push({
            type: 'line',
            line,
            originX: pt(originX),
            baselineY,
            ...(structId !== undefined ? { structId } : {}),
          });
        }
      }
      asm.cursorY -= pb.spacingAfterPt;
      if (pb.pageBreakAfter) asm.pendingPageBreak = true;
    } else if (block.kind === 'image') {
      const figId = builder ? createFigure(builder, block.altText, 'Image') : undefined;
      const emitImageAt = (x: number, topYUp: number, sink: Array<PageItem>) => {
        sink.push({
          type: 'image',
          x: pt(x),
          y: pt(asm.ctx.pageHeight - topYUp),
          width: pt(block.widthPt),
          height: pt(block.heightPt),
          imageResourceName: block.resourceName,
          ...(block.crop ? { crop: block.crop } : {}),
          ...(figId !== undefined ? { structId: figId } : {}),
        });
      };
      if (isOutOfFlowFloat(block.float)) {
        const fx = asm.floatX(block.float, block.widthPt);
        const fy = asm.floatTopYUp(block.float);
        emitImageAt(fx, fy, block.float.behind ? asm.floatsBehind : asm.floatsFront);
        if (block.float.wrap !== 'none') {
          asm.exclusions.push({
            x0: fx,
            x1: fx + block.widthPt,
            topYUp: fy,
            bottomYUp: fy - block.heightPt,
          });
        }
      } else {
        asm.cursorY -= block.spacingBeforePt;
        if (asm.cursorY - block.heightPt < asm.bottomLimit() && asm.colHasContent())
          asm.advanceColumn();
        asm.cursorY -= block.heightPt;
        emitImageAt(asm.colLeft(), asm.cursorY + block.heightPt, asm.current);
        asm.cursorY -= block.spacingAfterPt;
      }
    } else if (block.kind === 'shape') {
      // Shapes are atomic — never split across asm.pages.
      const figId = builder ? createFigure(builder, block.altText, 'Shape') : undefined;
      const emitShapeAt = (x: number, bottomYUp: number, sink: Array<PageItem>) => {
        const transform = flipTransform(
          buildShapeTransform(
            x,
            bottomYUp,
            block.widthPt,
            block.heightPt,
            block.rotation60k,
            block.flipH,
            block.flipV,
          ),
          asm.ctx.pageHeight,
        );
        sink.push({
          type: 'shape',
          shape: {
            paths: block.paths,
            ...(block.fillColorHex ? { fillColorHex: block.fillColorHex } : {}),
            ...(block.fillGradient ? { fillGradient: block.fillGradient } : {}),
            ...(block.stroke ? { stroke: block.stroke } : {}),
            ...(block.shadow ? { shadow: block.shadow } : {}),
            transform,
          },
          ...(figId !== undefined ? { structId: figId } : {}),
        });
        // Shape text: laid out axis-aligned, anchored vertically within the
        // inset rect, emitted as ordinary line commands so it rides the text
        // pass on top of the fill. (Rotated text boxes keep upright text.)
        if (block.textLines.length > 0) {
          const shapeBottom = bottomYUp;
          const shapeTop = bottomYUp + block.heightPt;
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
            sink.push({
              type: 'line',
              line,
              originX: pt(x + block.insetLeftPt + lineOffset),
              baselineY: pt(asm.ctx.pageHeight - (textY + lineDescent(line))),
              ...(figId !== undefined ? { structId: figId } : {}),
            });
          }
        }
      };
      if (isOutOfFlowFloat(block.float)) {
        const fx = asm.floatX(block.float, block.widthPt);
        const fy = asm.floatTopYUp(block.float);
        emitShapeAt(
          fx,
          fy - block.heightPt,
          block.float.behind ? asm.floatsBehind : asm.floatsFront,
        );
        if (block.float.wrap !== 'none') {
          asm.exclusions.push({
            x0: fx,
            x1: fx + block.widthPt,
            topYUp: fy,
            bottomYUp: fy - block.heightPt,
          });
        }
      } else {
        asm.cursorY -= block.spacingBeforePt;
        if (asm.cursorY - block.heightPt < asm.bottomLimit() && asm.colHasContent())
          asm.advanceColumn();
        asm.cursorY -= block.heightPt;
        const offset = alignmentOffset(block.resolvedAlignment, block.widthPt, asm.colWidth());
        emitShapeAt(asm.colLeft() + offset, asm.cursorY, asm.current);
        asm.cursorY -= block.spacingAfterPt;
      }
    } else if (block.kind === 'chart') {
      // Charts are atomic. Their primitives are in a local y-up frame; the
      // stored transform translates to the chart box's bottom-left (x, y in the
      // internal y-up cursor frame) composed with the page flip. The whole
      // chart is one Figure (alt = its title); its shapes + labels carry that id.
      const figId = builder ? createFigure(builder, block.altText, 'Chart') : undefined;
      const fig = figId !== undefined ? { structId: figId } : {};
      const emitChartAt = (x: number, bottomYUp: number, sink: Array<PageItem>) => {
        for (const s of block.layout.shapes) {
          sink.push({
            type: 'shape',
            shape: {
              paths: s.paths,
              ...(s.fillColorHex ? { fillColorHex: s.fillColorHex } : {}),
              ...(s.stroke ? { stroke: s.stroke } : {}),
              transform: flipTransform([1, 0, 0, 1, x, bottomYUp], asm.ctx.pageHeight),
            },
            ...fig,
          });
        }
        for (const t of block.layout.texts) {
          sink.push({
            type: 'line',
            line: t.line,
            originX: pt(x + t.x),
            baselineY: pt(asm.ctx.pageHeight - (bottomYUp + t.y)),
            ...(t.rotationDeg ? { rotationDeg: t.rotationDeg } : {}),
            ...fig,
          });
        }
      };
      if (isOutOfFlowFloat(block.float)) {
        const fx = asm.floatX(block.float, block.widthPt);
        const fy = asm.floatTopYUp(block.float);
        emitChartAt(
          fx,
          fy - block.heightPt,
          block.float.behind ? asm.floatsBehind : asm.floatsFront,
        );
        if (block.float.wrap !== 'none') {
          asm.exclusions.push({
            x0: fx,
            x1: fx + block.widthPt,
            topYUp: fy,
            bottomYUp: fy - block.heightPt,
          });
        }
      } else {
        asm.cursorY -= block.spacingBeforePt;
        if (asm.cursorY - block.heightPt < asm.bottomLimit() && asm.colHasContent())
          asm.advanceColumn();
        asm.cursorY -= block.heightPt;
        const offset = alignmentOffset(block.resolvedAlignment, block.widthPt, asm.colWidth());
        emitChartAt(asm.colLeft() + offset, asm.cursorY, asm.current);
        asm.cursorY -= block.spacingAfterPt;
      }
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
      const tableX = asm.colLeft() + block.xOffsetPt;
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
          row.heightPt > asm.ctx.pageContentHeight
            ? splitRowIntoChunks(row, asm.ctx.pageContentHeight)
            : [row];

        for (let ci = 0; ci < chunks.length; ci++) {
          const chunk = chunks[ci]!;
          // A manual <rowBreaks> break (first chunk only) forces a new page even
          // when the row would fit; an overflow break starts one when it won't.
          const forcedBreak =
            ci === 0 && row.breakBefore && !isLeadingHeader && asm.current.length > 0;
          const overflow = asm.cursorY - chunk.heightPt < asm.bottomLimit() && asm.colHasContent();
          if (forcedBreak || overflow) {
            if (forcedBreak) asm.flushPage();
            else asm.advanceColumn();
            // Re-emit the header rows on the fresh page (visual repetition →
            // artifacts, no structIds), but only when the breaking row is not
            // itself a header and the row still fits beneath the repeated band.
            if (
              !isLeadingHeader &&
              headerRows.length > 0 &&
              asm.cursorY - headerHeightPt - chunk.heightPt >= asm.bottomLimit()
            ) {
              for (const hr of headerRows) {
                emitRowChunk(
                  asm.current,
                  hr,
                  tableX,
                  asm.cursorY,
                  asm.ctx.pageHeight,
                  colCount,
                  undefined,
                );
                asm.cursorY -= hr.heightPt;
              }
            }
          }
          // A vertical merge's box for vertical alignment: measured only for a
          // row the page takes whole, and only against the room left on the
          // page, so a merge that a break cuts aligns inside the part that is
          // actually here.
          const mergedHeights =
            chunks.length === 1 && row.cells.some((c) => c.mergeRole === 'start')
              ? row.cells.map((c) =>
                  c.mergeRole === 'start'
                    ? mergedBoxHeightPt(block.rows, ri, c.colStart, asm.cursorY - asm.bottomLimit())
                    : undefined,
                )
              : undefined;
          emitRowChunk(
            asm.current,
            chunk,
            tableX,
            asm.cursorY,
            asm.ctx.pageHeight,
            colCount,
            cellStructIds,
            mergedHeights,
          );
          asm.cursorY -= chunk.heightPt;
          // Only the table's last row paints its bottom edge — every other
          // horizontal rule is the next row's top (see emitCellBorders). That
          // convention has no notion of a page: when the next row starts a new
          // page, the edge it owed this one is drawn up there, and the table is
          // left hanging open at the bottom of the page it just left. Look
          // ahead and close it. tdf58243.xlsx shows the seam plainly — the
          // vertical rules run past the last row into white space.
          if (ci === chunks.length - 1 && ri < block.rows.length - 1) {
            const next = block.rows[ri + 1]!;
            const nextHeight = Math.min(next.heightPt, asm.ctx.pageContentHeight);
            const nextForced =
              next.breakBefore && ri + 1 >= headerRows.length && asm.current.length > 0;
            const nextOverflow =
              asm.cursorY - nextHeight < asm.bottomLimit() && asm.colHasContent();
            if (nextForced || nextOverflow) {
              emitRowBottomEdge(asm.current, chunk, tableX, asm.cursorY, asm.ctx.pageHeight);
            }
          }
        }
      }
    }
  }

  // Trailing content on the last in-progress page.
  asm.flushPage();
  // A body that produced no flushable content (e.g. text lives only in the
  // header/footer bands — a header/footer-only document) must still emit one
  // page so those bands render, instead of falling back to a blank page.
  if (asm.pages.length === 0) asm.flushPage(true);

  // Every page exists now, so PAGE and NUMPAGES both have values: render the
  // dynamic bands and splice them where the static band would have sat
  // (header before the body content, footer after).
  for (const d of asm.dynBands) {
    const cmds = d.render(d.pageNumber, asm.pages.length);
    const page = asm.pages[d.pageIdx]!;
    if (d.position === 'header') page.commands.unshift(...cmds);
    else page.commands.push(...cmds);
  }
  for (const page of asm.pages) coalesceVerticalFills(page.commands);
  return asm.pages;
}

// Join a column of same-coloured cell fills into one rectangle.
//
// Runs ACROSS a row are already merged where they are emitted, for the reason
// given there: two rects composited separately share an edge that lands
// mid-pixel, and each side contributes partial coverage where one whole one
// belongs — a block of one colour shows a pale grid of seams exactly on its
// cell boundaries. Down the page each row still painted its own slice and
// leaned on FILL_SEAM_PT, which is a fifteenth of a point: a fifth of a device
// pixel at 300 DPI, not enough to make the boundary pixel opaque. 51710.xlsx
// paints its column A grey down 46 pages and showed a rung at every row.
//
// Rows are emitted in order, so a run is a stretch of consecutive fills with
// the same left edge, width and colour whose boxes meet.
function coalesceVerticalFills(commands: Array<PageItem>): void {
  const touching = (a: FillItem, b: FillItem): boolean =>
    Math.abs(a.x - b.x) < 0.01 &&
    Math.abs(a.width - b.width) < 0.01 &&
    a.fillColorHex === b.fillColorHex &&
    Math.min(a.y + a.height, b.y + b.height) >= Math.max(a.y, b.y) - 0.01;

  const open = new Map<string, FillItem>();
  const dead = new Set<PageItem>();
  for (const item of commands) {
    if (item.type !== 'fill') continue;
    const fill = item as FillItem;
    const k = `${fill.x.toFixed(2)}|${fill.width.toFixed(2)}|${fill.fillColorHex}`;
    const prev = open.get(k);
    if (prev && touching(prev, fill)) {
      const top = Math.max(prev.y + prev.height, fill.y + fill.height);
      const bottom = Math.min(prev.y, fill.y);
      (prev as { y: Pt }).y = pt(bottom);
      (prev as { height: Pt }).height = pt(top - bottom);
      dead.add(item);
      continue;
    }
    open.set(k, fill);
  }
  if (dead.size === 0) return;
  const kept = commands.filter((c) => !dead.has(c));
  commands.length = 0;
  commands.push(...kept);
}

interface FillItem {
  readonly type: 'fill';
  x: Pt;
  y: Pt;
  width: Pt;
  height: Pt;
  readonly fillColorHex: string;
}

// §14.8.5.2 /RowSpan — how many rows a vertical-merge origin spans. Walk down
// from the origin row at the same column, counting vMerge continuation cells
// (mergeRole middle/end) until the group ends. A standalone cell spans 1.
// How tall a vertical merge's box is FROM this row down — the height its own
// row and its continuation rows add up to, stopping at whatever still fits the
// page. A merge cut by a page break has only the part that is on the page, and
// aligning its text inside the whole box would put it past the paper.
function mergedBoxHeightPt(
  rows: ReadonlyArray<RowLayout>,
  startRowIdx: number,
  colStart: number,
  availablePt: number,
): number {
  let height = rows[startRowIdx]!.heightPt;
  for (let r = startRowIdx + 1; r < rows.length; r++) {
    const cont = rows[r]!.cells.find(
      (c) => c.colStart === colStart && (c.mergeRole === 'middle' || c.mergeRole === 'end'),
    );
    if (!cont) break;
    if (height + rows[r]!.heightPt > availablePt) break;
    height += rows[r]!.heightPt;
    if (cont.mergeRole === 'end') break;
  }
  return height;
}

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
  out: Array<PageItem>,
  row: RowLayout,
  marginLeft: number,
  cursorY: number,
  pageHeight: number,
  colCount: number,
  cellStructIds?: ReadonlyArray<number | undefined>,
  mergedHeights?: ReadonlyArray<number | undefined>,
): void {
  const rowTop = cursorY;
  const rowBottom = cursorY - row.heightPt;
  // Cell fills FIRST, and consecutive cells of one colour as a single
  // rectangle. Two abutting fills are composited separately, so a shared edge
  // landing mid-pixel takes each side's partial coverage instead of one whole
  // one, and a block of same-filled cells shows a pale grid of seams exactly
  // where its cell boundaries are. Excel and LibreOffice paint the run once and
  // have no seams to show; so does this now. (Row to row the same edge is still
  // shared, which is what FILL_SEAM_PT covers.)
  {
    let i = 0;
    while (i < row.cells.length) {
      const cell = row.cells[i]!;
      // A vertical merge's continuation rows paint too. They were skipped here
      // because they arrived with no shading of their own, which made a merged
      // block exactly one row tall however many rows it spanned; they now carry
      // the merge's fill, and each row painting its own slice (plus the seam
      // below) gives the block its full height with no join to show for it.
      if (!cell.shadingColorHex) {
        i++;
        continue;
      }
      // A text-overflow span is wider than the cell's own paint box.
      let width = cell.paintWidthPt ?? cell.widthPt;
      let j = i + 1;
      while (j < row.cells.length && cell.paintWidthPt === undefined) {
        const next = row.cells[j]!;
        if (next.mergeRole === 'middle' || next.mergeRole === 'end') break;
        if (next.shadingColorHex !== cell.shadingColorHex) break;
        width += next.paintWidthPt ?? next.widthPt;
        j++;
      }
      out.push({
        type: 'fill',
        x: pt(marginLeft + (row.columnXOffsets[i] ?? 0)),
        y: pt(pageHeight - rowBottom - row.heightPt - FILL_SEAM_PT),
        width: pt(width),
        height: pt(row.heightPt + FILL_SEAM_PT),
        fillColorHex: cell.shadingColorHex,
      });
      i = j;
    }
  }
  for (let i = 0; i < row.cells.length; i++) {
    const cell = row.cells[i]!;
    const cellX = marginLeft + (row.columnXOffsets[i] ?? 0);
    const structId = cellStructIds?.[i];
    // Conditional-format data bar (E-SHEET SC1c): a fraction-width fill over the
    // shading, under the text. Pushed after the shading fill so it paints on top.
    if (cell.dataBar && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      const start = Math.max(0, Math.min(1, cell.dataBar.startFraction ?? 0));
      const barWidth = cell.widthPt * Math.max(0, Math.min(1, cell.dataBar.fraction));
      if (barWidth > 0) {
        // §18.3.1.28 — Excel paints a data bar as a GRADIENT that fades away
        // from the axis the bar grows out of: solid at the axis end, white at
        // the tip. Drawn flat, databar.xlsx's five gauges read as blocks where
        // both references read as bars. A negative bar grows leftwards, so its
        // solid end is its right one.
        const barX = cellX + cell.widthPt * start;
        const barBottomYUp = rowBottom;
        out.push({
          type: 'shape',
          shape: {
            paths: [rectAtPath(0, 0, barWidth, row.heightPt)],
            // The solid approximation writers without gradients paint (and the
            // one PDF/A falls back to) is the bar's own colour.
            fillColorHex: cell.dataBar.colorHex,
            fillGradient: {
              kind: 'linear' as const,
              angle: cell.dataBar.negative ? 180 : 0,
              stops: [
                { offset: 0, colorHex: cell.dataBar.colorHex },
                { offset: 1, colorHex: 'FFFFFF' },
              ],
            },
            transform: flipTransform([1, 0, 0, 1, barX, barBottomYUp], pageHeight),
          },
        });
      }
    }
    // Conditional-format icon (E-SHEET SC1c): a vector glyph in the left gutter,
    // vertically centred in the row. Painted in the shapes pass (over fills,
    // under text); its local y-up frame is flipped onto the page.
    if (cell.icon && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      const iconSize = Math.min(CF_ICON_SIZE_PT, Math.max(0, row.heightPt - 1));
      if (iconSize > 0) {
        const iconX = cellX + (CF_ICON_GUTTER_PT - iconSize) / 2;
        const iconBottomYUp = rowBottom + (row.heightPt - iconSize) / 2;
        const iconTransform = flipTransform([1, 0, 0, 1, iconX, iconBottomYUp], pageHeight);
        for (const prim of buildCellIconShape(cell.icon, iconSize)) {
          out.push({
            type: 'shape',
            shape: {
              paths: prim.paths,
              ...(prim.fillColorHex ? { fillColorHex: prim.fillColorHex } : {}),
              ...(prim.stroke ? { stroke: prim.stroke } : {}),
              transform: iconTransform,
            },
          });
        }
      }
    }
    // Data-validation dropdown (E-SHEET SV1): a small button + ▾ at the cell's
    // right edge, painted in the shapes pass (over fills, under text). The cell
    // reserves a right gutter so it never covers the value.
    if (cell.dropdown && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      const btn = Math.min(DROPDOWN_BUTTON_PT, Math.max(0, row.heightPt - 2));
      if (btn > 1) {
        const btnX = cellX + cell.widthPt - btn - 1.5;
        const btnBottomYUp = rowBottom + (row.heightPt - btn) / 2;
        const btnTransform = flipTransform([1, 0, 0, 1, btnX, btnBottomYUp], pageHeight);
        for (const prim of buildDropdownPrims(btn)) {
          out.push({
            type: 'shape',
            shape: {
              paths: prim.paths,
              ...(prim.fillColorHex ? { fillColorHex: prim.fillColorHex } : {}),
              ...(prim.stroke ? { stroke: prim.stroke } : {}),
              transform: btnTransform,
            },
          });
        }
      }
    }
    // Sparkline (E-SHEET SC2): a mini chart filling the cell's content box,
    // painted in the shapes pass. Its local y-up frame is flipped onto the cell.
    if (
      cell.sparkline &&
      cell.sparkline.values.length > 0 &&
      cell.mergeRole !== 'middle' &&
      cell.mergeRole !== 'end'
    ) {
      const inset = 1.5;
      const sw = cell.widthPt - 2 * inset;
      const sh = row.heightPt - 2 * inset;
      if (sw > 0 && sh > 0) {
        const prims = buildSparkline(
          cell.sparkline.kind,
          cell.sparkline.values,
          sw,
          sh,
          cell.sparkline.colorHex,
        );
        const transform = flipTransform([1, 0, 0, 1, cellX + inset, rowBottom + inset], pageHeight);
        for (const prim of prims) {
          out.push({
            type: 'shape',
            shape: {
              paths: prim.paths,
              ...(prim.fillColorHex ? { fillColorHex: prim.fillColorHex } : {}),
              ...(prim.stroke ? { stroke: prim.stroke } : {}),
              transform,
            },
          });
        }
      }
    }
    emitCellBorders(
      out,
      cell,
      cellX,
      rowBottom,
      row.heightPt,
      pageHeight,
      row.rowIdx,
      row.rowCount,
      colCount,
    );
    // Diagonal cell strokes (Excel diagonal borders): a line across the cell box,
    // drawn in the shapes pass like the icon / dropdown glyphs. diagonalDown runs
    // top-left → bottom-right, diagonalUp bottom-left → top-right.
    const diagDown = cell.borders.diagonalDown;
    const diagUp = cell.borders.diagonalUp;
    if ((diagDown || diagUp) && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      const transform = flipTransform([1, 0, 0, 1, cellX, rowBottom], pageHeight);
      const w = cell.widthPt;
      const h = row.heightPt;
      const pushDiagonal = (b: Border, y0: number, y1: number): void => {
        out.push({
          type: 'shape',
          shape: {
            paths: [new PathBuilder().moveTo(0, y0).lineTo(w, y1).build()],
            stroke: { colorHex: b.colorHex ?? '000000', widthPt: b.width ?? 0.75 },
            transform,
          },
        });
      };
      if (diagDown) pushDiagonal(diagDown, h, 0);
      if (diagUp) pushDiagonal(diagUp, 0, h);
    }
    if (cell.mergeRole === 'middle' || cell.mergeRole === 'end') continue;
    // A box taller than its content: a spreadsheet cell sits at the BOTTOM of
    // it unless it says otherwise, which is what Excel and LibreOffice both do
    // and what a declared row height makes visible. A vertical merge's box is
    // taller than this one row — the caller measures it and passes it down, so
    // 50299.xlsx's `vertical="center"` label sits in the middle of a two-row
    // merge rather than at the top of its first row.
    const boxHeightPt = mergedHeights?.[i] ?? row.heightPt;
    const slack =
      cell.verticalAlign !== undefined
        ? boxHeightPt - cell.padTopPt - cell.contentHeightPt - cell.padBottomPt
        : 0;
    const vOffset =
      slack <= 0
        ? 0
        : cell.verticalAlign === 'bottom'
          ? slack
          : cell.verticalAlign === 'center'
            ? slack / 2
            : 0;
    let textY = rowTop - cell.padTopPt - vOffset;
    for (const line of cell.lines) {
      const h = computeLineHeight(line, line.resolved);
      textY -= h;
      const offset = alignmentOffset(
        line.resolved.alignment,
        line.contentWidthPt,
        cell.widthPt - cell.padLeftPt - cell.padRightPt,
      );
      // A paragraph's own left indent counts inside a cell too. The body path
      // has always added it; this one did not, so §18.8.1's `indent` reached the
      // model and stopped there — 45544.xlsx indents five rows under "Not
      // Seeking Employment" and we drew them flush with it.
      const indentLeft =
        line.resolved.indentLeft + (line.firstLine ? line.resolved.indentFirstLine : 0);
      out.push({
        type: 'line',
        line,
        originX: pt(cellX + cell.padLeftPt + indentLeft + offset),
        baselineY: pt(pageHeight - (textY + lineDescent(line))),
        // The cell's own box. For a vertical merge that is the MERGED height,
        // not this row's: the text is centred over the whole box (see
        // mergedHeights), so a one-row clip would fall entirely above it.
        ...(cell.clipped
          ? {
              clip: {
                x: pt(cellX),
                y: pt(pageHeight - rowTop),
                width: pt(cell.widthPt),
                height: pt(boxHeightPt),
              },
            }
          : {}),
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
          emitRowChunk(out, nrow, nestedX, textY, pageHeight, nt.colCount, nestedIds);
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

// How far a row's fill runs past its own box, to close the seam against the row
// below. A fifteenth of a point: invisible where the colours differ, and the
// difference between a flat field and a chequerboard where they do not.
const FILL_SEAM_PT = 0.07;

// Each shared edge between adjacent cells is the same physical line, so we
// render it exactly once. Convention: every cell paints its top + left; the
// last row paints its bottom and the last spanned column paints its right.
function emitCellBorders(
  out: Array<PageItem>,
  cell: CellLayout,
  cellX: number,
  cellY: number,
  rowHeight: number,
  pageHeight: number,
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
      x: pt(cellX),
      y: pt(pageHeight - cellY - rowHeight),
      // A rule belongs to the CELL, not to the run of empty neighbours its text
      // borrowed. The fill beside it already knows that; the border did not, so
      // 53734.xlsx framed one bold cell in a box twice its width, running out
      // past its own green fill and past where both references close it.
      width: pt(cell.paintWidthPt ?? cell.widthPt),
      height: pt(rowHeight),
      borderSizePt: sz,
      borderColorHex: border.colorHex ?? '000000',
      ...(border.style !== 'single' ? { borderStyle: border.style } : {}),
    });
  };
  pushSide('top', cell.borders.top);
  pushSide('left', cell.borders.left);
  if (rowIdx === rowCount - 1) pushSide('bottom', cell.borders.bottom);
  if (cell.colStart + cell.colSpan - 1 === colCount - 1) pushSide('right', cell.borders.right);
}

// Wide enough that the line breaker never breaks: a non-wrapping cell is laid
// out as one line and then cut to its box. Not Infinity — the breaker does
// arithmetic with this width.
const NO_WRAP_MEASURE_WIDTH = 1e6;

/**
 * One line, trimmed to the glyph that STRADDLES `widthPt` — what a non-wrapping
 * spreadsheet cell shows. Excel and LibreOffice paint that glyph and cut it
 * through the middle, and so do we now: the caller marks the cell and the
 * emitters clip the line to its box (TextLineItem.clip). Stopping at the last
 * glyph that wholly fits is what dropped a character off every cut cell.
 */
function clipLineToWidth(
  lines: ReadonlyArray<Line>,
  widthPt: number,
  hashOnOverflow = false,
): Array<Line> {
  const line = lines[0];
  if (!line) return [];
  if (line.contentWidthPt <= widthPt) return [line];
  // A number under its own format is never shown truncated — see
  // CellProperties.hashOnOverflow.
  if (hashOnOverflow) return [hashFill(line, widthPt)];
  const tokens: Array<Token> = [];
  let used = 0;
  for (const token of line.tokens) {
    const advance = token.widthPt;
    if (used + advance <= widthPt) {
      tokens.push(token);
      used += advance;
      continue;
    }
    // The token that straddles the edge: keep the characters that fit. Only
    // text can be cut — an image or a formula is atomic and simply drops.
    if (token.kind === 'text') {
      const room = widthPt - used;
      let text = '';
      let textWidth = 0;
      for (const ch of token.text) {
        if (textWidth >= room) break;
        text += ch;
        textWidth = token.font.measure.textWidthPt(text, token.fontSizePt);
      }
      if (text.length > 0) {
        tokens.push({ ...token, text, widthPt: textWidth });
        used += textWidth;
      }
    }
    break;
  }
  return [{ ...line, tokens, contentWidthPt: used }];
}

/**
 * The line as a row of `#` filling `widthPt` — what Excel and LibreOffice put in
 * a cell whose number does not fit. At least one, so the cell is never blank:
 * a column too narrow even for a single `#` still has to say something.
 */
function hashFill(line: Line, widthPt: number): Line {
  const first = line.tokens.find((t): t is TextToken => t.kind === 'text');
  if (!first) return line;
  const one = first.font.measure.textWidthPt('#', first.fontSizePt);
  const count = one > 0 ? Math.max(1, Math.floor(widthPt / one)) : 1;
  const text = '#'.repeat(count);
  const token: TextToken = {
    ...first,
    text,
    widthPt: first.font.measure.textWidthPt(text, first.fontSizePt),
  };
  return { ...line, tokens: [token], contentWidthPt: token.widthPt };
}

/**
 * The bottom edge of a row that ends a page — the rule the row below it owed it
 * and drew on the next page instead. See the call site: a table that spills over
 * a break is closed on both sides of it, the way Excel and Word print one.
 *
 * A cell whose bottom is suppressed (a vertical merge continuing past the break)
 * stays open; inventing an edge there would cut the merge in half.
 */
function emitRowBottomEdge(
  out: Array<PageItem>,
  row: RowLayout,
  marginLeft: number,
  rowBottom: number,
  pageHeight: number,
): void {
  for (let i = 0; i < row.cells.length; i++) {
    const cell = row.cells[i]!;
    const border = cell.borders.bottom;
    if (!border || border.style === 'none') continue;
    out.push({
      type: 'border',
      side: 'bottom',
      x: pt(marginLeft + (row.columnXOffsets[i] ?? 0)),
      y: pt(pageHeight - rowBottom - row.heightPt),
      width: pt(cell.widthPt),
      height: pt(row.heightPt),
      borderSizePt: border.width ?? DEFAULT_BORDER_SIZE_EIGHTH * EIGHTH_PT,
      borderColorHex: border.colorHex ?? '000000',
      ...(border.style !== 'single' ? { borderStyle: border.style } : {}),
    });
  }
}
