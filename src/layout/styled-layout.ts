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
  DocumentInfo,
  FloatAnchor,
  HeaderFooterReference,
  HeaderFooterType,
  ImageBlock,
  MathNode,
  Numbering,
  Paragraph,
  Run,
  Section,
  SectionColumns,
  SectionProperties,
  ShapeBlock,
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
import type { ResourceId } from '@/core/ir';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { StrokeStyle, VectorPath } from '@/core/vector';
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
import { FORCED_BREAK, breakLines } from '@/core/line-breaker';
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
import { buildChartScene } from '@/core/drawingml/chart-geometry';
import { layoutMath, mathGlyphSegments, variantStyle } from '@/layout/math-layout';
import { rectPath } from '@/core/drawingml/preset-geometry';
import {
  DEFAULT_INSET_LR_PT,
  DEFAULT_INSET_TB_PT,
  buildShapePaths,
  buildShapeTransform,
  buildStroke,
} from '@/core/drawingml/shape-render';
import { StructTreeBuilder } from '@/pdf/struct-tree';

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
  // §17.11 notes content by id. Footnotes render in a reserved band at the
  // bottom of the referencing page; endnotes flow after the body.
  readonly footnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  readonly endnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
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
  // PDF/UA-1 (ISO 14289-1): implies tagged; the XMP carries pdfuaid:part=1
  // and the document always gets a title (AT announces it). Combines freely
  // with pdfA level-a profiles.
  readonly pdfUA?: boolean;
  // Document natural language (BCP 47, e.g. "en-US", "ru-RU") for the tagged-PDF
  // catalog /Lang (§14.9.2). Defaults to "en-US". The docx converter fills this
  // from the document's default w:lang.
  readonly language?: string;
  // §7.6 PDF encryption (AES-256, R6). Only honoured on the ASYNC conversion
  // path (WebCrypto); mutually exclusive with pdfA (ISO 19005 forbids
  // /Encrypt) and with signatures (v1).
  readonly encrypt?: PdfEncryptOptions;
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

// PDF-only companion the same layout pass produces (oop-design A13): the
// logical-structure tree, per-section geometry (the emit fallback page), and
// the parsed PDF/A profile. Consumed only by emitStyledPdf; the SVG writer
// never sees it.
// §17.13.6.2 — a bookmark's GoTo destination: the page (0-based) and the
// y-up top of the anchoring paragraph's first line.
export interface BookmarkPosition {
  readonly pageIdx: number;
  readonly yTopPt: number;
}

export interface PdfLayoutAux {
  readonly structBuilder: StructTreeBuilder | undefined;
  readonly sectionCtxs: ReadonlyArray<SectionRenderCtx>;
  readonly pdfaProfile: PdfAProfile | undefined;
  readonly tagged: boolean;
  readonly bookmarks: ReadonlyMap<string, BookmarkPosition>;
}

// What layoutStyledDocument actually returns: the PageDoc with the PDF
// companion riding on `pdf`. Assignable to the narrow LaidOutDocument, so
// PageDoc-only consumers (writeSvg) take it as-is.
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
function assignNoteNumbers(body: ReadonlyArray<BodyElement>): {
  body: ReadonlyArray<BodyElement>;
  footnotes: ReadonlyMap<string, number>;
  endnotes: ReadonlyMap<string, number>;
  // Footnote ids whose references sit OUTSIDE top-level paragraphs (table
  // cells, shape text): greedy bottom-of-page placement only tracks paragraph
  // lines, so these notes flow after the body instead (documented v1).
  deferredFootnotes: ReadonlyArray<string>;
} {
  const footnotes = new Map<string, number>();
  const endnotes = new Map<string, number>();
  const paragraphFootnotes = new Set<string>();

  const numberRun = (run: Run, n: number): Run => ({
    ...run,
    text: String(n),
    properties: { ...run.properties, verticalAlign: 'superscript' },
  });

  const mapParagraph = (paragraph: Paragraph): Paragraph => {
    if (!paragraph.runs.some((r) => r.footnoteRef !== undefined || r.endnoteRef !== undefined)) {
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
          (r) => r.footnoteRef !== undefined || r.endnoteRef !== undefined,
        );
      }
      if (el.kind === 'table') {
        return el.table.rows.some((row) => row.cells.some((c) => hasRefs(c.content)));
      }
      if (el.kind === 'shape' && el.shape.text) return hasRefs(el.shape.text.content);
      return false;
    });
  if (!hasRefs(body)) {
    return { body, footnotes, endnotes, deferredFootnotes: [] };
  }

  const mapped = body.map((el) => mapElement(el, true));
  const deferredFootnotes = [...footnotes.keys()].filter((id) => !paragraphFootnotes.has(id));
  return { body: mapped, footnotes, endnotes, deferredFootnotes };
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

// What pagination needs to place footnotes: per-id content, numbers, and a
// per-section lazily-cached layout of each note at that section's width.
interface NotePlan {
  readonly numbers: ReadonlyMap<string, number>;
  readonly layout: (
    ctx: SectionRenderCtx,
    id: string,
  ) => { blocks: ReadonlyArray<LaidOutBlock>; heightPt: number } | undefined;
}

// Layout phase (the FlowDoc→PageDoc transform of ir-design §7): body →
// positioned pages (PageItems), font/image resources, logical structure.
export function layoutStyledDocument(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): LaidOutPdfDocument {
  const sectionList = resolveSectionList(body, options);

  const noteAssigned = assignNoteNumbers(applyNumbering(body, options.numbering));
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
      ctx.columns ? ctx.columns[0]!.widthPt : ctx.contentWidth,
      ctx.pageContentHeight,
    );
  });

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
  }

  const bookmarks = new Map<string, BookmarkPosition>();
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
    pdf: { structBuilder, sectionCtxs, pdfaProfile, tagged, bookmarks },
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
  // §17.6.4 multi-column sections: per-column x-offset (from marginLeft) and
  // width. Absent for single-column sections. Body blocks are laid out at the
  // FIRST column's width (explicit unequal widths degrade to flowing without
  // re-wrap); headers/footers and footnotes keep the full content width.
  readonly columns?: ReadonlyArray<{ readonly xOffsetPt: number; readonly widthPt: number }>;
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
    dims.pageHeight,
    dims.footerOffsetPt,
  );
  const columns = buildColumnGeometry(section.properties.columns, contentWidth);
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
    ...(columns ? { columns } : {}),
    headerSet,
    footerSet,
    titlePg: section.properties.titlePg === true,
    evenAndOddHeaders: section.properties.evenAndOddHeaders === true,
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
    const render = (c: ReadonlyArray<BodyElement>): Array<PageItem> => {
      const blocks = laidOutBlocksFor(c, options, fontResources, contentWidth);
      return markPagination(
        drawBlocksSequentially(blocks, marginLeft, pageHeight - headerOffsetPt, pageHeight),
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

function layoutFooterSet(
  section: SectionProperties,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  options: StyledRenderOptions,
  fontResources: ReadonlyMap<string, FontResource>,
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
      const blocks = laidOutBlocksFor(c, options, fontResources, contentWidth);
      const totalHeight = blocks.reduce(
        (sum, b) =>
          sum +
          (b.kind === 'paragraph' ? b.spacingBeforePt + b.heightPt + b.spacingAfterPt : b.heightPt),
        0,
      );
      return markPagination(
        drawBlocksSequentially(blocks, marginLeft, footerOffsetPt + totalHeight, pageHeight),
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
  if (band === 'first') return has(set.first) ? set.first : set.default;
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
  const fillColorHex = shape.fill.kind === 'solid' ? shape.fill.colorHex : undefined;
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
// [0,size]² frame; emitRowChunk composes the page-flip transform. v1 draws a
// single filled shape per family — circle (lights/symbols), triangle (arrows),
// square (flags); faithful glyph sets follow.
const CF_ICON_SIZE_PT = 9;
const CF_ICON_GUTTER_PT = 12;

function buildCellIconShape(
  icon: CellIcon,
  size: number,
): { paths: ReadonlyArray<VectorPath>; fillColorHex: string } {
  return { paths: [cellIconPath(icon.shape, size)], fillColorHex: icon.colorHex };
}

function cellIconPath(shape: CellIcon['shape'], s: number): VectorPath {
  switch (shape) {
    case 'square':
      return rectAtPath(s * 0.12, s * 0.12, s * 0.76, s * 0.76);
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
  for (const note of options.footnotes?.values() ?? []) visit(note);
  for (const note of options.endnotes?.values() ?? []) visit(note);

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
  // Tagged PDF: stamp every emitted line with this structure node (used by
  // the footnote band; header/footer bands stay artifact-marked instead).
  structId?: number,
): Array<PageItem> {
  const out: Array<PageItem> = [];
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
  for (const note of options.footnotes?.values() ?? []) visit(note);
  for (const note of options.endnotes?.values() ?? []) visit(note);

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
  const lines = wrap(tokens, firstLineWidth, otherWidth, resolved, options.hyphenator, lineWidths);

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
        ...(plan.run.href !== undefined ? { href: plan.run.href } : {}),
        ...(plan.run.footnoteRef !== undefined ? { footnoteRef: plan.run.footnoteRef } : {}),
        ...(plan.run.anchor !== undefined ? { anchor: plan.run.anchor } : {}),
        ...(plan.run.listMarker ? { listMarker: true } : {}),
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
        ...(plan.run.href !== undefined ? { href: plan.run.href } : {}),
        ...(plan.run.footnoteRef !== undefined ? { footnoteRef: plan.run.footnoteRef } : {}),
        ...(plan.run.anchor !== undefined ? { anchor: plan.run.anchor } : {}),
        ...(plan.run.listMarker ? { listMarker: true } : {}),
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
  // Float text wrapping: explicit per-line widths (the last reuses for the
  // tail, the Knuth-Plass convention). Overrides first/other when given.
  lineWidths?: ReadonlyArray<number>,
): Array<Line> {
  if (tokens.length === 0) return [];

  const widths = lineWidths ?? [firstLineWidth, otherWidth];
  const entries = paragraphItemStream(tokens, hyphenator);
  const { breaks } = breakLines(
    entries.map((e) => e.item),
    widths,
  );

  const lines: Array<Line> = [];
  let start = 0;
  let isFirst = true;
  let lineIdx = 0;
  for (const breakIdx of breaks) {
    const width = lineIdx < widths.length ? widths[lineIdx]! : widths[widths.length - 1]!;
    const line = lineFromRange(entries, start, breakIdx, width, isFirst, resolved);
    if (line) lines.push(line);
    start = breakIdx + 1;
    isFirst = false;
    lineIdx++;
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
  const padLeftBase = cellMar?.left ?? tableMar?.left ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
  const padRightPt = cellMar?.right ?? tableMar?.right ?? DEFAULT_CELL_PADDING_TWIPS * TWIP_TO_PT;
  // A conditional-format icon (E-SHEET SC1c) reserves a left gutter; the cell's
  // text is inset past it so the glyph and the value never overlap.
  const padLeftPt = padLeftBase + (cell.properties.icon ? CF_ICON_GUTTER_PT : 0);

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
    ...(cell.properties.dataBar ? { dataBar: cell.properties.dataBar } : {}),
    ...(cell.properties.icon ? { icon: cell.properties.icon } : {}),
    ...(cell.properties.sparkline ? { sparkline: cell.properties.sparkline } : {}),
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

// A9 (oop-design §4.3): the page-assembly state machine, extracted verbatim
// from paginateSections. One instance assembles all pages of a document:
// fields are the former local state, methods the former closures — the block
// loop drives it. Push order into `current`/`pages` IS the byte order of the
// emitted PDF, so the bodies move unchanged.
class PageAssembler {
  constructor(
    readonly sectionCtxs: ReadonlyArray<SectionRenderCtx>,
    readonly builder: StructTreeBuilder | undefined,
    readonly notes: NotePlan | undefined,
    readonly bookmarkPositions: Map<string, BookmarkPosition> | undefined,
  ) {
    this.ctx = sectionCtxs[0]!;
    this.cursorY = this.ctx.pageHeight - this.ctx.marginTop;
  }

  readonly pages: Array<LaidOutPage> = [];

  ctx: SectionRenderCtx;
  secIdx = 0;
  pageInSection = 0;
  globalPageIdx = 0;
  current: Array<PageItem> = [];
  pendingPageBreak = false;
  cursorY: number;

  // §17.6.4 multi-column flow: content fills column after column before the
  // page flushes. this.colStartLen marks where the this.current column's items begin in
  // `this.current` — the single-column "page has content" guard generalizes to
  // "this column has content" (identical when there is one column).
  colIdx = 0;
  colStartLen = 0;
  colLeft = (): number => this.ctx.marginLeft + (this.ctx.columns?.[this.colIdx]?.xOffsetPt ?? 0);
  colWidth = (): number => this.ctx.columns?.[this.colIdx]?.widthPt ?? this.ctx.contentWidth;
  colHasContent = (): boolean => this.current.length > this.colStartLen;
  // Overflow step: next column on this page, or a fresh page after the last.
  advanceColumn = (): void => {
    if (this.ctx.columns && this.colIdx + 1 < this.ctx.columns.length) {
      this.colIdx++;
      this.colStartLen = this.current.length;
      this.cursorY = this.ctx.pageHeight - this.ctx.marginTop;
    } else {
      this.flushPage();
    }
  };

  // §20.4.2.3 out-of-flow drawings (wrap 'none'): they render at their
  // anchored position without moving the cursor. behindDoc sinks below the
  // body text, everything else above it; both flush with the page.
  floatsBehind: Array<PageItem> = [];
  floatsFront: Array<PageItem> = [];
  // Side-wrapping floats (wrapSquare/tight/through): rectangles the body text
  // must flow around. Page-scoped, like the float graphics.
  exclusions: Array<{ x0: number; x1: number; topYUp: number; bottomYUp: number }> = [];
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
  // The drawing's TOP in the y-up cursor frame. paragraph/line-relative
  // offsets hang off the anchoring paragraph's this.current position.
  floatTopYUp = (f: FloatAnchor): number => {
    const v = f.posV;
    if (!v) return this.cursorY;
    if (v.relativeFrom === 'page') return this.ctx.pageHeight - (v.offsetPt ?? 0);
    if (v.relativeFrom === 'margin')
      return this.ctx.pageHeight - this.ctx.marginTop - (v.offsetPt ?? 0);
    return this.cursorY - (v.offsetPt ?? 0);
  };
  // Tagged PDF: the stack of open list levels (for L/LI/LBody nesting). Cleared
  // whenever a non-list-item block interrupts the run of list paragraphs.
  readonly listStack: Array<ListFrame> = [];

  // §17.11 footnotes: this.notes reserved for the CURRENT page (greedy — a line
  // carrying a reference pulls its note's height out of the page bottom, so
  // the line and its note land together). `this.placedNotes` is global: a note
  // renders once, on its first reference's page.
  pageNotes: Array<{ n: number; blocks: ReadonlyArray<LaidOutBlock>; heightPt: number }> = [];
  noteReserve = 0;
  readonly placedNotes = new Set<string>();
  // The page's usable bottom: the margin plus whatever the this.notes band has
  // claimed so far.
  bottomLimit = (): number => this.ctx.marginBottom + this.noteReserve;

  // The first exclusion a line spanning [yTop-h, yTop] collides with in the
  // current column (x-overlap with the column required).
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

  // Per-line geometry beside an exclusion: the narrowed width and the x shift.
  // Text goes on the WIDER side of the float (one side per line, v1).
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

  // Estimated per-line widths for a paragraph starting at startY: narrowed
  // while lines (estimated at the first line's height) overlap an exclusion,
  // then one full width Knuth-Plass reuses for the tail. Undefined when
  // nothing overlaps — the caller keeps the original block.
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

  // New (unplaced) footnotes referenced by a line's tokens, with their layout
  // at the this.current section's width.
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

  // The this.notes band for the flushing page: separator rule + each note's blocks
  // stacked inside the reserved area. Tagged: each note is a Note→P element.
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
          structId,
        ),
      );
      cursor -= note.heightPt;
    }
    return out;
  };

  // Dynamic PAGE/NUMPAGES bands re-render after pagination (both numbers are
  // known only then); each use records where its commands must be spliced.
  readonly dynBands: Array<{
    pageIdx: number;
    pageNumber: number;
    position: 'header' | 'footer';
    render: (pageNumber: number, totalPages: number) => Array<PageItem>;
  }> = [];

  flushPage = (force = false): void => {
    if (this.current.length === 0 && !force) return;
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
      asm.flushPage();
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
      if (asm.current.length > 0) asm.flushPage();
    }
    // A non-list-item block ends any open list run (tagged PDF).
    if (builder && !(block.kind === 'paragraph' && block.list)) asm.listStack.length = 0;
    if (block.kind === 'paragraph') {
      if (block.resolved.pageBreakBefore && asm.current.length > 0) asm.flushPage();
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
            ...(block.stroke ? { stroke: block.stroke } : {}),
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
          emitRowChunk(
            asm.current,
            chunk,
            tableX,
            asm.cursorY,
            asm.ctx.pageHeight,
            colCount,
            cellStructIds,
          );
          asm.cursorY -= chunk.heightPt;
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
  return asm.pages;
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
  out: Array<PageItem>,
  row: RowLayout,
  marginLeft: number,
  cursorY: number,
  pageHeight: number,
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
        x: pt(cellX),
        y: pt(pageHeight - rowBottom - row.heightPt),
        width: pt(cell.widthPt),
        height: pt(row.heightPt),
        fillColorHex: cell.shadingColorHex,
      });
    }
    // Conditional-format data bar (E-SHEET SC1c): a fraction-width fill over the
    // shading, under the text. Pushed after the shading fill so it paints on top.
    if (cell.dataBar && cell.mergeRole !== 'middle' && cell.mergeRole !== 'end') {
      const barWidth = cell.widthPt * Math.max(0, Math.min(1, cell.dataBar.fraction));
      if (barWidth > 0) {
        out.push({
          type: 'fill',
          x: pt(cellX),
          y: pt(pageHeight - rowBottom - row.heightPt),
          width: pt(barWidth),
          height: pt(row.heightPt),
          fillColorHex: cell.dataBar.colorHex,
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
        const { paths, fillColorHex } = buildCellIconShape(cell.icon, iconSize);
        out.push({
          type: 'shape',
          shape: {
            paths,
            fillColorHex,
            transform: flipTransform([1, 0, 0, 1, iconX, iconBottomYUp], pageHeight),
          },
        });
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
        originX: pt(cellX + cell.padLeftPt + offset),
        baselineY: pt(pageHeight - (textY + lineDescent(line))),
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
      width: pt(cell.widthPt),
      height: pt(rowHeight),
      borderSizePt: sz,
      borderColorHex: border.colorHex ?? '000000',
    });
  };
  pushSide('top', cell.borders.top);
  pushSide('left', cell.borders.left);
  if (rowIdx === rowCount - 1) pushSide('bottom', cell.borders.bottom);
  if (cell.colStart + cell.colSpan - 1 === colCount - 1) pushSide('right', cell.borders.right);
}
