// DOCX reader (ir-design §7): bytes → FlowDoc. The document-derived half of
// what convertDocxToPdfSync used to do inline — everything here comes from the
// .docx itself; caller-supplied conversion options (fonts, PDF/A, signature)
// stay with the converter/facade.

import type { ColorResolver } from '@/core/drawingml/colors';
import type {
  BodyElement,
  Chart,
  Comment,
  DocumentInfo,
  HeaderFooterReference,
  Numbering,
  Section,
  StyleSheet,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss, ResourceId } from '@/core/ir';
import type { CoreProperties } from '@/core/opc';
import type { HyperlinkResolver, ImageResolver, ParseContext, ResolvedDiagram } from '@/word';
import { poFindDescendant } from '@/core/po-helpers';
import { parseXml } from '@/pptx/pptx-reader';
import { bytesIncludePartName } from '@/core/bytes';
import { applyNumbering, applyNumberingToHeadersFooters } from '@/core/numbering';
import {
  EMPTY_STYLE_SHEET,
  resolveBodyStyles,
  resolveHeadersFootersStyles,
  resolveTableStyles,
} from '@/core/style-cascade';

import { FEATURES, ResourceStore } from '@/core/ir';
import { parseChart, withChartColorStyle } from '@/core/drawingml/chart-parser';
import { DEFAULT_THEME_PALETTE, makeColorResolver } from '@/core/drawingml/colors';
import {
  parseTheme,
  parseThemeEffectStyles,
  parseThemeLineWidths,
} from '@/core/drawingml/theme-parser';
import { OpcPackage, isOoxmlRel, parseCoreProperties } from '@/core/opc';
import {
  EMPTY_NUMBERING,
  EMPTY_SECTION,
  EMPTY_SETTINGS,
  HTML_AUTO_SPACING_PT,
  applyAuthorIds,
  bodyIndexForBlock,
  loadEmbeddedFonts,
  newBlockCounter,
  parseBackgroundColor,
  parseBackgroundFill,
  parseCommentThreads,
  parseDocument,
  parseHeaderFooter,
  parseNotes,
  parseNumbering,
  parsePeople,
  parseSections,
  parseSettings,
  parseStyles,
} from '@/word';

const STYLES_PART = 'word/styles.xml';
const FOOTNOTES_PART = 'word/footnotes.xml';
const ENDNOTES_PART = 'word/endnotes.xml';
const COMMENTS_PART = 'word/comments.xml';
const COMMENTS_EXTENDED_PART = 'word/commentsExtended.xml';
const PEOPLE_PART = 'word/people.xml';
const NUMBERING_PART = 'word/numbering.xml';
const SETTINGS_PART = 'word/settings.xml';
const CORE_PROPS_PART = 'docProps/core.xml';
// OPC §11.1 names the main document part through the package's `officeDocument`
// relationship, not by a fixed path — `word/document.xml` is only what every
// producer happens to choose. tdf104713_undefinedStyles.docx calls its
// `word/trial.xml` and hangs its footer off that part's own .rels, so every
// lookup keyed on the conventional name came back empty.
const MAIN_DOCUMENT_PART = 'word/document.xml';

const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const THEME_PART = 'word/theme/theme1.xml';

/**
 * Read a `.docx` package into the {@link FlowDoc} interlayer (ir-design §7): the
 * document-derived half of what the converter used to do inline. Parses the main
 * document, styles, numbering, notes, comments, headers/footers, charts, embedded
 * fonts and core properties, runs the body through the stage-6 FlowDoc transforms
 * (list markers then the style cascade), and returns the tree plus any
 * graceful-degradation losses. Caller-supplied conversion options (fonts, PDF/A,
 * signature) stay with the converter/facade.
 *
 * @param docx The raw `.docx` (OPC ZIP) bytes.
 * @returns The parsed {@link FlowDoc} and the losses recorded while reading.
 */
export function readDocx(docx: Uint8Array): ReadResult<FlowDoc> {
  const pkg = OpcPackage.open(docx);
  const main = pkg.getMainDocument();
  // Theme-backed colour resolver (schemeClr → hex); falls back to the built-in
  // Office palette when there is no theme part.
  const resolveColor = buildColorResolver(pkg, main.path);
  // §20.1.4.2.19 — the theme's own line weights, which a gallery-styled shape
  // indexes by `a:lnRef idx` for its outline.
  const themeData = loadTheme(pkg, main.path);
  const themeLineWidths = themeData ? parseThemeLineWidths(themeData) : undefined;
  const themeEffectStyles = themeData ? parseThemeEffectStyles(themeData) : undefined;
  // Content-addressed store for binary resources; the image resolver fills it
  // lazily as the parsers meet drawing relationships (identical bytes dedupe).
  const resources = new ResourceStore();
  const resolveImage = makeImageResolver(pkg, resources, main.path);
  const resolveHyperlink = makeHyperlinkResolver(pkg, main.path);
  // Graceful-degradation notices recorded while parsing the body (E-SMARTART
  // SA3: a SmartArt with no drawing override). Headers/footers and notes don't
  // resolve diagrams, so the sink rides only on the main-body context.
  const losses: Array<Loss> = [];
  // settings.xml is read before the body because §17.3.1.1's automatic
  // paragraph spacing depends on the document's compatibility mode.
  const settingsData = pkg.getPart(SETTINGS_PART);
  const settings = settingsData ? parseSettings(settingsData) : EMPTY_SETTINGS;
  const ctx: ParseContext = {
    resolveColor,
    ...(settings.compatibilityMode === undefined ? { autoSpacingPt: HTML_AUTO_SPACING_PT } : {}),
    ...(themeLineWidths && themeLineWidths.length > 0 ? { themeLineWidths } : {}),
    ...(themeEffectStyles && themeEffectStyles.length > 0 ? { themeEffectStyles } : {}),
    resolveImage,
    resolveHyperlink,
    resolveDiagram: makeDiagramResolver(pkg, main.path, resources),
    resolveChartPart: makeChartResolver(pkg, main.path),
    onLoss: (loss) => losses.push(loss),
    // Tracks open comment ranges across the body so runs carry commentRangeRefs.
    openCommentRanges: new Set<string>(),
  };
  // A paragraph carrying anchored drawings emits several body elements, while
  // parseSections counts source blocks — so the section boundaries have to be
  // translated, or a landscape first section covers three elements instead of
  // a hundred (fdo74605 draws its whole diagram on one landscape page).
  const blocks = newBlockCounter();
  const body = parseDocument(main.data, ctx, blocks);
  const backgroundColorHex = parseBackgroundColor(main.data);
  const backgroundFill = parseBackgroundFill(main.data, ctx.resolveImage);
  const rawSections = parseSections(main.data).map((s) => ({
    ...s,
    endIndex: bodyIndexForBlock(blocks, s.endIndex),
  }));

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseStyles(stylesData) : EMPTY_STYLE_SHEET;

  const numberingData = pkg.getPart(NUMBERING_PART);
  // §17.9.21 — a picture bullet's image is a relationship of the NUMBERING part.
  const numbering = numberingData
    ? parseNumbering(numberingData, makeImageResolver(pkg, resources, NUMBERING_PART))
    : EMPTY_NUMBERING;

  // §17.11 notes: parsed with per-part resolvers (their rels own their
  // images/links), then run through the same FlowDoc transforms as the body.
  const noteCtx = (part: string): ParseContext => ({
    resolveColor,
    resolveImage: makeImageResolver(pkg, resources, part),
    resolveHyperlink: makeHyperlinkResolver(pkg, part),
  });
  const footnotesData = pkg.getPart(FOOTNOTES_PART);
  const rawFootnotes = footnotesData
    ? parseNotes(footnotesData, 'w:footnotes', 'w:footnote', noteCtx(FOOTNOTES_PART))
    : undefined;
  const endnotesData = pkg.getPart(ENDNOTES_PART);
  const rawEndnotes = endnotesData
    ? parseNotes(endnotesData, 'w:endnotes', 'w:endnote', noteCtx(ENDNOTES_PART))
    : undefined;
  const commentsData = pkg.getPart(COMMENTS_PART);
  const commentsExtendedData = pkg.getPart(COMMENTS_EXTENDED_PART);
  const peopleData = pkg.getPart(PEOPLE_PART);
  let rawComments = commentsData
    ? parseCommentThreads(commentsData, commentsExtendedData, noteCtx(COMMENTS_PART))
    : undefined;
  // word/people.xml resolves each author to a presence identity (usually email).
  if (rawComments && peopleData) {
    rawComments = applyAuthorIds(rawComments, parsePeople(peopleData));
  }

  // evenAndOddHeaders lives in settings.xml; replicate the flag onto every
  // section so the renderer sees a per-section view of header bands.
  const sections: Array<Section> = withInheritedBands(
    rawSections.length > 0
      ? rawSections.map((s) => ({
          ...s,
          properties: settings.evenAndOddHeaders
            ? { ...s.properties, evenAndOddHeaders: true }
            : s.properties,
        }))
      : [
          {
            properties: settings.evenAndOddHeaders
              ? { ...EMPTY_SECTION, evenAndOddHeaders: true }
              : EMPTY_SECTION,
            endIndex: body.length,
          },
        ],
  ) as Array<Section>;

  const headersFooters = loadHeadersFootersForSections(pkg, sections, ctx, resources, main.path);
  const charts = loadCharts(pkg, resolveColor, chartOwningParts(pkg, sections, main.path));
  // The document's own embedded fonts (de-obfuscated). A run whose w:ascii
  // matches one renders with the real font instead of a substitute.
  const embeddedFonts = loadEmbeddedFonts(pkg);

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;
  const info = infoFromCore(coreProps);
  // Document language for the tagged-PDF /Lang.
  const language = detectDocxLanguage(stylesData, main.data);

  const doc: FlowDoc = {
    kind: 'flow',
    // Stage-6 FlowDoc transforms, in renderer order: list markers first, then
    // the style cascade — the tree carries final effective properties, so
    // every writer sees ready paragraphs. `numbering`/`styles` stay as raw
    // material for round-trip writers; render projections must NOT re-apply
    // them (the projector sends EMPTY_STYLE_SHEET).
    body: resolveBodyStyles(
      applyNumbering(resolveTableStyles(body, styles), numbering, styles),
      styles,
    ),
    sections,
    styles,
    numbering,
    ...(rawFootnotes && rawFootnotes.size > 0
      ? { footnotes: transformNotes(rawFootnotes, styles, numbering) }
      : {}),
    ...(rawEndnotes && rawEndnotes.size > 0
      ? { endnotes: transformNotes(rawEndnotes, styles, numbering) }
      : {}),
    ...(rawComments && rawComments.size > 0
      ? { comments: transformComments(rawComments, styles, numbering) }
      : {}),
    headersFooters: resolveHeadersFootersStyles(
      applyNumberingToHeadersFooters(headersFooters, numbering, styles),
      styles,
    ),
    charts,
    resources,
    ...(embeddedFonts.size > 0 ? { embeddedFonts } : {}),
    ...(info ? { info } : {}),
    ...(language ? { language } : {}),
    ...(settings.doNotExpandShiftReturn ? { doNotExpandShiftReturn: true } : {}),
    // §17.2.1 / §17.15.1.28 — the page background is a colour AND a flag: Word
    // keeps the colour in every document that ever had one and prints it only
    // when `w:displayBackgroundShape` is set.
    ...(settings.displayBackgroundShape && backgroundFill
      ? { pageBackgroundFill: backgroundFill }
      : {}),
    ...(settings.displayBackgroundShape && backgroundColorHex
      ? { pageBackgroundColorHex: backgroundColorHex }
      : {}),
    ...(settings.gutterAtTop ? { gutterAtTop: true } : {}),
  };
  return { doc, losses };
}

/**
 * The {@link DocumentReader} registration for `.docx`: its id, the FlowDoc
 * feature set it supports, the sniffer, and the {@link readDocx} entry point.
 */
export const docxReader: DocumentReader<FlowDoc> = {
  id: 'docx',
  produces: 'flow',
  supports: new Set([
    FEATURES.text,
    FEATURES.tables,
    FEATURES.tablesNested,
    FEATURES.lists,
    FEATURES.sections,
    FEATURES.headersFooters,
    FEATURES.images,
    FEATURES.shapes,
    FEATURES.charts,
    FEATURES.math,
    FEATURES.rtl,
    FEATURES.trackedChanges,
    FEATURES.fontsEmbedding,
  ]),
  // A docx is a ZIP whose central directory names the WordprocessingML parts —
  // the part names sit as plain bytes in the container, so a substring probe is
  // cheap and reliable without unzipping.
  //
  // The MAIN part's name is not fixed: OPC names it through the package's
  // `officeDocument` relationship, which the reader already follows.
  // tdf104713_undefinedStyles.docx calls it `word/trial.xml`, and a probe for
  // `word/document.xml` alone refused a file the reader behind it reads.
  // `word/styles.xml` is what every producer writes beside it, and neither an
  // xlsx (`xl/`) nor a pptx (`ppt/`) has anything under `word/`.
  sniff: (bytes) =>
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    (bytesIncludePartName(bytes, 'word/document.xml') ||
      bytesIncludePartName(bytes, 'word/styles.xml') ||
      bytesIncludePartName(bytes, 'word/_rels/')),
  read: (bytes) => readDocx(bytes),
};

// The document-derived half of the old converter mergeInfo: docProps/core.xml
// mapped into DocumentInfo. The converter spreads caller overrides on top.
function infoFromCore(core: CoreProperties | undefined): DocumentInfo | undefined {
  if (!core) return undefined;
  return {
    ...(core.title ? { title: core.title } : {}),
    ...(core.creator ? { author: core.creator } : {}),
    ...(core.subject ? { subject: core.subject } : {}),
    ...(core.keywords ? { keywords: core.keywords } : {}),
    ...(core.created ? { creationDate: core.created } : {}),
    ...(core.modified ? { modificationDate: core.modified } : {}),
  };
}

// SmartArt: a data relationship id (dgm:relIds @r:dm) → the diagram's
// pre-rendered drawing override (its dsp:spTree). Follows the doc part →
// diagrams/data#.xml → (rel type .../diagramDrawing) → diagrams/drawing#.xml.
// Undefined when the file ships no drawing override (E-SMARTART SA2).
function makeDiagramResolver(
  pkg: OpcPackage,
  partName: string,
  store: ResourceStore,
): (relId: string) => ResolvedDiagram | undefined {
  const cache = new Map<string, ResolvedDiagram | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    let resolved: ResolvedDiagram | undefined;
    const dataRel = pkg.getPartRelationships(partName).find((r) => r.id === relId);
    const data = dataRel ? pkg.resolveRelatedPart(partName, dataRel) : undefined;
    if (data) {
      const own = pkg
        .getPartRelationships(data.path)
        .find((r) => r.type.endsWith('/diagramDrawing'));
      const draw = own
        ? pkg.resolveRelatedPart(data.path, own)
        : drawingFromOwner(pkg, partName, data.path);
      if (draw) {
        for (const root of parseXml(draw.data)) {
          const spTree = poFindDescendant(root, 'dsp:spTree');
          if (spTree) {
            // A node's picture fill names a relationship of the DRAWING part,
            // not of the document (fdo74792 fills four nodes with clip art).
            resolved = { spTree, resolveImage: makeImageResolver(pkg, store, draw.path) };
            break;
          }
        }
      }
    }
    cache.set(relId, resolved);
    return resolved;
  };
}

// Half the SmartArt in the corpus (5 of 9 files, fdo73227 among them) gives the
// data part no .rels of its own and hangs the `.../2007/relationships/
// diagramDrawing` off the part that OWNS the diagram instead. The relationship
// then names no data part, so the pairing is the one Word writes into the file
// names: diagrams/data2.xml belongs with diagrams/drawing2.xml. A lone drawing
// pairs with whatever data part asked, index or no index.
function drawingFromOwner(
  pkg: OpcPackage,
  ownerPart: string,
  dataPart: string,
): { readonly path: string; readonly data: Uint8Array } | undefined {
  const drawings = pkg
    .getPartRelationships(ownerPart)
    .filter((r) => r.type.endsWith('/diagramDrawing'));
  if (drawings.length === 0) return undefined;
  const index = (path: string): string => /(\d*)\.[^.]+$/u.exec(path)?.[1] ?? '';
  const want = index(dataPart);
  const rel =
    drawings.find((r) => index(r.target) === want) ??
    (drawings.length === 1 ? drawings[0] : undefined);
  return rel ? pkg.resolveRelatedPart(ownerPart, rel) : undefined;
}

function makeImageResolver(
  pkg: OpcPackage,
  store: ResourceStore,
  partName: string = MAIN_DOCUMENT_PART,
): ImageResolver {
  // Relationship ids are scoped to their OWNING part (OPC §9.3) — a header's
  // rId must resolve against the header's own .rels, not the main document's
  // (oop-design §8, C5: shared-resolver bug fixed).
  const byRelId = new Map<string, Uint8Array>();
  for (const rel of pkg.getPartRelationships(partName)) {
    if (!isOoxmlRel(rel.type, 'image')) continue;
    const resolved = pkg.resolveRelatedPart(partName, rel);
    if (resolved) byRelId.set(rel.id, resolved.data);
  }
  const cache = new Map<string, ResourceId | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    const bytes = byRelId.get(relId);
    const id = bytes ? store.put(bytes) : undefined;
    cache.set(relId, id);
    return id;
  };
}

// Notes get the same FlowDoc transforms as the body: table styles, list
// markers, the resolved cascade (each note numbers its own lists, like a
// header/footer band).
function transformNotes(
  notes: Map<string, Array<BodyElement>>,
  styles: StyleSheet,
  numbering: Numbering,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  for (const content of notes.values()) resolveTableStyles(content, styles);
  return resolveHeadersFootersStyles(applyNumberingToHeadersFooters(notes, numbering), styles);
}

// Comments get the same FlowDoc transforms as notes, applied to each comment's
// content; its author/date metadata rides through unchanged (E-COMMENTS CM0).
function transformComments(
  comments: Map<string, Comment>,
  styles: StyleSheet,
  numbering: Numbering,
): ReadonlyMap<string, Comment> {
  const contentById = new Map<string, Array<BodyElement>>();
  for (const [id, c] of comments) contentById.set(id, [...c.content]);
  const transformed = transformNotes(contentById, styles, numbering);
  const out = new Map<string, Comment>();
  for (const [id, c] of comments) out.set(id, { ...c, content: transformed.get(id) ?? c.content });
  return out;
}

// §17.16.22 + OPC §9.3: hyperlink relationship ids are scoped to their OWNING
// part, and only TargetMode="External" targets are URLs (internal-mode
// hyperlink rels point at parts, not the web).
function makeHyperlinkResolver(
  pkg: OpcPackage,
  partName: string = MAIN_DOCUMENT_PART,
): HyperlinkResolver {
  const byRelId = new Map<string, string>();
  for (const rel of pkg.getPartRelationships(partName)) {
    if (rel.type === REL_HYPERLINK && rel.targetMode === 'External') {
      byRelId.set(rel.id, rel.target);
    }
  }
  return (relId) => byRelId.get(relId);
}

// Resolve & parse every chart part referenced by the main document, keyed by
// its relationship id (which ChartBlock.chartRelId points to).
// Charts live in every part that can hold a drawing, not just the body:
// chart-in-footer.docx anchors one in a footer, whose rels are its own (OPC
// §9.3). Key by the chart part's path — unique across owning parts, and the
// same key the .xlsx and .pptx readers use — and let each part's context map
// its local r:id onto it (see makeChartResolver).
function loadCharts(
  pkg: OpcPackage,
  resolveColor: ColorResolver,
  parts: ReadonlyArray<string>,
): ReadonlyMap<string, Chart> {
  const out = new Map<string, Chart>();
  for (const part of parts) {
    for (const rel of pkg.getPartRelationships(part)) {
      if (!isOoxmlRel(rel.type, 'chart')) continue;
      const resolved = pkg.resolveRelatedPart(part, rel);
      if (!resolved || out.has(resolved.path)) continue;
      const chart = parseChart(resolved.data, resolveColor);
      if (chart)
        out.set(resolved.path, withChartColorStyle(chart, pkg, resolved.path, resolveColor));
    }
  }
  return out;
}

// A drawing's `c:chart` `@r:id` → the chart part's path, i.e. the key
// {@link loadCharts} files it under. Absent ⇒ the parser keeps the raw rel id.
function makeChartResolver(
  pkg: OpcPackage,
  partName: string,
): (relId: string) => string | undefined {
  const byRelId = new Map<string, string>();
  for (const rel of pkg.getPartRelationships(partName)) {
    if (!isOoxmlRel(rel.type, 'chart')) continue;
    const resolved = pkg.resolveRelatedPart(partName, rel);
    if (resolved) byRelId.set(rel.id, resolved.path);
  }
  return (relId) => byRelId.get(relId);
}

// Theme colour resolver: merge the document's theme palette (if any) over the
// built-in Office defaults, so schemeClr references resolve to the document's
// actual accent colours and unspecified slots still have sensible values.
function buildColorResolver(pkg: OpcPackage, mainPart: string): ColorResolver {
  const themeData = loadTheme(pkg, mainPart);
  if (!themeData) return makeColorResolver(DEFAULT_THEME_PALETTE);
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const [slot, hex] of parseTheme(themeData)) palette.set(slot, hex);
  return makeColorResolver(palette);
}

function loadTheme(pkg: OpcPackage, mainPart: string): Uint8Array | undefined {
  for (const rel of pkg.getPartRelationships(mainPart)) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(mainPart, rel);
    if (resolved) return resolved.data;
  }
  return pkg.getPart(THEME_PART);
}

/**
 * §17.10.1 — a section that declares no header (or footer) of a given type
 * takes the one before it. endingSectionProps.docx puts its references on the
 * FIRST section and ends with a continuous one that states none, and reading
 * each section alone left the page with no header and no footer at all.
 *
 * @param sections The sections in document order.
 * @returns The same sections, each carrying the bands it inherits.
 */
function withInheritedBands(sections: ReadonlyArray<Section>): ReadonlyArray<Section> {
  const out: Array<Section> = [];
  let headers: ReadonlyArray<HeaderFooterReference> = [];
  let footers: ReadonlyArray<HeaderFooterReference> = [];
  // §17.10.6 — the flag that chooses BETWEEN the inherited bands comes with
  // them: endingSectionProps.docx marks its first section `titlePg`, and the
  // continuous section after it draws the same first-page footer.
  let titlePg = false;
  for (const section of sections) {
    const own = section.properties;
    // Word inherits per SECTION, not per type: a section that names any band
    // of its own starts a fresh set.
    const inherits = own.headers.length === 0 && own.footers.length === 0;
    headers = own.headers.length > 0 ? own.headers : headers;
    footers = own.footers.length > 0 ? own.footers : footers;
    titlePg = inherits ? titlePg : (own.titlePg ?? false);
    out.push(
      headers === own.headers && footers === own.footers && titlePg === (own.titlePg ?? false)
        ? section
        : {
            ...section,
            properties: { ...own, headers, footers, ...(titlePg ? { titlePg: true } : {}) },
          },
    );
  }
  return out;
}

function loadHeadersFootersForSections(
  pkg: OpcPackage,
  sections: ReadonlyArray<Section>,
  ctx: ParseContext,
  store: ResourceStore,
  mainPart: string,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  const wanted = new Set<string>();
  for (const s of sections) {
    for (const h of s.properties.headers) wanted.add(h.relationshipId);
    for (const f of s.properties.footers) wanted.add(f.relationshipId);
  }
  if (wanted.size === 0) return new Map();
  const rels = pkg.getPartRelationships(mainPart);
  if (rels.length === 0) return new Map();

  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const rel of rels) {
    if (!wanted.has(rel.id)) continue;
    if (!isOoxmlRel(rel.type, 'header') && !isOoxmlRel(rel.type, 'footer')) continue;
    const resolved = pkg.resolveRelatedPart(mainPart, rel);
    if (!resolved) continue;
    const hfCtx: ParseContext = {
      resolveColor: ctx.resolveColor,
      resolveImage: makeImageResolver(pkg, store, resolved.path),
      resolveHyperlink: makeHyperlinkResolver(pkg, resolved.path),
      resolveChartPart: makeChartResolver(pkg, resolved.path),
    };
    out.set(rel.id, parseHeaderFooter(resolved.data, hfCtx));
  }
  return out;
}

// The parts whose relationships may point at a chart: the body plus every
// header/footer the sections actually use.
function chartOwningParts(
  pkg: OpcPackage,
  sections: ReadonlyArray<Section>,
  mainPart: string,
): Array<string> {
  const wanted = new Set<string>();
  for (const s of sections) {
    for (const h of s.properties.headers) wanted.add(h.relationshipId);
    for (const f of s.properties.footers) wanted.add(f.relationshipId);
  }
  const parts = [mainPart];
  for (const rel of pkg.getPartRelationships(mainPart)) {
    if (!wanted.has(rel.id)) continue;
    if (!isOoxmlRel(rel.type, 'header') && !isOoxmlRel(rel.type, 'footer')) continue;
    const resolved = pkg.resolveRelatedPart(mainPart, rel);
    if (resolved) parts.push(resolved.path);
  }
  return parts;
}

// Best-effort document language for the tagged-PDF catalog /Lang. The default
// language lives in styles.xml docDefaults/rPrDefault/rPr/w:lang @w:val; fall
// back to the first w:lang in the document body. A cheap regex (cf.
// detectDocxFontFamily) — no need to resolve the cascade for a document hint.
function detectDocxLanguage(
  stylesData: Uint8Array | undefined,
  documentData: Uint8Array,
): string | undefined {
  const re = /<w:lang\b[^>]*\bw:val="([^"]+)"/;
  const decoder = new TextDecoder();
  for (const data of [stylesData, documentData]) {
    if (!data) continue;
    const m = re.exec(decoder.decode(data));
    if (m?.[1]) return m[1];
  }
  return undefined;
}
