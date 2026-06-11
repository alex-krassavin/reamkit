// DOCX reader (ir-design §7): bytes → FlowDoc. The document-derived half of
// what convertDocxToPdfSync used to do inline — everything here comes from the
// .docx itself; caller-supplied conversion options (fonts, PDF/A, signature)
// stay with the converter/facade.

import type { ColorResolver } from '@/core/drawingml/colors';
import type {
  BodyElement,
  Chart,
  DocumentInfo,
  Numbering,
  Section,
  StyleSheet,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { ResourceId } from '@/core/ir';
import type { CoreProperties } from '@/core/opc';
import type { HyperlinkResolver, ImageResolver, ParseContext } from '@/word';
import { bytesInclude } from '@/core/bytes';
import { applyNumbering, applyNumberingToHeadersFooters } from '@/core/numbering';
import {
  EMPTY_STYLE_SHEET,
  resolveBodyStyles,
  resolveHeadersFootersStyles,
  resolveTableStyles,
} from '@/core/style-cascade';

import { FEATURES, ResourceStore } from '@/core/ir';
import { parseChart } from '@/core/drawingml/chart-parser';
import { DEFAULT_THEME_PALETTE, makeColorResolver } from '@/core/drawingml/colors';
import { parseTheme } from '@/core/drawingml/theme-parser';
import { OpcPackage, parseCoreProperties } from '@/core/opc';
import {
  EMPTY_NUMBERING,
  EMPTY_SECTION,
  EMPTY_SETTINGS,
  loadEmbeddedFonts,
  parseDocument,
  parseHeaderFooter,
  parseNotes,
  parseNumbering,
  parseSections,
  parseSettings,
  parseStyles,
} from '@/word';

const STYLES_PART = 'word/styles.xml';
const FOOTNOTES_PART = 'word/footnotes.xml';
const ENDNOTES_PART = 'word/endnotes.xml';
const NUMBERING_PART = 'word/numbering.xml';
const SETTINGS_PART = 'word/settings.xml';
const CORE_PROPS_PART = 'docProps/core.xml';
const MAIN_DOCUMENT_PART = 'word/document.xml';

const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const THEME_PART = 'word/theme/theme1.xml';

export function readDocx(docx: Uint8Array): ReadResult<FlowDoc> {
  const pkg = OpcPackage.open(docx);
  const main = pkg.getMainDocument();
  // Theme-backed colour resolver (schemeClr → hex); falls back to the built-in
  // Office palette when there is no theme part.
  const resolveColor = buildColorResolver(pkg);
  // Content-addressed store for binary resources; the image resolver fills it
  // lazily as the parsers meet drawing relationships (identical bytes dedupe).
  const resources = new ResourceStore();
  const resolveImage = makeImageResolver(pkg, resources);
  const resolveHyperlink = makeHyperlinkResolver(pkg);
  const ctx: ParseContext = { resolveColor, resolveImage, resolveHyperlink };
  const body = parseDocument(main.data, ctx);
  const rawSections = parseSections(main.data);

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseStyles(stylesData) : EMPTY_STYLE_SHEET;

  const numberingData = pkg.getPart(NUMBERING_PART);
  const numbering = numberingData ? parseNumbering(numberingData) : EMPTY_NUMBERING;

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

  const settingsData = pkg.getPart(SETTINGS_PART);
  const settings = settingsData ? parseSettings(settingsData) : EMPTY_SETTINGS;

  // evenAndOddHeaders lives in settings.xml; replicate the flag onto every
  // section so the renderer sees a per-section view of header bands.
  const sections: Array<Section> =
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
        ];

  const headersFooters = loadHeadersFootersForSections(pkg, sections, ctx, resources);
  const charts = loadCharts(pkg, resolveColor);
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
    body: resolveBodyStyles(applyNumbering(resolveTableStyles(body, styles), numbering), styles),
    sections,
    styles,
    numbering,
    ...(rawFootnotes && rawFootnotes.size > 0
      ? { footnotes: transformNotes(rawFootnotes, styles, numbering) }
      : {}),
    ...(rawEndnotes && rawEndnotes.size > 0
      ? { endnotes: transformNotes(rawEndnotes, styles, numbering) }
      : {}),
    headersFooters: resolveHeadersFootersStyles(
      applyNumberingToHeadersFooters(headersFooters, numbering),
      styles,
    ),
    charts,
    resources,
    ...(embeddedFonts.size > 0 ? { embeddedFonts } : {}),
    ...(info ? { info } : {}),
    ...(language ? { language } : {}),
  };
  return { doc, losses: [] };
}

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
  // A docx is a ZIP whose central directory names word/document.xml — the part
  // names sit as plain bytes in the container, so a substring probe is cheap
  // and reliable without unzipping.
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesInclude(bytes, 'word/document.xml'),
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
    if (rel.type !== REL_IMAGE) continue;
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
function loadCharts(pkg: OpcPackage, resolveColor: ColorResolver): ReadonlyMap<string, Chart> {
  const out = new Map<string, Chart>();
  for (const rel of pkg.getPartRelationships(MAIN_DOCUMENT_PART)) {
    if (rel.type !== REL_CHART) continue;
    const resolved = pkg.resolveRelatedPart(MAIN_DOCUMENT_PART, rel);
    if (!resolved) continue;
    const chart = parseChart(resolved.data, resolveColor);
    if (chart) out.set(rel.id, chart);
  }
  return out;
}

// Theme colour resolver: merge the document's theme palette (if any) over the
// built-in Office defaults, so schemeClr references resolve to the document's
// actual accent colours and unspecified slots still have sensible values.
function buildColorResolver(pkg: OpcPackage): ColorResolver {
  const themeData = loadTheme(pkg);
  if (!themeData) return makeColorResolver(DEFAULT_THEME_PALETTE);
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const [slot, hex] of parseTheme(themeData)) palette.set(slot, hex);
  return makeColorResolver(palette);
}

function loadTheme(pkg: OpcPackage): Uint8Array | undefined {
  for (const rel of pkg.getPartRelationships(MAIN_DOCUMENT_PART)) {
    if (rel.type !== REL_THEME) continue;
    const resolved = pkg.resolveRelatedPart(MAIN_DOCUMENT_PART, rel);
    if (resolved) return resolved.data;
  }
  return pkg.getPart(THEME_PART);
}

function loadHeadersFootersForSections(
  pkg: OpcPackage,
  sections: ReadonlyArray<Section>,
  ctx: ParseContext,
  store: ResourceStore,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  const wanted = new Set<string>();
  for (const s of sections) {
    for (const h of s.properties.headers) wanted.add(h.relationshipId);
    for (const f of s.properties.footers) wanted.add(f.relationshipId);
  }
  if (wanted.size === 0) return new Map();
  const rels = pkg.getPartRelationships(MAIN_DOCUMENT_PART);
  if (rels.length === 0) return new Map();

  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const rel of rels) {
    if (!wanted.has(rel.id)) continue;
    if (rel.type !== REL_HEADER && rel.type !== REL_FOOTER) continue;
    const resolved = pkg.resolveRelatedPart(MAIN_DOCUMENT_PART, rel);
    if (!resolved) continue;
    const hfCtx: ParseContext = {
      resolveColor: ctx.resolveColor,
      resolveImage: makeImageResolver(pkg, store, resolved.path),
      resolveHyperlink: makeHyperlinkResolver(pkg, resolved.path),
    };
    out.set(rel.id, parseHeaderFooter(resolved.data, hfCtx));
  }
  return out;
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
