import type { BodyElement, Chart, Section } from '@/document-model';
import type { FontBytesByVariant } from '@/font';
import type { FamilyKey, FetchLike } from '@/fonts';
import type { ColorResolver } from '@/ooxml/drawingml/colors';
import type { CoreProperties } from '@/opc';
import type { DocumentInfo, SignatureOptions, StyledRenderOptions } from '@/pdf';
import { FontRegistry } from '@/font';
import { fetchFontSet, resolveFamilyKey } from '@/fonts';
import { parseChart } from '@/ooxml/drawingml/chart-parser';
import { DEFAULT_THEME_PALETTE, makeColorResolver } from '@/ooxml/drawingml/colors';
import { parseTheme } from '@/ooxml/drawingml/theme-parser';
import { OpcPackage, parseCoreProperties } from '@/opc';
import {
  EMPTY_NUMBERING,
  EMPTY_SECTION,
  EMPTY_SETTINGS,
  EMPTY_STYLE_SHEET,
  loadEmbeddedFonts,
  parseDocument,
  parseHeaderFooter,
  parseNumbering,
  parseSections,
  parseSettings,
  parseStyles,
} from '@/ooxml/wordproc';
import { renderStyledPdf, signPdf } from '@/pdf';

const STYLES_PART = 'word/styles.xml';
const NUMBERING_PART = 'word/numbering.xml';
const SETTINGS_PART = 'word/settings.xml';
const CORE_PROPS_PART = 'docProps/core.xml';
const MAIN_DOCUMENT_PART = 'word/document.xml';

const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const THEME_PART = 'word/theme/theme1.xml';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';

export interface ConvertDocxOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  readonly fontBytes?: Uint8Array;
  readonly fonts?: FontBytesByVariant;
  // Force a substitute font family for the auto-download path (e.g. 'serif').
  // Ignored when `fonts`/`fontBytes` are supplied.
  readonly fontFamily?: string;
  // Injectable fetch for the auto-download path (defaults to global fetch).
  readonly fontFetch?: FetchLike;
  // For PDF/A-3 only: embed the input .docx as an associated source file
  // (/AFRelationship /Source) so the archive carries its own source. Ignored
  // for other profiles (PDF/A-1/2 forbid arbitrary embedded files).
  readonly embedSource?: boolean;
  // Digitally sign the output (ISO 32000 §12.8). Requires the async
  // convertDocxToPdf (signing uses WebCrypto). Ignored by the sync converter.
  readonly signature?: SignatureOptions;
}

// Synchronous conversion. Requires the caller to supply `fonts`/`fontBytes`
// (no network). Use the async `convertDocxToPdf` for automatic font download.
export function convertDocxToPdfSync(docx: Uint8Array, options: ConvertDocxOptions): Uint8Array {
  const fonts: FontBytesByVariant | undefined =
    options.fonts ?? (options.fontBytes ? { regular: options.fontBytes } : undefined);
  if (!fonts) {
    throw new Error('convertDocxToPdfSync requires options.fonts or options.fontBytes');
  }
  const registry = FontRegistry.fromBytes(fonts);

  const pkg = OpcPackage.open(docx);
  const main = pkg.getMainDocument();
  // Theme-backed colour resolver (schemeClr → hex); falls back to the built-in
  // Office palette when there is no theme part.
  const resolveColor = buildColorResolver(pkg);
  const body = parseDocument(main.data, resolveColor);
  const rawSections = parseSections(main.data);

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseStyles(stylesData) : EMPTY_STYLE_SHEET;

  const numberingData = pkg.getPart(NUMBERING_PART);
  const numbering = numberingData ? parseNumbering(numberingData) : EMPTY_NUMBERING;

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

  const headersFooters = loadHeadersFootersForSections(pkg, sections, resolveColor);
  const images = loadImages(pkg);
  const charts = loadCharts(pkg, resolveColor);
  // The document's own embedded fonts (de-obfuscated). A run whose w:ascii
  // matches one renders with the real font instead of a substitute.
  const embeddedFonts = loadEmbeddedFonts(pkg);

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;

  const {
    fontBytes: _ignoreA,
    fonts: _ignoreB,
    info: callerInfo,
    embedSource,
    attachments: callerAttachments,
    signature: _ignoreSig,
    ...renderOptions
  } = options;
  void _ignoreA;
  void _ignoreB;
  void _ignoreSig;
  const info = mergeInfo(coreProps, callerInfo);
  // Document language for the tagged-PDF /Lang (caller's options.language, in
  // renderOptions, still wins by being spread last).
  const language = detectDocxLanguage(stylesData, main.data);
  // PDF/A-3 only: optionally embed the input .docx as an associated source file.
  const attachments = [...(callerAttachments ?? [])];
  if (embedSource && options.pdfA?.startsWith('PDF/A-3')) {
    attachments.push({
      name: 'source.docx',
      bytes: docx,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      relationship: 'Source',
      description: 'Source Word document',
    });
  }
  return renderStyledPdf(body, {
    registry,
    styles,
    numbering,
    sections,
    headersFooters,
    images,
    charts,
    ...(embeddedFonts.size > 0 ? { embeddedFonts } : {}),
    ...(info ? { info } : {}),
    ...(language ? { language } : {}),
    ...(attachments.length > 0 ? { attachments } : {}),
    ...renderOptions,
  });
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

// Convert a .docx to PDF, downloading an open substitute font automatically
// when the caller does not supply one. With `fonts`/`fontBytes` set, no network
// access occurs (it just delegates to the synchronous path).
// Distinct curated families the document references — always includes the sans
// default as a fallback for unstyled runs / math / charts. The async path
// fetches one substitute set per family so each run renders in the right one.
export function detectDocxFamilyKeys(docx: Uint8Array): Set<FamilyKey> {
  const keys = new Set<FamilyKey>(['roboto']);
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.open(docx);
  } catch {
    return keys;
  }
  const decoder = new TextDecoder('utf-8');
  const re = /<w:rFonts[^>]*\bw:ascii="([^"]+)"/g;
  for (const part of [STYLES_PART, MAIN_DOCUMENT_PART]) {
    const data = pkg.getPart(part);
    if (!data) continue;
    const xml = decoder.decode(data);
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) keys.add(resolveFamilyKey(m[1]));
  }
  return keys;
}

export async function convertDocxToPdf(
  docx: Uint8Array,
  options: ConvertDocxOptions = {},
): Promise<Uint8Array> {
  // A signature (SignatureOptions) is also a SignaturePlaceholder, so it doubles
  // as the render-time placeholder; the crypto runs afterwards on the result.
  const { signature } = options;
  const renderOptions: ConvertDocxOptions = signature
    ? { ...options, signaturePlaceholder: signature }
    : options;
  const pdf = await buildUnsignedDocxPdf(docx, renderOptions);
  return signature ? signPdf(pdf, signature) : pdf;
}

async function buildUnsignedDocxPdf(
  docx: Uint8Array,
  options: ConvertDocxOptions,
): Promise<Uint8Array> {
  if (options.fonts ?? options.fontBytes) {
    return convertDocxToPdfSync(docx, options);
  }
  const fetchOpt = options.fontFetch ? { fetch: options.fontFetch } : {};
  const keys = options.fontFamily
    ? new Set<FamilyKey>([resolveFamilyKey(options.fontFamily)])
    : detectDocxFamilyKeys(docx);

  // Single family (e.g. an all-sans document) → simple one-family path.
  if (keys.size <= 1) {
    const family = keys.values().next().value;
    const fonts = await fetchFontSet({ ...(family ? { family } : {}), ...fetchOpt });
    return convertDocxToPdfSync(docx, { ...options, fonts });
  }

  // Multiple families → fetch a substitute set for each, resolve per run.
  const registriesByFamily = new Map<FamilyKey, FontRegistry>();
  let baseBytes: FontBytesByVariant | undefined;
  for (const key of keys) {
    const bytes = await fetchFontSet({ family: key, ...fetchOpt });
    registriesByFamily.set(key, FontRegistry.fromBytes(bytes));
    if (key === 'roboto' || !baseBytes) baseBytes = bytes;
  }
  return convertDocxToPdfSync(docx, { ...options, fonts: baseBytes!, registriesByFamily });
}

// Best-effort detection of the document's primary font family, used to choose
// a substitute for auto-download. Prefers the document defaults' ascii font,
// then falls back to the most frequent run font. Cheap regex over the XML —
// no need to fully parse for this hint.
export function detectDocxFontFamily(docx: Uint8Array): string | undefined {
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.open(docx);
  } catch {
    return undefined;
  }
  const decoder = new TextDecoder('utf-8');
  const styles = pkg.getPart(STYLES_PART);
  if (styles) {
    const xml = decoder.decode(styles);
    const def = /<w:docDefaults>[\s\S]*?<w:rFonts[^>]*\bw:ascii="([^"]+)"/.exec(xml);
    if (def?.[1]) return def[1];
  }
  const main = pkg.getPart(MAIN_DOCUMENT_PART);
  if (main) {
    const xml = decoder.decode(main);
    const counts = new Map<string, number>();
    const re = /<w:rFonts[^>]*\bw:ascii="([^"]+)"/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) {
      const name = m[1]!;
      counts.set(name, (counts.get(name) ?? 0) + 1);
    }
    let best: string | undefined;
    let bestCount = 0;
    for (const [name, count] of counts) {
      if (count > bestCount) {
        best = name;
        bestCount = count;
      }
    }
    return best;
  }
  return undefined;
}

// Caller-provided `info` overrides any value from the docx core properties.
function mergeInfo(
  core: CoreProperties | undefined,
  caller: DocumentInfo | undefined,
): DocumentInfo | undefined {
  if (!core && !caller) return undefined;
  return {
    ...(core?.title ? { title: core.title } : {}),
    ...(core?.creator ? { author: core.creator } : {}),
    ...(core?.subject ? { subject: core.subject } : {}),
    ...(core?.keywords ? { keywords: core.keywords } : {}),
    ...(core?.created ? { creationDate: core.created } : {}),
    ...(core?.modified ? { modificationDate: core.modified } : {}),
    ...caller,
  };
}

function loadImages(pkg: OpcPackage): ReadonlyMap<string, Uint8Array> {
  const rels = pkg.getPartRelationships(MAIN_DOCUMENT_PART);
  if (rels.length === 0) return new Map();
  const out = new Map<string, Uint8Array>();
  for (const rel of rels) {
    if (rel.type !== REL_IMAGE) continue;
    const resolved = pkg.resolveRelatedPart(MAIN_DOCUMENT_PART, rel);
    if (!resolved) continue;
    out.set(rel.id, resolved.data);
  }
  return out;
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
  resolveColor: ColorResolver,
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
    out.set(rel.id, parseHeaderFooter(resolved.data, resolveColor));
  }
  // EMPTY_SECTION reference avoids "unused import" complaints in builds where
  // the variable is otherwise unused.
  void EMPTY_SECTION;
  return out;
}
