import type { FontBytesByVariant } from '@/core/font';
import type { FamilyKey, FetchLike } from '@/core/fonts';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { FontRegistry } from '@/core/font';
import { fetchFontSet, resolveFamilyKey } from '@/core/fonts';
import { OpcPackage } from '@/core/opc';
import { readDocx } from '@/word/docx-reader';
import { renderStyledPdf, signPdf } from '@/pdf';

const STYLES_PART = 'word/styles.xml';
const MAIN_DOCUMENT_PART = 'word/document.xml';

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

  // All document-derived state now comes from the docx reader (ir-design §7).
  const { doc: flow } = readDocx(docx);

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
  // Caller overrides spread over the document's own metadata.
  const info = flow.info || callerInfo ? { ...flow.info, ...callerInfo } : undefined;
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
  return renderStyledPdf(flow.body, {
    registry,
    styles: flow.styles,
    ...(flow.numbering ? { numbering: flow.numbering } : {}),
    sections: flow.sections,
    ...(flow.headersFooters ? { headersFooters: flow.headersFooters } : {}),
    resources: flow.resources,
    ...(flow.charts ? { charts: flow.charts } : {}),
    ...(flow.embeddedFonts ? { embeddedFonts: flow.embeddedFonts } : {}),
    ...(info ? { info } : {}),
    ...(flow.language ? { language: flow.language } : {}),
    ...(attachments.length > 0 ? { attachments } : {}),
    ...renderOptions,
  });
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
