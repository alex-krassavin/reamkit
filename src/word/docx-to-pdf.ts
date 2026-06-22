import type { BodyElement } from '@/core/document-model';
import type { FontBytesByVariant } from '@/core/font';
import type { FamilyKey, FetchLike } from '@/core/fonts';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { FontRegistry } from '@/core/font';
import { fetchFontSet, resolveFamilyKey } from '@/core/fonts';
import { OpcPackage } from '@/core/opc';
import { flowRenderOptions } from '@/core/converter/project';
import { readDocx } from '@/word/docx-reader';
import { renderStyledPdf, renderStyledPdfEncrypted, signPdf } from '@/pdf';

const STYLES_PART = 'word/styles.xml';
const MAIN_DOCUMENT_PART = 'word/document.xml';

/**
 * Options for the `.docx` → PDF converters. Extends the low-level
 * {@link StyledRenderOptions} (minus the font `registry` and `styles`, which the
 * converter builds itself) with font supply, substitute-font hints and two
 * source-touching conveniences.
 */
export interface ConvertDocxOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  /** A single regular-variant font as raw bytes. */
  readonly fontBytes?: Uint8Array;
  /** Explicit font bytes per variant (regular/bold/italic/bold-italic). */
  readonly fonts?: FontBytesByVariant;
  /**
   * Force a substitute font family for the auto-download path (e.g. `'serif'`).
   * Ignored when `fonts`/`fontBytes` are supplied.
   */
  readonly fontFamily?: string;
  /** Injectable `fetch` for the auto-download path (defaults to the global `fetch`). */
  readonly fontFetch?: FetchLike;
  /**
   * For PDF/A-3 only: embed the input `.docx` as an associated source file
   * (`/AFRelationship /Source`) so the archive carries its own source. Ignored
   * for other profiles (PDF/A-1/2 forbid arbitrary embedded files).
   */
  readonly embedSource?: boolean;
  /**
   * Digitally sign the output (ISO 32000 §12.8). Requires the async
   * {@link convertDocxToPdf} (signing uses WebCrypto); ignored by the sync converter.
   */
  readonly signature?: SignatureOptions;
}

/**
 * Synchronous one-shot conversion (requires `fonts`/`fontBytes`, no network).
 *
 * Internal since 1.0 — see the async variant above.
 */
export function convertDocxToPdfSync(docx: Uint8Array, options: ConvertDocxOptions): Uint8Array {
  const args = prepareDocxStyledRender(docx, options);
  return renderStyledPdf(args.body, args.styled);
}

// The shared docx → styled-render arguments (fonts, FlowDoc projection, info,
// PDF/A-3 source embedding): the sync converter renders them directly, the
// async path may route them through the encrypting renderer.
function prepareDocxStyledRender(
  docx: Uint8Array,
  options: ConvertDocxOptions,
): { body: ReadonlyArray<BodyElement>; styled: StyledRenderOptions } {
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
  return {
    body: flow.body,
    styled: {
      registry,
      ...flowRenderOptions(flow),
      ...(info ? { info } : {}),
      ...(attachments.length > 0 ? { attachments } : {}),
      ...renderOptions,
    },
  };
}

/**
 * The distinct curated font families the document references (scanning the
 * `w:ascii` of `styles.xml` + `document.xml`), always including the sans default
 * `'arimo'` as a fallback for unstyled runs / math / charts. The async path
 * fetches one substitute set per family so each run renders in the right one.
 *
 * @param docx The `.docx` bytes.
 * @returns The resolved {@link FamilyKey} set; just the default when the package
 *   cannot be opened.
 */
export function detectDocxFamilyKeys(docx: Uint8Array): Set<FamilyKey> {
  const keys = new Set<FamilyKey>(['arimo']);
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

/**
 * One-shot .docx → PDF (auto-downloads a substitute font set when the caller
 * supplies none; zero network with `fonts`/`fontBytes`).
 *
 * Internal since 1.0 — the public entry is `Ream.parse(bytes).convert('pdf')`;
 * the createConverter facade and the test suite drive this directly.
 */
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
  const withFonts =
    (options.fonts ?? options.fontBytes)
      ? options
      : { ...options, ...(await resolveDocxAutoFonts(docx, options)) };
  // §7.6: encryption needs WebCrypto, so it lives on this async path — the
  // sync core would throw on it.
  if (withFonts.encrypt) {
    const args = prepareDocxStyledRender(docx, withFonts);
    return renderStyledPdfEncrypted(args.body, args.styled);
  }
  return convertDocxToPdfSync(docx, withFonts);
}

/**
 * Substitute-font auto-download for a `.docx` without caller-supplied fonts:
 * detect the families used and fetch an open set per family (a single-family
 * document takes the simple one-family path). Shared by the converter and the
 * {@link Ream} facade — both work from the same source bytes.
 *
 * @param docx    The `.docx` bytes.
 * @param options Optional `fontFamily` override and injectable `fontFetch`.
 * @returns The base font bytes plus, for a multi-family document, a per-family
 *   registry map so each run resolves to its own substitute.
 */
export async function resolveDocxAutoFonts(
  docx: Uint8Array,
  options: { readonly fontFamily?: string; readonly fontFetch?: FetchLike } = {},
): Promise<{
  fonts: FontBytesByVariant;
  registriesByFamily?: Map<FamilyKey, FontRegistry>;
}> {
  const fetchOpt = options.fontFetch ? { fetch: options.fontFetch } : {};
  const keys = options.fontFamily
    ? new Set<FamilyKey>([resolveFamilyKey(options.fontFamily)])
    : detectDocxFamilyKeys(docx);

  // Single family (e.g. an all-sans document) → simple one-family path.
  if (keys.size <= 1) {
    const family = keys.values().next().value;
    return { fonts: await fetchFontSet({ ...(family ? { family } : {}), ...fetchOpt }) };
  }

  // Multiple families → fetch a substitute set for each, resolve per run.
  const registriesByFamily = new Map<FamilyKey, FontRegistry>();
  let baseBytes: FontBytesByVariant | undefined;
  for (const key of keys) {
    const bytes = await fetchFontSet({ family: key, ...fetchOpt });
    registriesByFamily.set(key, FontRegistry.fromBytes(bytes));
    if (key === 'arimo' || !baseBytes) baseBytes = bytes;
  }
  return { fonts: baseBytes!, registriesByFamily };
}

/**
 * Best-effort detection of the document's primary font family, used to choose a
 * substitute for auto-download. Prefers the document defaults' `w:ascii` font,
 * then falls back to the most frequent run font. A cheap regex over the XML — no
 * need to fully parse for this hint.
 *
 * @param docx The `.docx` bytes.
 * @returns The detected family name, or `undefined` when none is found.
 */
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
