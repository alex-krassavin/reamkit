import type { BodyElement } from '@/core/document-model';
// xlsx → PDF convenience converters: font acquisition (caller-supplied or
// auto-downloaded), optional signing, PDF/A-3 source embedding. All document
// understanding lives in the xlsx reader and the print model.

import type { FontBytesByVariant } from '@/core/font';
import type { FetchLike } from '@/core/fonts';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { FontRegistry } from '@/core/font';
import { fetchFontSet } from '@/core/fonts';
import { flowRenderOptions } from '@/core/converter/project';
import { readXlsx } from '@/excel/xlsx-reader';
import { renderStyledPdf, renderStyledPdfEncrypted, signPdf } from '@/pdf';

/**
 * Options for the xlsx → PDF convenience converters. Extends the low-level
 * {@link StyledRenderOptions} (minus the font `registry` and `styles`, which the
 * converter builds itself) with font acquisition and source-touching conveniences.
 */
export interface ConvertXlsxOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  /** Shorthand for supplying a single regular-variant font as raw bytes. */
  readonly fontBytes?: Uint8Array;
  /** Explicit font bytes per variant (regular/bold/italic/bold-italic). */
  readonly fonts?: FontBytesByVariant;
  /**
   * Force a substitute font family for the auto-download path. Ignored when
   * `fonts`/`fontBytes` are supplied.
   */
  readonly fontFamily?: string;
  /** Injectable `fetch` for the auto-download path (defaults to the global `fetch`). */
  readonly fontFetch?: FetchLike;
  /**
   * For PDF/A-3 only: embed the input `.xlsx` as an associated source file
   * (`/AFRelationship /Source`). Ignored for other profiles.
   */
  readonly embedSource?: boolean;
  /**
   * E-SHEET W9 — the reference date for conditional-format `timePeriod` rules and
   * `TODAY()`/`NOW()` in `expression` rules. An explicit input (never the wall
   * clock), so output stays deterministic; omitted ⇒ those clock-relative rules
   * no-op.
   */
  readonly now?: Date;
  /**
   * Digitally sign the output (ISO 32000 §12.8). Requires the async
   * {@link convertXlsxToPdf} (signing uses WebCrypto). Ignored by the sync
   * converter.
   */
  readonly signature?: SignatureOptions;
}

/**
 * Convert a .xlsx to PDF in one shot.
 *
 * @deprecated Use `Ream.parse(bytes).convert('pdf', options)` — one parse,
 * any number of targets. This function remains for 0.1.x compatibility.
 */
export async function convertXlsxToPdf(
  xlsx: Uint8Array,
  options: ConvertXlsxOptions = {},
): Promise<Uint8Array> {
  // A signature (SignatureOptions) is also a SignaturePlaceholder, so it doubles
  // as the render-time placeholder; the crypto runs afterwards on the result.
  const { signature } = options;
  const renderOptions: ConvertXlsxOptions = signature
    ? { ...options, signaturePlaceholder: signature }
    : options;
  const withFonts =
    (renderOptions.fonts ?? renderOptions.fontBytes)
      ? renderOptions
      : {
          ...renderOptions,
          fonts: await fetchFontSet({
            ...(options.fontFamily ? { family: options.fontFamily } : {}),
            ...(options.fontFetch ? { fetch: options.fontFetch } : {}),
          }),
        };
  // §7.6: encryption needs WebCrypto — async path only.
  if (withFonts.encrypt) {
    const args = prepareXlsxStyledRender(xlsx, withFonts);
    return renderStyledPdfEncrypted(args.body, args.styled);
  }
  const pdf = convertXlsxToPdfSync(xlsx, withFonts);
  return signature ? signPdf(pdf, signature) : pdf;
}

/**
 * Synchronous one-shot conversion (requires `fonts`/`fontBytes`).
 *
 * Internal since 1.0 — see the async variant above.
 */
export function convertXlsxToPdfSync(xlsx: Uint8Array, options: ConvertXlsxOptions): Uint8Array {
  const args = prepareXlsxStyledRender(xlsx, options);
  return renderStyledPdf(args.body, args.styled);
}

// The shared xlsx → styled-render arguments; see the docx twin.
function prepareXlsxStyledRender(
  xlsx: Uint8Array,
  options: ConvertXlsxOptions,
): { body: ReadonlyArray<BodyElement>; styled: StyledRenderOptions } {
  const fonts: FontBytesByVariant | undefined =
    options.fonts ?? (options.fontBytes ? { regular: options.fontBytes } : undefined);
  if (!fonts) {
    throw new Error('convertXlsxToPdfSync requires options.fonts or options.fontBytes');
  }
  const registry = FontRegistry.fromBytes(fonts);

  // All document-derived state now comes from the xlsx reader (ir-design §7). The
  // reference date (W9) feeds the conditional-format formula engine during the
  // SheetDoc → FlowDoc projection.
  const { doc: flow } = readXlsx(xlsx, options.now !== undefined ? { now: options.now } : {});

  const {
    fontBytes: _ignoreA,
    fonts: _ignoreB,
    section: sectionOverride,
    info: callerInfo,
    embedSource,
    now: _ignoreNow,
    attachments: callerAttachments,
    signature: _ignoreSig,
    // Spreadsheet geometry (row heights / column widths) is governed by the Calc
    // print model, not flowing-text leading — so the renderer-compat
    // layoutProfile (E-PARITY) does not apply to xlsx; drop it. Empirically the
    // flat row model already tracks Calc more closely than a font-metric one.
    layoutProfile: _ignoreProfile,
    ...renderOptions
  } = options;
  void _ignoreSig;
  void _ignoreA;
  void _ignoreB;
  void _ignoreProfile;
  void _ignoreNow;
  const section = sectionOverride ?? flow.section;
  // Caller overrides spread over the document's own metadata.
  const info = flow.info || callerInfo ? { ...flow.info, ...callerInfo } : undefined;
  // PDF/A-3 only: optionally embed the input .xlsx as an associated source file.
  const attachments = [...(callerAttachments ?? [])];
  if (embedSource && options.pdfA?.startsWith('PDF/A-3')) {
    attachments.push({
      name: 'source.xlsx',
      bytes: xlsx,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      relationship: 'Source',
      description: 'Source Excel workbook',
    });
  }
  return {
    body: flow.body,
    styled: {
      registry,
      ...flowRenderOptions(flow),
      ...(section ? { section } : {}),
      ...(info ? { info } : {}),
      ...(attachments.length > 0 ? { attachments } : {}),
      ...renderOptions,
    },
  };
}
