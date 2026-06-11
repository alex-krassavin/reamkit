// xlsx → PDF convenience converters: font acquisition (caller-supplied or
// auto-downloaded), optional signing, PDF/A-3 source embedding. All document
// understanding lives in the xlsx reader and the print model.

import type { FontBytesByVariant } from '@/core/font';
import type { FetchLike } from '@/core/fonts';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { FontRegistry } from '@/core/font';
import { fetchFontSet } from '@/core/fonts';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';
import { flowRenderOptions } from '@/core/converter/project';
import { readXlsx } from '@/excel/xlsx-reader';
import { renderStyledPdf, signPdf } from '@/pdf';

export interface ConvertXlsxOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  readonly fontBytes?: Uint8Array;
  readonly fonts?: FontBytesByVariant;
  // Force a substitute font family for the auto-download path. Ignored when
  // `fonts`/`fontBytes` are supplied.
  readonly fontFamily?: string;
  // Injectable fetch for the auto-download path (defaults to global fetch).
  readonly fontFetch?: FetchLike;
  // For PDF/A-3 only: embed the input .xlsx as an associated source file
  // (/AFRelationship /Source). Ignored for other profiles.
  readonly embedSource?: boolean;
  // Digitally sign the output (ISO 32000 §12.8). Requires the async
  // convertXlsxToPdf (signing uses WebCrypto). Ignored by the sync converter.
  readonly signature?: SignatureOptions;
}

// Convert a .xlsx to PDF, downloading an open substitute font automatically
// when the caller does not supply one (no network when `fonts` is set).
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
  let pdf: Uint8Array;
  if (renderOptions.fonts ?? renderOptions.fontBytes) {
    pdf = convertXlsxToPdfSync(xlsx, renderOptions);
  } else {
    const fonts = await fetchFontSet({
      ...(options.fontFamily ? { family: options.fontFamily } : {}),
      ...(options.fontFetch ? { fetch: options.fontFetch } : {}),
    });
    pdf = convertXlsxToPdfSync(xlsx, { ...renderOptions, fonts });
  }
  return signature ? signPdf(pdf, signature) : pdf;
}

// Synchronous conversion. Requires the caller to supply `fonts`/`fontBytes`.
export function convertXlsxToPdfSync(xlsx: Uint8Array, options: ConvertXlsxOptions): Uint8Array {
  const fonts: FontBytesByVariant | undefined =
    options.fonts ?? (options.fontBytes ? { regular: options.fontBytes } : undefined);
  if (!fonts) {
    throw new Error('convertXlsxToPdfSync requires options.fonts or options.fontBytes');
  }
  const registry = FontRegistry.fromBytes(fonts);

  // All document-derived state now comes from the xlsx reader (ir-design §7).
  const { doc: flow } = readXlsx(xlsx);

  const {
    fontBytes: _ignoreA,
    fonts: _ignoreB,
    section: sectionOverride,
    info: callerInfo,
    embedSource,
    attachments: callerAttachments,
    signature: _ignoreSig,
    ...renderOptions
  } = options;
  void _ignoreSig;
  void _ignoreA;
  void _ignoreB;
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
  return renderStyledPdf(flow.body, {
    registry,
    ...flowRenderOptions(flow),
    ...(section ? { section } : {}),
    ...(info ? { info } : {}),
    ...(attachments.length > 0 ? { attachments } : {}),
    ...renderOptions,
  });
}
