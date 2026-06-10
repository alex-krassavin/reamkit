// Public API for Ream.
//
// The package converts Word (.docx) and Excel (.xlsx) documents to PDF,
// implemented from the ECMA-376 and ISO 32000 specifications without wrapping
// LibreOffice, headless Office, or any commercial SDK.
//
// Typical use:
//
//   import { convertDocxToPdf } from 'reamkit';
//   const pdf = convertDocxToPdf(docxBytes, { fonts: { regular: robotoBytes } });
//
// Advanced consumers can drive the layout engine directly via renderStyledPdf,
// supply a hyphenator, or import the typed document model from the
// "reamkit/document-model" subpath.

// --- Converters (the main entry points) ---
// convertDocxToPdf / convertXlsxToPdf are async and download an open substitute
// font automatically when none is supplied. The *Sync variants are synchronous
// and require the caller to pass fonts.
export {
  convertDocxToPdf,
  convertXlsxToPdf,
  convertDocxToPdfSync,
  convertXlsxToPdfSync,
} from '@/converter';
export type { ConvertDocxOptions, ConvertXlsxOptions } from '@/converter';

// --- Remote fonts (used by the auto-download path; exported for customisation) ---
export { fetchFontSet, resolveFamilyKey } from '@/fonts';
export type { FamilyKey, FetchFontSetOptions, FetchLike } from '@/fonts';

// --- Low-level layout/PDF engine (for advanced rendering pipelines) ---
export { renderStyledPdf } from '@/pdf';
export type { StyledRenderOptions, DocumentInfo } from '@/pdf';

// --- Digital signatures (ISO 32000 §12.8) ---
export { signPdf } from '@/pdf';
export type { SignaturePlaceholder, SignerCredentials, SignatureOptions } from '@/pdf';

// --- Fonts (needed to build the `fonts` option, plus advanced font handling) ---
export { FontRegistry, parseTtf, subsetTtf } from '@/font';
export type { FontBytesByVariant, FontVariant, ParsedTtf } from '@/font';

// --- Hyphenation (opt-in; pass the result via options.hyphenator) ---
export {
  getHyphenator,
  createLanguageHyphenator,
  createHyphenator,
  splitPatternBundle,
} from '@/hyphenation';
export type { Hyphenator, HyphenatorOptions, SupportedLanguage } from '@/hyphenation';

// --- IR layer (@experimental — ir-design.md; may change in minor versions
// until the schema freezes against three-plus adapters) ---
export type {
  Pt,
  ResourceId,
  Feature,
  KnownFeature,
  Loss,
  LossReport,
  LossSeverity,
  NativeBag,
} from '@/ir';
export {
  FEATURES,
  ResourceStore,
  ConversionLossError,
  formatLoss,
  pt,
  twipsToPt,
  halfPtToPt,
  eighthPtToPt,
  emuToPt,
  pxToPt,
  inchToPt,
  mmToPt,
} from '@/ir';
export type { FlowDoc } from '@/ir/flow';
export type {
  DocumentReader,
  DocumentWriter,
  ReadOptions,
  ReadResult,
  WriteOptions,
  WriteResult,
} from '@/ir/adapters';
export { docxReader, readDocx } from '@/readers/docx-reader';
export { xlsxReader, readXlsx } from '@/readers/xlsx-reader';
export { createConverter } from '@/converter/facade';
export type {
  Converter,
  ConvertOptions,
  ConvertResult,
  CreateConverterOptions,
} from '@/converter/facade';
export type {
  PageItem,
  PageItemBase,
  TextLineItem,
  BorderItem,
  FillItem,
  ImageItem,
  ShapeItem,
  LaidOutPage,
} from '@/pdf/styled-page-renderer';
export {
  chainProviders,
  callerFontProvider,
  embeddedDocFontProvider,
  remoteFontProvider,
  localFontProvider,
  readOs2FsType,
  isEmbeddingRestricted,
  NO_FONT,
} from '@/fonts/provider';
export type { FontProvider, FontRequest, FontAnswer } from '@/fonts/provider';
