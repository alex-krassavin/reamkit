// Public API for Ream.
//
// The package converts Word (.docx) and Excel (.xlsx) documents to PDF (and an
// SVG page preview), implemented from the ECMA-376 and ISO 32000 specifications
// without wrapping LibreOffice, headless Office, or any commercial SDK.
//
// Typical use:
//
//   import { Ream } from 'reamkit';
//   const doc = Ream.parse(bytes);            // docx or xlsx — sniffed
//   const pdf = await doc.convert('pdf');     // one parse, any target
//
// Advanced consumers can build custom pipelines from the reader/writer
// interfaces, drive the layout engine directly via renderStyledPdf, supply a
// hyphenator, or import the typed document model from the
// "reamkit/document-model" subpath.

// --- Remote fonts (used by the auto-download path; exported for customisation) ---
export { fetchFontSet, resolveFamilyKey } from '@/core/fonts';
export type { FamilyKey, FetchFontSetOptions, FetchLike } from '@/core/fonts';

// --- Low-level layout/PDF engine (for advanced rendering pipelines) ---
export { renderStyledPdf } from '@/pdf';
export type { StyledRenderOptions, DocumentInfo } from '@/pdf';

// --- Digital signatures (ISO 32000 §12.8) ---
export { signPdf } from '@/pdf';
export type { SignaturePlaceholder, SignerCredentials, SignatureOptions } from '@/pdf';

// --- Fonts (needed to build the `fonts` option, plus advanced font handling) ---
export { FontRegistry, parseTtf, subsetTtf } from '@/core/font';
export type { FontBytesByVariant, FontVariant, ParsedTtf } from '@/core/font';

// --- Hyphenation (opt-in; pass the result via options.hyphenator) ---
export {
  getHyphenator,
  createLanguageHyphenator,
  createHyphenator,
  splitPatternBundle,
} from '@/core/hyphenation';
export type { Hyphenator, HyphenatorOptions, SupportedLanguage } from '@/core/hyphenation';

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
} from '@/core/ir';
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
} from '@/core/ir';
export type { FlowDoc } from '@/core/ir/flow';
export type {
  DocumentReader,
  DocumentWriter,
  ReadOptions,
  ReadResult,
  WriteOptions,
  WriteResult,
} from '@/core/ir/adapters';
export { docxReader, readDocx } from '@/word/docx-reader';
export { xlsxReader, readXlsx } from '@/excel/xlsx-reader';
// The object face: parse once into the FlowDoc interlayer, convert many.
export { Ream } from '@/core/converter/ream';
export type { ReamConvertOptions, ReamParseOptions, ReamTarget } from '@/core/converter/ream';
export { createConverter } from '@/core/converter/facade';
export type {
  Converter,
  ConvertOptions,
  ConvertResult,
  CreateConverterOptions,
} from '@/core/converter/facade';
export type {
  PageItem,
  PageItemBase,
  TextLineItem,
  BorderItem,
  FillItem,
  ImageItem,
  ShapeItem,
  LaidOutPage,
} from '@/layout/page-doc';
export {
  chainProviders,
  callerFontProvider,
  embeddedDocFontProvider,
  remoteFontProvider,
  localFontProvider,
  readOs2FsType,
  isEmbeddingRestricted,
  NO_FONT,
} from '@/core/fonts/provider';
export type { FontProvider, FontRequest, FontAnswer } from '@/core/fonts/provider';
export { svgWriter, writeSvg } from '@/svg/svg-writer';
export type { SvgWriteOptions } from '@/svg/svg-writer';
export { layoutStyledDocument } from '@/layout/styled-layout';
export type { LaidOutDocument } from '@/layout/page-doc';
