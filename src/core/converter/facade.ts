// Conversion facade (ir-design §7): the registry-driven entry point that
// composes readers → layout → writers. v0 dispatches format detection through
// reader sniffing and delegates the PDF pipeline to the existing converters
// (which themselves run on the readers) — so the facade is byte-identical to
// calling convertDocxToPdf/convertXlsxToPdf directly.
//
// The async boundary lives HERE (font fetching); readers/writers stay sync.

import type { ConvertDocxOptions } from '@/word/docx-to-pdf';
import type { DocumentReader } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { FontProvider } from '@/core/fonts/provider';
import type { Loss, LossReport } from '@/core/ir';
import type { FontBytesByVariant } from '@/core/font';

import { ConversionLossError, FEATURES } from '@/core/ir';
import { FontRegistry } from '@/core/font';
import { chainProviders } from '@/core/fonts/provider';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/pdf/styled-page-renderer';
import { writeSvg } from '@/svg/svg-writer';
import { convertDocxToPdf } from '@/word/docx-to-pdf';
import { convertXlsxToPdf } from '@/excel/xlsx-to-pdf';
import { docxReader } from '@/word/docx-reader';
import { xlsxReader } from '@/excel/xlsx-reader';

export interface ConvertOptions extends ConvertDocxOptions {
  /** Target format: 'pdf' (default) or 'svg' (page-stack preview). */
  readonly to?: 'pdf' | 'svg';
  /**
   * Strict mode (handoff v1 §5): throw ConversionLossError on the first
   * recorded loss instead of returning it in the report.
   */
  readonly strict?: boolean;
  /**
   * Font resolution chain (ir-design §8). When set (and no caller `fonts`),
   * the facade resolves the default font set through these providers — e.g.
   * [localFontProvider(), remoteFontProvider()] — and reports a 'substituted'
   * loss when anything below a caller/embedded answer wins. Local Font Access
   * stays strictly opt-in via this option (it can trigger a permission
   * prompt), which is why nothing wires it in by default.
   */
  readonly fontProviders?: ReadonlyArray<FontProvider>;
}

export interface ConvertResult {
  readonly bytes: Uint8Array;
  readonly losses: LossReport;
}

export interface Converter {
  /** The registered readers, in sniffing order. */
  readonly readers: ReadonlyArray<DocumentReader<FlowDoc>>;
  /** Detect the input format by reader sniffing; undefined when unknown. */
  detect: (bytes: Uint8Array) => DocumentReader<FlowDoc> | undefined;
  convert: (bytes: Uint8Array, options?: ConvertOptions) => Promise<ConvertResult>;
}

export interface CreateConverterOptions {
  /** Override / extend the reader registry (defaults to docx + xlsx). */
  readonly readers?: ReadonlyArray<DocumentReader<FlowDoc>>;
}

export const DEFAULT_READERS: ReadonlyArray<DocumentReader<FlowDoc>> = [docxReader, xlsxReader];

export function createConverter(opts: CreateConverterOptions = {}): Converter {
  const readers = opts.readers ?? DEFAULT_READERS;

  const detect = (bytes: Uint8Array): DocumentReader<FlowDoc> | undefined =>
    readers.find((r) => r.sniff(bytes));

  const convert = async (
    bytes: Uint8Array,
    options: ConvertOptions = {},
  ): Promise<ConvertResult> => {
    const { to = 'pdf', strict = false, fontProviders, ...rest } = options;
    const reader = detect(bytes);
    if (!reader) {
      throw new Error('Unrecognized input format (no registered reader sniffs these bytes)');
    }
    const losses: Array<Loss> = [];
    let conv = rest;
    // Resolve the default font set through the provider chain (unless the
    // caller supplied bytes directly — those always win).
    if (fontProviders && fontProviders.length > 0 && !rest.fonts && !rest.fontBytes) {
      const { fonts, loss } = await resolveFontsViaChain(fontProviders);
      if (fonts) conv = { ...rest, fonts };
      if (loss) losses.push(loss);
    }
    if (to === 'svg') {
      // FlowDoc → layout → svg writer: the stage-6 pipeline, no PDF involved.
      const fonts = conv.fonts ?? (conv.fontBytes ? { regular: conv.fontBytes } : undefined);
      if (!fonts) {
        throw new Error("to: 'svg' requires options.fonts/fontBytes or fontProviders");
      }
      const { doc: flow } = reader.read(bytes);
      const laid = layoutStyledDocument(flow.body, {
        registry: FontRegistry.fromBytes(fonts),
        ...flowRenderOptions(flow),
      });
      const svg = writeSvg(laid);
      losses.push(...svg.losses);
      if (strict && losses.length > 0) throw new ConversionLossError(losses[0]!);
      return { bytes: svg.bytes, losses };
    }
    const pdf =
      reader.id === 'xlsx'
        ? await convertXlsxToPdf(bytes, conv)
        : await convertDocxToPdf(bytes, conv);
    if (strict && losses.length > 0) throw new ConversionLossError(losses[0]!);
    return { bytes: pdf, losses };
  };

  return { readers, detect, convert };
}

// Resolve regular/bold/italic/boldItalic through the chain. v0 resolves the
// document-default family (per-run family-aware resolution folds in when the
// converters take providers natively). A 'remote' winner is a substitution.
export async function resolveFontsViaChain(
  providers: ReadonlyArray<FontProvider>,
): Promise<{ fonts?: FontBytesByVariant; loss?: Loss }> {
  const chain = chainProviders(providers);
  const ask = (bold: boolean, italic: boolean) => chain.resolve({ bold, italic });
  const [regular, bold, italic, boldItalic] = await Promise.all([
    ask(false, false),
    ask(true, false),
    ask(false, true),
    ask(true, true),
  ]);
  if (regular.kind !== 'bytes') return {};
  const fonts: FontBytesByVariant = {
    regular: regular.bytes,
    ...(bold.kind === 'bytes' ? { bold: bold.bytes } : {}),
    ...(italic.kind === 'bytes' ? { italic: italic.bytes } : {}),
    ...(boldItalic.kind === 'bytes' ? { boldItalic: boldItalic.bytes } : {}),
  };
  const loss: Loss | undefined =
    regular.providerId === 'remote' || regular.providerId === 'local'
      ? {
          severity: 'substituted',
          feature: FEATURES.fontsSubstitution,
          detail: `document fonts rendered with ${regular.faceName} (provider: ${regular.providerId})`,
        }
      : undefined;
  return { fonts, ...(loss ? { loss } : {}) };
}
