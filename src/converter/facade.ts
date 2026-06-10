// Conversion facade (ir-design §7): the registry-driven entry point that
// composes readers → layout → writers. v0 dispatches format detection through
// reader sniffing and delegates the PDF pipeline to the existing converters
// (which themselves run on the readers) — so the facade is byte-identical to
// calling convertDocxToPdf/convertXlsxToPdf directly.
//
// The async boundary lives HERE (font fetching); readers/writers stay sync.

import type { ConvertDocxOptions } from '@/converter/docx-to-pdf';
import type { DocumentReader } from '@/ir/adapters';
import type { FlowDoc } from '@/ir/flow';
import type { LossReport } from '@/ir';

import { ConversionLossError } from '@/ir';
import { convertDocxToPdf } from '@/converter/docx-to-pdf';
import { convertXlsxToPdf } from '@/converter/xlsx-to-pdf';
import { docxReader } from '@/readers/docx-reader';
import { xlsxReader } from '@/readers/xlsx-reader';

export interface ConvertOptions extends ConvertDocxOptions {
  /** Target format. v0 ships the PDF pipeline; more writers land at stage 6. */
  readonly to?: 'pdf';
  /**
   * Strict mode (handoff v1 §5): throw ConversionLossError on the first
   * recorded loss instead of returning it in the report.
   */
  readonly strict?: boolean;
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

const DEFAULT_READERS: ReadonlyArray<DocumentReader<FlowDoc>> = [docxReader, xlsxReader];

export function createConverter(opts: CreateConverterOptions = {}): Converter {
  const readers = opts.readers ?? DEFAULT_READERS;

  const detect = (bytes: Uint8Array): DocumentReader<FlowDoc> | undefined =>
    readers.find((r) => r.sniff(bytes));

  const convert = async (
    bytes: Uint8Array,
    options: ConvertOptions = {},
  ): Promise<ConvertResult> => {
    const { to = 'pdf', strict = false, ...rest } = options;
    void to; // single target in v0
    const reader = detect(bytes);
    if (!reader) {
      throw new Error('Unrecognized input format (no registered reader sniffs these bytes)');
    }
    const pdf =
      reader.id === 'xlsx'
        ? await convertXlsxToPdf(bytes, rest)
        : await convertDocxToPdf(bytes, rest);
    // v0: the PDF pipeline does not record losses yet (font substitution and
    // friends wire into LossReport with the FontProvider chain, stage 5).
    const losses: LossReport = [];
    if (strict && losses.length > 0) throw new ConversionLossError(losses[0]!);
    return { bytes: pdf, losses };
  };

  return { readers, detect, convert };
}
