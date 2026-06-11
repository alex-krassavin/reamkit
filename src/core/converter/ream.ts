// Ream — the object face of the library.
//
// One parse into the interlayer (FlowDoc), any number of conversions out of
// it, never re-reading the source:
//
//   const doc = Ream.parse(bytes);              // sniff → reader → FlowDoc
//   const pdf = await doc.convert('pdf', { fonts });
//   const svg = await doc.convert('svg', { fonts });
//
// The class is a thin GRASP Controller: readers parse, flowRenderOptions
// projects the FlowDoc, layout/emit and the svg writer do the work. It is
// the one deliberate composition root besides createConverter, so importing
// it pulls every format module — use the per-format functions when bundle
// size matters more than convenience.
//
// It keeps the source bytes for the two source-touching features only:
// substitute-font auto-detection (docx) and PDF/A-3 `embedSource`.

import type { ConvertResult } from '@/core/converter/facade';
import type { FontBytesByVariant } from '@/core/font';
import type { FetchLike } from '@/core/fonts';
import type { FontProvider } from '@/core/fonts/provider';
import type { Loss } from '@/core/ir';
import type { DocumentReader } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { DEFAULT_READERS, resolveFontsViaChain } from '@/core/converter/facade';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { fetchFontSet } from '@/core/fonts';
import { ConversionLossError } from '@/core/ir';
import { writeHtml } from '@/html/html-writer';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { renderStyledPdf, renderStyledPdfEncrypted, signPdf } from '@/pdf';
import { writeSvg } from '@/svg/svg-writer';
import { resolveDocxAutoFonts } from '@/word/docx-to-pdf';

export type ReamTarget = 'pdf' | 'svg' | 'html';

export interface ReamParseOptions {
  // Reader registry override — defaults to the built-in docx + xlsx readers.
  readonly readers?: ReadonlyArray<DocumentReader>;
}

export interface ReamConvertOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  readonly fonts?: FontBytesByVariant;
  readonly fontBytes?: Uint8Array;
  // Substitute family hint for the auto-download path.
  readonly fontFamily?: string;
  // Injectable fetch for the auto-download path (defaults to global fetch).
  readonly fontFetch?: FetchLike;
  // Font resolution chain (caller/embedded/local/remote) — used when
  // `fonts`/`fontBytes` are absent; a remote/local winner records a
  // substitution Loss.
  readonly fontProviders?: ReadonlyArray<FontProvider>;
  // Throw ConversionLossError on the first loss instead of reporting it.
  readonly strict?: boolean;
  // PDF/A-3 only: embed the parsed source file (/AFRelationship /Source).
  readonly embedSource?: boolean;
  // Digitally sign the output (ISO 32000 §12.8, WebCrypto).
  readonly signature?: SignatureOptions;
}

const SOURCE_MIME: Readonly<Record<string, string>> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

export class Ream {
  private constructor(
    // The interlayer itself — the parsed, format-neutral document tree.
    readonly flow: FlowDoc,
    // Losses recorded while reading the source.
    readonly losses: ReadonlyArray<Loss>,
    private readonly source: Uint8Array,
    private readonly readerId: string,
  ) {}

  // Sniff the format, parse once into the FlowDoc interlayer.
  static parse(bytes: Uint8Array, options: ReamParseOptions = {}): Ream {
    const readers = options.readers ?? DEFAULT_READERS;
    const reader = readers.find((r) => r.sniff(bytes));
    if (!reader) {
      throw new Error(
        `Unrecognized document format (readers: ${readers.map((r) => r.id).join(', ')})`,
      );
    }
    const { doc, losses } = reader.read(bytes);
    return new Ream(doc, losses, bytes, reader.id);
  }

  // Source format id ('docx', 'xlsx', …).
  get format(): string {
    return this.readerId;
  }

  async convert(to: ReamTarget, options: ReamConvertOptions = {}): Promise<Uint8Array> {
    return (await this.convertWithReport(to, options)).bytes;
  }

  async convertWithReport(
    to: ReamTarget,
    options: ReamConvertOptions = {},
  ): Promise<ConvertResult> {
    const losses: Array<Loss> = [...this.losses];

    if (to === 'html') {
      // Flow medium: no layout, no fonts to embed — zero I/O.
      const html = writeHtml(this.flow);
      losses.push(...html.losses);
      this.enforceStrict(options, losses);
      return { bytes: html.bytes, losses };
    }

    const { fonts, registriesByFamily } = await this.resolveFonts(options, losses);
    const registry = FontRegistry.fromBytes(fonts);

    if (to === 'svg') {
      const laid = layoutStyledDocument(this.flow.body, {
        registry,
        ...flowRenderOptions(this.flow),
      });
      const svg = writeSvg(laid);
      losses.push(...svg.losses);
      this.enforceStrict(options, losses);
      return { bytes: svg.bytes, losses };
    }

    const {
      fonts: _a,
      fontBytes: _b,
      fontFamily: _c,
      fontFetch: _d,
      fontProviders: _e,
      strict: _f,
      embedSource,
      signature,
      info: callerInfo,
      attachments: callerAttachments,
      ...renderOptions
    } = options;
    void _a;
    void _b;
    void _c;
    void _d;
    void _e;
    void _f;

    // Caller overrides spread over the document's own metadata.
    const info = this.flow.info || callerInfo ? { ...this.flow.info, ...callerInfo } : undefined;
    const attachments = [...(callerAttachments ?? [])];
    if (embedSource && options.pdfA?.startsWith('PDF/A-3')) {
      attachments.push({
        name: `source.${this.readerId}`,
        bytes: this.source,
        mimeType: SOURCE_MIME[this.readerId] ?? 'application/octet-stream',
        relationship: 'Source',
        description: 'Source document',
      });
    }

    const styled = {
      registry,
      ...(registriesByFamily ? { registriesByFamily } : {}),
      ...flowRenderOptions(this.flow),
      ...(info ? { info } : {}),
      ...(attachments.length > 0 ? { attachments } : {}),
      ...(signature ? { signaturePlaceholder: signature } : {}),
      ...renderOptions,
    };
    // §7.6: encryption runs on this async path (WebCrypto); the plain branch
    // stays the byte-stable sync render.
    let pdf = styled.encrypt
      ? await renderStyledPdfEncrypted(this.flow.body, styled)
      : renderStyledPdf(this.flow.body, styled);
    if (signature) pdf = await signPdf(pdf, signature);
    this.enforceStrict(options, losses);
    return { bytes: pdf, losses };
  }

  private async resolveFonts(
    options: ReamConvertOptions,
    losses: Array<Loss>,
  ): Promise<{
    fonts: FontBytesByVariant;
    registriesByFamily?: StyledRenderOptions['registriesByFamily'];
  }> {
    const explicit =
      options.fonts ?? (options.fontBytes ? { regular: options.fontBytes } : undefined);
    if (explicit) return { fonts: explicit };

    if (options.fontProviders && options.fontProviders.length > 0) {
      const { fonts, loss } = await resolveFontsViaChain(options.fontProviders);
      if (loss) losses.push(loss);
      if (fonts) return { fonts };
    }

    // Auto-download an open substitute set (per detected family for docx).
    if (this.readerId === 'docx') {
      return resolveDocxAutoFonts(this.source, options);
    }
    const fetchOpt = {
      ...(options.fontFamily ? { family: options.fontFamily } : {}),
      ...(options.fontFetch ? { fetch: options.fontFetch } : {}),
    };
    return { fonts: await fetchFontSet(fetchOpt) };
  }

  private enforceStrict(options: ReamConvertOptions, losses: ReadonlyArray<Loss>): void {
    if (options.strict && losses.length > 0) throw new ConversionLossError(losses[0]!);
  }
}
