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

import type { ConvertResult, SourceDoc } from '@/core/converter/facade';
import type { FontBytesByVariant } from '@/core/font';
import type { FetchLike } from '@/core/fonts';
import type { FontProvider } from '@/core/fonts/provider';
import type { Loss } from '@/core/ir';
import type { DocumentReader } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { SheetDoc } from '@/core/ir/sheet';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { DEFAULT_READERS, resolveFontsViaChain, toFlowDoc } from '@/core/converter/facade';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { fetchFontSet } from '@/core/fonts';
import { ConversionLossError } from '@/core/ir';
import { writeDocx } from '@/word/docx-writer';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { writeXlsx } from '@/excel/xlsx-writer';
import { writeHtml } from '@/html/html-writer';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { renderStyledPdf, renderStyledPdfEncrypted, signPdf } from '@/pdf';
import { writeSvg } from '@/svg/svg-writer';
import { resolveDocxAutoFonts } from '@/word/docx-to-pdf';

/** The output formats {@link Ream.convert} can produce. */
export type ReamTarget = 'pdf' | 'svg' | 'html' | 'docx' | 'xlsx';

/** Options for {@link Ream.parse}. */
export interface ReamParseOptions {
  /** Reader registry override; defaults to the built-in docx + xlsx readers. */
  readonly readers?: ReadonlyArray<DocumentReader<SourceDoc>>;
  /**
   * The user password for an encrypted PDF source (ISO 32000 §7.6). Defaults to
   * the empty string, which opens the common permissions-only encryption (EP14).
   */
  readonly password?: string;
}

/**
 * Options for {@link Ream.convert} and {@link Ream.convertWithReport}. Extends
 * the low-level {@link StyledRenderOptions} (minus the font `registry` and
 * `styles`, which Ream builds itself) with font resolution and source-touching
 * conveniences.
 */
export interface ReamConvertOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  /** Explicit font bytes per variant (regular/bold/italic/bold-italic). */
  readonly fonts?: FontBytesByVariant;
  /** Shorthand for supplying a single regular-variant font as raw bytes. */
  readonly fontBytes?: Uint8Array;
  /** Substitute family hint for the auto-download path. */
  readonly fontFamily?: string;
  /** Injectable `fetch` for the auto-download path (defaults to the global `fetch`). */
  readonly fontFetch?: FetchLike;
  /**
   * Font resolution chain (caller/embedded/local/remote), used when neither
   * `fonts` nor `fontBytes` is given. A remote or local winner records a
   * substitution {@link Loss}.
   */
  readonly fontProviders?: ReadonlyArray<FontProvider>;
  /** Throw {@link ConversionLossError} on the first loss instead of reporting it. */
  readonly strict?: boolean;
  /** PDF/A-3 only: embed the parsed source file (`/AFRelationship /Source`). */
  readonly embedSource?: boolean;
  /** Digitally sign the output (ISO 32000 §12.8, WebCrypto). */
  readonly signature?: SignatureOptions;
  /**
   * Reference date for spreadsheet conditional-format `timePeriod` rules and for
   * `TODAY()`/`NOW()` in `expression` rules (E-SHEET W9). Supplying it re-projects
   * a spreadsheet source so those clock-relative rules resolve against this date —
   * an explicit input, never the wall clock. Omitted, they no-op and the output is
   * unchanged.
   */
  readonly now?: Date;
}

/** OOXML / legacy MIME types by reader id, for the PDF/A-3 embedded source file. */
const SOURCE_MIME: Readonly<Record<string, string>> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc: 'application/msword',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppt: 'application/vnd.ms-powerpoint',
  pdf: 'application/pdf',
};

/**
 * The object face of the library: parse a document once into the format-neutral
 * {@link FlowDoc} interlayer, then convert it to any number of targets without
 * re-reading the source.
 *
 * ```ts
 * const doc = Ream.parse(bytes); // sniff → reader → FlowDoc
 * const pdf = await doc.convert('pdf', { fonts });
 * const svg = await doc.convert('svg', { fonts });
 * ```
 *
 * It is a thin GRASP Controller: readers parse, `flowRenderOptions` projects the
 * FlowDoc, and layout/emit plus the writers do the work. As a deliberate
 * composition root, importing it pulls in every format module — prefer the
 * per-format functions when bundle size matters more than convenience. The
 * source bytes are retained only for the two source-touching features: docx
 * substitute-font auto-detection and PDF/A-3 `embedSource`.
 */
export class Ream {
  /**
   * @param flow     The interlayer — the parsed, format-neutral document tree.
   * @param sheet    The native SpreadsheetML tree when the source is a spreadsheet
   *                 (xlsx); {@link Ream.flow} is its projection through the print model.
   * @param losses   Losses recorded while reading the source.
   * @param source   The original source bytes (kept only for docx auto-fonts and
   *                 PDF/A-3 `embedSource`).
   * @param readerId The id of the reader that parsed the source.
   */
  private constructor(
    readonly flow: FlowDoc,
    readonly sheet: SheetDoc | undefined,
    readonly losses: ReadonlyArray<Loss>,
    private readonly source: Uint8Array,
    private readonly readerId: string,
  ) {}

  /**
   * Sniff the format and parse the bytes once into the {@link FlowDoc} interlayer.
   *
   * @param bytes   The raw document bytes; the format is detected by sniffing.
   * @param options Optional reader-registry override and/or password for an
   *                encrypted source.
   * @returns A reusable {@link Ream} instance.
   * @throws Error when no registered reader recognizes the bytes.
   */
  static parse(bytes: Uint8Array, options: ReamParseOptions = {}): Ream {
    const readers = options.readers ?? DEFAULT_READERS;
    const reader = readers.find((r) => r.sniff(bytes));
    if (!reader) {
      throw new Error(
        `Unrecognized document format (readers: ${readers.map((r) => r.id).join(', ')})`,
      );
    }
    const { doc, losses } = reader.read(bytes, { password: options.password });
    // The reader's native tree — a SheetDoc for spreadsheets — is projected to
    // the FlowDoc the render path consumes; the SheetDoc is kept for inspection.
    const sheet = doc.kind === 'sheet' ? doc : undefined;
    return new Ream(toFlowDoc(doc), sheet, losses, bytes, reader.id);
  }

  /** The source format id (`'docx'`, `'xlsx'`, …). */
  get format(): string {
    return this.readerId;
  }

  /**
   * Convert the parsed document to `to` and return just the output bytes. A thin
   * wrapper over {@link Ream.convertWithReport} that drops the loss report.
   *
   * @param to      The target format.
   * @param options Font resolution and target-specific options.
   * @returns The encoded output bytes.
   */
  async convert(to: ReamTarget, options: ReamConvertOptions = {}): Promise<Uint8Array> {
    return (await this.convertWithReport(to, options)).bytes;
  }

  /**
   * Convert the parsed document to `to`, returning the output bytes together with
   * the accumulated {@link Loss} report (read-time losses plus any added while
   * writing). HTML, DOCX and XLSX are produced straight from the interlayer — no
   * layout, no fonts, zero I/O; SVG and PDF run the layout engine and resolve
   * fonts first.
   *
   * @param to      The target format. `'xlsx'` requires a spreadsheet source.
   * @param options Font resolution and target-specific options.
   * @returns The encoded bytes and the loss report.
   * @throws Error when `to` is `'xlsx'` but the source has no grid.
   * @throws ConversionLossError when `options.strict` is set and any loss was recorded.
   */
  async convertWithReport(
    to: ReamTarget,
    options: ReamConvertOptions = {},
  ): Promise<ConvertResult> {
    const losses: Array<Loss> = [...this.losses];

    // W9: a caller-supplied reference date re-projects the sheet so conditional-
    // format timePeriod / TODAY() rules resolve against it. Without it (or for a
    // non-sheet source) the parse-time flow — byte-identical to before — is used.
    const flow =
      this.sheet && options.now ? projectSheetDoc(this.sheet, { now: options.now }) : this.flow;

    if (to === 'html') {
      // Flow medium: no layout, no fonts to embed — zero I/O.
      const html = writeHtml(flow);
      losses.push(...html.losses);
      this.enforceStrict(options, losses);
      return { bytes: html.bytes, losses };
    }

    if (to === 'docx') {
      // Flow medium too: the writer re-serializes the interlayer — no layout,
      // no fonts, zero I/O. Output is denormalized (resolved properties as
      // direct formatting) but valid; see docx-writer.ts.
      const docx = writeDocx(flow);
      losses.push(...docx.losses);
      this.enforceStrict(options, losses);
      return { bytes: docx.bytes, losses };
    }

    if (to === 'xlsx') {
      // The native grid medium (E-SHEET SD1): the writer consumes the SheetDoc
      // directly — a docx (no grid) cannot be written to xlsx. Zero I/O.
      if (!this.sheet) {
        throw new Error("convert('xlsx') requires a spreadsheet source; this document has no grid");
      }
      const xlsx = writeXlsx(this.sheet);
      losses.push(...xlsx.losses);
      this.enforceStrict(options, losses);
      return { bytes: xlsx.bytes, losses };
    }

    const { fonts, registriesByFamily } = await this.resolveFonts(options, losses);
    const registry = FontRegistry.fromBytes(fonts);

    if (to === 'svg') {
      const laid = layoutStyledDocument(flow.body, {
        registry,
        ...flowRenderOptions(flow),
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
    const info = flow.info || callerInfo ? { ...flow.info, ...callerInfo } : undefined;
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
      ...flowRenderOptions(flow),
      ...(info ? { info } : {}),
      ...(attachments.length > 0 ? { attachments } : {}),
      ...(signature ? { signaturePlaceholder: signature } : {}),
      ...renderOptions,
    };
    // §7.6: encryption runs on this async path (WebCrypto); the plain branch
    // stays the byte-stable sync render.
    let pdf = styled.encrypt
      ? await renderStyledPdfEncrypted(flow.body, styled)
      : renderStyledPdf(flow.body, styled);
    if (signature) pdf = await signPdf(pdf, signature);
    this.enforceStrict(options, losses);
    return { bytes: pdf, losses };
  }

  /**
   * Resolve the font set for a layout/PDF conversion. Explicit `fonts`/`fontBytes`
   * win; otherwise the provider chain is tried, then an open substitute set is
   * auto-downloaded (per detected family for docx). Any substitution is appended
   * to `losses`.
   *
   * @param options The convert options carrying the font preferences.
   * @param losses  The mutable loss list a substitution is appended to.
   * @returns The font bytes and, for docx, optional per-family registries.
   */
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

  /**
   * In strict mode, throw {@link ConversionLossError} for the first recorded loss.
   *
   * @param options The convert options (checked for `strict`).
   * @param losses  The losses accumulated so far.
   * @throws ConversionLossError when `options.strict` is set and `losses` is non-empty.
   */
  private enforceStrict(options: ReamConvertOptions, losses: ReadonlyArray<Loss>): void {
    if (options.strict && losses.length > 0) throw new ConversionLossError(losses[0]!);
  }
}
