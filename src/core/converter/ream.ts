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
import type { FamilyKey, FetchLike, SubstituteKey } from '@/core/fonts';
import type { FontProvider } from '@/core/fonts/provider';
import type { Loss } from '@/core/ir';
import type { DocumentReader } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { SheetDoc } from '@/core/ir/sheet';
import type { SignatureOptions, StyledRenderOptions } from '@/pdf';
import { DEFAULT_READERS, resolveFontsViaChain, toFlowDoc } from '@/core/converter/facade';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry, createFontMeasure } from '@/core/font';
import { fetchFontSet, fetchScriptFont, resolveFamilyKey } from '@/core/fonts';
import { scriptsInFlow } from '@/core/fonts/scripts';
import { familiesInFlow } from '@/core/fonts/families';
import { ConversionLossError } from '@/core/ir';
import { decryptPackage, isEncryptedPackage } from '@/core/crypto/offcrypto';
import { writeDocx } from '@/word/docx-writer';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { writeXlsx } from '@/excel/xlsx-writer';
import { writeHtml } from '@/html/html-writer';
import { writeMarkdown } from '@/markdown/markdown-writer';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { renderStyledPdf, renderStyledPdfEncrypted, signPdf } from '@/pdf';
import { writeSvg } from '@/svg/svg-writer';

/** The output formats {@link Ream.convert} can produce. */
export type ReamTarget = 'pdf' | 'svg' | 'html' | 'md' | 'docx' | 'xlsx';

/**
 * The point size Excel's column-width unit is quoted at — its default theme
 * font is 11 pt, and 8.43 of its digits are the documented 64 px column.
 */
const DEFAULT_WORKBOOK_FONT_PT = 11;

/** Options for {@link Ream.parse}. */
export interface ReamParseOptions {
  /** Reader registry override; defaults to the built-in docx + xlsx readers. */
  readonly readers?: ReadonlyArray<DocumentReader<SourceDoc>>;
  /**
   * The password an encrypted source is opened with: a PDF's user password
   * (ISO 32000 §7.6) or an OOXML package's (ECMA-376 §2.3, MS-OFFCRYPTO).
   * Defaults to the empty string, which opens a PDF's common permissions-only
   * encryption (EP14); an encrypted OOXML package always names a password.
   */
  readonly password?: string;
  /**
   * PDF only: what to read out of the page. `'flow'` (the default) reconstructs
   * a re-flowable document — paragraphs and tables in reading order, from the
   * structure tree where the file has one. `'positional'` keeps the page as a
   * page: every line stands where its glyphs stand, beside the rules and fills
   * the page draws. A form or a drawing needs the second; a report to be edited
   * or converted to markdown needs the first.
   */
  readonly pdfLayout?: 'flow' | 'positional';
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
  /**
   * §18.3.1.34 `&F` — the workbook's file name, for a spreadsheet whose header
   * or footer prints it. A byte-oriented reader cannot know it; supplied here,
   * the code resolves, and omitted it is dropped exactly as before.
   */
  readonly fileName?: string;
  /**
   * Markdown only: how a picture reaches the output — inlined as a `data:` URI
   * (the default), named under `./media/` for a caller that writes the bytes
   * itself, or dropped. See {@link MarkdownWriteOptions}.
   */
  readonly images?: 'dataUri' | 'link' | 'drop';
  /**
   * Markdown only: what a page break becomes — nothing (the default), or the
   * `---` thematic break a slide deck wants between its slides. See
   * {@link MarkdownWriteOptions}.
   */
  readonly pageBreaks?: 'rule' | 'drop';
  /**
   * Markdown from a SPREADSHEET only: open each sheet with a heading carrying
   * its tab name. On by default — markdown has no pages to tell one sheet from
   * the next by, so without them a workbook is a pile of tables with nothing to
   * say which is which. Set `false` for the bare tables.
   */
  readonly sheetNames?: boolean;
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
    // ECMA-376 §2.3 — an encrypted package is an OLE container holding the whole
    // OPC zip as ciphertext, which every reader correctly declines. Given the
    // password it opens here, and what comes out is an ordinary document.
    let source = bytes;
    if (isEncryptedPackage(bytes)) {
      if (options.password === undefined) {
        throw new Error(
          'Password-protected OOXML document (ECMA-376 §2.3 EncryptedPackage) — ' +
            'pass the password as `password`',
        );
      }
      source = decryptPackage(bytes, options.password);
    }
    const reader = readers.find((r) => r.sniff(source));
    if (!reader) {
      throw new Error(
        `Unrecognized document format (readers: ${readers.map((r) => r.id).join(', ')})`,
      );
    }
    const { doc, losses } = reader.read(source, {
      password: options.password,
      ...(options.pdfLayout ? { pdfLayout: options.pdfLayout } : {}),
    });
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
   * writing). HTML, Markdown, DOCX and XLSX are produced straight from the
   * interlayer — no layout, no fonts, zero I/O; SVG and PDF run the layout
   * engine and resolve fonts first.
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
      this.sheet && (options.now || options.fileName)
        ? projectSheetDoc(this.sheet, {
            ...(options.now ? { now: options.now } : {}),
            ...(options.fileName ? { fileName: options.fileName } : {}),
          })
        : this.flow;

    if (to === 'html') {
      // Flow medium: no layout, no fonts to embed — zero I/O.
      const html = writeHtml(flow);
      losses.push(...html.losses);
      this.enforceStrict(options, losses);
      return { bytes: html.bytes, losses };
    }

    if (to === 'md') {
      // Flow medium as well, and the narrowest of them: markdown keeps the
      // document's structure and drops its geometry, reporting each omission.
      //
      // A workbook re-projects, because markdown has no pages to tell one sheet
      // from the next by: the projection is asked for the tab NAMES, which it
      // withholds from a printed page because Excel and Calc print none.
      const source =
        this.sheet && options.sheetNames !== false
          ? projectSheetDoc(this.sheet, {
              ...(options.now ? { now: options.now } : {}),
              ...(options.fileName ? { fileName: options.fileName } : {}),
              sheetHeadings: true,
            })
          : flow;
      const markdown = writeMarkdown(source, {
        ...(options.images ? { images: options.images } : {}),
        ...(options.pageBreaks ? { pageBreaks: options.pageBreaks } : {}),
      });
      losses.push(...markdown.losses);
      this.enforceStrict(options, losses);
      return { bytes: markdown.bytes, losses };
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

    const { fonts, registriesByFamily } = await this.resolveFonts(options, losses, flow);
    const registry = FontRegistry.fromBytes(fonts);

    // §18.3.1.13 measures a column in Maximum Digit Widths — of the font it is
    // drawn in. The parse-time projection has no font, so a paginated target
    // re-projects the grid against the face it is about to render with;
    // otherwise every column is laid out to one font's digit and filled with
    // another's, and the text that does not fit is clipped away.
    const paginated = this.sheet
      ? projectSheetDoc(this.sheet, {
          ...(options.now ? { now: options.now } : {}),
          ...(options.fileName ? { fileName: options.fileName } : {}),
          digitWidthPt: createFontMeasure(registry.resolveByStyle(false, false).parsed).textWidthPt(
            '0',
            DEFAULT_WORKBOOK_FONT_PT,
          ),
        })
      : flow;

    if (to === 'svg') {
      const laid = layoutStyledDocument(paginated.body, {
        registry,
        ...flowRenderOptions(paginated),
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
    const info = paginated.info || callerInfo ? { ...paginated.info, ...callerInfo } : undefined;
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
      ...flowRenderOptions(paginated),
      ...(info ? { info } : {}),
      ...(attachments.length > 0 ? { attachments } : {}),
      ...(signature ? { signaturePlaceholder: signature } : {}),
      ...renderOptions,
    };
    // §7.6: encryption runs on this async path (WebCrypto); the plain branch
    // stays the byte-stable sync render.
    let pdf = styled.encrypt
      ? await renderStyledPdfEncrypted(paginated.body, styled)
      : renderStyledPdf(paginated.body, styled);
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
    flow: FlowDoc,
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

    // Auto-download an open substitute set, one per family the document names
    // (families-in-flow, so every format is served the same way — a deck whose
    // theme is Times and whose body is Calibri gets both).
    const fetchOpt = options.fontFetch ? { fetch: options.fontFetch } : {};
    const keys = options.fontFamily
      ? new Set<FamilyKey>([resolveFamilyKey(options.fontFamily)])
      : familiesInFlow(flow);
    if (keys.size <= 1 && scriptsInFlow(flow).scripts.size === 0) {
      const family = keys.values().next().value;
      return { fonts: await fetchFontSet({ ...(family ? { family } : {}), ...fetchOpt }) };
    }
    const registriesByFamily = new Map<SubstituteKey, FontRegistry>();
    let baseBytes: FontBytesByVariant | undefined;
    for (const key of keys) {
      const bytes = await fetchFontSet({ family: key, ...fetchOpt });
      registriesByFamily.set(key, FontRegistry.fromBytes(bytes));
      // The registry a run falls back to when its family is none of these.
      if (key === 'arimo' || !baseBytes) baseBytes = bytes;
    }
    await this.addScriptFonts(registriesByFamily, flow, options, losses);
    return { fonts: baseBytes!, registriesByFamily };
  }

  /**
   * Fetch one face per writing system the document holds text in — the curated
   * families are Latin, and Han, Kana, Hangul, Arabic and the geometric symbols
   * are a notdef box in every one of them.
   *
   * Only the regular weight is fetched: Noto Sans SC is ten megabytes, and a
   * bold run in it is better stroked (see `SyntheticFace`) than downloaded four
   * times over. A face that fails to arrive is a recorded loss, not a throw —
   * the rest of the document still renders.
   *
   * @param into    The registry map the run resolver looks in.
   * @param flow    The parsed document.
   * @param options The convert options (for the injectable `fetch`).
   * @param losses  Where a face that could not be fetched records itself.
   */
  private async addScriptFonts(
    into: Map<SubstituteKey, FontRegistry>,
    flow: FlowDoc,
    options: ReamConvertOptions,
    losses: Array<Loss>,
  ): Promise<void> {
    const { scripts } = scriptsInFlow(flow);
    for (const script of scripts) {
      const bytes = await fetchScriptFont(script, options.fontFetch);
      if (!bytes) {
        losses.push({
          feature: 'font',
          detail: `no face could be fetched for the ${script} script; its text draws as boxes`,
          severity: 'substituted',
        });
        continue;
      }
      into.set(script, FontRegistry.fromBytes(bytes));
    }
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
