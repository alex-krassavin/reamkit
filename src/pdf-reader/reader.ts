// E-PDF EP5 — the pdfReader adapter. Implements the DocumentReader contract so
// `Ream.parse(pdfBytes)` sniffs %PDF-, parses the objects (EP1), extracts text
// (EP2) and reconstructs a FlowDoc — via the tagged structure tree when present
// (EP3) or the layout heuristic otherwise (EP4). Raster images are lifted back
// out and placed in reading order (EP6); true vector graphics are not.

import { PdfFile } from './document';
import { collectPageVectors } from './vector';
import { extractPageText } from './text';
import { reconstructByLayout } from './layout';
import { reconstructTaggedPdf } from './tagged';
import type { StreamFilters } from './document';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import { FEATURES } from '@/core/ir';

// Sniff the %PDF- header, tolerating a few junk bytes before it (some producers
// prepend a BOM or whitespace) — §7.5.2.
function sniffPdf(bytes: Uint8Array): boolean {
  const limit = Math.min(bytes.length - 5, 1024);
  for (let i = 0; i <= limit; i++) {
    if (
      bytes[i] === 0x25 && // %
      bytes[i + 1] === 0x50 && // P
      bytes[i + 2] === 0x44 && // D
      bytes[i + 3] === 0x46 && // F
      bytes[i + 4] === 0x2d // -
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Parse PDF bytes into a {@link FlowDoc} (E-PDF EP5): reconstruct via the tagged
 * structure tree when present (EP3), else the layout heuristic (EP4), lifting
 * raster images back into reading order (EP6). Records losses for encryption that
 * could not be opened, heuristic (untagged) reconstruction, and the vector
 * regions that are not reconstructed (clipping paths, bare `sh` shadings).
 *
 * @param bytes    The complete PDF file bytes.
 * @param password The user password for an encrypted source; the empty string
 *                 opens permissions-only encryption.
 * @param filters  Decoders for `/Filter` names this reader does not implement
 *                 (§7.4); see {@link StreamFilters}.
 * @param layout   `'auto'` (the default) lets the FILE decide: a page that is
 *                 mostly marks is reproduced, one that is mostly lines is re-set,
 *                 and the reader records which it chose. `'flow'` reads a
 *                 re-flowable document out of the page —
 *                 paragraphs and tables in reading order, from the structure
 *                 tree where there is one. `'positional'` keeps the page: every
 *                 line stands where its glyphs do, beside the artwork, which is
 *                 what a form or a drawing needs and what a paragraph cannot be.
 * @returns The reconstructed FlowDoc and its accumulated {@link Loss} report.
 */
export function readPdf(
  bytes: Uint8Array,
  password = '',
  layout: 'flow' | 'positional' | 'auto' = 'auto',
  filters: StreamFilters = {},
): ReadResult<FlowDoc> {
  const file = PdfFile.parse(bytes, password, filters);
  const losses: Array<Loss> = [];

  if (file.encryptionUnsupported) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail:
        'encrypted PDF — the user password was missing or incorrect, or the handler is unsupported',
    });
  }

  // A placed reconstruction never consults the structure tree: the tree names a
  // reading order, and a reading order is the one thing a placed page does not
  // have. The words go where the glyphs are.
  // §7.4 — a stream filter nothing here can undo leaves that stream unread.
  // Reported before anything else: when it is the cross-reference stream, every
  // page of the file is missing and no other loss explains why.
  for (const name of file.unknownFilters) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail: `PDF stream filter /${name} is not supported; streams using it are unreadable`,
    });
  }

  // Which reading the FILE asks for, where the caller did not say.
  const reading = layout === 'auto' ? readingOf(file) : layout;
  if (layout === 'auto') {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        reading === 'positional'
          ? 'PDF read as a PAGE (placed): its artwork outweighs its prose, so every line stands where its glyphs stand — pass pdfLayout: "flow" to re-set it as a document instead'
          : 'PDF read as a DOCUMENT (flowing): its prose outweighs its artwork, so the words re-set and the artwork takes its turn in reading order — pass pdfLayout: "positional" to reproduce the page instead',
    });
  }
  const tagged = reading === 'positional' ? undefined : reconstructTaggedPdf(file);
  const reconstruction = tagged ?? reconstructByLayout(file, reading);
  if (!tagged && reading !== 'positional') {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'untagged PDF — text and headings reconstructed heuristically from glyph positions; structure is approximate',
    });
  }
  // Per-image losses from EP6 (undecodable colour spaces, dropped alpha, …).
  losses.push(...reconstruction.losses);
  // Filled paths (EP10), stroked lines (EP11), shading-pattern gradients
  // (EP16c) and the clipping paths that bound them are all lifted; a bare `sh`
  // shading, which paints a region rather than filling a path, is not.
  losses.push({
    severity: 'dropped',
    feature: FEATURES.images,
    detail: 'PDF bare-shading (sh) vector regions are not reconstructed',
  });
  return { doc: reconstruction.doc, losses };
}

/**
 * Which reading a PDF asks for: a DOCUMENT to re-set, or a PAGE to reproduce.
 *
 * The two cannot be mixed. A flowing reading moves the words, and words that
 * move cannot agree with rules that do not — anchored artwork over reflowed
 * text puts every label on the wrong box. A placed reading keeps both where
 * they were drawn, and reflows nothing.
 *
 * The file says which it is by what is on it. A form or a drawing is mostly
 * MARKS — 160F-2019.pdf sets 28 numbered rows in 355 ruled boxes — and a paper
 * is mostly LINES, with a rule or two between them. So the marks are counted
 * against the lines, on the MEDIAN page rather than the worst, since one plan
 * folded into a report does not make the report a plan.
 *
 * A threshold is a guess, and this one is stated rather than hidden: the reader
 * records which reading it took and why, and the caller can name the other.
 */
function readingOf(file: PdfFile): 'flow' | 'positional' {
  /** Below this many marks a page is prose with decoration, whatever the ratio. */
  const ENOUGH_MARKS = 20;
  /** Twice as many marks as lines is a page that is drawn rather than written. */
  const DRAWN = 2;
  const ratios: Array<number> = [];
  for (const page of file.pages()) {
    let marks = 0;
    let lines = 0;
    try {
      marks = collectPageVectors(file, page, []).vectors.length;
      // Baselines, not runs: a line broken into twenty runs is still one line,
      // and counting runs would make ordinary justified prose look drawn.
      const ys = new Set<number>();
      for (const run of extractPageText(file, page)) ys.add(Math.round(run.y));
      lines = ys.size;
    } catch {
      continue;
    }
    if (marks < ENOUGH_MARKS) {
      ratios.push(0);
      continue;
    }
    ratios.push(lines > 0 ? marks / lines : DRAWN);
  }
  if (ratios.length === 0) return 'flow';
  const sorted = [...ratios].sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)] ?? 0;
  return median >= DRAWN ? 'positional' : 'flow';
}

/**
 * The `pdfReader` adapter: a {@link DocumentReader} that sniffs the `%PDF-`
 * header and parses the bytes into a {@link FlowDoc} (E-PDF EP5).
 */
export const pdfReader: DocumentReader<FlowDoc> = {
  id: 'pdf',
  produces: 'flow',
  supports: new Set([FEATURES.text, FEATURES.tables, FEATURES.lists, FEATURES.images]),
  sniff: sniffPdf,
  read: (bytes, opts) =>
    readPdf(
      bytes,
      typeof opts?.password === 'string' ? opts.password : '',
      opts?.pdfLayout === 'positional'
        ? 'positional'
        : opts?.pdfLayout === 'flow'
          ? 'flow'
          : 'auto',
      isFilters(opts?.filters) ? opts.filters : {},
    ),
};

/** A caller's `filters` option, when it is the shape the reader can use. */
function isFilters(value: unknown): value is StreamFilters {
  return typeof value === 'object' && value !== null;
}
