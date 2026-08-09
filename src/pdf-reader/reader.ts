// E-PDF EP5 — the pdfReader adapter. Implements the DocumentReader contract so
// `Ream.parse(pdfBytes)` sniffs %PDF-, parses the objects (EP1), extracts text
// (EP2) and reconstructs a FlowDoc — via the tagged structure tree when present
// (EP3) or the layout heuristic otherwise (EP4). Raster images are lifted back
// out and placed in reading order (EP6); true vector graphics are not.

import { PdfFile } from './document';
import { reconstructByLayout } from './layout';
import { reconstructTaggedPdf } from './tagged';
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
 * @param layout   `'flow'` reads a re-flowable document out of the page —
 *                 paragraphs and tables in reading order, from the structure
 *                 tree where there is one. `'positional'` keeps the page: every
 *                 line stands where its glyphs do, beside the artwork, which is
 *                 what a form or a drawing needs and what a paragraph cannot be.
 * @returns The reconstructed FlowDoc and its accumulated {@link Loss} report.
 */
export function readPdf(
  bytes: Uint8Array,
  password = '',
  layout: 'flow' | 'positional' = 'flow',
): ReadResult<FlowDoc> {
  const file = PdfFile.parse(bytes, password);
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

  const tagged = layout === 'positional' ? undefined : reconstructTaggedPdf(file);
  const reconstruction = tagged ?? reconstructByLayout(file, layout);
  if (!tagged && layout !== 'positional') {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'untagged PDF — text and headings reconstructed heuristically from glyph positions; structure is approximate',
    });
  }
  // Per-image losses from EP6 (undecodable colour spaces, dropped alpha, …).
  losses.push(...reconstruction.losses);
  // Filled paths (EP10), stroked lines (EP11) and shading-pattern gradients
  // (EP16c) are lifted on the untagged path; clipping paths and bare `sh`
  // shadings are not.
  losses.push({
    severity: 'dropped',
    feature: FEATURES.images,
    detail: 'PDF clipping paths and bare-shading (sh) vector regions are not reconstructed',
  });
  return { doc: reconstruction.doc, losses };
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
      opts?.pdfLayout === 'positional' ? 'positional' : 'flow',
    ),
};
