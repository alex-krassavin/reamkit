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

export function readPdf(bytes: Uint8Array): ReadResult<FlowDoc> {
  const file = PdfFile.parse(bytes);
  const losses: Array<Loss> = [];

  const tagged = reconstructTaggedPdf(file);
  const reconstruction = tagged ?? reconstructByLayout(file);
  if (!tagged) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'untagged PDF — text and headings reconstructed heuristically from glyph positions; structure is approximate',
    });
  }
  // Per-image losses from EP6 (undecodable colour spaces, dropped alpha, …).
  losses.push(...reconstruction.losses);
  // Raster images are lifted; true vector graphics (paths, shadings) are not.
  losses.push({
    severity: 'dropped',
    feature: FEATURES.images,
    detail: 'PDF vector graphics are not reconstructed',
  });
  return { doc: reconstruction.doc, losses };
}

export const pdfReader: DocumentReader<FlowDoc> = {
  id: 'pdf',
  produces: 'flow',
  supports: new Set([FEATURES.text, FEATURES.tables, FEATURES.lists, FEATURES.images]),
  sniff: sniffPdf,
  read: (bytes) => readPdf(bytes),
};
