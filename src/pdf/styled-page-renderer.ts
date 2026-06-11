// PDF composition root for the styled pipeline: layout (src/layout/) + emit
// (styled-page-emitter) glued over a PdfDocument. The engine and the PageDoc
// schema themselves moved to src/layout/ at stage 6.4 — this module keeps the
// historical import path alive and owns the one place where both halves meet.

import type { BodyElement } from '@/core/document-model';
import type { StyledRenderOptions } from '@/layout/styled-layout';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { emitStyledPdf } from '@/pdf/styled-page-emitter';
import { PdfDocument } from '@/pdf/writer';

export function renderStyledPdf(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): Uint8Array {
  const laid = layoutStyledDocument(body, options);
  const doc = new PdfDocument();
  return emitStyledPdf(laid, options, doc);
}

// Compatibility re-exports — the schema and the engine live in src/layout/.
export * from '@/layout/page-doc';
export * from '@/layout/styled-layout';
