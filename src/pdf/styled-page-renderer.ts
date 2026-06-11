// PDF composition root for the styled pipeline: layout (src/layout/) + emit
// (styled-page-emitter) glued over a PdfDocument. The engine and the PageDoc
// schema themselves moved to src/layout/ at stage 6.4 — this module keeps the
// historical import path alive and owns the one place where both halves meet.

import type { BodyElement } from '@/core/document-model';
import type { StyledRenderOptions } from '@/layout/styled-layout';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { emitStyledPdf, emitStyledPdfEncrypted } from '@/pdf/styled-page-emitter';
import { PdfDocument } from '@/pdf/writer';

export function renderStyledPdf(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): Uint8Array {
  if (options.encrypt) {
    throw new Error('PDF encryption requires the async conversion path (WebCrypto)');
  }
  const laid = layoutStyledDocument(body, options);
  const doc = new PdfDocument();
  return emitStyledPdf(laid, options, doc);
}

// §7.6 — the encrypting variant (async: WebCrypto). Validates the standing
// conflicts here so every converter shares them: PDF/A forbids /Encrypt
// (ISO 19005), signatures + encryption is out of scope in v1, and PDF/UA
// requires assistive technology to keep extraction access (bit 10).
export async function renderStyledPdfEncrypted(
  body: ReadonlyArray<BodyElement>,
  options: StyledRenderOptions,
): Promise<Uint8Array> {
  const encrypt = options.encrypt;
  if (!encrypt) return renderStyledPdf(body, options);
  if (options.pdfA) throw new Error('PDF/A forbids encryption (ISO 19005)');
  if (options.signaturePlaceholder) {
    throw new Error('encryption combined with a digital signature is not supported (v1)');
  }
  const effective =
    options.pdfUA === true
      ? { ...encrypt, permissions: { ...encrypt.permissions, contentAccessibility: true } }
      : encrypt;
  const laid = layoutStyledDocument(body, options);
  const doc = new PdfDocument();
  return emitStyledPdfEncrypted(laid, options, doc, effective);
}

// Compatibility re-exports — the schema and the engine live in src/layout/.
export * from '@/layout/page-doc';
export * from '@/layout/styled-layout';
