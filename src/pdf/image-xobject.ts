// PDF Image XObjects (ISO 32000-1 §8.9.5).
//
// The prepare/add split (oop-design §3.1): `prepareImage` (now in
// core/images — pure decode/validate, no PdfDocument) produces the
// ready-to-emit stream bytes; `addImage` here only creates the PDF objects.
// Layout probes with prepareImage, the emit phase replays the prepared
// result, and other writers (SVG) reuse the mime/dimensions.

import type { PdfRef } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import type { EmbedImageOptions, PreparedImage } from '@/core/images';
import { name, stream } from '@/pdf/objects';
import { prepareImage } from '@/core/images';

// Compatibility re-exports — the decode experts moved to core/images.
export { detectImageFormat, prepareImage } from '@/core/images';
export type { EmbedImageOptions, ImageFormat, PreparedImage } from '@/core/images';

export interface EmbeddedImage {
  readonly ref: PdfRef;
  readonly widthPx: number;
  readonly heightPx: number;
}

export function addImage(doc: PdfDocument, prepared: PreparedImage): EmbeddedImage {
  // The soft mask object precedes the color image — the order the pre-split
  // embed produced.
  let smaskRef: PdfRef | undefined;
  if (prepared.smaskData) {
    smaskRef = doc.add(
      stream(
        {
          Type: name('XObject'),
          Subtype: name('Image'),
          Width: prepared.widthPx,
          Height: prepared.heightPx,
          ColorSpace: name('DeviceGray'),
          BitsPerComponent: 8,
          Filter: name('FlateDecode'),
        },
        prepared.smaskData,
      ),
    );
  }
  const entries: Record<string, unknown> = {
    Type: name('XObject'),
    Subtype: name('Image'),
    Width: prepared.widthPx,
    Height: prepared.heightPx,
    ...(prepared.colorSpace ? { ColorSpace: name(prepared.colorSpace) } : {}),
    ...(prepared.bitsPerComponent !== undefined
      ? { BitsPerComponent: prepared.bitsPerComponent }
      : {}),
    Filter: name(prepared.filter),
  };
  if (smaskRef) entries['SMask'] = smaskRef;
  const ref = doc.add(stream(entries as Parameters<typeof stream>[0], prepared.data));
  return { ref, widthPx: prepared.widthPx, heightPx: prepared.heightPx };
}

export function embedImage(
  doc: PdfDocument,
  bytes: Uint8Array,
  options: EmbedImageOptions = {},
): EmbeddedImage {
  return addImage(doc, prepareImage(bytes, options));
}
