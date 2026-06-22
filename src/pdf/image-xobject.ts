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

/** A reference to an embedded image XObject plus its pixel dimensions. */
export interface EmbeddedImage {
  /** Indirect reference to the image XObject, for a page's `/XObject` resources. */
  readonly ref: PdfRef;
  readonly widthPx: number;
  readonly heightPx: number;
}

/**
 * Create the PDF image-XObject objects from an already-prepared image (ISO
 * 32000-1 §8.9.5), emitting an `/SMask` soft-mask object first when present.
 * The decode/validate work happens earlier in `prepareImage`; this only writes
 * the objects.
 *
 * @param prepared The decoded, ready-to-emit image stream.
 * @returns The image reference and its pixel dimensions.
 */
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

/**
 * Decode raw image `bytes` and embed them in one step — a convenience over
 * `prepareImage` + {@link addImage}.
 *
 * @param bytes   The raw image file bytes (PNG/JPEG/…).
 * @param options Decode/embed options.
 * @returns The image reference and its pixel dimensions.
 */
export function embedImage(
  doc: PdfDocument,
  bytes: Uint8Array,
  options: EmbedImageOptions = {},
): EmbeddedImage {
  return addImage(doc, prepareImage(bytes, options));
}
