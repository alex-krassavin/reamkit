// §8.9.6.4 — colour key masking: an image's `/Mask` given as a range of sample
// values per colour component, and a sample every range holds is not painted.

import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { prepareImage } from '@/core/images';
import { PdfFile } from '@/pdf-reader/document';
import { decodePdfImage } from '@/pdf-reader/image-decode';
import { name, stream } from '@/pdf/objects';

const FILE = PdfFile.parse(new TextEncoder().encode('%PDF-1.7\ntrailer\n<< /Size 1 >>\n%%EOF\n'));

/** The alpha the image comes back with, sample by sample; `undefined` for none. */
function alphaOf(
  entries: Record<string, number | ReadonlyArray<number>>,
): Array<number> | undefined {
  // Three samples in a row: full red, full green, full blue.
  const decoded = decodePdfImage(
    FILE,
    stream(
      { Width: 3, Height: 1, ColorSpace: name('DeviceRGB'), BitsPerComponent: 8, ...entries },
      Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255]),
    ),
  );
  if (!decoded.ok) throw new Error('the image decodes');
  const alpha = prepareImage(decoded.bytes).smaskData;
  return alpha === undefined ? undefined : [...unzlibSync(alpha)];
}

describe('colour key masking (§8.9.6.4)', () => {
  it('leaves unpainted the samples every range holds', () => {
    // colorkeymask.pdf keys out every sample of full red, and painted whole
    // its red bar stood beside the green and blue ones the page shows.
    expect(alphaOf({ Mask: [255, 255, 0, 255, 0, 255] })).toEqual([0, 255, 255]);
  });

  it('holds the samples to the ranges before any /Decode', () => {
    const inverted = { Mask: [255, 255, 0, 255, 0, 255], Decode: [1, 0, 1, 0, 1, 0] };
    expect(alphaOf(inverted)).toEqual([0, 255, 255]);
  });

  it('paints the whole image where no sample is keyed out', () => {
    expect(alphaOf({ Mask: [128, 128, 128, 128, 128, 128] })).toBeUndefined();
  });
});
