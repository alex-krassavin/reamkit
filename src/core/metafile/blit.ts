// MS-WMF §2.1.1.31 — what a blit's ternary raster operation asks for, and the
// one idiom that needs two records to say it.
//
// A blit combines the source, the destination and the current brush through a
// boolean function named by its ROP code. Only a handful of the 256 appear in
// documents, and they divide into: paint the source (SRCCOPY), paint the brush
// (PATCOPY and the two constants), and the AND/OR pair a transparent picture is
// drawn with — a monochrome mask blitted with SRCAND to knock a hole in the
// ground, then the picture itself ORed into the hole with SRCPAINT.
//
// That pair is one bitmap with an alpha channel, and this is where the two
// records are put back together: both readers hand their blits to a blitter,
// which holds a mask until it sees what it belongs to.

import type { DibImage } from '@/core/metafile/dib';
import type { PictureImage } from '@/core/metafile/picture';
import { dibToPng, maskedDib } from '@/core/metafile/dib';

/** How a blit's ROP wants the source drawn. */
export type BlitMode =
  /** SRCCOPY and its neighbours: the source replaces what is under it. */
  | 'opaque'
  /** SRCAND: a mask, held until the picture it belongs to arrives. */
  | 'mask'
  /** SRCPAINT: ORed in, so the source's BLACK leaves the ground showing. */
  | 'or'
  /** A ROP that ignores the source, or one whose result is not a picture. */
  | 'skip';

/** The destination rectangle a blit paints, in the metafile's own units. */
export interface BlitRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * What a ternary raster operation asks a reader to draw.
 *
 * @param rop The ROP code, as the record states it.
 */
export function blitMode(rop: number): BlitMode {
  switch (rop >>> 0) {
    case 0x00cc0020: // SRCCOPY
    case 0x00e20746: // DSPDxax — the source through the brush, drawn as itself
      return 'opaque';
    case 0x008800c6: // SRCAND
      return 'mask';
    case 0x00ee0086: // SRCPAINT
      return 'or';
    case 0x00660046: // SRCINVERT — an XOR, which is a rubber band, not a picture
    case 0x00f00021: // PATCOPY, and the constants: no source at all
    case 0x00000042:
    case 0x00ff0062:
      return 'skip';
    default:
      // The rest are rare enough that no corpus file draws one. A ROP that
      // reads the source is drawn as itself; one that does not is not a
      // picture at all. MS-WMF §2.1.1.31: the source bits of the code are the
      // ones that move when the source does.
      return ((rop >>> 2) & 0x00330000) !== (rop & 0x00330000) ? 'opaque' : 'skip';
  }
}

/**
 * A sink for a metafile's blits that puts the AND/OR pair back together.
 *
 * @param emit Where a finished picture goes — the reader's primitive list.
 */
export function makeBlitter(emit: (prim: PictureImage) => void): {
  blit: (img: DibImage, dest: BlitRect, rop: number) => void;
} {
  let pending: { img: DibImage; dest: BlitRect } | undefined;
  const place = (img: DibImage, dest: BlitRect): void => {
    emit({ kind: 'image', ...dest, png: dibToPng(img) });
  };
  return {
    blit: (img, dest, rop) => {
      const mode = blitMode(rop);
      if (mode === 'skip') return;
      if (mode === 'mask') {
        // A lone mask is a black silhouette, which is worse than nothing: it is
        // dropped if the picture it belongs to never comes.
        pending = { img, dest };
        return;
      }
      const mask = pending;
      pending = undefined;
      if (mode === 'or') {
        // Its own black is the transparent part either way — through the mask
        // when there is one, and by the OR's own arithmetic when there is not.
        place(mask && sameRect(mask.dest, dest) ? maskedDib(img, mask.img) : keyBlack(img), dest);
        return;
      }
      place(img, dest);
    },
  };
}

function sameRect(a: BlitRect, b: BlitRect): boolean {
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

// An OR blit leaves the destination showing wherever the source is black, so
// black IS the transparency when no mask says otherwise.
function keyBlack(img: DibImage): DibImage {
  const rgba = new Uint8Array(img.rgba);
  for (let i = 0; i < rgba.length; i += 4) {
    if (rgba[i] === 0 && rgba[i + 1] === 0 && rgba[i + 2] === 0) rgba[i + 3] = 0;
  }
  return { width: img.width, height: img.height, rgba };
}
