// A TrueType program that NUMBERS its glyphs, as Quartz subsets a chart's
// face: one Macintosh `cmap` subtable whose codes count up from 33 in the order
// the chart first drew each glyph, and no glyph names. Built from Roboto with
// its `cmap` replaced and its `post` table renamed out of reach.

import { readFileSync } from 'node:fs';

import { parseTtf } from '@/core/font';

export const ROBOTO = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));

/** The glyph index Roboto uses for one character. */
export function robotoGlyph(ch: string): number {
  return parseTtf(ROBOTO).glyphForCodepoint(ch.codePointAt(0)!);
}

/**
 * Roboto with its `cmap` one (1,0) subtable in format 6 that maps codes from
 * `first` onto the glyphs of `chars`, in that order.
 *
 * @param chars     The characters whose glyphs the codes reach.
 * @param first     The first code.
 * @param keepNames Whether the `post` table's glyph names stay readable.
 */
export function numberedProgram(chars: string, first = 33, keepNames = false): Uint8Array {
  const gids = [...chars].map(robotoGlyph);
  // version, one record: platform 1, encoding 0, at offset 12 — then format 6.
  const words = [0, 1, 1, 0, 0, 12, 6, 10 + 2 * gids.length, 0, first, gids.length, ...gids];
  const table = Uint8Array.from(words.flatMap((w) => [w >> 8, w & 0xff]));
  const start = Math.ceil(ROBOTO.length / 4) * 4;
  const out = new Uint8Array(start + table.length);
  out.set(ROBOTO);
  out.set(table, start);
  const view = new DataView(out.buffer);
  const count = view.getUint16(4);
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    const tag = String.fromCharCode(...out.subarray(at, at + 4));
    if (tag === 'cmap') {
      view.setUint32(at + 8, start);
      view.setUint32(at + 12, table.length);
    } else if (tag === 'post' && !keepNames) {
      out[at] = 'z'.charCodeAt(0);
    }
  }
  return out;
}

/** The `/ToUnicode` Quartz writes for such a face: a CMap that maps nothing. */
export const EMPTY_TO_UNICODE = [
  '/CIDInit /ProcSet findresource begin',
  '12 dict begin',
  'begincmap',
  '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
  '/CMapName /Adobe-Identity-UCS def',
  '/CMapType 2 def',
  '1 begincodespacerange',
  '<00> <FF>',
  'endcodespacerange',
  'endcmap',
  'CMapName currentdict /CMap defineresource pop',
  'end',
  'end',
].join('\n');
