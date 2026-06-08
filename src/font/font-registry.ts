// Registry of font variants. The renderer resolves a (bold, italic) pair to
// the closest available variant — when an exact match is missing, the
// registry falls back through bold → italic → regular.

import type { ParsedTtf } from '@/font/ttf-parser';
import { parseTtf } from '@/font/ttf-parser';

export type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic';

export interface FontBytesByVariant {
  readonly regular: Uint8Array;
  readonly bold?: Uint8Array;
  readonly italic?: Uint8Array;
  readonly boldItalic?: Uint8Array;
}

export class FontRegistry {
  private constructor(private readonly fonts: ReadonlyMap<FontVariant, ParsedTtf>) {}

  static fromBytes(input: FontBytesByVariant): FontRegistry {
    const map = new Map<FontVariant, ParsedTtf>();
    map.set('regular', parseTtf(input.regular));
    if (input.bold) map.set('bold', parseTtf(input.bold));
    if (input.italic) map.set('italic', parseTtf(input.italic));
    if (input.boldItalic) map.set('boldItalic', parseTtf(input.boldItalic));
    return new FontRegistry(map);
  }

  resolveByStyle(bold: boolean, italic: boolean): { variant: FontVariant; parsed: ParsedTtf } {
    const candidates: Array<FontVariant> =
      bold && italic
        ? ['boldItalic', 'bold', 'italic', 'regular']
        : bold
          ? ['bold', 'regular']
          : italic
            ? ['italic', 'regular']
            : ['regular'];
    for (const v of candidates) {
      const parsed = this.fonts.get(v);
      if (parsed) return { variant: v, parsed };
    }
    throw new Error('FontRegistry has no regular font');
  }

  entries(): IterableIterator<[FontVariant, ParsedTtf]> {
    return this.fonts.entries();
  }

  hasVariant(v: FontVariant): boolean {
    return this.fonts.has(v);
  }
}
