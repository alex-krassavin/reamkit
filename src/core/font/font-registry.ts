// Registry of font variants. The renderer resolves a (bold, italic) pair to
// the closest available variant — when an exact match is missing, the
// registry falls back through bold → italic → regular.

import type { ParsedTtf } from '@/core/font/ttf-parser';
import { parseTtf } from '@/core/font/ttf-parser';

export type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic';

// One expert for face fallback (oop-design §8, A4): the candidate cascade a
// missing variant degrades through. Shared by the registry and every font
// provider — three hand-rolled copies had already drifted (the remote
// provider lost the boldItalic→bold/italic degradation).
export function pickVariant(
  has: (variant: FontVariant) => boolean,
  bold: boolean,
  italic: boolean,
): FontVariant | undefined {
  const candidates: ReadonlyArray<FontVariant> =
    bold && italic
      ? ['boldItalic', 'bold', 'italic', 'regular']
      : bold
        ? ['bold', 'regular']
        : italic
          ? ['italic', 'regular']
          : ['regular'];
  return candidates.find(has);
}

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
    const variant = pickVariant((v) => this.fonts.has(v), bold, italic);
    const parsed = variant ? this.fonts.get(variant) : undefined;
    if (!variant || !parsed) throw new Error('FontRegistry has no regular font');
    return { variant, parsed };
  }

  entries(): IterableIterator<[FontVariant, ParsedTtf]> {
    return this.fonts.entries();
  }

  hasVariant(v: FontVariant): boolean {
    return this.fonts.has(v);
  }
}
