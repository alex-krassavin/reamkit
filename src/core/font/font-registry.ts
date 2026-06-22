// Registry of font variants. The renderer resolves a (bold, italic) pair to
// the closest available variant — when an exact match is missing, the
// registry falls back through bold → italic → regular.

import type { ParsedTtf } from '@/core/font/ttf-parser';
import { parseTtf } from '@/core/font/ttf-parser';

/** A font style slot: regular, bold, italic or bold-italic. */
export type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic';

/**
 * The single owner of face fallback (oop-design §8, A4): the candidate cascade a
 * missing variant degrades through (`boldItalic → bold → italic → regular`),
 * shared by the registry and every font provider.
 *
 * @param has    Predicate — does a given variant exist?
 * @param bold   Whether bold was requested.
 * @param italic Whether italic was requested.
 * @returns The best available variant, or `undefined` when even `regular` is missing.
 */
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

/** Raw TTF/OTF bytes per style variant; only `regular` is required. */
export interface FontBytesByVariant {
  /** The regular (upright, normal-weight) face — required. */
  readonly regular: Uint8Array;
  /** The bold face, if available. */
  readonly bold?: Uint8Array;
  /** The italic face, if available. */
  readonly italic?: Uint8Array;
  /** The bold-italic face, if available. */
  readonly boldItalic?: Uint8Array;
}

/**
 * A registry of parsed font variants. The renderer resolves a `(bold, italic)`
 * request to the closest available face via {@link pickVariant}, degrading
 * through `bold → italic → regular` when an exact match is missing.
 */
export class FontRegistry {
  private constructor(private readonly fonts: ReadonlyMap<FontVariant, ParsedTtf>) {}

  /**
   * Parse a set of font bytes into a registry.
   *
   * @param input The bytes per variant (`regular` required).
   * @returns A registry holding the parsed faces.
   */
  static fromBytes(input: FontBytesByVariant): FontRegistry {
    const map = new Map<FontVariant, ParsedTtf>();
    map.set('regular', parseTtf(input.regular));
    if (input.bold) map.set('bold', parseTtf(input.bold));
    if (input.italic) map.set('italic', parseTtf(input.italic));
    if (input.boldItalic) map.set('boldItalic', parseTtf(input.boldItalic));
    return new FontRegistry(map);
  }

  /**
   * Resolve a style request to the closest available face.
   *
   * @param bold   Whether bold was requested.
   * @param italic Whether italic was requested.
   * @returns The chosen variant and its parsed font.
   * @throws Error when the registry has no usable (regular) font.
   */
  resolveByStyle(bold: boolean, italic: boolean): { variant: FontVariant; parsed: ParsedTtf } {
    const variant = pickVariant((v) => this.fonts.has(v), bold, italic);
    const parsed = variant ? this.fonts.get(variant) : undefined;
    if (!variant || !parsed) throw new Error('FontRegistry has no regular font');
    return { variant, parsed };
  }

  /** Iterate the `[variant, parsed font]` pairs the registry holds. */
  entries(): IterableIterator<[FontVariant, ParsedTtf]> {
    return this.fonts.entries();
  }

  /** Whether a given variant is present. */
  hasVariant(v: FontVariant): boolean {
    return this.fonts.has(v);
  }
}
