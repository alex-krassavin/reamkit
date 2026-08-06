// ECMA-376 §17.3.2.26 `w:rFonts` + §17.18.88 ST_Theme — the typeface a document
// names by pointing at its theme instead of spelling it out.
//
// WordprocessingML has its own spelling of the indirection DrawingML writes as
// `+mn-lt`: a slot name in a parallel attribute (`w:asciiTheme="minorHAnsi"`)
// beside the literal one. It resolves against the same `a:fontScheme` — the
// theme's two fonts, each in three scripts — so this is only the vocabulary,
// and `resolveThemeFont` is where the DrawingML half lives.
//
// It is not a corner: 828 of the corpus's 1121 .docx name a font this way and
// 414 name it NO other way, their theme reading Calibri. Left unresolved they
// come out in whatever the default is — a sans where LibreOffice sets Carlito,
// which is a wider line, a different break and a different page.

import type { ThemeFontSlots, ThemeFonts } from '@/core/drawingml/theme-parser';

// §17.18.88 — the eight values, and the (font, script) each one names. `Ascii`
// and `HAnsi` are both the latin slot: the theme has one latin font and Word
// spends two attributes pointing at it.
const THEME_SLOTS: ReadonlyMap<string, readonly ['major' | 'minor', keyof ThemeFontSlots]> =
  new Map([
    ['majorascii', ['major', 'latin']],
    ['majorhansi', ['major', 'latin']],
    ['majoreastasia', ['major', 'ea']],
    ['majorbidi', ['major', 'cs']],
    ['minorascii', ['minor', 'latin']],
    ['minorhansi', ['minor', 'latin']],
    ['minoreastasia', ['minor', 'ea']],
    ['minorbidi', ['minor', 'cs']],
  ]);

/**
 * Resolve a `w:rFonts` theme attribute to the typeface it stands for.
 *
 * @param value The attribute as written (`minorHAnsi`, `majorBidi`, …).
 * @param fonts The theme's font scheme, when the part was read.
 * @returns The typeface, or `undefined` for an unknown slot, a missing theme,
 *          or a slot the theme leaves empty.
 */
export function resolveWordThemeFont(
  value: string | undefined,
  fonts: ThemeFonts | undefined,
): string | undefined {
  if (value === undefined || !fonts) return undefined;
  const slot = THEME_SLOTS.get(value.trim().toLowerCase());
  if (!slot) return undefined;
  const [which, script] = slot;
  return fonts[which][script];
}
