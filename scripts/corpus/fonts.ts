// What the corpus renders with — and why it is no longer Roboto.
//
// Every number the pixel harness prints is a distance to LibreOffice's own
// render, so the typeface counts twice: once for the shapes, and once for the
// line breaks that decide where a page ends. Audited over the 1896 cached
// reference renders (the `/BaseFont` of every face they embed), LibreOffice on
// this host drew the corpus with:
//
//   Carlito 875 files · Arial 268 · Times New Roman 168 · Liberation Serif 84
//   Arial Unicode 71 · Liberation Sans 57 · OpenSymbol 46 · Caladea 44
//   Symbol 34 · Wingdings 28 · Courier New 18 · Tahoma 18 · Arial Black 17 …
//
// — its own bundled substitutes (Carlito for Calibri, Caladea for Cambria,
// Liberation for the Croscore families) plus whatever the system holds: macOS
// ships the real Arial, Times New Roman, Courier New, Verdana and Tahoma.
//
// Our own substitution table already aims at that set — calibri→Carlito,
// cambria→Caladea, arial→Arimo, times→Tinos, courier→Cousine — and Arimo, Tinos
// and Cousine ARE the Liberation designs, metric-compatible with Arial, Times
// and Courier. Roboto is none of them. Worse, passing explicit `fonts` at all
// short-circuits the entire font pipeline: the provider chain, the substitution
// table and the per-family registries never run, so the harness measured a
// document rendered in one family and called the difference layout.
//
// So the sweep renders through the library's own automatic path, and the CDN is
// visited once: every file it returns is kept under `corpus/.fonts`, keyed by
// its name, the way `.lo-cache` keeps the reference renders. `CORPUS_FONTS=
// roboto` restores the old single-family measurement for a side-by-side.

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { FetchLike } from '@/core/fonts';
import type { FontBytesByVariant } from '@/core/font';
import { Ream } from '@/core/converter/ream';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');

/** Where the downloaded substitutes live, so a sweep runs offline after the first. */
export const FONT_CACHE_DIR = resolve(root, 'corpus/.fonts');

/** The old measurement's single family, kept for `CORPUS_FONTS=roboto` and for CJK. */
export const ROBOTO: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(
    readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
  ),
};

/** Whether the sweep is pinned to the old single-family render. */
export const EXPLICIT_FONTS = process.env.CORPUS_FONTS === 'roboto';

const bufferOf = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/**
 * A `fetch` for the font path that answers from disk when it can. A miss is
 * fetched once and written under {@link FONT_CACHE_DIR} by the file's own name
 * (`Carlito_400Regular.ttf`), which is unique across the curated families.
 */
export function cachedFontFetch(): FetchLike {
  return async (url: string) => {
    mkdirSync(FONT_CACHE_DIR, { recursive: true });
    const file = resolve(FONT_CACHE_DIR, basename(new URL(url).pathname));
    if (existsSync(file)) {
      const bytes = new Uint8Array(readFileSync(file));
      return { ok: true, arrayBuffer: () => Promise.resolve(bufferOf(bytes)) };
    }
    const res = await fetch(url);
    if (!res.ok) return { ok: false, arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)) };
    const buf = await res.arrayBuffer();
    writeFileSync(file, Buffer.from(buf));
    return { ok: true, arrayBuffer: () => Promise.resolve(buf) };
  };
}

/**
 * The font options every corpus render passes: nothing but a cached `fetch`, so
 * the library resolves the document's own families — or the pinned set under
 * `CORPUS_FONTS=roboto`.
 */
export function corpusFontOptions(): { fonts: FontBytesByVariant } | { fontFetch: FetchLike } {
  return EXPLICIT_FONTS ? { fonts: ROBOTO } : { fontFetch: cachedFontFetch() };
}

// Neither Roboto nor the curated substitutes have Hangul, Kana or Han, so a CJK
// document renders as a page of tofu and every real difference hides behind it —
// 1_NoIden.xlsx is seven rows of Korean. LibreOffice substitutes a system face;
// so does the harness, when the document needs one and the host has one. The
// library itself does not: a registry is ONE family in four weights, so
// per-script fallback is the caller's business (and, for Ream, F4's).
const CJK_FACES = [
  '/System/Library/Fonts/Supplemental/AppleGothic.ttf',
  '/System/Library/Fonts/Supplemental/AppleMyungjo.ttf',
  '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
];

const NEEDS_CJK =
  /[\u1100-\u11FF\u3040-\u30FF\u3130-\u318F\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF]/u;

/**
 * The one face to render `input` with when its text is CJK, or `undefined` when
 * the ordinary path will do.
 *
 * @param input Path to the source document.
 * @param note  Where to say which face was chosen (the diff tool prints it).
 */
export function cjkFontsFor(
  input: string,
  note?: (msg: string) => void,
): FontBytesByVariant | undefined {
  let text = '';
  try {
    text = JSON.stringify(Ream.parse(new Uint8Array(readFileSync(input))).flow.body);
  } catch {
    return undefined;
  }
  if (!NEEDS_CJK.test(text)) return undefined;
  const face = CJK_FACES.find((p) => existsSync(p));
  if (!face) {
    note?.('note: document has CJK text and no CJK face was found on this host');
    return undefined;
  }
  note?.(`note: CJK text — rendering with ${basename(face)}`);
  // One face for every weight: the substitute has no bold, and pairing a Latin
  // bold with a CJK regular would measure one script in the other's metrics.
  return { regular: new Uint8Array(readFileSync(face)) };
}
