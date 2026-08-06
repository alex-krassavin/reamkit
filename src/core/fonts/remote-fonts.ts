// Remote font provider — downloads open-licensed fonts so the converters can
// "just work" without the caller supplying font bytes.
//
// Word/Excel documents reference fonts by name (Calibri, Arial, Times…), most
// of which are proprietary and can't be bundled. We map each referenced family
// to an open substitute hosted on a CDN and fetch the raw TrueType file at
// runtime. Source: @expo-google-fonts packages served as static .ttf via
// jsDelivr (full character sets — Latin, Cyrillic, Greek — in a single file,
// parseable directly by our TTF parser).
//
// Two tiers, mirroring LibreOffice's substitution:
//   • Metric-compatible twins — open fonts engineered to reproduce a specific
//     proprietary font's advance widths 1:1, so text breaks into lines at the
//     same points and visual parity is much closer:
//       Calibri → Carlito   Cambria → Caladea
//       Arial   → Arimo     Times New Roman → Tinos   Courier New → Cousine
//     (the Croscore + Carlito/Caladea set LibreOffice bundles for this purpose).
//   • Class fallback — for families without a known twin, a reasonable style
//     match (serif → Tinos, mono → Cousine, else → Arimo). Widths are only
//     approximate here; the twins above are width-exact.
//
// fetch() is used (universal in Node 18+, browsers, and edge runtimes); a
// custom implementation can be injected for tests / offline use.

import type { FontBytesByVariant, FontVariant } from '@/core/font';

const CDN_BASE = 'https://cdn.jsdelivr.net/npm/@expo-google-fonts';

const VARIANT_SUFFIX: Record<FontVariant, string> = {
  regular: '400Regular',
  bold: '700Bold',
  italic: '400Regular_Italic',
  boldItalic: '700Bold_Italic',
};

/** A curated open substitute family (the Croscore + Carlito/Caladea set, like LibreOffice). */
export type FamilyKey = 'arimo' | 'tinos' | 'cousine' | 'carlito' | 'caladea';

/**
 * A WRITING SYSTEM the curated five cannot draw at all. They are Latin families
 * — Greek, Cyrillic and Hebrew ride along, everything else is a notdef box —
 * and 2145 characters of the corpus fall outside them: Han, Kana, Hangul,
 * Arabic, the geometric symbols a form's checkbox is made of.
 *
 * These are fetched only when a document actually holds such a character, and
 * only in the regular weight: Noto Sans SC is ten megabytes, and a bold run in
 * it is better stroked (see `SyntheticFace`) than downloaded four times over.
 */
export type ScriptKey = 'jp' | 'kr' | 'sc' | 'arabic' | 'hebrew' | 'thai' | 'symbols';

/** Either kind of substitute — a Latin family or a per-script face. */
export type SubstituteKey = FamilyKey | ScriptKey;

interface ScriptFamily {
  readonly pkg: string;
  readonly file: string;
}

const SCRIPTS: Record<ScriptKey, ScriptFamily> = {
  jp: { pkg: 'noto-sans-jp', file: 'NotoSansJP' },
  kr: { pkg: 'noto-sans-kr', file: 'NotoSansKR' },
  sc: { pkg: 'noto-sans-sc', file: 'NotoSansSC' },
  arabic: { pkg: 'noto-sans-arabic', file: 'NotoSansArabic' },
  hebrew: { pkg: 'noto-sans-hebrew', file: 'NotoSansHebrew' },
  thai: { pkg: 'noto-sans-thai', file: 'NotoSansThai' },
  symbols: { pkg: 'noto-sans-symbols-2', file: 'NotoSansSymbols2' },
};

/** Whether a substitute key names a writing system rather than a Latin family. */
export function isScriptKey(key: SubstituteKey): key is ScriptKey {
  return key in SCRIPTS;
}

/**
 * Fetch the ONE face a writing system is drawn with.
 *
 * @param script The writing system.
 * @param fetchImpl Injectable `fetch` (defaults to the global one).
 * @returns The regular face, or `undefined` when the download fails.
 */
export async function fetchScriptFont(
  script: ScriptKey,
  fetchImpl: FetchLike = globalThis.fetch.bind(globalThis),
): Promise<FontBytesByVariant | undefined> {
  const family = SCRIPTS[script];
  const url = `${CDN_BASE}/${family.pkg}/400Regular/${family.file}_400Regular.ttf`;
  const bytes = await fetchTtf(url, fetchImpl, false);
  return bytes ? { regular: bytes } : undefined;
}

interface CuratedFamily {
  readonly pkg: string; // @expo-google-fonts package name
  readonly file: string; // capitalised file prefix
  // Newer @expo-google-fonts packages nest each variant in its own folder
  // (`/400Regular/Carlito_400Regular.ttf`) rather than flat at the package root.
  readonly nested?: boolean;
}

const FAMILIES: Record<FamilyKey, CuratedFamily> = {
  arimo: { pkg: 'arimo', file: 'Arimo' },
  tinos: { pkg: 'tinos', file: 'Tinos' },
  cousine: { pkg: 'cousine', file: 'Cousine' },
  carlito: { pkg: 'carlito', file: 'Carlito', nested: true },
  caladea: { pkg: 'caladea', file: 'Caladea' },
};

// Metric-compatible twins: an open font engineered to reproduce the named
// proprietary font's advance widths 1:1. Matched before the class fallback so
// e.g. Cambria resolves to its exact twin Caladea, not the generic serif Tinos.
const EXACT: Record<string, FamilyKey> = {
  calibri: 'carlito',
  // A theme's HEADING font in 240 of the corpus's documents. It is Calibri's
  // own light weight, so Carlito is its twin too — read as an unknown name it
  // went to the generic sans and set every heading in Arial's widths.
  'calibri light': 'carlito',
  cambria: 'caladea',
  arial: 'arimo',
  helvetica: 'arimo',
  'liberation sans': 'arimo',
  'times new roman': 'tinos',
  times: 'tinos',
  'liberation serif': 'tinos',
  'courier new': 'cousine',
  courier: 'cousine',
  'liberation mono': 'cousine',
};

// Class fallback for families without an exact twin — a style match only
// (widths approximate). Anything not serif/mono falls through to sans (Arimo).
const SERIF = new Set([
  'georgia',
  'garamond',
  'book antiqua',
  'palatino',
  'pt serif',
  'minion pro',
  'serif',
]);

const MONO = new Set([
  'consolas',
  'monaco',
  'menlo',
  'lucida console',
  'dejavu sans mono',
  'monospace',
]);

// The words a family name ends with to say which MEMBER of the family it is —
// a weight, a width or a slant. Each maps to what the substitute can do about
// it: take its bold cut, take its italic, squeeze it. `none` is a face no
// curated family ships (light, medium, thin), and the regular one is the honest
// answer there — which is what LibreOffice does with them too.
const WEIGHT_WORDS: ReadonlyMap<string, 'bold' | 'italic' | 'narrow' | 'none'> = new Map([
  ['black', 'bold'],
  ['heavy', 'bold'],
  ['bold', 'bold'],
  ['semibold', 'bold'],
  ['demibold', 'bold'],
  ['demi', 'bold'],
  ['italic', 'italic'],
  ['oblique', 'italic'],
  ['light', 'none'],
  ['thin', 'none'],
  ['medium', 'none'],
  ['regular', 'none'],
  ['book', 'none'],
  ['narrow', 'narrow'],
  ['condensed', 'narrow'],
  ['cond', 'narrow'],
  ['expanded', 'none'],
  ['extra', 'none'],
  ['ultra', 'none'],
  ['pro', 'none'],
]);

/** A family name, read: which substitute it maps to and what it says about the face. */
export interface FamilyStyle {
  readonly key: FamilyKey;
  /** The NAME asks for a heavy face (`Arial Black`, `DIN-Bold`). */
  readonly bold?: boolean;
  /** …or a slanted one (`Frutiger Oblique`). */
  readonly italic?: boolean;
  /**
   * …or a narrow one (`Arial Narrow`), as the fraction of the normal advance it
   * sets at. No curated family ships a condensed cut, so the substitute is
   * squeezed instead — Arial Narrow's own widths are 82 % of Arial's, which is
   * also what LibreOffice's Liberation Sans Narrow reproduces.
   */
  readonly widthScale?: number;
}

/** Arial Narrow's advance widths, as a fraction of Arial's. */
const NARROW_SCALE = 0.82;

/**
 * Map a document-referenced font family to a curated open substitute: an exact
 * metric twin when one is known (e.g. Calibri → Carlito), otherwise a
 * serif/mono/sans style fallback.
 *
 * The name may also carry the FACE — `Times New Roman Bold` is the Times
 * family, and read whole it matches no twin at all and fell through to the
 * generic sans, which is how a Times heading came out in a grotesque.
 *
 * @param name The referenced family name (case-insensitive); empty ⇒ Arimo.
 * @returns The chosen family and what the name said about the face.
 */
export function resolveFamilyStyle(name: string | undefined): FamilyStyle {
  if (!name) return { key: 'arimo' };
  // PostScript spells the face with a hyphen (`CenturySchoolbook-Bold`), and a
  // stray comma is how some producers separate it (`Arial,Bold`).
  const words = name
    .trim()
    .toLowerCase()
    .split(/[\s\-_,]+/u)
    .filter((w) => w !== '');
  let bold = false;
  let italic = false;
  let narrow = false;
  // Strip the face words off the END — a family may be NAMED for one of them
  // (Book Antiqua), and only a trailing word is the face.
  while (words.length > 1) {
    const word = WEIGHT_WORDS.get(words[words.length - 1]!);
    if (word === undefined) break;
    if (word === 'bold') bold = true;
    if (word === 'italic') italic = true;
    if (word === 'narrow') narrow = true;
    words.pop();
  }
  // The whole name first, then the words a foundry prefix or a modifier may be
  // hiding the family behind: `Adobe Garamond Pro` is a Garamond.
  const tries = [words.join(' '), words[words.length - 1]!, words[0]!];
  let key: FamilyKey = 'arimo';
  for (const n of tries) {
    const found = EXACT[n] ?? (MONO.has(n) ? 'cousine' : SERIF.has(n) ? 'tinos' : undefined);
    if (found) {
      key = found;
      break;
    }
  }
  return {
    key,
    ...(bold ? { bold } : {}),
    ...(italic ? { italic } : {}),
    ...(narrow ? { widthScale: NARROW_SCALE } : {}),
  };
}

/**
 * The curated substitute for a family name — {@link resolveFamilyStyle} without
 * the face it also carries.
 *
 * @param name The referenced family name (case-insensitive); empty ⇒ Arimo.
 * @returns The chosen {@link FamilyKey}.
 */
export function resolveFamilyKey(name: string | undefined): FamilyKey {
  return resolveFamilyStyle(name).key;
}

function fontUrl(family: CuratedFamily, variant: FontVariant): string {
  const suffix = VARIANT_SUFFIX[variant];
  const leaf = `${family.file}_${suffix}.ttf`;
  // Nested packages place each variant under its own suffix-named folder.
  return family.nested
    ? `${CDN_BASE}/${family.pkg}/${suffix}/${leaf}`
    : `${CDN_BASE}/${family.pkg}/${leaf}`;
}

/** A minimal `fetch`-like function (injectable for tests / offline use). */
export type FetchLike = (
  url: string,
) => Promise<{ ok: boolean; arrayBuffer: () => Promise<ArrayBuffer> }>;

// Cache by URL so repeated conversions (and the four variants of one family)
// don't re-download. Stores the in-flight promise to dedupe concurrent fetches.
const cache = new Map<string, Promise<Uint8Array | undefined>>();

async function fetchTtf(
  url: string,
  fetchImpl: FetchLike,
  required: boolean,
): Promise<Uint8Array | undefined> {
  let pending = cache.get(url);
  if (!pending) {
    pending = (async () => {
      const res = await fetchImpl(url);
      if (!res.ok) return undefined;
      const buf = await res.arrayBuffer();
      const bytes = new Uint8Array(buf);
      // Guard against CDN error pages masquerading as 200: a real sfnt starts
      // with 0x00010000 (TrueType) or 'OTTO'/'true'/'ttcf'.
      if (bytes.length < 4) return undefined;
      const sig = (bytes[0]! << 24) | (bytes[1]! << 16) | (bytes[2]! << 8) | bytes[3]!;
      const ok =
        sig === 0x00010000 || sig === 0x4f54544f || sig === 0x74727565 || sig === 0x74746366;
      return ok ? bytes : undefined;
    })();
    cache.set(url, pending);
  }
  const result = await pending;
  if (!result && required) {
    cache.delete(url);
    throw new Error(`Failed to download font from ${url}`);
  }
  return result;
}

/** Options for {@link fetchFontSet}. */
export interface FetchFontSetOptions {
  /** Document font family name to substitute, or a curated {@link FamilyKey} directly. */
  readonly family?: string | FamilyKey;
  /** Injectable fetch (defaults to the global `fetch`); lets tests run offline. */
  readonly fetch?: FetchLike;
}

/**
 * Download a full variant set (regular required; bold/italic/bold-italic
 * best-effort) for the open family that best matches the requested name.
 *
 * @param options The family to substitute and an optional fetch override.
 * @returns The downloaded font bytes per variant.
 * @throws Error when the required regular face cannot be downloaded.
 */
export async function fetchFontSet(options: FetchFontSetOptions = {}): Promise<FontBytesByVariant> {
  const fetchImpl: FetchLike = options.fetch ?? ((url) => fetch(url));
  const key: FamilyKey =
    options.family && options.family in FAMILIES
      ? (options.family as FamilyKey)
      : resolveFamilyKey(options.family);
  const family = FAMILIES[key];

  const regular = await fetchTtf(fontUrl(family, 'regular'), fetchImpl, true);
  const [bold, italic, boldItalic] = await Promise.all([
    fetchTtf(fontUrl(family, 'bold'), fetchImpl, false),
    fetchTtf(fontUrl(family, 'italic'), fetchImpl, false),
    fetchTtf(fontUrl(family, 'boldItalic'), fetchImpl, false),
  ]);

  return {
    regular: regular!,
    ...(bold ? { bold } : {}),
    ...(italic ? { italic } : {}),
    ...(boldItalic ? { boldItalic } : {}),
  };
}

/** Test / diagnostic helper: clear the in-memory font-download cache. */
export function clearFontCache(): void {
  cache.clear();
}
