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

export type FamilyKey = 'arimo' | 'tinos' | 'cousine' | 'carlito' | 'caladea';

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

// Map a document-referenced font family to a curated open substitute: an exact
// metric twin when one is known, otherwise a serif/mono/sans style fallback.
export function resolveFamilyKey(name: string | undefined): FamilyKey {
  if (!name) return 'arimo';
  const n = name.trim().toLowerCase();
  const exact = EXACT[n];
  if (exact) return exact;
  if (MONO.has(n)) return 'cousine';
  if (SERIF.has(n)) return 'tinos';
  return 'arimo';
}

function fontUrl(family: CuratedFamily, variant: FontVariant): string {
  const suffix = VARIANT_SUFFIX[variant];
  const leaf = `${family.file}_${suffix}.ttf`;
  // Nested packages place each variant under its own suffix-named folder.
  return family.nested
    ? `${CDN_BASE}/${family.pkg}/${suffix}/${leaf}`
    : `${CDN_BASE}/${family.pkg}/${leaf}`;
}

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

export interface FetchFontSetOptions {
  // Document font family name to substitute, or a curated key directly.
  readonly family?: string | FamilyKey;
  // Injectable fetch (defaults to global fetch). Lets tests run offline.
  readonly fetch?: FetchLike;
}

// Download a full variant set (regular required; bold/italic/boldItalic
// best-effort) for the family that best matches the requested name.
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

// Test/diagnostic helper: clear the in-memory download cache.
export function clearFontCache(): void {
  cache.clear();
}
