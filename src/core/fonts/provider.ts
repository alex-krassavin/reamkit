// FontProvider chain (ir-design §8, handoff v1 §6) — font resolution as an
// adapter pipeline. A request (family, bold, italic) walks the providers; the
// first byte-level answer wins. Everything below an exact caller/embedded hit
// is a substitution the conversion reports as a Loss (severity 'substituted').
//
// Chain order (the facade builds it): embedded-in-document → caller bytes →
// local (Local Font Access, opt-in) → remote CDN substitutes. The 'metrics'
// answer kind (canvas-metrics measuring without bytes) is declared per the
// design but intentionally unimplemented — it is the experimental stage and
// lands behind a flag once the chain is stable.

import type { FamilyKey, FetchLike } from '@/core/fonts/remote-fonts';
import type { FontBytesByVariant, FontRegistry, FontVariant } from '@/core/font';

import { pickVariant } from '@/core/font';
import { fetchFontSet, resolveFamilyKey } from '@/core/fonts/remote-fonts';

/** A font resolution request: the wanted family and style. */
export interface FontRequest {
  /** Family name as the document references it (e.g. "Times New Roman"). */
  readonly family?: string;
  /** Whether a bold face is wanted. */
  readonly bold: boolean;
  /** Whether an italic face is wanted. */
  readonly italic: boolean;
}

/** A provider's answer: resolved bytes, or `none` to fall through to the next provider. */
export type FontAnswer =
  | {
      readonly kind: 'bytes';
      readonly bytes: Uint8Array;
      /** Face the bytes actually are (for substitution reporting). */
      readonly faceName: string;
      /** Which provider answered (chain bookkeeping). */
      readonly providerId: string;
    }
  | { readonly kind: 'none' };

/** The shared `none` answer — a provider that cannot satisfy a request. */
export const NO_FONT: FontAnswer = { kind: 'none' };

/** One stage in the font-resolution chain (ir-design §8). */
export interface FontProvider {
  /** Provider id: `'embedded'` | `'caller'` | `'local'` | `'remote'` | a custom id. */
  readonly id: string;
  /** Resolve a request to bytes, or {@link NO_FONT} to defer to the next provider. */
  resolve: (req: FontRequest) => Promise<FontAnswer>;
}

/** First byte-level answer wins; 'none' falls through to the next provider. */
export function chainProviders(providers: ReadonlyArray<FontProvider>): FontProvider {
  return {
    id: 'chain',
    resolve: async (req) => {
      for (const p of providers) {
        const answer = await p.resolve(req);
        if (answer.kind !== 'none') return answer;
      }
      return NO_FONT;
    },
  };
}

function variantFor(req: FontRequest): FontVariant {
  if (req.bold && req.italic) return 'boldItalic';
  if (req.bold) return 'bold';
  if (req.italic) return 'italic';
  return 'regular';
}

/**
 * Caller-supplied bytes — answers every family (the caller chose these fonts
 * deliberately). Falls back through bold/italic → regular like the registry.
 */
export function callerFontProvider(fonts: FontBytesByVariant): FontProvider {
  return {
    id: 'caller',
    resolve: (req) => {
      const v = variantFor(req);
      const picked = pickVariant((x) => fonts[x] !== undefined, req.bold, req.italic) ?? 'regular';
      const bytes = fonts[picked] ?? fonts.regular;
      return Promise.resolve({
        kind: 'bytes',
        bytes,
        faceName: `caller:${v}`,
        providerId: 'caller',
      });
    },
  };
}

/**
 * Fonts embedded in the source document itself (docx fontTable), keyed by the
 * normalized family name — an exact-name match, never a substitution.
 */
export function embeddedDocFontProvider(embedded: ReadonlyMap<string, FontRegistry>): FontProvider {
  return {
    id: 'embedded',
    resolve: (req) => {
      const name = req.family?.trim().toLowerCase();
      const registry = name ? embedded.get(name) : undefined;
      if (!registry) return Promise.resolve(NO_FONT);
      const { parsed } = registry.resolveByStyle(req.bold, req.italic);
      return Promise.resolve({
        kind: 'bytes',
        bytes: parsed.raw,
        faceName: req.family!,
        providerId: 'embedded',
      });
    },
  };
}

/**
 * Open CDN substitutes (Arimo / Tinos / Cousine / Carlito / Caladea — the
 * LibreOffice metric-compatible mapping).
 * Always answers; the chain reports it as a substitution.
 */
export function remoteFontProvider(options: { readonly fetch?: FetchLike } = {}): FontProvider {
  const cache = new Map<FamilyKey, Promise<FontBytesByVariant>>();
  return {
    id: 'remote',
    resolve: async (req) => {
      const family = resolveFamilyKey(req.family ?? '');
      let set = cache.get(family);
      if (!set) {
        set = fetchFontSet({ family, ...(options.fetch ? { fetch: options.fetch } : {}) });
        cache.set(family, set);
      }
      const fonts = await set;
      // The shared cascade restores boldItalic→bold/italic degradation here —
      // fetchFontSet is best-effort, so a face can legitimately be missing.
      const picked = pickVariant((x) => fonts[x] !== undefined, req.bold, req.italic) ?? 'regular';
      const bytes = fonts[picked] ?? fonts.regular;
      return { kind: 'bytes', bytes, faceName: family, providerId: 'remote' };
    },
  };
}

// ---------------------------------------------------------------------------
// Local Font Access (Chromium): real system-font bytes, with the OS/2 fsType
// licensing gate. https://wicg.github.io/local-font-access/
// ---------------------------------------------------------------------------

interface LocalFontData {
  readonly family: string;
  readonly fullName: string;
  readonly style: string;
  blob: () => Promise<{ arrayBuffer: () => Promise<ArrayBuffer> }>;
}

type QueryLocalFonts = (options?: {
  readonly postscriptNames?: ReadonlyArray<string>;
}) => Promise<ReadonlyArray<LocalFontData>>;

/**
 * OS/2 fsType embedding permissions (OpenType §OS/2). Returns undefined when
 * the table is absent/corrupt. The licensing nibble: 0 = installable,
 * 2 = RESTRICTED (no embedding), 4 = preview & print, 8 = editable.
 */
export function readOs2FsType(bytes: Uint8Array): number | undefined {
  if (bytes.length < 12) return undefined;
  const u16 = (o: number): number => (bytes[o]! << 8) | bytes[o + 1]!;
  const u32 = (o: number): number => u16(o) * 0x10000 + u16(o + 2);
  const numTables = u16(4);
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16;
    if (rec + 16 > bytes.length) return undefined;
    const tag = String.fromCharCode(bytes[rec]!, bytes[rec + 1]!, bytes[rec + 2]!, bytes[rec + 3]!);
    if (tag !== 'OS/2') continue;
    const offset = u32(rec + 8);
    if (offset + 10 > bytes.length) return undefined;
    return u16(offset + 8);
  }
  return undefined;
}

/** True when the fsType licensing nibble forbids embedding entirely. */
export function isEmbeddingRestricted(fsType: number | undefined): boolean {
  if (fsType === undefined) return false; // absent table ⇒ installable default
  return (fsType & 0x000f) === 0x0002;
}

/**
 * System fonts via the Local Font Access API (Chromium 103+, permission
 * prompt). OPT-IN by design: never wired in implicitly. Returns 'none' when
 * the API is unavailable (Node, Safari, Firefox), the family is not installed,
 * or the face's OS/2 fsType marks embedding as restricted — embedding a
 * restricted font into a PDF would violate its license, so the chain falls
 * through to a substitute instead.
 */
export function localFontProvider(): FontProvider {
  const query = (globalThis as { queryLocalFonts?: QueryLocalFonts }).queryLocalFonts;
  return {
    id: 'local',
    resolve: async (req) => {
      if (!query || !req.family) return NO_FONT;
      let faces: ReadonlyArray<LocalFontData>;
      try {
        faces = await query();
      } catch {
        return NO_FONT; // permission denied / API error → substitute
      }
      const family = req.family.trim().toLowerCase();
      const wantBold = req.bold;
      const wantItalic = req.italic;
      const candidates = faces.filter((f) => f.family.trim().toLowerCase() === family);
      if (candidates.length === 0) return NO_FONT;
      const styleMatches = (f: LocalFontData): boolean => {
        const s = f.style.toLowerCase();
        return (
          s.includes('bold') === wantBold &&
          (s.includes('italic') || s.includes('oblique')) === wantItalic
        );
      };
      const face = candidates.find(styleMatches) ?? candidates[0]!;
      try {
        const blob = await face.blob();
        const bytes = new Uint8Array(await blob.arrayBuffer());
        if (isEmbeddingRestricted(readOs2FsType(bytes))) return NO_FONT;
        return { kind: 'bytes', bytes, faceName: face.fullName, providerId: 'local' };
      } catch {
        return NO_FONT;
      }
    },
  };
}
