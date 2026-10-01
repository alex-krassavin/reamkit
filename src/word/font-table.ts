// ECMA-376 Part 1 §17.8 — word/fontTable.xml + embedded (obfuscated) fonts.
//
// A document may embed its own font binaries (word/fonts/fontN.odttf) so it
// renders with the exact fonts the author used. Each is "obfuscated": the first
// 32 bytes are XOR'd with the 16-byte fontKey GUID (applied in reverse byte
// order, repeated twice). De-obfuscating restores a normal sfnt. Using these
// avoids substitution entirely → glyph-exact output.

import type { FontBytesByVariant } from '@/core/font';
import type { OpcPackage } from '@/core/opc';
import { FontRegistry } from '@/core/font';
import { isWordChar, tagMatches } from '@/core/opc/tag-scan';

const FONT_TABLE_PART = 'word/fontTable.xml';
const REL_FONT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';

type EmbedVariant = 'regular' | 'bold' | 'italic' | 'boldItalic';
interface EmbedRef {
  readonly rId: string;
  readonly fontKey: string;
}
interface FontTableEntry {
  readonly name: string;
  readonly embeds: Partial<Record<EmbedVariant, EmbedRef>>;
}

const VARIANT_BY_TAG: Record<string, EmbedVariant> = {
  Regular: 'regular',
  Bold: 'bold',
  Italic: 'italic',
  BoldItalic: 'boldItalic',
};

/**
 * §17.8.1 — restore an obfuscated embedded font by XOR-ing its first 32 bytes
 * with the `fontKey` GUID bytes in reverse order, recovering a normal sfnt.
 *
 * @param data    The obfuscated `.odttf` bytes.
 * @param fontKey The 16-byte `w:fontKey` GUID (with or without braces/dashes).
 * @returns The de-obfuscated bytes, or `data` unchanged when `fontKey` is not a GUID.
 */
export function deobfuscateEmbeddedFont(data: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, '');
  if (hex.length !== 32) return data; // not a GUID → assume already plain
  const key = new Uint8Array(16);
  for (let i = 0; i < 16; i++) key[i] = parseInt(hex.substring(i * 2, i * 2 + 2), 16);
  key.reverse();
  const out = new Uint8Array(data);
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) out[i] = out[i]! ^ key[i % 16]!;
  return out;
}

/**
 * §17.8.1 — obfuscate a font for embedding. The obfuscation is an XOR of the
 * first 32 bytes with the `fontKey` GUID's bytes in reverse order, so it is its
 * own inverse: this is {@link deobfuscateEmbeddedFont} run on a plain font.
 *
 * @param data    The plain sfnt bytes.
 * @param fontKey The GUID the font table states as its `w:fontKey`.
 * @returns The bytes to store as the `.odttf` part.
 */
export function obfuscateEmbeddedFont(data: Uint8Array, fontKey: string): Uint8Array {
  return deobfuscateEmbeddedFont(data, fontKey);
}

/**
 * Parse `word/fontTable.xml` for its embedded-font references (§17.8). Returns
 * one `FontTableEntry` per `w:font` that carries at least one `w:embed*` child,
 * each naming the relationship id + `w:fontKey` for a Regular/Bold/Italic/
 * BoldItalic face. A regex pass, not a full XML parse — only the embed refs matter.
 */
export function parseFontTable(data: Uint8Array): Array<FontTableEntry> {
  const xml = new TextDecoder('utf-8').decode(data);
  const out: Array<FontTableEntry> = [];
  // No font ends past the last `</w:font>`. Cut there, the lazy search for the
  // end of each one finds one; searched for in the whole part, a `w:font`
  // that never closed — a self-closing one, say — read on to the end of the
  // part from every `<w:font` before it.
  const closed = xml.slice(0, xml.lastIndexOf('</w:font>') + '</w:font>'.length);
  for (const fm of tagMatches(closed, '<w:font', true, FONT)) {
    const name = fm[1]!;
    const embeds: FontTableEntry['embeds'] = {};
    for (const em of embedRefs(fm[2]!)) {
      const variant = VARIANT_BY_TAG[em.tag];
      if (variant) embeds[variant] = { rId: em.rId, fontKey: em.fontKey };
    }
    if (Object.keys(embeds).length > 0) out.push({ name, embeds });
  }
  return out;
}

/** A named `w:font` and everything up to its close. Sticky, for tagMatches. */
const FONT = /<w:font\b[^>]*\bw:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/y;

const EMBED_HEAD = /<w:embed(Regular|Bold|Italic|BoldItalic)\b/y;

/** One `w:embed*` reference as found: which face, its relationship and its key. */
interface EmbedMatch extends EmbedRef {
  readonly tag: string;
}

/**
 * The faces a font's element embeds — what
 * `/<w:embed(Regular|Bold|Italic|BoldItalic)\b[^>]*\br:id="([^"]+)"[^>]*\bw:fontKey="([^"]+)"/g`
 * finds in it, in one pass. The expression gave the tag back one `r:id` at a
 * time and read the rest of it again for a `w:fontKey` after each, quadratic
 * in a tag with many; and searching, it read each tag again from every
 * `<w:embed` inside it. Where a start fails, every later one before the tag's
 * `>` does too, so the search moves past it.
 */
function embedRefs(inner: string): Array<EmbedMatch> {
  const out: Array<EmbedMatch> = [];
  let at = inner.indexOf('<w:embed');
  while (at >= 0) {
    EMBED_HEAD.lastIndex = at;
    const head = EMBED_HEAD.exec(inner);
    if (head === null) {
      at = inner.indexOf('<w:embed', at + 1);
      continue;
    }
    const from = EMBED_HEAD.lastIndex;
    const tagEnd = endOfStretch(inner, from);
    const found = embedRef(inner, from, tagEnd);
    if (found) {
      out.push({ tag: head[1]!, rId: found.rId, fontKey: found.fontKey });
      at = inner.indexOf('<w:embed', found.end);
    } else if (tagEnd < inner.length) {
      at = inner.indexOf('<w:embed', tagEnd + 1);
    } else {
      break;
    }
  }
  return out;
}

/** The first `>` at or after `from`, or the end of the text. */
function endOfStretch(s: string, from: number): number {
  const gt = s.indexOf('>', from);
  return gt < 0 ? s.length : gt;
}

/**
 * The reference one `<w:embed…` makes, its attributes read from `from` up to
 * the tag's `>` at `tagEnd`. The first `[^>]*` gives back the last `r:id`
 * first; the second takes the last `w:fontKey` before the `>` that follows the
 * `r:id`'s value — and that one is the same for every `r:id` whose value ends
 * in one stretch, so it is looked for once per stretch.
 */
function embedRef(
  s: string,
  from: number,
  tagEnd: number,
): { rId: string; fontKey: string; end: number } | undefined {
  const keys = new Map<number, { at: number; value: string; end: number } | undefined>();
  const lastKey = (stretchEnd: number): { at: number; value: string; end: number } | undefined => {
    if (keys.has(stretchEnd)) return keys.get(stretchEnd);
    const start = s.lastIndexOf('>', stretchEnd - 1) + 1;
    let key: { at: number; value: string; end: number } | undefined;
    for (
      let f = s.lastIndexOf('w:fontKey="', stretchEnd - 11);
      f >= start;
      f = s.lastIndexOf('w:fontKey="', f - 1)
    ) {
      const close = s.indexOf('"', f + 11);
      if (!isWordChar(s.charCodeAt(f - 1)) && close > f + 11) {
        key = { at: f, value: s.slice(f + 11, close), end: close + 1 };
        break;
      }
      if (f === 0) break;
    }
    keys.set(stretchEnd, key);
    return key;
  };
  for (let p = s.lastIndexOf('r:id="', tagEnd - 6); p >= from; p = s.lastIndexOf('r:id="', p - 1)) {
    if (isWordChar(s.charCodeAt(p - 1))) continue;
    const close = s.indexOf('"', p + 6);
    if (close <= p + 6) continue;
    // A value that ends inside the tag leaves the rest of the tag as its
    // stretch; only the last `r:id`'s value can run on past the tag's `>`.
    const key = lastKey(close < tagEnd ? tagEnd : endOfStretch(s, close + 1));
    if (key !== undefined && key.at > close) {
      return { rId: s.slice(p + 6, close), fontKey: key.value, end: key.end };
    }
  }
  return undefined;
}

/**
 * De-obfuscate and load every embedded font in the package, building one
 * {@link FontRegistry} per font keyed by its normalized (trimmed, lower-cased)
 * name so a run's `w:ascii` can match it. Each candidate face is validated
 * against the sfnt signature; a font lacking a usable Regular face, or whose
 * faces fail to parse, is skipped (the family then falls back to substitution).
 *
 * @param pkg The opened OPC package for the `.docx`.
 * @returns A map from normalized font name to its registry (empty when nothing embeds).
 */
export function loadEmbeddedFonts(pkg: OpcPackage): Map<string, FontRegistry> {
  const out = new Map<string, FontRegistry>();
  const ftData = pkg.getPart(FONT_TABLE_PART);
  if (!ftData) return out;
  const relById = new Map(pkg.getPartRelationships(FONT_TABLE_PART).map((r) => [r.id, r]));

  for (const entry of parseFontTable(ftData)) {
    const bytes: { -readonly [K in keyof FontBytesByVariant]?: FontBytesByVariant[K] } = {};
    for (const variant of ['regular', 'bold', 'italic', 'boldItalic'] as const) {
      const ref = entry.embeds[variant];
      if (!ref) continue;
      const rel = relById.get(ref.rId);
      if (!rel || rel.type !== REL_FONT) continue;
      const resolved = pkg.resolveRelatedPart(FONT_TABLE_PART, rel);
      if (!resolved) continue;
      try {
        const ttf = deobfuscateEmbeddedFont(resolved.data, ref.fontKey);
        // A real sfnt starts with 0x00010000 / 'OTTO' / 'true' / 'ttcf'.
        const sig = ((ttf[0]! << 24) | (ttf[1]! << 16) | (ttf[2]! << 8) | ttf[3]!) >>> 0;
        if (sig === 0x00010000 || sig === 0x4f54544f || sig === 0x74727565 || sig === 0x74746366) {
          bytes[variant] = ttf;
        }
      } catch {
        // skip an undecodable face; the family may still have others
      }
    }
    if (bytes.regular) {
      try {
        out.set(
          entry.name.trim().toLowerCase(),
          FontRegistry.fromBytes({ ...bytes, regular: bytes.regular }),
        );
      } catch {
        // FontRegistry/parse failure for this font → skip it (fall back to substitution)
      }
    }
  }
  return out;
}
