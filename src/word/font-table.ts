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

// §17.8.1 — restore an obfuscated embedded font by XOR-ing the first 32 bytes
// with the fontKey GUID bytes in reverse order.
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

export function parseFontTable(data: Uint8Array): Array<FontTableEntry> {
  const xml = new TextDecoder('utf-8').decode(data);
  const out: Array<FontTableEntry> = [];
  const fontRe = /<w:font\b[^>]*\bw:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/g;
  let fm: RegExpExecArray | null;
  while ((fm = fontRe.exec(xml)) !== null) {
    const name = fm[1]!;
    const inner = fm[2]!;
    const embeds: FontTableEntry['embeds'] = {};
    const embedRe =
      /<w:embed(Regular|Bold|Italic|BoldItalic)\b[^>]*\br:id="([^"]+)"[^>]*\bw:fontKey="([^"]+)"/g;
    let em: RegExpExecArray | null;
    while ((em = embedRe.exec(inner)) !== null) {
      const variant = VARIANT_BY_TAG[em[1]!];
      if (variant) embeds[variant] = { rId: em[2]!, fontKey: em[3]! };
    }
    if (Object.keys(embeds).length > 0) out.push({ name, embeds });
  }
  return out;
}

// Build a registry per embedded font, keyed by the normalized font name so a
// run's w:ascii can match it. A font lacking a usable Regular face is skipped.
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
