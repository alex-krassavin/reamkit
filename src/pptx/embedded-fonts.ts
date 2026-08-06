// ECMA-376 §15.2.13 / §19.2.1.13 `p:embeddedFontLst` — the faces a deck brings
// with it, so it reads the same on a machine that has none of them.
//
// Word and PowerPoint embed a font in DIFFERENT wrappers, and the difference
// decides how far this can go. A .docx part is the font itself with its first
// 32 bytes obfuscated against the GUID in the part's name (`font-table.ts`
// undoes it in four lines). A .pptx `fntdata` part is an EOT container —
// Embedded OpenType, the format Internet Explorer shipped — and every one of
// the eleven parts in the corpus sets its TTEMBED_TTCOMPRESSED flag, which
// means the font inside is packed with MicroType Express. That is a codec of
// its own (a bespoke LZ over the glyph tables), not a header to skip.
//
// So this reads the container, uses the font when it is stored plainly, and
// records a loss when it is compressed — the deck then renders in a substitute,
// which is what it did before, but no longer silently.

import type { FontBytesByVariant } from '@/core/font';
import type { Loss } from '@/core/ir';
import type { OpcPackage } from '@/core/opc';
import type { PoNode } from '@/core/po-helpers';
import { FontRegistry } from '@/core/font';
import { poAttr, poChildren, poIs } from '@/core/po-helpers';

const REL_FONT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';

/** §19.2.1.13 — the four faces a `p:embeddedFont` may name, by element. */
const VARIANTS = [
  ['p:regular', 'regular'],
  ['p:bold', 'bold'],
  ['p:italic', 'italic'],
  ['p:boldItalic', 'boldItalic'],
] as const;

// The EOT header's fixed prefix: its own size, the font's, the version, and the
// flags (§2.1 of the EOT submission). 0x00020002 is the version every producer
// writes; TTEMBED_TTCOMPRESSED says the font data is MicroType Express.
const EOT_VERSIONS = new Set([0x00010000, 0x00020001, 0x00020002]);
const TTEMBED_TTCOMPRESSED = 0x00000004;
const TTEMBED_XORENCRYPTDATA = 0x10000000;

/** What an `fntdata` part turned out to hold. */
type FaceBytes =
  | { readonly kind: 'font'; readonly bytes: Uint8Array }
  | { readonly kind: 'compressed' }
  | { readonly kind: 'unreadable' };

/**
 * Read one `fntdata` part: the font itself when the container stores it
 * plainly, and what stopped it otherwise.
 *
 * @param data The part's bytes.
 */
export function readFntData(data: Uint8Array): FaceBytes {
  // A part may be the bare font — nothing in the spec forbids it, and a reader
  // that insists on the wrapper would drop one.
  if (isSfnt(data)) return { kind: 'font', bytes: data };
  if (data.length < 82) return { kind: 'unreadable' };
  const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const eotSize = v.getUint32(0, true);
  const version = v.getUint32(8, true);
  const flags = v.getUint32(12, true);
  if (eotSize !== data.length || !EOT_VERSIONS.has(version)) return { kind: 'unreadable' };
  if ((flags & (TTEMBED_TTCOMPRESSED | TTEMBED_XORENCRYPTDATA)) !== 0)
    return { kind: 'compressed' };
  // The font sits after the header's variable tail: four length-prefixed UTF-16
  // strings (family, style, version, full name), each followed by a padding
  // word, and then the root string of a v2 header.
  let at = 82; // past the fixed fields, at FamilyNameSize
  for (let i = 0; i < 4; i++) {
    if (at + 2 > data.length) return { kind: 'unreadable' };
    at += 2 + v.getUint16(at, true) + 2;
  }
  if (version !== 0x00010000) {
    if (at + 2 > data.length) return { kind: 'unreadable' };
    at += 2 + v.getUint16(at, true);
    // A v2 header adds a signature and two more optional strings; the font data
    // is what remains, and its size is stated up front.
    at = data.length - v.getUint32(4, true);
  }
  const bytes = data.subarray(at);
  return isSfnt(bytes) ? { kind: 'font', bytes } : { kind: 'unreadable' };
}

/**
 * §19.2.1.13 — the fonts a deck embeds, by normalized family name.
 *
 * @param pkg      The opened `.pptx` package.
 * @param presPath The presentation part's path (relationships resolve from it).
 * @param pres     Its parsed `p:presentation` element.
 * @param onLoss   Where a face that cannot be unpacked records itself.
 * @returns The families whose faces were readable (empty when none were).
 */
export function loadPptxEmbeddedFonts(
  pkg: OpcPackage,
  presPath: string,
  pres: PoNode | undefined,
  onLoss: (loss: Loss) => void,
): Map<string, FontRegistry> {
  const out = new Map<string, FontRegistry>();
  const list = pres ? poChildren(pres).find((c) => poIs(c, 'p:embeddedFontLst')) : undefined;
  if (!list) return out;
  const relById = new Map(pkg.getPartRelationships(presPath).map((r) => [r.id, r]));

  for (const font of poChildren(list).filter((c) => poIs(c, 'p:embeddedFont'))) {
    const bytes: { -readonly [K in keyof FontBytesByVariant]?: FontBytesByVariant[K] } = {};
    let compressed = false;
    for (const [tag, variant] of VARIANTS) {
      const node = poChildren(font).find((c) => poIs(c, tag));
      const rId = node ? poAttr(node, 'r:id') : undefined;
      const rel = rId !== undefined ? relById.get(rId) : undefined;
      if (!rel || rel.type !== REL_FONT) continue;
      const part = pkg.resolveRelatedPart(presPath, rel);
      if (!part) continue;
      const face = readFntData(part.data);
      if (face.kind === 'compressed') compressed = true;
      if (face.kind === 'font') bytes[variant] = face.bytes;
    }
    const typeface = poChildren(font).find((c) => poIs(c, 'p:font'));
    const name = (typeface ? poAttr(typeface, 'typeface') : undefined)?.trim() ?? '';
    if (bytes.regular) {
      try {
        out.set(name.toLowerCase(), FontRegistry.fromBytes({ ...bytes, regular: bytes.regular }));
        continue;
      } catch {
        // An unparseable face leaves the family to substitution, as before.
      }
    }
    if (compressed) {
      onLoss({
        feature: 'font',
        detail:
          `the deck embeds ${name || 'a font'}, packed with MicroType Express ` +
          `(EOT TTEMBED_TTCOMPRESSED); it renders in a substitute`,
        severity: 'substituted',
      });
    }
  }
  return out;
}

// The four signatures a real sfnt starts with.
function isSfnt(data: Uint8Array): boolean {
  if (data.length < 4) return false;
  const sig = ((data[0]! << 24) | (data[1]! << 16) | (data[2]! << 8) | data[3]!) >>> 0;
  return sig === 0x00010000 || sig === 0x4f54544f || sig === 0x74727565 || sig === 0x74746366;
}
