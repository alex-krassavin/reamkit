// ECMA-376 Part 1 §17.8.1 — the faces a document is set in, EMBEDDED in it.
//
// A family the reader's machine lacks is substituted, and a substitute is a
// fraction wider or narrower than the face the document was set in: its lines
// break elsewhere and a page runs onto the next. Where the source carried its
// faces (a PDF's programs, see `FlowDoc.faceOutlines`), each family the runs
// name is written into the package as fonts of its own — a Regular, Bold,
// Italic and BoldItalic slot per family, each an obfuscated TrueType file
// (`word/fonts/fontN.odttf`) the font table points to.
//
// The licence a face states decides whether it travels (OS/2 `fsType`, see
// `editableEmbedding`): a restricted face, or one allowed only in documents
// opened read-only, stays out, and a reader substitutes for it as before.

import type { FaceFamily, FaceGlyph, FaceOutlines } from '@/core/ir/flow';
import type { BuiltGlyph, BuiltKernPair } from '@/core/font';
import type { Loss } from '@/core/ir';
import type { OpcPart, Relationship } from '@/core/opc';
import { MAX_BUILT_GLYPHS, buildTrueType, editableEmbedding } from '@/core/font';
import { FEATURES } from '@/core/ir';
import { md5 } from '@/core/crypto/primitives';
import { obfuscateEmbeddedFont } from '@/word/font-table';

/** §17.8.1 — the content type of an obfuscated embedded font. */
export const OBFUSCATED_FONT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.obfuscatedFont';

const REL_FONT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';

/** The faces written into a package, and how the font table reaches them. */
export interface EmbeddedFaces {
  /** Family → the `w:embed*` elements its `w:font` carries (§17.8.3.3–6). */
  readonly elements: ReadonlyMap<string, string>;
  /** The obfuscated font parts, under `word/fonts/`. */
  readonly parts: ReadonlyArray<OpcPart>;
  /** The font table's relationships to them. */
  readonly relationships: ReadonlyArray<Relationship>;
}

/** §17.8.3 — the four faces a family may embed, in the order CT_Font lists them. */
const SLOTS = ['Regular', 'Bold', 'Italic', 'BoldItalic'] as const;
type Slot = (typeof SLOTS)[number];

/**
 * Build the embedded fonts for the faces the runs name (§17.8.1).
 *
 * The faces are gathered by family and by slot — the style each face IS. Two
 * faces in one slot (a SemiBold and a Bold, both written bold) become one font:
 * the face with the most characters, and the other's characters it lacks.
 *
 * A face its licence keeps out is reported, not written.
 *
 * @param used     The run font names the document's runs carry.
 * @param outlines Run font name → the face's outlines.
 * @param families Run font name → the family a word processor knows it by.
 * @param losses   Appended to for every face left out.
 * @returns The fonts to write and the font table's references to them.
 */
export function embedFaces(
  used: Iterable<string>,
  outlines: ReadonlyMap<string, FaceOutlines>,
  families: ReadonlyMap<string, FaceFamily>,
  losses: Array<Loss>,
): EmbeddedFaces {
  const bySlot = new Map<string, Map<Slot, Array<FaceOutlines>>>();
  for (const name of new Set(used)) {
    const face = outlines.get(name);
    const family = families.get(name)?.family;
    if (!face || family === undefined) continue;
    if (!editableEmbedding(face.fsType)) {
      losses.push({
        severity: 'degraded',
        feature: FEATURES.fontsEmbedding,
        detail: `font ${face.postScriptName} not embedded: its licence (OS/2 fsType ${face.fsType}) does not allow it in an editable document`,
      });
      continue;
    }
    const slot: Slot = face.bold
      ? face.italic
        ? 'BoldItalic'
        : 'Bold'
      : face.italic
        ? 'Italic'
        : 'Regular';
    let slots = bySlot.get(family);
    if (!slots) bySlot.set(family, (slots = new Map()));
    slots.set(slot, [...(slots.get(slot) ?? []), face]);
  }

  const elements = new Map<string, string>();
  const parts: Array<OpcPart> = [];
  const relationships: Array<Relationship> = [];
  for (const family of [...bySlot.keys()].sort()) {
    const slots = bySlot.get(family)!;
    for (const slot of SLOTS) {
      const faces = [...(slots.get(slot) ?? [])].sort((a, b) => b.glyphs.size - a.glyphs.size);
      const first = faces[0];
      if (!first) continue;
      const glyphs = new Map(first.glyphs);
      const kerning = new Map(first.kerning);
      const ligatures = new Map(first.ligatures);
      for (const other of faces.slice(1)) {
        for (const [char, glyph] of other.glyphs) if (!glyphs.has(char)) glyphs.set(char, glyph);
        for (const [pair, value] of other.kerning ?? []) {
          if (!kerning.has(pair)) kerning.set(pair, value);
        }
        for (const [letters, glyph] of other.ligatures ?? []) {
          if (!ligatures.has(letters)) ligatures.set(letters, glyph);
        }
      }
      const built = builtGlyphs(glyphs);
      if (built.length + ligatures.size > MAX_BUILT_GLYPHS) {
        losses.push({
          severity: 'degraded',
          feature: FEATURES.fontsEmbedding,
          detail: `font ${first.postScriptName} not embedded: ${built.length + ligatures.size} glyphs is more than an embedded subset holds`,
        });
        continue;
      }
      const program = buildTrueType({
        family,
        bold: slot === 'Bold' || slot === 'BoldItalic',
        italic: slot === 'Italic' || slot === 'BoldItalic',
        postScriptName: first.postScriptName,
        glyphs: built,
        ascent: first.ascent,
        descent: first.descent,
        ...(first.capHeight !== undefined ? { capHeight: first.capHeight } : {}),
        ...(first.xHeight !== undefined ? { xHeight: first.xHeight } : {}),
        italicAngle: first.italicAngle,
        fixedPitch: first.fixedPitch,
        fsType: statedFsType(first.fsType),
        kerning: kernPairs(kerning),
        ligatures: [...ligatures].map(([letters, glyph]) => ({
          codePoints: [...letters].map((c) => c.codePointAt(0)!),
          outline: glyph.outline,
          advance: glyph.advance,
        })),
      });
      const fontKey = fontKeyOf(program);
      const index = parts.length + 1;
      const target = `fonts/font${index}.odttf`;
      const id = `rId${index}`;
      parts.push({
        path: `word/${target}`,
        data: obfuscateEmbeddedFont(program, fontKey),
        contentType: OBFUSCATED_FONT_CONTENT_TYPE,
      });
      relationships.push({ id, type: REL_FONT, target, targetMode: 'Internal' });
      elements.set(
        family,
        `${elements.get(family) ?? ''}<w:embed${slot} r:id="${id}" w:fontKey="${fontKey}" w:subsetted="1"/>`,
      );
    }
  }
  return { elements, parts, relationships };
}

/**
 * A face's characters as the glyphs of a font: characters that share one glyph
 * (the space and the no-break space) select the same one, and the glyphs stand
 * in the order of their first character.
 */
function builtGlyphs(glyphs: ReadonlyMap<string, FaceGlyph>): Array<BuiltGlyph> {
  const byGlyph = new Map<FaceGlyph, Array<number>>();
  for (const [char, glyph] of glyphs) {
    const cp = char.codePointAt(0);
    if (cp === undefined) continue;
    byGlyph.set(glyph, [...(byGlyph.get(glyph) ?? []), cp]);
  }
  return [...byGlyph]
    .map(([glyph, codePoints]) => ({
      codePoints: codePoints.sort((a, b) => a - b),
      outline: glyph.outline,
      advance: glyph.advance,
    }))
    .sort((a, b) => a.codePoints[0]! - b.codePoints[0]!);
}

/** The face's kerning pairs as the font states them: by code point, left then right. */
function kernPairs(kerning: ReadonlyMap<string, number>): Array<BuiltKernPair> {
  const out: Array<BuiltKernPair> = [];
  for (const [pair, value] of kerning) {
    const [left, right] = [...pair].map((c) => c.codePointAt(0));
    if (left !== undefined && right !== undefined) out.push({ left, right, value });
  }
  return out;
}

/**
 * The `fsType` the built font states: the program's own, less the bits the
 * format reserves; Editable where the program stated none — which is the use
 * the font is put to here, and no more than that.
 */
function statedFsType(fsType: number | undefined): number {
  if (fsType === undefined) return EDITABLE_EMBEDDING;
  return fsType & DEFINED_FS_TYPE_BITS;
}

/** OS/2 `fsType` — Editable embedding, and every bit the format defines. */
const EDITABLE_EMBEDDING = 0x0008;
const DEFINED_FS_TYPE_BITS = 0x030e;

/**
 * §17.8.1 — the GUID a font is obfuscated with, drawn from the font's own bytes
 * so the same document always writes the same package.
 */
function fontKeyOf(program: Uint8Array): string {
  const hex = [...md5(program)].map((b) => b.toString(16).padStart(2, '0').toUpperCase());
  const group = (from: number, to: number): string => hex.slice(from, to).join('');
  return `{${group(0, 4)}-${group(4, 6)}-${group(6, 8)}-${group(8, 10)}-${group(10, 16)}}`;
}
