// §9.9 — the font programs a PDF carries, lifted so the rebuilt document can be
// set in the SAME faces the page was.
//
// A reconstructed page has always been re-typeset in whatever face the host
// supplies, and a substituted face is never the one the author used: its glyphs
// are a fraction wider or narrower, so every word drifts from where the page put
// it and the drift accumulates along the line. On Brotli-Prototype-FileA.pdf
// that is most of what still separates our render from the file's — the geometry
// agrees and the words do not sit on it.
//
// A PDF embeds its faces (§9.9 `/FontFile2` for TrueType), and the layout
// already prefers a document's own font over any substitution
// (`FlowDoc.embeddedFonts`). This is the missing half: read them out.

import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { PdfFile, PdfPage } from './document';
import type { Loss } from '@/core/ir';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { FontRegistry, parseTtf } from '@/core/font';
import { FEATURES } from '@/core/ir';

const MAX_FORM_DEPTH = 8;

/**
 * Every embedded TrueType program the pages use, by the name a run will ask for
 * — the `/BaseFont` lowercased, with the subset prefix dropped.
 *
 * Each PDF font object becomes its OWN entry rather than being folded into a
 * family: `Arial-BoldMT` is a different program from `ArialMT` and the run that
 * uses it names it exactly, so nothing has to guess at weights or match faces
 * up. A face the reader cannot parse is left out and the writer substitutes for
 * it as before.
 *
 * Only `/FontFile2` is read. A CFF program (`/FontFile3`) and a Type 1 one
 * (`/FontFile`) are different formats that the layout's parser does not take;
 * those keep their substitute.
 *
 * And only one whose OUTLINES this pipeline can carry: a program with no
 * `glyf`/`loca` cannot be subset, and a face offered here is one the writer
 * will be asked to embed. bug1186827.pdf ships an OpenType/CFF program under
 * `/FontFile2`, which parses like any other and then killed the whole
 * conversion — "Subsetting requires a TrueType font with glyf+loca tables" —
 * rather than losing one face.
 *
 * @param file   The owning file.
 * @param pages  The pages whose fonts are wanted.
 * @param losses Appended to for a face the page embeds and this cannot carry.
 * @returns Name → a one-face registry holding that program.
 */
export function collectEmbeddedFonts(
  file: PdfFile,
  pages: ReadonlyArray<PdfPage>,
  losses?: Array<Loss>,
): Map<string, FontRegistry> {
  const out = new Map<string, FontRegistry>();
  const seen = new Set<PdfDict>();
  const visiting = new Set<PdfStream>();

  const addFont = (fontDict: PdfDict): void => {
    if (seen.has(fontDict)) return;
    seen.add(fontDict);
    const name = embeddedFontName(file, fontDict);
    if (name === undefined || out.has(name)) return;
    const program = fontProgram(file, fontDict);
    if (!program) return;
    try {
      const parsed = parseTtf(program);
      if (!parsed.tables.has('glyf') || !parsed.tables.has('loca')) {
        losses?.push({
          severity: 'degraded',
          feature: FEATURES.text,
          detail: `embedded font ${name} has no TrueType outlines (a CFF program under /FontFile2); its text is re-set in a substitute face`,
        });
        return;
      }
      out.set(name, FontRegistry.fromBytes({ regular: program }));
    } catch {
      // A program the parser will not take is a program the page keeps to
      // itself; the writer substitutes, as it did before any of this.
    }
  };

  const walk = (resources: PdfDict | undefined, depth: number): void => {
    if (!resources) return;
    const fonts = file.get(resources, 'Font');
    if (fonts instanceof Map) {
      for (const value of fonts.values()) {
        const dict = file.resolve(value);
        if (dict instanceof Map) addFont(dict);
      }
    }
    if (depth >= MAX_FORM_DEPTH) return;
    // §8.8 — a form draws with resources of its own, and a drawing keeps most
    // of its lettering inside them.
    const xobjects = file.get(resources, 'XObject');
    if (!(xobjects instanceof Map)) return;
    for (const value of xobjects.values()) {
      const stream = file.resolve(value);
      if (!(stream instanceof PdfStream) || visiting.has(stream)) continue;
      const subtype = file.get(stream.dict, 'Subtype');
      if (!(subtype instanceof PdfName) || subtype.value !== 'Form') continue;
      visiting.add(stream);
      const own = file.get(stream.dict, 'Resources');
      walk(own instanceof Map ? own : resources, depth + 1);
      visiting.delete(stream);
    }
  };

  for (const page of pages) walk(page.resources, 0);
  return out;
}

/**
 * The name a run set in `fontDict` will ask for: its `/BaseFont` without the
 * six-capital subset prefix (§9.6.4), lowercased.
 *
 * @param file     The owning file.
 * @param fontDict A `/Font` dictionary.
 * @returns The name, or `undefined` when the font states none.
 */
export function embeddedFontName(file: PdfFile, fontDict: PdfDict): string | undefined {
  const base = file.resolve(fontDict.get('BaseFont') ?? PDF_NULL);
  if (!(base instanceof PdfName)) return undefined;
  const name = base.value.replace(/^[A-Z]{6}\+/u, '').trim();
  return name.length > 0 ? name.toLowerCase() : undefined;
}

/**
 * Whether this font carries a program this pipeline can set the page in.
 *
 * Only `/FontFile2` is liftable, and a run in such a face is filed under the
 * face's own name — so that name must be left exactly as it is. Every other
 * face is SUBSTITUTED, and its name is free to say what it is.
 *
 * @param file     The document.
 * @param fontDict The font dictionary.
 */
export function hasLiftableProgram(file: PdfFile, fontDict: PdfDict): boolean {
  return fontProgram(file, fontDict) !== undefined;
}

/** §9.9 `/FontFile2` — the TrueType program, off the font or its descendant. */
function fontProgram(file: PdfFile, fontDict: PdfDict): Uint8Array | undefined {
  const owner = descendant(file, fontDict) ?? fontDict;
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return undefined;
  const program = file.resolve(descriptor.get('FontFile2') ?? PDF_NULL);
  if (!(program instanceof PdfStream)) return undefined;
  const bytes = file.streamData(program);
  return bytes.length > 0 ? bytes : undefined;
}

/** §9.7.4 — a `/Type0` font's descendant CIDFont, which owns the descriptor. */
function descendant(file: PdfFile, fontDict: PdfDict): PdfDict | undefined {
  const list = file.resolve(fontDict.get('DescendantFonts') ?? PDF_NULL);
  if (!Array.isArray(list)) return undefined;
  const first: PdfValue = file.resolve(list[0] ?? PDF_NULL);
  return first instanceof Map ? first : undefined;
}
