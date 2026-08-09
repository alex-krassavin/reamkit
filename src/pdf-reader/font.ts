// E-PDF EP2 — font resolution. Builds a ContentFont (the interpreter's decode +
// advance hooks) from a /Font dictionary: Unicode from the /ToUnicode CMap, and
// glyph advances from a simple font's /Widths or a composite font's /W array.

import { parseToUnicodeCMap } from './cmap';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { ContentFont } from './content';
import type { PdfFile } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

/**
 * Build a {@link ContentFont} (the interpreter's decode + advance hooks) from a
 * `/Font` dictionary (E-PDF EP2). Unicode comes from the `/ToUnicode` CMap;
 * glyph advances from a simple font's `/Widths` (§9.6.2.1) or a composite
 * `/Type0` font's descendant `/W` array (§9.7.4.3). A code with no `/ToUnicode`
 * entry decodes to its Latin-1 character for a simple font, or to nothing for a
 * composite one.
 *
 * @param file     The owning {@link PdfFile}, used to resolve indirect references.
 * @param fontDict The `/Font` dictionary.
 * @returns The decode/advance hooks plus the code width (1 or 2 bytes per code).
 */
export function buildContentFont(file: PdfFile, fontDict: PdfDict): ContentFont {
  const isType0 = asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) === 'Type0';

  let toUnicode: ReadonlyMap<number, string> = new Map();
  let codeBytes: 1 | 2 = isType0 ? 2 : 1;
  const tu = file.resolve(fontDict.get('ToUnicode') ?? PDF_NULL);
  if (tu instanceof PdfStream) {
    const parsed = parseToUnicodeCMap(file.streamData(tu));
    toUnicode = parsed.map;
    // §9.6: a SIMPLE font's glyphs are always selected by one-byte codes —
    // only a composite Type0 font takes its code width from a CMap. The
    // `codespacerange` of a /ToUnicode belongs to that CMap's own convention,
    // and Distiller writes `<0000> <FFFF>` there whatever the font is. Read as
    // two, an Arial subset's every string came apart into pairs of bytes:
    // 160F-2019.pdf's "rémunérations brutes" arrived as "isr".
    if (isType0) codeBytes = parsed.codeBytes;
  }

  const width = isType0 ? cidWidths(file, fontDict) : simpleWidths(file, fontDict);
  const bytesPerCode = codeBytes;
  const style = faceStyle(file, fontDict, isType0);

  return {
    bytesPerCode,
    // Map each code to Unicode; an unmapped code in a simple font falls back to
    // its Latin-1 character, a composite font's to nothing (no sensible guess).
    decode: (codes) =>
      codes
        .map((c) => toUnicode.get(c) ?? (bytesPerCode === 1 ? String.fromCharCode(c) : ''))
        .join(''),
    width,
    ...style,
  };
}

/** §9.8.2 `/Flags` — bit 7 is Italic, bit 19 ForceBold (bits numbered from 1). */
const FLAG_ITALIC = 1 << 6;
const FLAG_FORCE_BOLD = 1 << 18;

/** §9.8.1 `/FontWeight` — 400 is normal, 700 bold; 600 is where "bold" begins. */
const BOLD_WEIGHT = 600;

/**
 * §9.8.1 — whether the face a run is shown in is bold or slanted.
 *
 * A descriptor is the witness where there is one, and the ONLY witness: it
 * states `/FontWeight`, `/ItalicAngle` and the `/Flags` bits, so one that gives
 * neither a weight nor the ForceBold bit is saying the face is not bold.
 * ArabicCIDTrueType.pdf shows why that matters — two of its four faces are
 * called `NewBasrahBold` and `DamascusBold`, which is the family's own name and
 * not a weight, and reading the name over the descriptor set two lines heavy
 * that no reader sets heavy.
 *
 * The name is read only where no descriptor exists at all, which is the
 * standard-14 case (§9.6.2.2): `Helvetica-BoldOblique` has nothing else to go
 * on. The subset prefix (`ISVAYD+`) is dropped first — six arbitrary capitals
 * may spell anything.
 */
function faceStyle(
  file: PdfFile,
  fontDict: PdfDict,
  isType0: boolean,
): { bold?: boolean; italic?: boolean } {
  const owner = isType0 ? descendantFont(file, fontDict) : fontDict;
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
  if (descriptor instanceof Map) {
    const flags = asNumber(file.resolve(descriptor.get('Flags') ?? PDF_NULL), 0);
    const weight = asNumber(file.resolve(descriptor.get('FontWeight') ?? PDF_NULL), 0);
    const slant = asNumber(file.resolve(descriptor.get('ItalicAngle') ?? PDF_NULL), 0);
    const bold = weight >= BOLD_WEIGHT || (flags & FLAG_FORCE_BOLD) !== 0;
    const italic = slant !== 0 || (flags & FLAG_ITALIC) !== 0;
    return { ...(bold ? { bold } : {}), ...(italic ? { italic } : {}) };
  }
  const name = asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)).replace(
    /^[A-Z]{6}\+/u,
    '',
  );
  const bold = /bold|black|heavy/iu.test(name);
  const italic = /italic|oblique/iu.test(name);
  return { ...(bold ? { bold } : {}), ...(italic ? { italic } : {}) };
}

/** §9.7.4 — a `/Type0` font's one descendant CIDFont, which owns the descriptor. */
function descendantFont(file: PdfFile, fontDict: PdfDict): PdfDict {
  const descFonts = file.resolve(fontDict.get('DescendantFonts') ?? PDF_NULL);
  const first = Array.isArray(descFonts) ? file.resolve(descFonts[0] ?? PDF_NULL) : PDF_NULL;
  return first instanceof Map ? first : new Map<string, PdfValue>();
}

// §9.6.2.1 — a simple font's /Widths array is indexed by (code − /FirstChar).
function simpleWidths(file: PdfFile, fontDict: PdfDict): (code: number) => number {
  const first = asNumber(file.resolve(fontDict.get('FirstChar') ?? PDF_NULL), 0);
  const widthsVal = file.resolve(fontDict.get('Widths') ?? PDF_NULL);
  const widths = Array.isArray(widthsVal) ? widthsVal : [];
  const descriptor = file.resolve(fontDict.get('FontDescriptor') ?? PDF_NULL);
  const missing =
    descriptor instanceof Map
      ? asNumber(file.resolve(descriptor.get('MissingWidth') ?? PDF_NULL), 0)
      : 0;
  return (code) => {
    const w = widths[code - first];
    return typeof w === 'number' ? w : missing > 0 ? missing : 500;
  };
}

// §9.7.4.3 — a composite font's widths live on its descendant CIDFont as /DW
// (default) plus a /W array. With Identity encoding the CID equals the code.
function cidWidths(file: PdfFile, fontDict: PdfDict): (cid: number) => number {
  const descFonts = file.resolve(fontDict.get('DescendantFonts') ?? PDF_NULL);
  const desc0 = Array.isArray(descFonts) ? file.resolve(descFonts[0] ?? PDF_NULL) : PDF_NULL;
  const cidFont = desc0 instanceof Map ? desc0 : new Map<string, PdfValue>();
  const dw = asNumber(file.resolve(cidFont.get('DW') ?? PDF_NULL), 1000);
  const wMap = parseCidW(file, file.resolve(cidFont.get('W') ?? PDF_NULL));
  return (cid) => wMap.get(cid) ?? (dw || 1000);
}

// The /W array is a sequence of `c [w0 w1 …]` (per-CID widths from c) or
// `cFirst cLast w` (one width across a CID range).
function parseCidW(file: PdfFile, wVal: PdfValue): Map<number, number> {
  const out = new Map<number, number>();
  if (!Array.isArray(wVal)) return out;
  let i = 0;
  while (i < wVal.length) {
    const c = file.resolve(wVal[i++]!);
    if (typeof c !== 'number') break;
    const next = file.resolve(wVal[i] ?? PDF_NULL);
    if (Array.isArray(next)) {
      i++;
      next.forEach((w, k) => {
        if (typeof w === 'number') out.set(c + k, w);
      });
    } else if (typeof next === 'number') {
      i++;
      const w = file.resolve(wVal[i++] ?? PDF_NULL);
      if (typeof w === 'number') {
        for (let cc = c; cc <= next && cc - c < 65_536; cc++) out.set(cc, w);
      }
    } else {
      break;
    }
  }
  return out;
}

function asName(v: PdfValue): string {
  return v instanceof PdfName ? v.value : '';
}

function asNumber(v: PdfValue, dflt: number): number {
  return typeof v === 'number' ? v : dflt;
}
