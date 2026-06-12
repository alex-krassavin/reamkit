// E-PDF EP2 — font resolution. Builds a ContentFont (the interpreter's decode +
// advance hooks) from a /Font dictionary: Unicode from the /ToUnicode CMap, and
// glyph advances from a simple font's /Widths or a composite font's /W array.

import { parseToUnicodeCMap } from './cmap';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { ContentFont } from './content';
import type { PdfFile } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

export function buildContentFont(file: PdfFile, fontDict: PdfDict): ContentFont {
  const isType0 = asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) === 'Type0';

  let toUnicode: ReadonlyMap<number, string> = new Map();
  let codeBytes: 1 | 2 = isType0 ? 2 : 1;
  const tu = file.resolve(fontDict.get('ToUnicode') ?? PDF_NULL);
  if (tu instanceof PdfStream) {
    const parsed = parseToUnicodeCMap(file.streamData(tu));
    toUnicode = parsed.map;
    codeBytes = parsed.codeBytes;
  }

  const width = isType0 ? cidWidths(file, fontDict) : simpleWidths(file, fontDict);
  const bytesPerCode = codeBytes;

  return {
    bytesPerCode,
    // Map each code to Unicode; an unmapped code in a simple font falls back to
    // its Latin-1 character, a composite font's to nothing (no sensible guess).
    decode: (codes) =>
      codes
        .map((c) => toUnicode.get(c) ?? (bytesPerCode === 1 ? String.fromCharCode(c) : ''))
        .join(''),
    width,
  };
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
