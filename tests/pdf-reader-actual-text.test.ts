// ISO 32000-1 §14.9.4 — replacement text. A marked-content sequence whose
// property list holds /ActualText says what its glyphs stand for, and what they
// decode to is not it: a ligature, a hyphen put in to break a word, a tab drawn
// as a space, the short i of हि shown before the ha it follows. The text stands
// where the glyphs stood, so lines, word gaps and the reading are built as for
// any run.

import { describe, expect, it } from 'vitest';

import type { ContentFont, TextRun } from '@/pdf-reader/content';
import { Ream } from '@/core/converter/ream';
import { interpretContent } from '@/pdf-reader/content';

const encode = (cs: string): Uint8Array => Uint8Array.from(cs, (c) => c.charCodeAt(0));

const run = (
  cs: string,
  fonts = new Map<string, ContentFont>(),
  named = new Map<string, string>(),
): Array<TextRun> =>
  interpretContent(encode(cs), fonts, undefined, undefined, undefined, undefined, undefined, named)
    .texts;

/** A one-byte font, each glyph a full em wide, its codes read as Latin-1. */
const WIDE: ContentFont = {
  bytesPerCode: 1,
  decode: (codes) => String.fromCharCode(...codes),
  width: () => 1000,
};

describe('a marked-content sequence with replacement text (§14.9.4)', () => {
  it('reads as its text, not as its glyphs', () => {
    const runs = run(
      'BT /F1 12 Tf 10 20 Td /Span <</ActualText <FEFF00660069>>> BDC (X) Tj EMC ET',
    );
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ text: 'fi', fontKey: 'F1' });
  });

  it('stands where its first glyph stood and reaches to where its last one ended', () => {
    // Several show operators and a change of face inside one sequence: the
    // text takes the first glyph's face and size, and the whole extent.
    const runs = run(
      'BT /F1 12 Tf 10 20 Td /Span <</ActualText <FEFF0939093F>>> BDC (a) Tj ' +
        '/F2 10 Tf [(b) -100 (c)] TJ EMC (d) Tj ET',
      new Map([['F2', WIDE]]),
    );
    expect(runs.map((r) => r.text)).toEqual(['हि', 'd']);
    expect(runs[0]).toMatchObject({ fontKey: 'F1' });
    expect(runs[0]!.fontSizePt).toBeCloseTo(12);
    expect(runs[0]!.x).toBeCloseTo(10);
    expect(runs[0]!.y).toBeCloseTo(20);
    // a is half an em of 12, b and c an em of 10 each, the nudge a tenth of one.
    expect(runs[0]!.endX).toBeCloseTo(37);
    expect(runs[1]!.x).toBeCloseTo(37);
  });

  it('breaks a word where the page breaks it, and reads it whole', () => {
    // The standard's own example: "Drucker" hyphenated as "Druk-ker", the k
    // and the hyphen standing for a c.
    const runs = run(
      'BT /F1 12 Tf 0 100 Td (Dru) Tj /Span <</ActualText (c)>> BDC (k-) Tj EMC ' +
        '0 -14 Td (ker) Tj ET',
    );
    expect(runs.map((r) => r.text)).toEqual(['Dru', 'c', 'ker']);
    expect(runs[1]!.x).toBeCloseTo(18);
    expect(runs[1]!.endX).toBeCloseTo(30);
  });

  it('says the glyphs stand for nothing when the text is empty', () => {
    const runs = run('BT /F1 12 Tf 0 0 Td (hy) Tj /Span <</ActualText ()>> BDC (-) Tj EMC ET');
    expect(runs.map((r) => r.text)).toEqual(['hy']);
  });

  it('covers a sequence inside it with its own text', () => {
    const runs = run(
      'BT /F1 12 Tf 0 0 Td /Span <</ActualText (abc)>> BDC (a) Tj ' +
        '/Span <</ActualText (X)>> BDC (b) Tj EMC (c) Tj EMC (d) Tj ET',
    );
    expect(runs.map((r) => r.text)).toEqual(['abc', 'd']);
  });

  it('closes on its own EMC, not on one of a sequence inside it', () => {
    const runs = run(
      'BT /F1 12 Tf 0 0 Td /Span <</ActualText (fi)>> BDC /Lang BMC (f) Tj EMC (i) Tj EMC ' +
        '(n) Tj ET',
    );
    expect(runs.map((r) => r.text)).toEqual(['fi', 'n']);
  });

  it('reads UTF-16 behind its mark, and PDFDocEncoding otherwise (§7.9.2.2)', () => {
    const utf16 = run(
      'BT /F1 12 Tf 0 0 Td /Span <</ActualText (\\376\\377\\000\\351)>> BDC (e) Tj EMC ET',
    );
    expect(utf16[0]!.text).toBe('é');
    // 0x93 is the fi ligature in PDFDocEncoding, and a C1 control in Latin-1.
    const doc = run('BT /F1 12 Tf 0 0 Td /Span <</ActualText (\\223)>> BDC (f) Tj EMC ET');
    expect(doc[0]!.text).toBe('\ufb01');
  });

  it('takes the text a property list named in the resources holds', () => {
    const runs = run(
      'BT /F1 12 Tf 0 0 Td /Span /AT1 BDC (X) Tj EMC ET',
      undefined,
      new Map([['AT1', 'fl']]),
    );
    expect(runs.map((r) => r.text)).toEqual(['fl']);
  });

  it('gives no text where it shows no glyph', () => {
    // Word brackets a tab it drew nothing for, and a figure may carry text of
    // its own: neither is a word on the line.
    expect(run('BT /F1 12 Tf 0 0 Td /Span <</ActualText (\\t)>> BDC EMC (a) Tj ET')).toHaveLength(
      1,
    );
  });

  it('reads a tab or a line end in it as a space, leaving the distance to the page', () => {
    // Word draws a tab as a space glyph named `\t`. The next word stands where
    // the page set it; taken for a tab it went to the next default stop.
    const runs = run(
      'BT /F1 12 Tf 0 0 Td (OK:) Tj /Span <</ActualText <FEFF0009>>> BDC ( ) Tj EMC ' +
        '/Span <</ActualText (a\\nb)>> BDC (ab) Tj EMC ET',
    );
    expect(runs.map((r) => r.text)).toEqual(['OK:', ' ', 'a b']);
  });

  it('writes a glyph that has no character, rather than drawing it', () => {
    // §9.6.6 — a code the face cannot name is drawn as its outline. The
    // replacement text names it, and drawn as well it would stand twice.
    const square = [
      { op: 'move' as const, x: 0, y: 0 },
      { op: 'line' as const, x: 500, y: 0 },
      { op: 'line' as const, x: 500, y: 500 },
      { op: 'close' as const },
    ];
    const blank: ContentFont = {
      bytesPerCode: 1,
      decode: () => '',
      width: () => 500,
      outline: { matrix: [0.001, 0, 0, 0.001, 0, 0], path: () => square },
    };
    const fonts = new Map([['F1', blank]]);
    const replaced = interpretContent(
      encode('BT /F1 12 Tf 0 0 Td /Span <</ActualText (fi)>> BDC (\\001) Tj EMC ET'),
      fonts,
    );
    expect(replaced.texts.map((r) => r.text)).toEqual(['fi']);
    expect(replaced.outlines).toHaveLength(0);
    // With no text to replace it, it is drawn as before.
    expect(
      interpretContent(encode('BT /F1 12 Tf 0 0 Td (\\001) Tj ET'), fonts).outlines,
    ).toHaveLength(1);
  });
});

describe('a sequence nested in a structure content sequence (§14.6)', () => {
  it('belongs to the structure element the one around it does', () => {
    // Word and Ream both set a /Span with replacement text inside a paragraph's
    // `/P <</MCID n>> BDC`; untagged, its words dropped out of a tagged reading.
    const runs = run(
      '/P <</MCID 3>> BDC BT /F1 12 Tf 0 0 Td /Span <</ActualText (fi)>> BDC (X) Tj EMC ' +
        '/Span <</Lang (en-US)>> BDC (ne) Tj EMC (!) Tj ET EMC',
    );
    expect(runs.map((r) => [r.text, r.mcid])).toEqual([
      ['fi', 3],
      ['ne', 3],
      ['!', 3],
    ]);
  });

  it('leaves an artifact inside it an artifact', () => {
    const runs = run(
      '/P <</MCID 3>> BDC BT /F1 12 Tf 0 0 Td /Artifact BMC (1) Tj EMC ' +
        '/Artifact <</Type /Pagination>> BDC (2) Tj EMC ET EMC',
    );
    expect(runs.map((r) => r.mcid)).toEqual([undefined, undefined]);
  });
});

/** A one-page file of `content` in Helvetica, with `properties` as its /Properties. */
function pagePdf(content: string, properties: string): Uint8Array {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] /Contents 4 0 R ' +
      `/Resources << /Font << /F0 5 0 R >> /Properties ${properties} >> >>`,
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /ActualText <FEFF00660069> >>',
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return encode(pdf);
}

describe('replacement text on a page', () => {
  it('is read from the page’s /Properties, directly or by reference', async () => {
    const content =
      'BT /F0 12 Tf 20 40 Td (A ) Tj /Span /Lig BDC (X) Tj EMC (eld ) Tj /Span /Ref BDC (Y) Tj EMC (at) Tj ET';
    const pdf = pagePdf(content, '<< /Lig << /ActualText (fi) >> /Ref 6 0 R >>');
    const md = new TextDecoder().decode(await Ream.parse(pdf).convert('md'));
    expect(md.trim()).toBe('A field fiat');
  });
});
