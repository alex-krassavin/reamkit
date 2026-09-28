// §9.6.2.1 — the family a PostScript face name is made of, which is the name a
// word processor finds a font by: a PDF names `Inter-SemiBold`, a .docx names
// `Inter` and says bold beside it.

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';
import { familyOfFace } from '@/pdf-reader/font';

describe('the family a PDF face belongs to', () => {
  it('drops the style a face name ends in', () => {
    expect(familyOfFace('Inter-SemiBold')).toBe('Inter');
    expect(familyOfFace('Inter-Medium')).toBe('Inter');
    expect(familyOfFace('Calibri,Bold')).toBe('Calibri');
    expect(familyOfFace('NimbusRomNo9L-ReguItal')).toBe('Nimbus Rom No9L');
  });

  it('drops the subset tag and the vendor’s suffix', () => {
    expect(familyOfFace('ABCDEF+ArialMT')).toBe('Arial');
    expect(familyOfFace('Arial-BoldMT')).toBe('Arial');
  });

  it('drops what a producer writes after the face', () => {
    // bug900822.pdf's .docx asked for "Courier New,Bold Win Char Set FFFF",
    // which no machine has, over the Courier New that every machine does.
    expect(familyOfFace('CourierNew,Bold-WinCharSetFFFF')).toBe('Courier New');
    expect(familyOfFace('ArialUnicodeMS-WinCharSetFFFF-H2')).toBe('Arial Unicode MS');
    expect(familyOfFace('Calibri,Bold-OneByteIdentityH')).toBe('Calibri');
    expect(familyOfFace('*Arial-68771-Identity-H')).toBe('Arial');
    // …and the slant or stretch an interpreter synthesised.
    expect(familyOfFace('NimbusRomNo9L-Regu-Slant_167')).toBe('Nimbus Rom No9L');
    expect(familyOfFace('LucidaTypewriter-Extend_850')).toBe('Lucida Typewriter');
    // A name in another encoding, read as Latin-1, is letters all the same.
    expect(familyOfFace('ËÎÌå')).toBe('ËÎÌå');
  });

  it('drops the CMap a composite font’s name ends in (§9.7.6.1)', () => {
    const content = 'BT /F1 12 Tf 20 100 Td <0001> Tj ET';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        '/Resources << /Font << /F1 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
      '<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3-UniJIS-UCS2-H ' +
        '/Encoding /UniJIS-UCS2-H /DescendantFonts [6 0 R] >>',
      '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 ' +
        '/CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> >>',
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
    const flow = Ream.parse(new TextEncoder().encode(pdf)).flow;
    expect([...(flow.faceFamilies?.values() ?? [])].map((f) => f.family)).toEqual([
      'Heisei Min W3',
    ]);
  });

  it('parts the words a family’s name runs together', () => {
    expect(familyOfFace('TimesNewRomanPS-BoldMT')).toBe('Times New Roman');
    expect(familyOfFace('SegoeUI')).toBe('Segoe UI');
    expect(familyOfFace('IBMPlexSans-Regular')).toBe('IBM Plex Sans');
    expect(familyOfFace('DejaVuSans-Bold')).toBe('DejaVu Sans');
  });

  it('takes off a style run on to the family with no separator', () => {
    expect(familyOfFace('SUBSET+CalibriBold')).toBe('Calibri');
    expect(familyOfFace('ArialBoldItalic')).toBe('Arial');
  });

  it('keeps a hyphen that is part of the family, not a style after it', () => {
    // `MS-Mincho` is a family of its own; "Mincho" is no weight.
    expect(familyOfFace('MS-Mincho')).toBe('MS Mincho');
  });
});
