// §9.6.2.1 — the family a PostScript face name is made of, which is the name a
// word processor finds a font by: a PDF names `Inter-SemiBold`, a .docx names
// `Inter` and says bold beside it.

import { describe, expect, it } from 'vitest';

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

  it('parts the words a family’s name runs together', () => {
    expect(familyOfFace('TimesNewRomanPS-BoldMT')).toBe('Times New Roman');
    expect(familyOfFace('SegoeUI')).toBe('Segoe UI');
    expect(familyOfFace('IBMPlexSans-Regular')).toBe('IBM Plex Sans');
    expect(familyOfFace('DejaVuSans-Bold')).toBe('DejaVu Sans');
  });

  it('keeps a hyphen that is part of the family, not a style after it', () => {
    // `MS-Mincho` is a family of its own; "Mincho" is no weight.
    expect(familyOfFace('MS-Mincho')).toBe('MS Mincho');
  });
});
