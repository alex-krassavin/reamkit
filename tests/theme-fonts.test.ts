// E-FONT F1 — §17.3.2.26 `w:rFonts`: the typeface a document points at in its
// theme instead of spelling out. 828 of the corpus's 1121 .docx name a font
// that way and 414 name it no other way, their theme reading Calibri; left
// unresolved every one of them came out in the default sans.

import { describe, expect, it } from 'vitest';

import { zipSync } from 'fflate';

import { detectDocxFamilyKeys, detectDocxFontFamily } from '@/word/docx-to-pdf';
import { parseThemeFonts } from '@/core/drawingml/theme-parser';
import { readDocx } from '@/word/docx-reader';
import { resolveWordThemeFont } from '@/word/theme-fonts';

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

const THEME =
  `<?xml version="1.0"?>\n<a:theme xmlns:a="${A_NS}" name="doc"><a:themeElements>` +
  `<a:fontScheme name="doc">` +
  `<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Meiryo"/><a:cs typeface="Arial"/></a:minorFont>` +
  `</a:fontScheme></a:themeElements></a:theme>`;

/** A .docx whose only font statement is the one `rPr` makes. */
function docx(rPr: string, opts: { theme?: boolean } = {}): Uint8Array {
  const rels = (inner: string): Uint8Array =>
    enc.encode(`<?xml version="1.0"?>\n<Relationships xmlns="${PKG}">${inner}</Relationships>`);
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `</Types>`,
    ),
    '_rels/.rels': rels(
      `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`,
    ),
    'word/document.xml': enc.encode(
      `<?xml version="1.0"?>\n<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
        `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>hi</w:t></w:r></w:p></w:body></w:document>`,
    ),
    'word/styles.xml': enc.encode(
      `<?xml version="1.0"?>\n<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>${rPr}</w:rPr>` +
        `</w:rPrDefault></w:docDefaults></w:styles>`,
    ),
  };
  if (opts.theme !== false) files['word/theme/theme1.xml'] = enc.encode(THEME);
  return zipSync(files);
}

const THEMED = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>';

describe('a font named by theme slot (E-FONT F1)', () => {
  it('resolves each slot to the typeface the scheme names', () => {
    const fonts = parseThemeFonts(enc.encode(THEME));
    expect(resolveWordThemeFont('minorHAnsi', fonts)).toBe('Calibri');
    expect(resolveWordThemeFont('minorAscii', fonts)).toBe('Calibri');
    expect(resolveWordThemeFont('majorHAnsi', fonts)).toBe('Cambria');
    expect(resolveWordThemeFont('minorEastAsia', fonts)).toBe('Meiryo');
    expect(resolveWordThemeFont('minorBidi', fonts)).toBe('Arial');
    // An unknown slot names nothing rather than itself.
    expect(resolveWordThemeFont('somethingElse', fonts)).toBeUndefined();
    expect(resolveWordThemeFont('minorHAnsi', undefined)).toBeUndefined();
  });

  it('reaches the run properties a document is read into', () => {
    const { doc } = readDocx(docx(THEMED));
    expect(doc.styles.defaultRunProperties.fontFamily?.ascii).toBe('Calibri');
    const para = doc.body[0];
    expect(
      para?.kind === 'paragraph' ? para.paragraph.runs[0]?.properties.fontFamily : undefined,
    ).toMatchObject({ ascii: 'Calibri' });
  });

  it('names the substitute the auto-download path fetches', () => {
    // Without the theme this document declared no family at all, so the
    // downloader took the default sans — Arial's widths for Calibri's text.
    expect(detectDocxFontFamily(docx(THEMED))).toBe('Calibri');
    expect([...detectDocxFamilyKeys(docx(THEMED))]).toContain('carlito');
    // A document with no theme part is back to naming nothing.
    expect(detectDocxFontFamily(docx(THEMED, { theme: false }))).toBeUndefined();
  });

  it('reads a DrawingML token written into the literal attribute', () => {
    // fdo74605.docx spells `w:cs="+mn-cs"` — the OTHER dialect's token, in the
    // slot meant for a name. No substitution table knows a family beginning
    // with `+`, so it has to resolve here or travel on as a typeface.
    const { doc } = readDocx(docx('<w:rFonts w:ascii="Verdana" w:cs="+mn-cs"/>'));
    expect(doc.styles.defaultRunProperties.fontFamily).toEqual({ ascii: 'Verdana', cs: 'Arial' });
  });

  it('lets a spelled-out font win over the slot beside it', () => {
    const both = '<w:rFonts w:ascii="Verdana" w:asciiTheme="minorHAnsi"/>';
    expect(detectDocxFontFamily(docx(both))).toBe('Verdana');
    expect(readDocx(docx(both)).doc.styles.defaultRunProperties.fontFamily?.ascii).toBe('Verdana');
  });
});
