import { describe, expect, it } from 'vitest';

import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { parseStyles } from '@/word';

const encoder = new TextEncoder();

function parse(stylesXml: string) {
  return parseStyles(
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${stylesXml}</w:styles>`,
    ),
  );
}

describe('parseStyles', () => {
  it('parses docDefaults', () => {
    const sheet = parse(
      '<w:docDefaults>' +
        '<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault>' +
        '</w:docDefaults>',
    );
    expect(sheet.defaultRunProperties).toEqual({ fontSizePt: halfPtToPt(24) });
    expect(sheet.defaultParagraphProperties).toEqual({ alignment: 'both' });
  });

  it('parses a paragraph style with basedOn', () => {
    const sheet = parse(
      '<w:style w:type="paragraph" w:styleId="Heading1">' +
        '<w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>' +
        '</w:style>',
    );
    const h1 = sheet.styles.get('Heading1');
    expect(h1).toBeDefined();
    expect(h1!.type).toBe('paragraph');
    expect(h1!.basedOn).toBe('Normal');
    expect(h1!.runProperties).toEqual({ bold: true, fontSizePt: halfPtToPt(32) });
    expect(h1!.paragraphProperties).toEqual({ alignment: 'center' });
  });

  it('parses default flag', () => {
    const sheet = parse(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr/></w:style>',
    );
    expect(sheet.styles.get('Normal')!.isDefault).toBe(true);
  });

  it('ignores styles with unknown type', () => {
    const sheet = parse('<w:style w:type="bogus" w:styleId="X"><w:rPr/></w:style>');
    expect(sheet.styles.size).toBe(0);
  });

  it('parses multiple styles and a character style', () => {
    const sheet = parse(
      '<w:style w:type="paragraph" w:styleId="Normal"><w:rPr/></w:style>' +
        '<w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style>',
    );
    expect(sheet.styles.size).toBe(2);
    expect(sheet.styles.get('Emphasis')!.type).toBe('character');
    expect(sheet.styles.get('Emphasis')!.runProperties).toEqual({ italic: true });
  });
});

describe("parseStyles — Word's built-in headings, known by name (§17.7.4.9)", () => {
  // Word 16 for Mac, `outline level of paragraph format` over a probe
  // (2026-11-20): the NAME decides, never the id.
  const sheet = parse(
    '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="X5"><w:name w:val="heading 5"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="Fancy"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/>' +
      '<w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>' +
      '<w:style w:type="character" w:styleId="Heading1Char"><w:name w:val="heading 1"/></w:style>',
  );

  it('gives a "heading N" style level N-1 when it states none', () => {
    expect(sheet.styles.get('Heading4')!.builtInOutlineLevel).toBe(3);
    expect(sheet.styles.get('Heading4')!.paragraphProperties.outlineLevel).toBe(3);
  });

  it('reads the name under any id', () => {
    expect(sheet.styles.get('X5')!.name).toBe('heading 5');
    expect(sheet.styles.get('X5')!.paragraphProperties.outlineLevel).toBe(4);
    expect(sheet.styles.get('1')!.builtInOutlineLevel).toBe(0);
  });

  it("fixes the level over the style's own w:outlineLvl", () => {
    expect(sheet.styles.get('1')!.paragraphProperties.outlineLevel).toBe(0);
  });

  it('leaves a heading-like id under another name alone', () => {
    expect(sheet.styles.get('Heading6')!.builtInOutlineLevel).toBeUndefined();
    expect(sheet.styles.get('Heading6')!.paragraphProperties.outlineLevel).toBeUndefined();
  });

  it('applies to paragraph styles only', () => {
    expect(sheet.styles.get('Heading1Char')!.builtInOutlineLevel).toBeUndefined();
  });
});
