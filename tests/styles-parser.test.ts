import { describe, expect, it } from 'vitest';

import { parseStyles } from '@/ooxml/wordproc';

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
    expect(sheet.defaultRunProperties).toEqual({ fontSizeHalfPoints: 24 });
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
    expect(h1!.runProperties).toEqual({ bold: true, fontSizeHalfPoints: 32 });
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
