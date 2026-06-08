import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import { parseRunProperties } from '@/ooxml/wordproc';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

function parseRpr(inner: string): unknown {
  const xml = `<root xmlns:w="ns">${inner}</root>`;
  const tree = parser.parse(xml) as { root?: { 'w:rPr'?: unknown } };
  return tree.root?.['w:rPr'];
}

describe('parseRunProperties', () => {
  it('returns empty object when rPr is absent', () => {
    expect(parseRunProperties(parseRpr(''))).toEqual({});
  });

  it('treats a self-closing toggle as true', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:b/></w:rPr>'))).toEqual({ bold: true });
  });

  it('parses explicit val=false toggle', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:b w:val="false"/></w:rPr>'))).toEqual({
      bold: false,
    });
    expect(parseRunProperties(parseRpr('<w:rPr><w:i w:val="0"/></w:rPr>'))).toEqual({
      italic: false,
    });
  });

  it('parses italic, strike, font size, color together', () => {
    const rPr = parseRpr(
      '<w:rPr><w:i/><w:strike/><w:sz w:val="28"/><w:color w:val="ff0000"/></w:rPr>',
    );
    expect(parseRunProperties(rPr)).toEqual({
      italic: true,
      strike: true,
      fontSizeHalfPoints: 28,
      colorHex: 'FF0000',
    });
  });

  it('parses underline style', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:u w:val="single"/></w:rPr>'))).toEqual({
      underline: 'single',
    });
  });

  it('rejects malformed color values', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:color w:val="auto"/></w:rPr>'))).toEqual({});
    expect(parseRunProperties(parseRpr('<w:rPr><w:color w:val="ZZZZZZ"/></w:rPr>'))).toEqual({});
  });

  it('parses rFonts ascii + hAnsi', () => {
    expect(
      parseRunProperties(parseRpr('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>')),
    ).toEqual({
      fontFamily: { ascii: 'Arial', hAnsi: 'Arial' },
    });
  });

  it('parses style reference and vertical align', () => {
    expect(
      parseRunProperties(
        parseRpr('<w:rPr><w:rStyle w:val="Emphasis"/><w:vertAlign w:val="superscript"/></w:rPr>'),
      ),
    ).toEqual({ styleId: 'Emphasis', verticalAlign: 'superscript' });
  });

  it('parses w:rtl toggle', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:rtl/></w:rPr>'))).toEqual({ rtl: true });
    expect(parseRunProperties(parseRpr('<w:rPr><w:rtl w:val="false"/></w:rPr>'))).toEqual({
      rtl: false,
    });
  });
});
