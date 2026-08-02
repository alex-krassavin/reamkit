import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { parseRunProperties } from '@/word';

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
      fontSizePt: halfPtToPt(28),
      colorHex: 'FF0000',
    });
  });

  it('parses underline style', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:u w:val="single"/></w:rPr>'))).toEqual({
      underline: 'single',
    });
  });

  it('gives the underline its own colour (§17.3.2.40)', () => {
    // A themed colour is written alongside the resolved hex it stands for, so
    // reading `w:color` is enough. Ignored, Test_CharUnderlineThemeColor.docx
    // drew a gold rule under black text in black.
    expect(
      parseRunProperties(
        parseRpr('<w:rPr><w:u w:val="single" w:color="C49A00" w:themeColor="accent1"/></w:rPr>'),
      ),
    ).toEqual({ underline: 'single', underlineColorHex: 'C49A00' });
    // "auto" is the text's own colour, which is what an absent one means.
    expect(
      parseRunProperties(parseRpr('<w:rPr><w:u w:val="single" w:color="auto"/></w:rPr>')),
    ).toEqual({ underline: 'single' });
  });

  it('reads the capitals toggles (§17.3.2.5 / §17.3.2.33)', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:caps/></w:rPr>'))).toEqual({ caps: true });
    expect(parseRunProperties(parseRpr('<w:rPr><w:smallCaps/></w:rPr>'))).toEqual({
      smallCaps: true,
    });
    // §17.17.4 — an explicit off is off, not "inherit".
    expect(parseRunProperties(parseRpr('<w:rPr><w:caps w:val="0"/></w:rPr>'))).toEqual({
      caps: false,
    });
  });

  it('rejects malformed color values', () => {
    expect(parseRunProperties(parseRpr('<w:rPr><w:color w:val="ZZZZZZ"/></w:rPr>'))).toEqual({});
  });

  it('§17.3.2.6 — `auto` is the automatic colour, and overrides the style', () => {
    // Not "inherit": a run that names it takes black back from a style that
    // lends it something else.
    expect(parseRunProperties(parseRpr('<w:rPr><w:color w:val="auto"/></w:rPr>'))).toEqual({
      colorHex: '000000',
    });
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
