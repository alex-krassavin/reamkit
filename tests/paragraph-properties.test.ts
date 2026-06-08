import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import { parseParagraphProperties } from '@/ooxml/wordproc';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

function parsePpr(inner: string): unknown {
  const xml = `<root xmlns:w="ns">${inner}</root>`;
  const tree = parser.parse(xml) as { root?: { 'w:pPr'?: unknown } };
  return tree.root?.['w:pPr'];
}

describe('parseParagraphProperties', () => {
  it('returns empty object when pPr is absent', () => {
    expect(parseParagraphProperties(parsePpr(''))).toEqual({});
  });

  it('parses style reference', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')),
    ).toEqual({ styleId: 'Heading1' });
  });

  it('parses outline level (w:outlineLvl)', () => {
    expect(parseParagraphProperties(parsePpr('<w:pPr><w:outlineLvl w:val="2"/></w:pPr>'))).toEqual({
      outlineLevel: 2,
    });
  });

  it('parses each alignment value', () => {
    for (const a of ['left', 'right', 'center', 'both', 'distribute'] as const) {
      expect(parseParagraphProperties(parsePpr(`<w:pPr><w:jc w:val="${a}"/></w:pPr>`))).toEqual({
        alignment: a,
      });
    }
  });

  it('ignores unknown alignment value', () => {
    expect(parseParagraphProperties(parsePpr('<w:pPr><w:jc w:val="bogus"/></w:pPr>'))).toEqual({});
  });

  it('parses spacing before/after/line + lineRule', () => {
    const result = parseParagraphProperties(
      parsePpr(
        '<w:pPr><w:spacing w:before="240" w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>',
      ),
    );
    expect(result).toEqual({
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      spacingLineTwips: 276,
      spacingLineRule: 'auto',
    });
  });

  it('parses indent left/right/firstLine', () => {
    const result = parseParagraphProperties(
      parsePpr('<w:pPr><w:ind w:left="720" w:right="0" w:firstLine="360"/></w:pPr>'),
    );
    expect(result).toEqual({
      indentLeftTwips: 720,
      indentRightTwips: 0,
      indentFirstLineTwips: 360,
    });
  });

  it('parses w:hanging as negative indentFirstLineTwips', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>')),
    ).toEqual({ indentLeftTwips: 720, indentFirstLineTwips: -360 });
  });

  it('prefers w:firstLine over w:hanging when both are set', () => {
    expect(
      parseParagraphProperties(
        parsePpr('<w:pPr><w:ind w:left="720" w:firstLine="180" w:hanging="360"/></w:pPr>'),
      ),
    ).toEqual({ indentLeftTwips: 720, indentFirstLineTwips: 180 });
  });

  it('parses w:numPr (numId + ilvl)', () => {
    expect(
      parseParagraphProperties(
        parsePpr('<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="5"/></w:numPr></w:pPr>'),
      ),
    ).toEqual({ numbering: { numId: '5', ilvl: 2 } });
  });

  it('parses nested rPr (paragraph-mark run properties)', () => {
    const result = parseParagraphProperties(parsePpr('<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>'));
    expect(result).toEqual({ runProperties: { bold: true } });
  });

  it('parses w:bidi toggle (RTL base direction)', () => {
    expect(parseParagraphProperties(parsePpr('<w:pPr><w:bidi/></w:pPr>'))).toEqual({ bidi: true });
    expect(parseParagraphProperties(parsePpr('<w:pPr><w:bidi w:val="0"/></w:pPr>'))).toEqual({
      bidi: false,
    });
  });
});
