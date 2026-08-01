import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { parseParagraphProperties } from '@/word';

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
      spacingBefore: twipsToPt(240),
      spacingAfter: twipsToPt(120),
      spacingLine: twipsToPt(276),
      spacingLineRule: 'auto',
    });
  });

  it('parses indent left/right/firstLine', () => {
    const result = parseParagraphProperties(
      parsePpr('<w:pPr><w:ind w:left="720" w:right="0" w:firstLine="360"/></w:pPr>'),
    );
    expect(result).toEqual({
      indentLeft: twipsToPt(720),
      indentRight: twipsToPt(0),
      indentFirstLine: twipsToPt(360),
    });
  });

  it('parses w:hanging as negative indentFirstLine', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>')),
    ).toEqual({ indentLeft: twipsToPt(720), indentFirstLine: twipsToPt(-360) });
  });

  it('prefers w:firstLine over w:hanging when both are set', () => {
    expect(
      parseParagraphProperties(
        parsePpr('<w:pPr><w:ind w:left="720" w:firstLine="180" w:hanging="360"/></w:pPr>'),
      ),
    ).toEqual({ indentLeft: twipsToPt(720), indentFirstLine: twipsToPt(180) });
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

// ECMA-376 Part 1 §17.3.1.37 `w:tabs` — a tab advances to a POSITION, not by a
// fixed amount. Measured as ordinary whitespace it collapsed to a space: the
// page numbers of FDO77715.docx's index sat against their titles with no dot
// leader between, and its header ran its left and right halves together.
describe('tab stops (§17.3.1.37)', () => {
  it('reads a stop’s position, alignment and leader', () => {
    expect(
      parseParagraphProperties(
        parsePpr(
          '<w:pPr><w:tabs><w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs></w:pPr>',
        ),
      ).tabs,
    ).toEqual([{ positionPt: twipsToPt(9360), alignment: 'right', leader: 'dot' }]);
  });

  it('sorts the stops by position, whatever order they are written in', () => {
    expect(
      parseParagraphProperties(
        parsePpr(
          '<w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/>' +
            '<w:tab w:val="center" w:pos="4680"/></w:tabs></w:pPr>',
        ),
      ).tabs?.map((t) => t.positionPt),
    ).toEqual([twipsToPt(4680), twipsToPt(9360)]);
  });

  it('keeps only the stops that place text', () => {
    // §17.18.90 — `bar` draws a rule and advances nothing; `clear` removes an
    // inherited stop. Neither positions anything, so neither becomes one.
    expect(
      parseParagraphProperties(
        parsePpr(
          '<w:pPr><w:tabs><w:tab w:val="bar" w:pos="1440"/>' +
            '<w:tab w:val="clear" w:pos="2880"/>' +
            '<w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>',
        ),
      ).tabs,
    ).toEqual([{ positionPt: twipsToPt(720), alignment: 'left' }]);
  });

  it('ignores a stop that names no position', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:tabs><w:tab w:val="left"/></w:tabs></w:pPr>'))
        .tabs,
    ).toBeUndefined();
  });

  it('takes `start` and `end` as the left and right they are', () => {
    expect(
      parseParagraphProperties(
        parsePpr(
          '<w:pPr><w:tabs><w:tab w:val="start" w:pos="720"/>' +
            '<w:tab w:val="end" w:pos="1440"/></w:tabs></w:pPr>',
        ),
      ).tabs?.map((t) => t.alignment),
    ).toEqual(['left', 'right']);
  });
});
