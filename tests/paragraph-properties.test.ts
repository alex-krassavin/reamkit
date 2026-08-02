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

// §17.3.1.24 `w:pBdr` — rules around the paragraph, spelled exactly as a cell's
// are but with a `w:space` in POINTS. Read nowhere, Test_ThemeBorderColor.docx
// lost the two coloured rules that are the whole of its page.
describe('paragraph borders (§17.3.1.24)', () => {
  it('reads each edge with its width, colour and standoff', () => {
    const p = parseParagraphProperties(
      parsePpr(
        '<w:pPr><w:pBdr>' +
          '<w:top w:val="single" w:sz="48" w:space="1" w:color="DE81E1"/>' +
          '<w:bottom w:val="double" w:sz="8" w:space="4" w:color="90ABF0"/>' +
          '</w:pBdr></w:pPr>',
      ),
    );
    expect(p.borders?.top).toEqual({
      style: 'single',
      width: eighthPtToPt(48),
      spacePt: 1,
      colorHex: 'DE81E1',
    });
    expect(p.borders?.bottom).toEqual({
      style: 'double',
      width: eighthPtToPt(8),
      spacePt: 4,
      colorHex: '90ABF0',
    });
    expect(p.borders?.left).toBeUndefined();
  });

  it('takes start and end as the left and right they are', () => {
    const p = parseParagraphProperties(
      parsePpr(
        '<w:pPr><w:pBdr><w:start w:val="single" w:sz="4"/><w:end w:val="single" w:sz="4"/></w:pBdr></w:pPr>',
      ),
    );
    expect(p.borders?.left?.style).toBe('single');
    expect(p.borders?.right?.style).toBe('single');
  });

  it('draws a pattern it cannot spell as a solid rule', () => {
    // §17.18.2 names some hundred and eighty patterns. Rejecting the ones we
    // cannot draw exactly lost SdtContent.docx the `thickThinSmallGap` rule
    // under its header; a solid rule of the stated width is far closer.
    expect(
      parseParagraphProperties(
        parsePpr(
          '<w:pPr><w:pBdr><w:bottom w:val="thickThinSmallGap" w:sz="24" w:color="622423"/></w:pBdr></w:pPr>',
        ),
      ).borders?.bottom,
    ).toEqual({ style: 'single', width: eighthPtToPt(24), colorHex: '622423' });
  });

  it('records nil as a rule that is explicitly absent', () => {
    // Not "unspecified": an edge spelled `nil` overrides the one a style would
    // otherwise lend it, and the drawing skips a `none` as it always has.
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:pBdr><w:top w:val="nil"/></w:pBdr></w:pPr>'))
        .borders?.top,
    ).toEqual({ style: 'none' });
  });

  it('ignores an edge with no style, and an auto colour', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:pBdr><w:top w:sz="4"/></w:pBdr></w:pPr>'))
        .borders,
    ).toBeUndefined();
    expect(
      parseParagraphProperties(
        parsePpr('<w:pPr><w:pBdr><w:top w:val="single" w:color="auto"/></w:pBdr></w:pPr>'),
      ).borders?.top?.colorHex,
    ).toBeUndefined();
  });
});

// §17.3.1.31 `w:pPr/w:shd` — the paragraph's own background. Read nowhere,
// Test_ThemeTextParaBackgroundColor.docx printed three bare lines where
// LibreOffice fills three bands behind them.
describe('paragraph shading (§17.3.1.31)', () => {
  it('reads a direct fill', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:shd w:val="clear" w:fill="F4E7D3"/></w:pPr>'))
        .shading,
    ).toEqual({ colorHex: 'F4E7D3' });
  });

  it('leaves an auto or absent fill unshaded', () => {
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:shd w:val="clear" w:fill="auto"/></w:pPr>'))
        .shading,
    ).toBeUndefined();
    expect(
      parseParagraphProperties(parsePpr('<w:pPr><w:shd w:val="pct25"/></w:pPr>')).shading,
    ).toBeUndefined();
  });
});
