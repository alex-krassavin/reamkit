import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Numbering } from '@/core/document-model';
import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { parseNumbering } from '@/word';
import { NumberingState } from '@/core/numbering';
import { readDocx } from '@/word/docx-reader';

const encoder = new TextEncoder();

function parse(xml: string): Numbering {
  return parseNumbering(
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</w:numbering>`,
    ),
  );
}

describe('parseNumbering', () => {
  it('a level whose bullet is a picture carries the image and its size', () => {
    // §17.9.9 / §17.9.21 — the `w:lvlText` glyph is only the fallback beside it.
    const numbering = parse(`
      <w:numPicBullet w:numPicBulletId="3">
        <w:pict>
          <v:shape style="width:11.25pt;height:0.25in" o:bullet="t">
            <v:imagedata r:id="rId7"/>
          </v:shape>
        </w:pict>
      </w:numPicBullet>
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="bullet"/>
          <w:lvlText w:val="&#xF0B7;"/>
          <w:lvlPicBulletId w:val="3"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const lvl = numbering.abstractNums.get('0')!.levels.get(0)!;
    expect(lvl.picBullet?.widthPt).toBeCloseTo(11.25, 6);
    expect(lvl.picBullet?.heightPt).toBeCloseTo(18, 6); // 0.25in
  });

  it('parses a single-level decimal abstractNum + num link', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
          <w:start w:val="1"/>
          <w:numFmt w:val="decimal"/>
          <w:lvlText w:val="%1."/>
          <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    expect(numbering.abstractNums.size).toBe(1);
    expect(numbering.numInstances.size).toBe(1);
    const abstractNum = numbering.abstractNums.get('0')!;
    const level = abstractNum.levels.get(0)!;
    expect(level.format).toBe('decimal');
    expect(level.lvlText).toBe('%1.');
    expect(level.start).toBe(1);
    expect(level.paragraphProperties.indentLeft).toBe(twipsToPt(720));
    expect(level.paragraphProperties.indentFirstLine).toBe(twipsToPt(-360));
    expect(numbering.numInstances.get('1')!.abstractNumId).toBe('0');
  });

  // §17.9.25 — a level that states no `w:start` starts at ZERO, and the second
  // item is 1: FDO74105.docx omits it and its list runs 0, 1, 2, …
  it('starts a level with no w:start at zero, and counts on from there', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    expect(numbering.abstractNums.get('0')!.levels.get(0)!.start).toBe(0);
    const state = new NumberingState();
    const ref = { numId: '1', ilvl: 0 };
    expect(state.resolveMarker(numbering, ref)).toBe('0.');
    expect(state.resolveMarker(numbering, ref)).toBe('1.');
    expect(state.resolveMarker(numbering, ref)).toBe('2.');
  });

  it('parses multi-level numbering with different formats', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2."/></w:lvl>
        <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
    `);
    const abstractNum = numbering.abstractNums.get('0')!;
    expect(abstractNum.levels.get(0)!.format).toBe('upperRoman');
    expect(abstractNum.levels.get(1)!.format).toBe('upperLetter');
    expect(abstractNum.levels.get(2)!.format).toBe('decimal');
  });
});

describe('NumberingState marker generation', () => {
  it('produces sequential decimal markers for a single level', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    const state = new NumberingState();
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('1.');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('2.');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('3.');
  });

  it('resets deeper levels when a shallower level advances', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    const state = new NumberingState();
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('1');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 1 })).toBe('1.a');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 1 })).toBe('1.b');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('2');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 1 })).toBe('2.a');
  });

  it('formats roman numerals (lower and upper)', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1)"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
    `);
    const state = new NumberingState();
    for (let i = 0; i < 4; i++) state.resolveMarker(numbering, { numId: '1', ilvl: 0 });
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('V.');
    expect(state.resolveMarker(numbering, { numId: '2', ilvl: 0 })).toBe('i)');
    expect(state.resolveMarker(numbering, { numId: '2', ilvl: 0 })).toBe('ii)');
  });

  it('treats lvlText as bullet character when numFmt=bullet', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    const state = new NumberingState();
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('•');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('•');
  });

  it('substitutes Symbol-font private-use bullets with U+2022', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    const state = new NumberingState();
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('•');
  });

  it('respects per-level start value', () => {
    const numbering = parse(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    `);
    const state = new NumberingState();
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('5.');
    expect(state.resolveMarker(numbering, { numId: '1', ilvl: 0 })).toBe('6.');
  });
});

describe('the ideograph and decimal variants (§17.18.59)', () => {
  const runFmt = (fmt: string, count: number): Array<string> => {
    const numbering = parse(
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
        `<w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
    );
    const state = new NumberingState();
    return Array.from(
      { length: count },
      () => state.resolveMarker(numbering, { numId: '1', ilvl: 0 }) ?? '',
    );
  };

  // Every expectation here is what LibreOffice draws for the matching
  // cjklist*.docx: the cycles fall back to digits past their end, the counting
  // systems elide the leading 一 of the tens place and the formal one does not.
  it('numbers the twelve branches, then in digits', () => {
    expect(runFmt('ideographZodiac', 14)).toEqual([
      '子',
      '丑',
      '寅',
      '卯',
      '辰',
      '巳',
      '午',
      '未',
      '申',
      '酉',
      '戌',
      '亥',
      '13',
      '14',
    ]);
  });

  it('numbers the ten stems, then in digits', () => {
    expect(runFmt('ideographTraditional', 12).slice(8)).toEqual(['壬', '癸', '11', '12']);
  });

  it('counts 十一 and writes the formal numerals 壹拾壹', () => {
    expect(runFmt('taiwaneseCountingThousand', 21).slice(9)).toEqual([
      '十',
      '十一',
      '十二',
      '十三',
      '十四',
      '十五',
      '十六',
      '十七',
      '十八',
      '十九',
      '二十',
      '二十一',
    ]);
    expect(runFmt('ideographLegalTraditional', 12).slice(8)).toEqual([
      '玖',
      '壹拾',
      '壹拾壹',
      '壹拾貳',
    ]);
  });

  it('writes the digital systems digit by digit', () => {
    expect(runFmt('koreanDigital2', 12).slice(8)).toEqual(['九', '一零', '一一', '一二']);
    expect(runFmt('decimalFullWidth', 11).slice(8)).toEqual(['９', '１０', '１１']);
  });

  it('pads a decimalZero and suffixes an ordinal', () => {
    expect(runFmt('decimalZero', 11).slice(8)).toEqual(['09', '10', '11']);
    expect(runFmt('ordinal', 4)).toEqual(['1st', '2nd', '3rd', '4th']);
    expect(runFmt('ordinal', 22).slice(10)).toEqual([
      '11th',
      '12th',
      '13th',
      '14th',
      '15th',
      '16th',
      '17th',
      '18th',
      '19th',
      '20th',
      '21st',
      '22nd',
    ]);
  });
});

describe('numbering a paragraph carries through its style (§17.9.24)', () => {
  it('numbers a heading whose w:numPr lives in the style, not the paragraph', () => {
    // chtoutline.docx numbers Heading 1 "第 %1 章" from the style alone, and
    // reading the paragraph's own properties dropped the chapter number.
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
      '<w:lvlText w:val="Chapter %1."/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const stylesXml =
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Test</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Again</w:t></w:r></w:p>';
    const { doc } = readDocx(buildDocxFromBody(body, { numberingXml, stylesXml }));
    const markers = doc.body.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs.filter((r) => r.listMarker).map((r) => r.text) : [],
    );
    expect(markers).toEqual(['Chapter 1.\t', 'Chapter 2.\t']);
  });
});
