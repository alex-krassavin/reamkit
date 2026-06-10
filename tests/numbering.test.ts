import { describe, expect, it } from 'vitest';

import type { Numbering } from '@/document-model';
import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/ir';

import { parseNumbering } from '@/ooxml/wordproc';
import { NumberingState } from '@/numbering';

const encoder = new TextEncoder();

function parse(xml: string): Numbering {
  return parseNumbering(
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</w:numbering>`,
    ),
  );
}

describe('parseNumbering', () => {
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
