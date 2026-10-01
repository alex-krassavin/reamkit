// ECMA-376 Part 1 §17.4.38 — in a document Word sets, a horizontal table
// border takes the room it is wide: the edge above a row in that row, the
// table's bottom edge in its last one. Measured in Word for Mac by the
// positions of 11pt Calibri paragraphs, averaged over 15–20 rows.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BorderItem, PageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** 11pt Calibri's line, as Word sets it (see word-line-heights.test.ts). */
const LINE = (11 * 2500) / 2048;

const para = (text: string): string =>
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** Table borders of `eighths` of a point, the outer top and bottom as asked. */
function borders(eighths: number, top = true, bottom = true): string {
  const rule = `w:val="single" w:sz="${String(eighths)}" w:space="0" w:color="000000"`;
  const nil = 'w:val="nil"';
  return (
    `<w:tblBorders><w:top ${top ? rule : nil}/><w:left ${rule}/>` +
    `<w:bottom ${bottom ? rule : nil}/><w:right ${rule}/>` +
    `<w:insideH ${rule}/><w:insideV ${rule}/></w:tblBorders>`
  );
}

/** A one-column table of the given rows, no cell margins top or bottom. */
function table(rows: ReadonlyArray<string>, tblBorders: string): string {
  const margins =
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${tblBorders}${margins}</w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    rows
      .map(
        (r) =>
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(r)}</w:tc></w:tr>`,
      )
      .join('') +
    '</w:tbl>'
  );
}

const SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

function layout(body: string, asWord = true): ReadonlyArray<PageItem> {
  const flow = Ream.parse(buildDocxFromBody(body + SECTION)).flow;
  const { typesetBy: _word, ...plain } = flowRenderOptions(flow);
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...(asWord ? flowRenderOptions(flow) : plain),
  });
  return laid.pages[0]!.commands;
}

/** The baseline of the line holding `text`. */
function baseline(items: ReadonlyArray<PageItem>, text: string): number {
  const hit = items.find(
    (c): c is TextLineItem =>
      c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === text),
  );
  if (!hit) throw new Error(`no line "${text}"`);
  return hit.baselineY;
}

describe('a Word table’s rows (§17.4.38)', () => {
  const rows = Array.from({ length: 5 }, (_, i) => `r${String(i)}`);
  const pitch = (eighths: number): number => {
    const items = layout(table(rows, borders(eighths)));
    return (baseline(items, 'r4') - baseline(items, 'r1')) / 3;
  };

  it('stand a border’s width apart beyond their text', () => {
    // Word: 13.44 with no borders, 13.92 under 0.5pt, 14.94 under 1.5pt, 16.43 under 3pt.
    expect(pitch(0)).toBeCloseTo(LINE, 3);
    expect(pitch(4)).toBeCloseTo(LINE + 0.5, 3);
    expect(pitch(12)).toBeCloseTo(LINE + 1.5, 3);
    expect(pitch(24)).toBeCloseTo(LINE + 3, 3);
  });

  it('give the table’s own top and bottom borders their room, each whole', () => {
    // Word: a one-row table between two paragraphs adds 1.01 under 0.5pt
    // borders, 5.98 under 3pt ones, 2.98 when only its top or its bottom is 3pt.
    const between = (tblBorders: string): number => {
      const items = layout(para('above') + table(['cell'], tblBorders) + para('below'));
      return baseline(items, 'below') - baseline(items, 'above') - 2 * LINE;
    };
    expect(between(borders(0))).toBeCloseTo(0, 3);
    expect(between(borders(4))).toBeCloseTo(1, 3);
    expect(between(borders(24))).toBeCloseTo(6, 3);
    expect(between(borders(24, true, false))).toBeCloseTo(3, 3);
    expect(between(borders(24, false, true))).toBeCloseTo(3, 3);
  });

  it('set their text below the border, and draw it down the middle of its room', () => {
    const items = layout(para('above') + table(['cell'], borders(24)));
    const top = items.find((c): c is BorderItem => c.type === 'border' && c.side === 'top')!;
    // The paragraph above ends at its line; the rule's 3pt start there and the
    // text's line after them.
    const tableTop = baseline(items, 'above') + (11 * 550) / 2048;
    expect(top.y).toBeCloseTo(tableTop + 1.5, 3);
    expect(baseline(items, 'cell') - baseline(items, 'above')).toBeCloseTo(LINE + 3, 3);
  });

  it('take no room for their borders where the document is not Word’s', () => {
    const items = layout(table(rows, borders(24)), false);
    expect((baseline(items, 'r4') - baseline(items, 'r1')) / 3).toBeCloseTo(11 * 1.2, 3);
  });
});
