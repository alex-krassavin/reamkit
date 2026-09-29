// [MS-DOCX] compatibilityMode — Word 2010 and earlier place a table by its
// first cell's text, not its edge: the table stands out past its indent by
// that cell's left margin. Word 2013 (mode 15) places the edge. Measured in
// Word for Mac by where the first cell's text starts on the page.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const MARGIN = 72;

const para = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

interface TableOptions {
  /** `w:tblPr` children besides the width. */
  readonly tblPr?: string;
  /** The first cell's own `w:tcMar`. */
  readonly firstCellMargin?: number;
}

/** A two-column table whose first cell says "first". */
function table(options: TableOptions = {}): string {
  const tcMar =
    options.firstCellMargin === undefined
      ? ''
      : `<w:tcMar><w:left w:w="${String(options.firstCellMargin)}" w:type="dxa"/></w:tcMar>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>${options.tblPr ?? ''}</w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcMar}</w:tcPr>${para('first')}</w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para('second')}</w:tc>` +
    '</w:tr></w:tbl>'
  );
}

/** Where the first cell's text starts, in a document of `mode` (none: states no mode). */
function firstCellX(tbl: string, mode?: number, asWord = true): number {
  const docx = buildDocxFromBody(
    para('body') +
      tbl +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    mode === undefined
      ? {}
      : {
          settingsXml:
            '<w:compat><w:compatSetting w:name="compatibilityMode" ' +
            `w:uri="http://schemas.microsoft.com/office/word" w:val="${String(mode)}"/></w:compat>`,
        },
  );
  const flow = Ream.parse(docx).flow;
  const { typesetBy: _word, ...plain } = flowRenderOptions(flow);
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...(asWord ? flowRenderOptions(flow) : plain),
  });
  const line = laid.pages[0]!.commands.find(
    (c): c is TextLineItem =>
      c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === 'first'),
  );
  return line!.originX;
}

describe('a table in a document Word 2010 lays out', () => {
  it('stands out by its first cell’s margin, its text at the indent', () => {
    // Word: the text at 72 in modes 14 and none, at 77.85 in mode 15.
    expect(firstCellX(table(), 14)).toBeCloseTo(MARGIN, 3);
    expect(firstCellX(table())).toBeCloseTo(MARGIN, 3);
    expect(firstCellX(table(), 15)).toBeCloseTo(MARGIN + 5.4, 3);
  });

  it('by the first cell’s own margin, and the table’s where the cell has none', () => {
    // Word: 72 in mode 14 for both; 87.05 and 92.1 in mode 15.
    const tableMargin =
      '<w:tblCellMar><w:left w:w="300" w:type="dxa"/><w:right w:w="300" w:type="dxa"/></w:tblCellMar>';
    expect(firstCellX(table({ tblPr: tableMargin }), 14)).toBeCloseTo(MARGIN, 3);
    expect(firstCellX(table({ tblPr: tableMargin }), 15)).toBeCloseTo(MARGIN + 15, 3);
    expect(firstCellX(table({ firstCellMargin: 400 }), 14)).toBeCloseTo(MARGIN, 3);
    expect(firstCellX(table({ firstCellMargin: 400 }), 15)).toBeCloseTo(MARGIN + 20, 3);
  });

  it('measures its indent to that text', () => {
    // Word: 108 in mode 14, 113.85 in mode 15.
    const indent = '<w:tblInd w:w="720" w:type="dxa"/>';
    expect(firstCellX(table({ tblPr: indent }), 14)).toBeCloseTo(MARGIN + 36, 3);
    expect(firstCellX(table({ tblPr: indent }), 15)).toBeCloseTo(MARGIN + 36 + 5.4, 3);
  });

  it('leaves a centred table where it is', () => {
    const centred = '<w:jc w:val="center"/>';
    expect(firstCellX(table({ tblPr: centred }), 14)).toBeCloseTo(
      firstCellX(table({ tblPr: centred }), 15),
      3,
    );
  });

  it('is Word’s alone: a table that is not a Word document’s stays at its edge', () => {
    expect(firstCellX(table(), 14, false)).toBeCloseTo(MARGIN + 5.4, 3);
  });
});
