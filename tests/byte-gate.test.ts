// Byte gate — the project's core invariant, automated (oop-design.md §7).
//
// A PURE refactoring must keep PDF output byte-for-byte identical. The corpus
// pipeline diffs against LibreOffice through rasterization, and the other
// "byte-identical" tests check determinism within one revision — neither
// catches object renumbering BETWEEN revisions. These snapshots do.
//
// If this test fails on your change:
//   - refactoring? → you changed the bytes; find out why and undo it;
//   - deliberate output change? → review the diff, then `vitest run
//     tests/byte-gate.test.ts -u` and call out the gate update in the commit.

import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { buildTinyPng } from './fixtures/build-png';
import { convertDocxToPdfSync, convertXlsxToPdfSync } from '@/core/converter';

const FIXTURE_DIR = 'tests/fixtures/byte-gate';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

describe('byte gate: static fixtures', () => {
  // Synthetic documents checked into the repo (text, tables, lists,
  // headers/footers, sheet grids and number formats).
  for (const name of readdirSync(FIXTURE_DIR).sort()) {
    it(name, () => {
      const bytes = new Uint8Array(readFileSync(join(FIXTURE_DIR, name)));
      const pdf = name.endsWith('.xlsx')
        ? convertXlsxToPdfSync(bytes, { fonts })
        : convertDocxToPdfSync(bytes, { fonts });
      expect(sha256(pdf)).toMatchSnapshot();
    });
  }
});

describe('byte gate: images and PDF/A emit paths', () => {
  // Opaque PNG + semi-transparent PNG (SMask) inline in one paragraph; the
  // image pipeline is the most order-sensitive part of the emit phase.
  const opaque = buildTinyPng(4, 4, [0, 200, 0, 255]);
  const translucent = buildTinyPng(4, 4, [200, 0, 0, 128]);
  const body =
    '<w:p><w:r><w:t>images</w:t></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId21"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const docx = buildDocxFromBody(body, {
    images: {
      rId20: { contentType: 'image/png', bytes: opaque, extension: 'png' },
      rId21: { contentType: 'image/png', bytes: translucent, extension: 'png' },
    },
  });

  it('images.docx (plain: SMask kept)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-1b: alpha flattened)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-1b' }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-2b: transparency group)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-2b' }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-1a: tagged structure tree)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-1a' }))).toMatchSnapshot();
  });
});

// E-SHEET SA0 — the byte-zero safety net for the SheetDoc refactor (SA1/SA2).
// These synthesized fixtures exercise the full worksheet → print-model → FlowDoc
// projection surface, so relocating it behind a SheetDoc boundary can be proven
// byte-for-byte identical. Each renders to PDF and snapshots its hash; the same
// FlowDoc feeds the SVG/HTML writers, so PDF identity implies theirs.
describe('byte gate: xlsx feature surface (E-SHEET SA0)', () => {
  // One style sheet referenced by index: bold/coloured font, solid fill, thick
  // borders, a custom number format, centered alignment, a built-in date format.
  const STYLES = `
    <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
    <fonts count="2">
      <font><sz val="11"/><name val="Calibri"/></font>
      <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
    </fonts>
    <fills count="3">
      <fill><patternFill patternType="none"/></fill>
      <fill><patternFill patternType="gray125"/></fill>
      <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    </fills>
    <borders count="2">
      <border/>
      <border>
        <left style="thick"><color rgb="FFD92020"/></left>
        <right style="thick"><color rgb="FFD92020"/></right>
        <top style="thick"><color rgb="FFD92020"/></top>
        <bottom style="thick"><color rgb="FFD92020"/></bottom>
      </border>
    </borders>
    <cellXfs count="7">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
      <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
      <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
      <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="center"/></xf>
      <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    </cellXfs>`;

  const BAR_CHART =
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>' +
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/>' +
    '<c:order val="0"/><c:val><c:numRef><c:numCache><c:ptCount val="2"/>' +
    '<c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef>' +
    '</c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

  const grid = (rows: number, cols: number): Array<Array<number>> =>
    Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => r * cols + c + 1),
    );

  const cases: Record<string, Uint8Array> = {
    merges: buildXlsx({
      rows: [
        ['Merged header', null, null],
        ['a', 'b', 'c'],
        ['x', 'y', 'z'],
      ],
      mergeRefs: ['A1:C1', 'A2:A3'],
    }),
    'cell-styles': buildXlsx({
      rows: [
        [
          { value: 'Bold', styleIndex: 1 },
          { value: 'Filled', styleIndex: 2 },
        ],
        [
          { value: 'Bordered', styleIndex: 3 },
          { value: 1234.5, styleIndex: 4 },
        ],
        [{ value: 'Centered', styleIndex: 5 }],
      ],
      stylesXml: STYLES,
    }),
    tracks: buildXlsx({
      rows: [['narrow', 'wide']],
      columns: [
        { min: 1, max: 1, widthChars: 5 },
        { min: 2, max: 2, widthChars: 30 },
      ],
      rowHeights: [{ row: 0, heightPt: 40 }],
    }),
    'fit-to-page': buildXlsx({
      rows: grid(40, 12),
      fitToPage: true,
      pageSetup: { fitToWidth: 1, fitToHeight: 1 },
    }),
    scale: buildXlsx({ rows: grid(10, 6), pageSetup: { scale: 75 } }),
    'print-area': buildXlsx({
      rows: grid(5, 5),
      definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Sheet1!$A$1:$B$2' }],
    }),
    'print-titles': buildXlsx({
      rows: grid(60, 4),
      definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 0, value: 'Sheet1!$1:$1' }],
    }),
    'row-breaks': buildXlsx({ rows: grid(6, 3), rowBreaks: [2] }),
    centering: buildXlsx({
      rows: [['centered']],
      printOptions: { horizontalCentered: true, verticalCentered: true },
    }),
    gridlines: buildXlsx({ rows: grid(3, 3), printOptions: { gridLines: true } }),
    'multi-sheet': buildXlsx({
      sheets: [
        { name: 'One', rows: [['first sheet']] },
        { name: 'Two', rows: [['second sheet']] },
      ],
    }),
    landscape: buildXlsx({ rows: grid(3, 8), pageSetup: { orientation: 'landscape' } }),
    'date-1904': buildXlsx({
      rows: [[{ value: 40000, styleIndex: 6 }]],
      stylesXml: STYLES,
      date1904: true,
    }),
    'sheet-chart': buildXlsx({
      rows: [
        ['A', 4],
        ['B', 9],
      ],
      sheetChart: { chartXml: BAR_CHART },
    }),
  };

  for (const [name, bytes] of Object.entries(cases)) {
    it(`xlsx: ${name}`, () => {
      expect(sha256(convertXlsxToPdfSync(bytes, { fonts }))).toMatchSnapshot();
    });
  }
});
