// Generate a starter validation corpus of docx/xlsx files into corpus/inputs/.
//
// Each document declares an explicit A4 page (so LibreOffice and our renderer
// agree on geometry) and requests the Roboto font (installed in the user font
// dir, so LibreOffice substitutes the same glyphs we embed). This keeps the
// visual diff meaningful rather than dominated by font/paper mismatches.

import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { buildDocxFromBody } from '../../tests/fixtures/build-docx';
import { buildXlsx } from '../../tests/fixtures/build-xlsx';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(here, '../../corpus/inputs');
mkdirSync(outDir, { recursive: true });

// A4 section with 1-inch margins.
const A4_SECT = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
const RF = `<w:rFonts w:ascii="Roboto" w:hAnsi="Roboto" w:cs="Roboto"/>`;

// A run that names Roboto so LibreOffice uses the same font we embed.
function run(text: string, extraRpr = ''): string {
  return `<w:r><w:rPr>${RF}${extraRpr}</w:rPr><w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
}
function para(text: string, pPr = '', extraRpr = ''): string {
  return `<w:p><w:pPr>${pPr}</w:pPr>${run(text, extraRpr)}</w:p>`;
}
function escape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.';

const DECIMAL_NUMBERING = `
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;

interface Doc {
  readonly name: string;
  readonly bytes: Uint8Array;
}

const docs: Array<Doc> = [];

// 1. Plain multi-paragraph text.
docs.push({
  name: 'text-basic.docx',
  bytes: buildDocxFromBody(
    `${para('Document Title', '<w:jc w:val="center"/>', '<w:b/><w:sz w:val="36"/>')}` +
      para(LOREM) +
      para(LOREM) +
      A4_SECT,
  ),
});

// 2. Inline styling: bold / italic / size / colour.
docs.push({
  name: 'text-styled.docx',
  bytes: buildDocxFromBody(
    `<w:p><w:pPr></w:pPr>` +
      run('Regular ') +
      run('bold ', '<w:b/>') +
      run('italic ', '<w:i/>') +
      run('big ', '<w:sz w:val="40"/>') +
      run('red', '<w:color w:val="CC0000"/>') +
      `</w:p>` +
      A4_SECT,
  ),
});

// 3. Justified paragraph (Knuth-Plass vs LO line breaking).
docs.push({
  name: 'text-justified.docx',
  bytes: buildDocxFromBody(para(LOREM + ' ' + LOREM, '<w:jc w:val="both"/>') + A4_SECT),
});

// 4. A simple bordered table.
docs.push({
  name: 'table-basic.docx',
  bytes: buildDocxFromBody(
    `<w:tbl>
      <w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/>
      </w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>${cell('Name', true)}${cell('Qty', true)}${cell('Price', true)}</w:tr>
      <w:tr>${cell('Widget')}${cell('10')}${cell('19.99')}</w:tr>
      <w:tr>${cell('Gadget')}${cell('3')}${cell('149.00')}</w:tr>
    </w:tbl>${A4_SECT}`,
  ),
});

// 5. Numbered + bullet list (needs numbering.xml).
docs.push({
  name: 'list-basic.docx',
  bytes: buildDocxFromBody(
    listPara('First item', '1', 0) +
      listPara('Second item', '1', 0) +
      listPara('Third item', '1', 0) +
      A4_SECT,
    { numberingXml: DECIMAL_NUMBERING },
  ),
});

// 6. Headers / footers.
docs.push({
  name: 'header-footer.docx',
  bytes: buildDocxFromBody(
    para('Body content on the page.') +
      `<w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:footerReference r:id="rId11" w:type="default"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`,
    {
      headerXml: para('HEADER'),
      footerXml: para('FOOTER'),
    },
  ),
});

// 7. xlsx grid.
docs.push({
  name: 'sheet-basic.xlsx',
  bytes: buildXlsx({
    rows: [
      ['Region', 'Q1', 'Q2', 'Total'],
      ['North', 100, 120, 220],
      ['South', 90, 110, 200],
      ['East', 130, 140, 270],
    ],
    columns: [
      { min: 1, max: 1, widthChars: 14 },
      { min: 2, max: 4, widthChars: 10 },
    ],
  }),
});

// 8. xlsx with number formats.
docs.push({
  name: 'sheet-formats.xlsx',
  bytes: buildXlsx({
    rows: [[{ value: 1234567, styleIndex: 1 }], [{ value: 0.875, styleIndex: 2 }]],
    stylesXml: `
      <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
      <fonts count="1"><font/></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="3" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>`,
  }),
});

function cell(text: string, bold = false): string {
  return `<w:tc><w:p><w:pPr></w:pPr>${run(text, bold ? '<w:b/>' : '')}</w:p></w:tc>`;
}

function listPara(text: string, numId: string, ilvl: number): string {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${run(text)}</w:p>`;
}

for (const doc of docs) {
  writeFileSync(resolve(outDir, doc.name), doc.bytes);
  console.log(`wrote corpus/inputs/${doc.name} (${doc.bytes.byteLength} bytes)`);
}
console.log(`\n${docs.length} documents written.`);
