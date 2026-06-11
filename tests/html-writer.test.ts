import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { createConverter } from '@/core/converter/facade';
import { Ream } from '@/core/converter/ream';
import { ConversionLossError } from '@/core/ir';
import { htmlWriter, writeHtml } from '@/html/html-writer';
import { readDocx } from '@/word/docx-reader';

const decode = (b: Uint8Array) => new TextDecoder().decode(b);

describe('html writer (FlowDoc adapter)', () => {
  it('converts docx → flowed HTML with no fonts and no I/O', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:r><w:rPr><w:i/><w:u w:val="single"/><w:color w:val="FF0000"/></w:rPr><w:t>styled run</w:t></w:r></w:p>',
    );
    // No fonts anywhere in the options — html conversion must not need them.
    const html = decode(await Ream.parse(docx).convert('html'));
    expect(html.startsWith('<!DOCTYPE html>')).toBe(true);
    expect(html).toContain('<h1'); // outlineLvl 0 → h1, like tagged PDF
    expect(html).toContain('Title');
    expect(html).toContain('font-weight:700');
    expect(html).toContain('text-align:center');
    expect(html).toContain('font-style:italic');
    expect(html).toContain('text-decoration-line:underline');
    expect(html).toContain('color:#FF0000');
    expect(html).toContain('</html>');
  });

  it('materialized list markers ride along as text', async () => {
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const li = (t: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(li('first') + li('second'), { numberingXml });
    const html = decode(await Ream.parse(docx).convert('html'));
    // applyNumbering prepends "1.\t" / "2.\t"; the tab renders as a .tab gap.
    expect(html).toContain('1.<span class="tab"></span>');
    expect(html).toContain('2.<span class="tab"></span>');
    expect(html).toContain('second');
  });

  it('tables carry colspan/rowspan, shading and borders', async () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblBorders>' +
      '<w:top w:val="single" w:sz="6"/><w:bottom w:val="single" w:sz="6"/>' +
      '<w:left w:val="single" w:sz="6"/><w:right w:val="single" w:sz="6"/>' +
      '<w:insideH w:val="single" w:sz="6"/><w:insideV w:val="single" w:sz="6"/>' +
      '</w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      // row 1: [A spans 2 cols][B starts a vertical merge]
      '<w:tr>' +
      '<w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:fill="FFCC00"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '</w:tr>' +
      // row 2: [c][d][vMerge continue]
      '<w:tr>' +
      '<w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
      '</w:tr></w:tbl>';
    const html = decode(await Ream.parse(buildDocxFromBody(tbl)).convert('html'));
    expect(html).toContain('<table');
    expect(html).toContain('<colgroup>');
    expect(html).toContain('colspan="2"');
    expect(html).toContain('rowspan="2"');
    expect(html).toContain('background-color:#FFCC00');
    expect(html).toContain('border-top:');
    // The vMerge continuation cell must not be emitted as its own <td>.
    expect(html.match(/<td/g)!.length).toBe(4);
  });

  it('renders an xlsx grid as a table', async () => {
    const xlsx = buildXlsx([
      ['name', 'value'],
      ['answer', '42'],
    ]);
    const html = decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain('<table');
    expect(html).toContain('answer');
    expect(html).toContain('42');
  });

  it('reports headers/footers as a dropped loss; strict throws', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>body</w:t></w:r></w:p>' +
        '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>',
      { headerXml: '<w:p><w:r><w:t>running head</w:t></w:r></w:p>' },
    );
    const doc = Ream.parse(docx);
    const { bytes, losses } = await doc.convertWithReport('html');
    expect(decode(bytes)).toContain('body');
    expect(decode(bytes)).not.toContain('running head');
    expect(losses.some((l) => l.feature === 'headersFooters' && l.severity === 'dropped')).toBe(
      true,
    );
    await expect(doc.convert('html', { strict: true })).rejects.toThrow(ConversionLossError);
  });

  it('is deterministic and exposed as a flow adapter', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>same bytes</w:t></w:r></w:p>');
    const a = await Ream.parse(docx).convert('html');
    const b = await Ream.parse(docx).convert('html');
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);

    expect(htmlWriter.id).toBe('html');
    expect(htmlWriter.consumes).toBe('flow');
    const { doc } = readDocx(docx);
    const direct = htmlWriter.write(doc);
    expect(decode(direct.bytes)).toContain('same bytes');
  });

  it('escapes markup-significant characters', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>a &lt;b&gt; &amp; "c"</w:t></w:r></w:p>');
    const html = decode(await Ream.parse(docx).convert('html'));
    expect(html).toContain('a &lt;b&gt; &amp; "c"');
    expect(html).not.toContain('a <b>');
  });

  it('converts through the createConverter facade', async () => {
    const conv = createConverter();
    const { bytes, losses } = await conv.convert(
      buildDocxFromBody('<w:p><w:r><w:t>facade html</w:t></w:r></w:p>'),
      { to: 'html' },
    );
    expect(decode(bytes)).toContain('facade html');
    expect(losses).toEqual([]);
  });

  it('writeHtml works on a hand-built raw FlowDoc (resolves the cascade itself)', () => {
    const { doc } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
    // Replace the body with a RAW (unresolved) paragraph — the writer must
    // resolve it over the empty sheet exactly like the PDF layout would.
    const raw = {
      ...doc,
      body: [
        {
          kind: 'paragraph' as const,
          paragraph: {
            properties: { alignment: 'right' as const },
            runs: [{ text: 'raw tree', properties: { bold: true } }],
          },
        },
      ],
    };
    const html = decode(writeHtml(raw).bytes);
    expect(html).toContain('raw tree');
    expect(html).toContain('font-weight:700');
    expect(html).toContain('text-align:right');
  });
});
