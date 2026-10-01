import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { createConverter } from '@/core/converter/facade';
import { Ream } from '@/core/converter/ream';
import { ConversionLossError, pt } from '@/core/ir';
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
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
      '<w:lvlText w:val="%1."/></w:lvl>' +
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

  it("lays each sheet out whole, under its tab's name", async () => {
    // Thirty 20-character columns are some 4000pt: printed, the sheet is cut
    // into column bands a page wide, and the page breaks between them and
    // between the sheets. On a screen it is one table, and every tab shows —
    // the empty one too.
    const wide = Array.from({ length: 30 }, (_, i) => `c${String(i + 1)}`);
    const xlsx = buildXlsx({
      sheets: [
        { name: 'Wide', rows: [wide], columns: [{ min: 1, max: 30, widthChars: 20 }] },
        { name: 'Empty', rows: [] },
        { name: 'R&D', rows: [['x']] },
      ],
    });
    const html = decode(await Ream.parse(xlsx).convert('html'));
    expect(
      [...html.matchAll(/<section class="sheet" data-sheet="([^"]*)">/gu)].map((m) => m[1]),
    ).toEqual(['Wide', 'Empty', 'R&amp;D']);
    const wideSheet = html.slice(
      html.indexOf('data-sheet="Wide"'),
      html.indexOf('data-sheet="Empty"'),
    );
    expect(wideSheet.match(/<table/gu)).toHaveLength(1);
    expect(wideSheet).toContain('c30');
    expect(html).not.toContain('break-before:page');
    expect(html).toContain('<h2 class="sheet-name">R&amp;D</h2>');
    // No page, so no page-wide column for the sheet to be squeezed into.
    expect(html).toContain('<article>');
  });

  it("draws the window's gridlines, but not over a fill nor where the sheet hides them", async () => {
    const STYLES = `
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      stylesXml: STYLES,
      sheets: [
        { name: 'Grid', rows: [['plain', { value: 'filled', styleIndex: 1 }]] },
        { name: 'Clean', rows: [['plain']], hideGridLines: true },
      ],
    });
    const html = decode(await Ream.parse(xlsx).convert('html'));
    const sheet = (name: string): string =>
      html.slice(html.indexOf(`data-sheet="${name}"`)).split('</section>')[0]!;
    // Three edges round the plain cell: the fourth is the yellow one's, and a
    // fill covers the gridlines on every side of it.
    expect(sheet('Grid').match(/0\.5pt solid #D4D4D4/gu)).toHaveLength(3);
    expect(
      /<td style="([^"]*)">\s*<p[^>]*><span[^>]*>filled/u.exec(sheet('Grid'))?.[1],
    ).not.toContain('D4D4D4');
    expect(sheet('Clean')).not.toContain('D4D4D4');
  });

  it('flags a noted cell in its corner, and shows the note on hover', async () => {
    const xlsx = buildXlsx({
      rows: [['noted', 'plain']],
      comments: [{ ref: 'A1', author: 'Ada', text: 'look here' }],
    });
    const html = decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toMatch(/<td title="Ada: look here" style="[^"]*position:relative/u);
    expect(html).toContain('border-top:4.5pt solid #FF0000');
    expect(html.match(/<td title=/gu)).toHaveLength(1);
  });

  it('leaves an edge to the neighbour that rules it', async () => {
    // A gridline and a thin rule come out the same pixel wide, and collapsed
    // borders then go to the cell above or to the left: a cell ruled all
    // round lost its top and left sides to the grey grid next to them.
    const STYLES = `
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="2"><border/><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      stylesXml: STYLES,
      rows: [
        ['', 'above'],
        ['beside', { value: 'ruled', styleIndex: 1 }],
      ],
    });
    const html = decode(await Ream.parse(xlsx).convert('html'));
    const styleOf = (word: string): string =>
      new RegExp(`<td style="([^"]*)">\\s*<p[^>]*><span[^>]*>${word}<`, 'u').exec(html)?.[1] ?? '';
    expect(styleOf('ruled')).toContain('border-top:0.75pt solid #000000');
    expect(styleOf('above')).not.toContain('border-bottom');
    expect(styleOf('beside')).not.toContain('border-right');
    // The edges nobody rules keep their gridline.
    expect(styleOf('above')).toContain('border-top:0.5pt solid #D4D4D4');
  });

  it('sets a grid row at its height and a cell at the bottom of it', async () => {
    const xlsx = buildXlsx({
      rows: [['tall'], ['plain']],
      rowHeights: [{ row: 0, heightPt: 30, customHeight: true }],
      columns: [{ min: 1, max: 1, widthChars: 10 }],
    });
    const html = decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain('<tr style="height:30pt">');
    expect(html).toContain('<tr style="height:15pt">');
    expect(html).toMatch(/<td style="[^"]*vertical-align:bottom/u);
    // A fixed table is given the width of its grid, or a browser ignores the
    // fixed layout and sizes the columns to their content.
    expect(html).toMatch(/<table style="width:[\d.]+pt;table-layout:fixed">/u);
  });

  it("places a sheet's drawing on its surface where it is anchored", () => {
    const { doc } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
    const html = decode(
      writeHtml({
        ...doc,
        sections: [
          { properties: { headers: [], footers: [] }, endIndex: 2, sheet: { name: 'Plot' } },
        ],
        body: [
          {
            kind: 'shape' as const,
            shape: {
              float: {
                wrap: 'none' as const,
                posH: { relativeFrom: 'margin' as const, offsetPt: pt(100) },
                posV: { relativeFrom: 'margin' as const, offsetPt: pt(50) },
              },
              width: pt(80),
              height: pt(40),
              geometry: { kind: 'preset' as const, preset: 'rect' },
              fill: { kind: 'solid' as const, colorHex: 'FF0000' },
              paragraphProperties: {},
            },
          },
          {
            kind: 'paragraph' as const,
            paragraph: { properties: {}, runs: [{ text: 'under it', properties: {} }] },
          },
        ],
      }).bytes,
    );
    expect(html).toContain(
      '<div class="drawing" style="position:absolute;left:100pt;top:50pt;width:80pt;height:40pt">',
    );
    // The surface reaches the drawing's far corner, so the next sheet starts below it.
    expect(html).toContain('<div class="surface" style="min-width:180pt;min-height:90pt">');
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

  it('emits hyperlinks as <a> with a scheme allowlist', async () => {
    const docx = buildDocxFromBody(
      '<w:p>' +
        '<w:hyperlink r:id="rId30"><w:r><w:t>safe link</w:t></w:r></w:hyperlink>' +
        '<w:hyperlink r:id="rId31"><w:r><w:t>evil link</w:t></w:r></w:hyperlink>' +
        '<w:hyperlink w:anchor="bm1"><w:r><w:t>internal</w:t></w:r></w:hyperlink>' +
        '</w:p>',
      {
        hyperlinks: {
          rId30: 'https://reamkit.dev/?a=1&b=2',
          rId31: 'javascript:alert(1)',
        },
      },
    );
    const { bytes, losses } = await Ream.parse(docx).convertWithReport('html');
    const html = decode(bytes);
    expect(html).toContain('<a href="https://reamkit.dev/?a=1&amp;b=2">');
    expect(html).toContain('safe link');
    expect(html).toContain('evil link'); // the text survives, the link does not
    expect(html).not.toContain('javascript:');
    expect(html).toContain('internal'); // anchor-only hyperlink → plain text
    expect(losses.some((l) => l.feature === 'hyperlinks' && l.severity === 'degraded')).toBe(true);
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

  it('sets a right-to-left paragraph to the sides its start and end are', () => {
    // §17.3.1.13 — in a `w:bidi` paragraph "right" is the END of the line,
    // which is the left of the page, and the indent named "left" is the one at
    // the start, on the right: the layout crosses them over, and so must CSS,
    // whose sides are the page's whatever `dir` says.
    const { doc } = readDocx(
      buildDocxFromBody(
        '<w:p><w:pPr><w:bidi/><w:ind w:left="720"/><w:jc w:val="right"/></w:pPr><w:r><w:t>שלום</w:t></w:r></w:p>',
      ),
    );
    const html = decode(writeHtml(doc).bytes);
    expect(html).toMatch(/<p dir="rtl" style="[^"]*text-align:left/u);
    expect(html).toMatch(/<p dir="rtl" style="[^"]*margin-right:36pt/u);
    expect(html).not.toMatch(/<p dir="rtl" style="[^"]*margin-left:36pt/u);
  });

  it("keeps the whole of a run's style when the run names its font", async () => {
    // The family is a CSS string inside a double-quoted attribute: written in
    // double quotes it ended the attribute, and a browser dropped the size,
    // weight and colour after it.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic"/><w:b/>' +
        '<w:color w:val="4E5B6F"/><w:sz w:val="50"/></w:rPr><w:t>Budget</w:t></w:r></w:p>',
    );
    const html = decode(await Ream.parse(docx).convert('html'));
    expect(/<span style="([^"]*)">Budget<\/span>/u.exec(html)?.[1]).toBe(
      "font-family:'Century Gothic',sans-serif;font-size:25pt;font-weight:700;color:#4E5B6F",
    );
  });

  it('follows a family with its metric twin and its class, and escapes the name', () => {
    const { doc } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
    const run = (text: string, ascii: string) => ({ text, properties: { fontFamily: { ascii } } });
    const html = decode(
      writeHtml({
        ...doc,
        body: [
          {
            kind: 'paragraph' as const,
            paragraph: {
              properties: {},
              runs: [
                run('calibri', 'Calibri'),
                run('cambria', 'Cambria'),
                run('courier', 'Courier New'),
                run('quoted', 'O\'Brien "Sans"'),
              ],
            },
          },
        ],
      }).bytes,
    );
    const family = (text: string): string | undefined =>
      new RegExp(`<span style="font-family:([^"]*?);font-size:[^"]*">${text}</span>`, 'u').exec(
        html,
      )?.[1];
    expect(family('calibri')).toBe("'Calibri',Carlito,sans-serif");
    expect(family('cambria')).toBe("'Cambria',Caladea,serif");
    expect(family('courier')).toBe("'Courier New',monospace");
    // A quote of either kind stays inside the name: escaped for CSS, then for the attribute.
    expect(family('quoted')).toBe("'O\\'Brien &quot;Sans&quot;',sans-serif");
  });

  // ── charts and shapes as inline SVG ──────────────────────────────────────

  const C_NS =
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

  const chartDrawing = (rId: string): string => `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="5486400" cy="3200400"/>
      <wp:docPr id="1" name="Chart 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="${rId}"/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;

  const BAR_CHART = `<c:chartSpace ${C_NS}>
    <c:chart>
      <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:plotArea><c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart></c:plotArea>
    </c:chart>
  </c:chartSpace>`;

  const PIE_CHART = `<c:chartSpace ${C_NS}>
    <c:chart><c:plotArea><c:pieChart>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
    </c:pieChart></c:plotArea></c:chart>
  </c:chartSpace>`;

  const shapeDrawing = (spPrInner: string, body = ''): string => `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="1828800" cy="914400"/>
      <wp:docPr id="2" name="Shape 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>${spPrInner}</wps:spPr>
            ${body}<wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;

  it('renders a bar chart as inline SVG with anchored labels', async () => {
    const docx = buildDocxFromBody(chartDrawing('rId5'), { charts: { rId5: BAR_CHART } });
    const html = decode(await Ream.parse(docx).convert('html'));
    expect(html).toContain('<svg viewBox="0 0 432 252"'); // 5486400/3200400 EMU
    expect(html).toContain('aria-label="Quarterly Sales"');
    expect(html).toContain('fill="#4472C4"'); // series bars
    expect(html).toMatch(/<rect [^>]*fill="#4472C4"/);
    expect(html).toContain('>Q1</text>'); // category label, browser-rendered
    expect(html).toContain('text-anchor');
    // The graphics flip wrapper (y-up scene → y-down viewport).
    expect(html).toContain('<g transform="matrix(1 0 0 -1 0 252)">');
  });

  it('renders pie wedges as bezier paths', async () => {
    const docx = buildDocxFromBody(chartDrawing('rId6'), { charts: { rId6: PIE_CHART } });
    const html = decode(await Ream.parse(docx).convert('html'));
    expect(html).toMatch(/<path d="M [^"]*C [^"]*Z" fill="#/); // center→arc→close
  });

  it('a chart without its part is a dropped loss', async () => {
    const docx = buildDocxFromBody(chartDrawing('rId9'));
    const flow = Ream.parse(docx).flow;
    const { losses, bytes } = writeHtml(flow);
    expect(losses.some((l) => l.feature === 'charts' && l.severity === 'dropped')).toBe(true);
    expect(decode(bytes)).not.toContain('<svg');
  });

  it('renders shape geometry with fill, stroke and rotation', async () => {
    const spPr =
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>' +
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>';
    const html = decode(await Ream.parse(buildDocxFromBody(shapeDrawing(spPr))).convert('html'));
    expect(html).toContain('<svg viewBox="0 0 144 72"');
    expect(html).toMatch(/<path d="M [^"]*Z" fill="#4472C4"[^>]*stroke="#2F528F" stroke-width="1"/);
    expect(html).toContain('transform="matrix(');
  });

  it('overlays text-box content inside the shape with its anchor', async () => {
    const spPr = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    const body =
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed text</w:t></w:r></w:p></w:txbxContent></wps:txbx>';
    const withAnchor = shapeDrawing(spPr, body).replace(
      '<wps:bodyPr/>',
      '<wps:bodyPr anchor="ctr"/>',
    );
    const html = decode(await Ream.parse(buildDocxFromBody(withAnchor)).convert('html'));
    expect(html).toContain('position:relative;width:144pt;height:72pt');
    expect(html).toContain('justify-content:center');
    expect(html).toContain('boxed text');
  });
});
