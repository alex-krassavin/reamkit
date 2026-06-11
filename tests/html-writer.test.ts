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
