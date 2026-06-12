import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { Ream } from '@/core/converter/ream';
import { OpcPackage } from '@/core/opc';
import { writeDocx } from '@/word/docx-writer';
import { readDocx } from '@/word/docx-reader';

const decode = (b: Uint8Array) => new TextDecoder().decode(b);

const BODY =
  '<w:p><w:r><w:t>first paragraph</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">spaced  &amp; escaped &lt;text&gt;</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/></w:sectPr>';

describe('docx writer (E-DOCX D2 skeleton)', () => {
  it('round-trips text and page geometry through its own reader', () => {
    const { doc: flow } = readDocx(buildDocxFromBody(BODY));
    const { bytes, losses } = writeDocx(flow);
    expect(losses).toHaveLength(0);

    const { doc: again } = readDocx(bytes);
    const texts = again.body.flatMap((el) =>
      el.kind === 'paragraph' ? [el.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    expect(texts).toEqual(['first paragraph', 'spaced  & escaped <text>']);
    // Page geometry survives (A4 portrait, custom margins).
    const sect = again.sections[0]?.properties ?? again.section;
    expect(sect?.pageSize?.width).toBeCloseTo(11906 / 20, 1);
    expect(sect?.margins?.left).toBeCloseTo(54, 1);
  });

  it('is exposed as the docx target on Ream.convert', async () => {
    const src = buildDocxFromBody('<w:p><w:r><w:t>via facade</w:t></w:r></w:p>');
    const out = await Ream.parse(src).convert('docx');
    const pkg = OpcPackage.open(out);
    expect(decode(pkg.getMainDocument().data)).toContain('via facade');
  });

  it('is deterministic', () => {
    const { doc: flow } = readDocx(buildDocxFromBody(BODY));
    expect([...writeDocx(flow).bytes]).toEqual([...writeDocx(flow).bytes]);
  });

  it('round-trips run formatting (bold/italic/size/color/valign) as direct rPr', () => {
    const body =
      '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/>' +
      '<w:ind w:left="720" w:firstLine="360"/></w:pPr>' +
      '<w:r><w:rPr><w:b/><w:i/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr>' +
      '<w:t>styled</w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r></w:p>';
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);

    const para = again.body[0];
    if (para?.kind !== 'paragraph') throw new Error('expected paragraph');
    const pPr = para.paragraph.properties as Record<string, unknown>;
    expect(pPr.alignment).toBe('center');
    expect(pPr.spacingBefore).toBeCloseTo(12, 1); // 240 twips
    expect(pPr.spacingAfter).toBeCloseTo(6, 1);
    expect(pPr.indentLeft).toBeCloseTo(36, 1); // 720 twips
    expect(pPr.indentFirstLine).toBeCloseTo(18, 1);

    const r0 = para.paragraph.runs[0]!.properties as Record<string, unknown>;
    expect(r0.bold).toBe(true);
    expect(r0.italic).toBe(true);
    expect(r0.fontSizePt).toBeCloseTo(14, 1); // sz=28 half-points
    expect(r0.colorHex).toBe('FF0000');
    const r1 = para.paragraph.runs[1]!.properties as Record<string, unknown>;
    expect(r1.verticalAlign).toBe('superscript');
  });

  it('omits default-valued properties (no rPr/pPr noise for plain text)', () => {
    const { doc: flow } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>plain</w:t></w:r></w:p>'));
    const xml = new TextDecoder().decode(
      OpcPackage.open(writeDocx(flow).bytes).getMainDocument().data,
    );
    expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">plain</w:t></w:r></w:p>');
    expect(xml).not.toContain('<w:rPr>');
    expect(xml).not.toContain('<w:pPr>');
  });

  it('round-trips a numbered list: numPr + numbering.xml, markers regenerated', () => {
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
      '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>' +
      '<w:lvlText w:val="%2)"/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const li = (ilvl: number, t: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${t}</w:t></w:r></w:p>`;
    const body = li(0, 'first') + li(0, 'second') + li(1, 'nested');
    const { doc: flow } = readDocx(buildDocxFromBody(body, { numberingXml }));
    const { bytes } = writeDocx(flow);

    // The package carries numbering.xml.
    const pkg = OpcPackage.open(bytes);
    expect(pkg.getPart('word/numbering.xml')).toBeDefined();

    const { doc: again } = readDocx(bytes);
    const paras = again.body.flatMap((el) => (el.kind === 'paragraph' ? [el.paragraph] : []));
    // Each paragraph keeps its list reference (numId/ilvl).
    expect(paras[0]!.properties.numbering).toMatchObject({ numId: '1', ilvl: 0 });
    expect(paras[2]!.properties.numbering).toMatchObject({ numId: '1', ilvl: 1 });
    // Markers regenerate identically — and exactly once (no doubling).
    expect(paras[0]!.runs.map((r) => r.text).join('')).toBe('1.\tfirst');
    expect(paras[2]!.runs.map((r) => r.text).join('')).toBe('a)\tnested');
  });

  it('omits numbering.xml when the document has no lists', () => {
    const { doc: flow } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>plain</w:t></w:r></w:p>'));
    const pkg = OpcPackage.open(writeDocx(flow).bytes);
    expect(pkg.getPart('word/numbering.xml')).toBeUndefined();
  });

  it('round-trips an external hyperlink: w:hyperlink r:id + external rel', () => {
    const body =
      '<w:p><w:r><w:t>see </w:t></w:r>' +
      '<w:hyperlink r:id="rId50"><w:r><w:t>the site</w:t></w:r></w:hyperlink>' +
      '<w:r><w:t> now</w:t></w:r></w:p>';
    const { doc: flow } = readDocx(
      buildDocxFromBody(body, { hyperlinks: { rId50: 'https://reamkit.dev' } }),
    );
    const { bytes } = writeDocx(flow);
    const { doc: again } = readDocx(bytes);

    const para = again.body[0];
    if (para?.kind !== 'paragraph') throw new Error('expected paragraph');
    const linked = para.paragraph.runs.find((r) => r.text === 'the site');
    expect(linked?.href).toBe('https://reamkit.dev');
    // The surrounding runs are not linked.
    expect(para.paragraph.runs.find((r) => r.text === 'see ')?.href).toBeUndefined();
  });

  it('round-trips an internal anchor link: w:hyperlink w:anchor (no rel)', () => {
    const body =
      '<w:p><w:hyperlink w:anchor="target"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>' +
      '<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>here</w:t></w:r>' +
      '<w:bookmarkEnd w:id="1"/></w:p>';
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);

    const linkPara = again.body[0];
    if (linkPara?.kind !== 'paragraph') throw new Error('expected paragraph');
    expect(linkPara.paragraph.runs[0]!.anchor).toBe('target');
    expect(linkPara.paragraph.runs[0]!.href).toBeUndefined();
    // The bookmark survives onto its paragraph.
    const targetPara = again.body[1];
    if (targetPara?.kind !== 'paragraph') throw new Error('expected paragraph');
    expect(targetPara.paragraph.bookmarks).toContain('target');
  });

  it('numbering and hyperlink relationships share one rId space', () => {
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/>' +
      '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const body =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      '<w:r><w:t>item</w:t></w:r></w:p>' +
      '<w:p><w:hyperlink r:id="rId50"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>';
    const { doc: flow } = readDocx(
      buildDocxFromBody(body, { numberingXml, hyperlinks: { rId50: 'https://x.test' } }),
    );
    const { doc: again } = readDocx(writeDocx(flow).bytes);
    // Both survive — distinct rIds in document.xml.rels.
    const linkPara = again.body[1];
    expect(linkPara?.kind === 'paragraph' && linkPara.paragraph.runs[0]!.href).toBe(
      'https://x.test',
    );
    const listPara = again.body[0];
    expect(listPara?.kind === 'paragraph' && listPara.paragraph.properties.numbering?.numId).toBe(
      '1',
    );
  });

  it('round-trips a table: grid, borders, shading, header row', () => {
    const body =
      '<w:tbl><w:tblPr>' +
      '<w:tblBorders><w:top w:val="single" w:sz="8" w:color="FF0000"/>' +
      '<w:insideH w:val="single" w:sz="4"/></w:tblBorders>' +
      '</w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
      '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>';
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);

    const tbl = again.body.find((el) => el.kind === 'table');
    if (tbl?.kind !== 'table') throw new Error('expected table');
    const t = tbl.table;
    expect(t.grid.map((w) => Math.round(w))).toEqual([100, 150]); // 2000/3000 twips
    expect(t.properties.borders?.top).toMatchObject({ style: 'single', colorHex: 'FF0000' });
    expect(t.properties.borders?.top?.width).toBeCloseTo(1, 1); // sz=8 eighths
    expect(t.rows[0]!.properties.isHeader).toBe(true);
    expect(t.rows[0]!.cells[0]!.properties.shading?.colorHex).toBe('4472C4');
    const cellText = (r: number, c: number) => {
      const cell = t.rows[r]!.cells[c]!.content[0]!;
      return cell.kind === 'paragraph' ? cell.paragraph.runs.map((x) => x.text).join('') : '';
    };
    expect(cellText(0, 0)).toBe('H1');
    expect(cellText(1, 1)).toBe('b');
  });

  it('round-trips column/row spans (gridSpan + vMerge)', () => {
    const body =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>';
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);
    const tbl = again.body.find((el) => el.kind === 'table');
    if (tbl?.kind !== 'table') throw new Error('expected table');
    const t = tbl.table;
    expect(t.rows[0]!.cells[0]!.properties.colSpan).toBe(2);
    expect(t.rows[1]!.cells[0]!.properties.merge).toBe('start');
    expect(t.rows[2]!.cells[0]!.properties.merge).toBe('end');
  });

  it('round-trips a nested table (table inside a cell)', () => {
    const inner =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const body =
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr><w:tc>${inner}<w:p/></w:tc></w:tr></w:tbl>`;
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);
    const outer = again.body.find((el) => el.kind === 'table');
    if (outer?.kind !== 'table') throw new Error('expected table');
    const nested = outer.table.rows[0]!.cells[0]!.content.find((el) => el.kind === 'table');
    expect(nested?.kind).toBe('table');
    if (nested?.kind !== 'table') throw new Error('expected nested table');
    const innerCell = nested.table.rows[0]!.cells[0]!.content[0]!;
    expect(innerCell.kind === 'paragraph' && innerCell.paragraph.runs[0]!.text).toBe('inner');
  });

  it('round-trips an image: media part + blip rel, dimensions and alt text', () => {
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
    const drawing =
      '<w:r><w:drawing><wp:inline ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="914400" cy="685800"/><wp:docPr id="1" name="Pic" descr="a red square"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="685800"/></a:xfrm></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const { doc: flow } = readDocx(
      buildDocxFromBody(`<w:p>${drawing}</w:p>`, {
        images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
      }),
    );
    const { bytes, losses } = writeDocx(flow);
    expect(losses).toHaveLength(0);

    const pkg = OpcPackage.open(bytes);
    // The media part is present and the bytes survive verbatim.
    expect(pkg.getPart('word/media/image1.png')).toBeDefined();
    expect([...pkg.getPart('word/media/image1.png')!]).toEqual([...png]);

    const { doc: again } = readDocx(bytes);
    const el = again.body.find((b) => b.kind === 'image');
    if (el?.kind !== 'image') throw new Error('expected image block');
    expect(el.image.resource).toBeDefined(); // resolved to a stored resource
    expect(el.image.width).toBeCloseTo(72, 1); // 914400 EMU = 72pt
    expect(el.image.height).toBeCloseTo(54, 1); // 685800 EMU
    expect(el.image.altText).toBe('a red square');
  });

  it('deduplicates a resource used twice into one media part', () => {
    const png = buildTinyPng(1, 1, [0, 255, 0, 255]);
    const d = (rId: string) =>
      `<w:r><w:drawing><wp:inline ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="P"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill>` +
      `<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const { doc: flow } = readDocx(
      buildDocxFromBody(`<w:p>${d('rId20')}</w:p><w:p>${d('rId21')}</w:p>`, {
        images: {
          rId20: { contentType: 'image/png', bytes: png, extension: 'png' },
          rId21: { contentType: 'image/png', bytes: png, extension: 'png' },
        },
      }),
    );
    const pkg = OpcPackage.open(writeDocx(flow).bytes);
    // Same bytes ⇒ one content-addressed resource ⇒ a single media part.
    expect(pkg.getPart('word/media/image1.png')).toBeDefined();
    expect(pkg.getPart('word/media/image2.png')).toBeUndefined();
  });

  it('round-trips a header and footer through their parts', () => {
    const body =
      '<w:p><w:r><w:t>body</w:t></w:r></w:p>' +
      '<w:sectPr>' +
      '<w:headerReference w:type="default" r:id="rId10"/>' +
      '<w:footerReference w:type="default" r:id="rId11"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
    const { doc: flow } = readDocx(
      buildDocxFromBody(body, {
        headerXml: '<w:p><w:r><w:t>page header</w:t></w:r></w:p>',
        footerXml: '<w:p><w:r><w:t>page footer</w:t></w:r></w:p>',
      }),
    );
    const { bytes } = writeDocx(flow);
    const pkg = OpcPackage.open(bytes);
    expect(decode(pkg.getPart('word/header1.xml'))).toContain('page header');
    expect(decode(pkg.getPart('word/footer1.xml'))).toContain('page footer');

    // Re-read resolves the parts through their relationships.
    const { doc: again } = readDocx(bytes);
    const sect = again.sections[0]?.properties ?? again.section;
    const headerRel = sect?.headers.find((h) => h.type === 'default')?.relationshipId;
    expect(headerRel).toBeDefined();
    const headerContent = again.headersFooters?.get(headerRel!);
    const headerText =
      headerContent?.[0]?.kind === 'paragraph'
        ? headerContent[0].paragraph.runs.map((r) => r.text).join('')
        : '';
    expect(headerText).toBe('page header');
  });

  it('round-trips multi-column geometry (w:cols) and titlePg', () => {
    const body =
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:cols w:num="2" w:space="708"/><w:titlePg/></w:sectPr>';
    const { doc: flow } = readDocx(buildDocxFromBody(body));
    const { doc: again } = readDocx(writeDocx(flow).bytes);
    const sect = again.sections[0]?.properties ?? again.section;
    expect(sect?.columns?.count).toBe(2);
    expect(sect?.columns?.spacePt).toBeCloseTo(708 / 20, 1);
    expect(sect?.titlePg).toBe(true);
  });

  // An inline picture with no bytes: still a clean dropped loss to assert on.
  const IMAGE_BODY =
    '<w:p><w:r><w:drawing><wp:inline ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Pic"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';

  it('reports unwritten block kinds as losses (v0)', () => {
    const { doc: flow } = readDocx(buildDocxFromBody(IMAGE_BODY));
    const { losses } = writeDocx(flow);
    expect(losses.some((l) => l.feature === 'images' && l.severity === 'dropped')).toBe(true);
  });

  it('strict mode surfaces writer losses', async () => {
    await expect(
      Ream.parse(buildDocxFromBody(IMAGE_BODY)).convert('docx', { strict: true }),
    ).rejects.toThrow();
  });
});
