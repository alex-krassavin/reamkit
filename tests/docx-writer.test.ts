import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
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

  it('reports unwritten block kinds as losses (v0)', () => {
    const tbl =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc: flow } = readDocx(buildDocxFromBody(tbl));
    const { losses } = writeDocx(flow);
    expect(losses.some((l) => l.feature === 'tables' && l.severity === 'dropped')).toBe(true);
  });

  it('strict mode surfaces writer losses', async () => {
    const tbl =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    await expect(
      Ream.parse(buildDocxFromBody(tbl)).convert('docx', { strict: true }),
    ).rejects.toThrow();
  });
});
