import { describe, expect, it } from 'vitest';

import { buildDocxFromBody, buildRichDocx } from './fixtures/build-docx';
import { OpcPackage } from '@/opc';
import { parseDocument } from '@/ooxml/wordproc';

function parse(paragraphs: Parameters<typeof buildRichDocx>[0]) {
  const docx = buildRichDocx(paragraphs);
  const pkg = OpcPackage.open(docx);
  return parseDocument(pkg.getMainDocument().data);
}

function paragraphs(body: ReturnType<typeof parse>) {
  return body.filter((b) => b.kind === 'paragraph').map((b) => b.paragraph);
}

describe('parseDocument: end-to-end', () => {
  it('extracts plain paragraphs with no properties', () => {
    const body = parse([{ runs: [{ text: 'Hello' }] }, { runs: [{ text: 'World' }] }]);
    expect(paragraphs(body)).toEqual([
      { properties: {}, runs: [{ text: 'Hello', properties: {} }] },
      { properties: {}, runs: [{ text: 'World', properties: {} }] },
    ]);
  });

  it('extracts run properties (bold + size + color) on the right runs', () => {
    const body = parse([
      {
        runs: [
          { text: 'Plain ' },
          {
            text: 'Bold',
            rPrXml: '<w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="ff0000"/></w:rPr>',
          },
          { text: ' tail' },
        ],
      },
    ]);
    expect(paragraphs(body)).toEqual([
      {
        properties: {},
        runs: [
          { text: 'Plain ', properties: {} },
          {
            text: 'Bold',
            properties: { bold: true, fontSizeHalfPoints: 28, colorHex: 'FF0000' },
          },
          { text: ' tail', properties: {} },
        ],
      },
    ]);
  });

  it('extracts paragraph properties (alignment, style, spacing)', () => {
    const body = parse([
      {
        pPrXml:
          '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/>' +
          '<w:spacing w:before="240" w:after="120"/></w:pPr>',
        runs: [{ text: 'Title' }],
      },
    ]);
    expect(paragraphs(body)[0]!.properties).toEqual({
      styleId: 'Heading1',
      alignment: 'center',
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
    });
  });

  it('combines paragraph and run properties on the same paragraph', () => {
    const body = parse([
      {
        pPrXml: '<w:pPr><w:jc w:val="right"/></w:pPr>',
        runs: [{ text: 'Right ', rPrXml: '<w:rPr><w:i/></w:rPr>' }, { text: 'aligned' }],
      },
    ]);
    expect(paragraphs(body)[0]).toEqual({
      properties: { alignment: 'right' },
      runs: [
        { text: 'Right ', properties: { italic: true } },
        { text: 'aligned', properties: {} },
      ],
    });
  });

  it('preserves Cyrillic text through the parser', () => {
    const body = parse([{ runs: [{ text: 'Привет, мир!' }] }]);
    expect(paragraphs(body)[0]!.runs[0]!.text).toBe('Привет, мир!');
  });
});

// §17.13.5 — revision tracking. We render the "accepted" / final document:
// inserted (w:ins) and moved-in (w:moveTo) content shows; deleted (w:del) and
// moved-out (w:moveFrom) content is dropped (its text lives in w:delText).
// Regression for the POI-corpus delins.docx, which lost all 67 w:ins runs
// (474 → 1326 extracted chars after the fix).
describe('parseDocument: tracked changes (accept-all)', () => {
  it('shows w:ins / w:moveTo runs and omits w:del / w:moveFrom runs', () => {
    const A = `w:author="a" w:date="2020-01-01T00:00:00Z"`;
    const body =
      `<w:p>` +
      `<w:r><w:t xml:space="preserve">Keep </w:t></w:r>` +
      `<w:ins w:id="1" ${A}><w:r><w:t>ins</w:t></w:r></w:ins>` +
      `<w:del w:id="2" ${A}><w:r><w:delText>DEL</w:delText></w:r></w:del>` +
      `<w:moveTo w:id="3" ${A}><w:r><w:t xml:space="preserve"> mv</w:t></w:r></w:moveTo>` +
      `<w:moveFrom w:id="4" ${A}><w:r><w:delText>MF</w:delText></w:r></w:moveFrom>` +
      `</w:p>`;
    const els = parseDocument(OpcPackage.open(buildDocxFromBody(body)).getMainDocument().data);
    expect(els).toHaveLength(1);
    const para = els[0]!.kind === 'paragraph' ? els[0]!.paragraph : null;
    const text = (para?.runs ?? []).map((r) => r.text).join('');
    expect(text).toBe('Keep ins mv'); // ins + moveTo kept, del + moveFrom dropped
    expect(text).not.toContain('DEL');
    expect(text).not.toContain('MF');
  });
});

// §17.3.3.1 — a forced page break vs a soft line break.
describe('parseDocument: page break', () => {
  it('flags w:br w:type="page" as a page break; plain w:br stays a newline', () => {
    const body = `<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t><w:br w:type="page"/></w:r></w:p>`;
    const els = parseDocument(OpcPackage.open(buildDocxFromBody(body)).getMainDocument().data);
    const para = els[0]!.kind === 'paragraph' ? els[0]!.paragraph : null;
    const run = para!.runs[0]!;
    expect(run.text).toBe('A\nB'); // plain w:br → newline (soft break)
    expect(run.pageBreak).toBe(true); // w:br w:type="page" → page-break flag
  });
});
