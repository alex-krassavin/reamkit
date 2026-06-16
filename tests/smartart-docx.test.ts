// E-SMARTART SA2/SA3 — a SmartArt diagram inline in a docx paragraph. The reader
// follows the data relationship (dgm:relIds @r:dm) → diagrams/data1.xml →
// diagrams/drawing1.xml and renders the override's dsp:sp nodes as floating
// shapes anchored to the paragraph's column origin. SA3 adds scheme-colour
// resolution through the document theme, an explicit loss when no override
// ships, and an end-to-end render (PDF + HTML).

import { readFileSync } from 'node:fs';

import { zipSync } from 'fflate';

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const srgb = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
// accent1 scheme fill markup — for the SA3 theme-resolution test.
const SCHEME_ACCENT1 = `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`;

function smartArtDocx(
  withOverride = true,
  opts: { readonly fillA?: string; readonly theme?: string } = {},
): Uint8Array {
  const drawing =
    `<w:drawing><wp:inline xmlns:wp="${WP_NS}">` +
    `<wp:extent cx="5486400" cy="2743200"/><wp:docPr id="1" name="Diagram"/>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${DGM}">` +
    `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rId100"/>` +
    `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
    `<w:p><w:r>${drawing}</w:r></w:p></w:body></w:document>`;

  const node = (text: string, x: number, fill: string): string =>
    `<dsp:sp><dsp:spPr>` +
    `<a:xfrm><a:off x="${x}" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `${fill}</dsp:spPr>` +
    `<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
  const drawingXml =
    `<?xml version="1.0"?>\n<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A_NS}"><dsp:spTree>` +
    node('NodeA', 0, opts.fillA ?? srgb('4472C4')) +
    node('NodeB', 2743200, srgb('ED7D31')) +
    `</dsp:spTree></dsp:drawing>`;
  const rels = (inner: string): Uint8Array =>
    enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="${PKG}">${inner}</Relationships>`,
    );

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `</Types>`,
    ),
    '_rels/.rels': rels(
      `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`,
    ),
    'word/document.xml': enc.encode(document),
  };
  if (withOverride) {
    files['word/_rels/document.xml.rels'] = rels(
      `<Relationship Id="rId100" Type="${R_NS}/diagramData" Target="diagrams/data1.xml"/>`,
    );
    files['word/diagrams/data1.xml'] = enc.encode(
      `<?xml version="1.0"?>\n<dgm:dataModel xmlns:dgm="${DGM}"/>`,
    );
    files['word/diagrams/_rels/data1.xml.rels'] = rels(
      `<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/>`,
    );
    files['word/diagrams/drawing1.xml'] = enc.encode(drawingXml);
  }
  if (opts.theme) {
    files['word/theme/theme1.xml'] = enc.encode(
      `<?xml version="1.0"?>\n<a:theme xmlns:a="${A_NS}" name="doc"><a:themeElements>` +
        `<a:clrScheme name="doc">${opts.theme}</a:clrScheme></a:themeElements></a:theme>`,
    );
  }
  return zipSync(files);
}

function shapeText(shape: { text?: { content: ReadonlyArray<unknown> } }): string {
  return (shape.text?.content ?? [])
    .flatMap((c) =>
      (c as { kind: string; paragraph?: { runs: Array<{ text: string }> } }).kind === 'paragraph'
        ? (c as { paragraph: { runs: Array<{ text: string }> } }).paragraph.runs.map((r) => r.text)
        : [],
    )
    .join('');
}

describe('SmartArt diagrams in docx (E-SMARTART SA2/SA3)', () => {
  it('renders the drawing-override nodes as floating shapes', () => {
    const shapes = Ream.parse(smartArtDocx()).flow.body.filter((e) => e.kind === 'shape');
    const texts = shapes.map((s) => shapeText(s.shape));
    expect(texts).toContain('NodeA');
    expect(texts).toContain('NodeB');
  });

  it('anchors each node column-relative at its diagram offset', () => {
    const shapes = Ream.parse(smartArtDocx()).flow.body.filter((e) => e.kind === 'shape');
    const xs = shapes
      .map((s) => Math.round(s.shape.float?.posH?.offsetPt ?? -1))
      .sort((a, b) => a - b);
    // NodeA at 0; NodeB at 2743200 EMU = 216pt, both relative to the column.
    expect(xs).toEqual([0, 216]);
    expect(shapes.every((s) => s.shape.float?.posH?.relativeFrom === 'column')).toBe(true);
  });

  it('resolves a node scheme-colour fill through the document theme (SA3)', () => {
    // NodeA fills with accent1; the document theme maps accent1 → FF8800. The
    // shared ColorResolver that styles ordinary drawings styles diagrams too.
    const docx = smartArtDocx(true, {
      fillA: SCHEME_ACCENT1,
      theme: `<a:accent1><a:srgbClr val="FF8800"/></a:accent1>`,
    });
    const shapes = Ream.parse(docx).flow.body.filter((e) => e.kind === 'shape');
    const nodeA = shapes.find((s) => shapeText(s.shape) === 'NodeA');
    expect(nodeA?.shape.fill.colorHex).toBe('FF8800'); // the theme accent1, not the default 4472C4
  });

  it('flows the diagram nodes through to PDF and HTML (SA3 demo)', async () => {
    const docx = smartArtDocx();
    const html = new TextDecoder().decode(await Ream.parse(docx).convert('html'));
    expect(html).toContain('NodeA');
    expect(html).toContain('NodeB');

    const file = PdfFile.parse(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('')
      .replace(/\s/g, '');
    expect(text).toContain('NodeA');
    expect(text).toContain('NodeB');
  });

  it('keeps the paragraph and records a loss when no drawing override ships (SA3)', () => {
    const doc = Ream.parse(smartArtDocx(false));
    expect(doc.flow.body.filter((e) => e.kind === 'shape')).toHaveLength(0);
    // SA3: the dropped diagram is reported, not silently swallowed.
    const loss = doc.losses.find((l) => l.feature === 'shapes.smartArt');
    expect(loss?.severity).toBe('dropped');
  });
});
