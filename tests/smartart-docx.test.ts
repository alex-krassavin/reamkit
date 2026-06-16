// E-SMARTART SA2 — a SmartArt diagram inline in a docx paragraph. The reader
// follows the data relationship (dgm:relIds @r:dm) → diagrams/data1.xml →
// diagrams/drawing1.xml and renders the override's dsp:sp nodes as floating
// shapes anchored to the paragraph's column origin.

import { zipSync } from 'fflate';

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

function smartArtDocx(withOverride = true): Uint8Array {
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
    `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></dsp:spPr>` +
    `<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
  const drawingXml =
    `<?xml version="1.0"?>\n<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A_NS}"><dsp:spTree>` +
    node('NodeA', 0, '4472C4') +
    node('NodeB', 2743200, 'ED7D31') +
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

describe('SmartArt diagrams in docx (E-SMARTART SA2)', () => {
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

  it('keeps the paragraph (no shapes) when no drawing override ships', () => {
    const doc = Ream.parse(smartArtDocx(false));
    expect(doc.flow.body.filter((e) => e.kind === 'shape')).toHaveLength(0);
  });
});
