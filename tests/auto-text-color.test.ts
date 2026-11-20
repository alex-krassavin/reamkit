// ECMA-376 Part 1 §17.3.2.6 — the AUTOMATIC text colour, which Word sets
// white over a dark background (issue #137: words naming no colour in a cell
// filled 25397D came out black). Every expectation below is what Word for Mac
// drew on our own probes (see src/core/style-cascade/auto-text-color.ts).

import { readFileSync } from 'node:fs';

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { TextToken } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { isDarkBehindAutoText } from '@/core/style-cascade';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const SHD = (fill: string, val = 'clear', color = 'auto') =>
  `<w:shd w:val="${val}" w:color="${color}" w:fill="${fill}"/>`;

/** A one-cell table: the cell's `tcPr` extras, then the paragraph inside it. */
const cell = (tcPr: string, paragraph: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>${tcPr}</w:tcPr>${paragraph}</w:tc></w:tr>` +
  '</w:tbl><w:p/>';

const para = (text: string, rPr = '', pPr = '') =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** The colour every word of the page is drawn in, by its text. */
function drawnColours(body: string, numberingXml?: string): Map<string, string> {
  const flow = Ream.parse(buildDocxFromBody(body, numberingXml ? { numberingXml } : {})).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out = new Map<string, string>();
  for (const c of laid.pages[0]!.commands) {
    if (c.type !== 'line') continue;
    for (const t of c.line.tokens) {
      if (t.kind === 'text' && !t.isSpace) out.set(t.text, t.resolvedRun.colorHex);
    }
  }
  return out;
}

const colourOf = (body: string, text: string, numberingXml?: string): string | undefined =>
  drawnColours(body, numberingXml).get(text);

describe('the background Word turns automatic text white on (§17.3.2.6)', () => {
  it('is one whose Rec. 601 luma is below 0x3C00 on a 16-bit scale', () => {
    // Each pair straddles the switch in Word: white on the first, black on the second.
    const pairs = [
      ['3B3B3B', '3C3C3C'],
      ['C70000', 'C80000'],
      ['006500', '006600'],
      ['0034FF', '0036FF'],
      ['6500FF', '6800FF'],
      ['661980', '681880'],
      ['263A7D', '253B7D'],
    ] as const;
    for (const [dark, light] of pairs) {
      expect(isDarkBehindAutoText(dark), dark).toBe(true);
      expect(isDarkBehindAutoText(light), light).toBe(false);
    }
    // The issue's fill, pure blue, and two that Word keeps black on.
    expect(isDarkBehindAutoText('25397D')).toBe(true);
    expect(isDarkBehindAutoText('0000FF')).toBe(true);
    expect(isDarkBehindAutoText('7030A0')).toBe(false);
    expect(isDarkBehindAutoText('0070C0')).toBe(false);
  });
});

describe('automatic text over a dark background (issue #137)', () => {
  it('is drawn white in a dark cell, and black in a light one', () => {
    expect(colourOf(cell(SHD('25397D'), para('Inspection')), 'Inspection')).toBe('FFFFFF');
    expect(colourOf(cell(SHD('FFC000'), para('Amber')), 'Amber')).toBe('000000');
  });

  it('is white when it says `auto`, and black when it says 000000', () => {
    expect(colourOf(cell(SHD('25397D'), para('auto', '<w:color w:val="auto"/>')), 'auto')).toBe(
      'FFFFFF',
    );
    expect(colourOf(cell(SHD('25397D'), para('black', '<w:color w:val="000000"/>')), 'black')).toBe(
      '000000',
    );
  });

  it('is judged by the nearest background: run, then paragraph, then cell', () => {
    expect(colourOf(cell(SHD('25397D'), para('run', SHD('FFFFFF'))), 'run')).toBe('000000');
    expect(colourOf(cell(SHD('25397D'), para('para', '', SHD('FFFFFF'))), 'para')).toBe('000000');
    expect(colourOf(cell(SHD('FFFFFF'), para('dark', SHD('000000'))), 'dark')).toBe('FFFFFF');
    expect(colourOf(para('band', '', SHD('25397D')), 'band')).toBe('FFFFFF');
  });

  it('is not judged by a highlight, which shows over the shading all the same', () => {
    const yellow = '<w:highlight w:val="yellow"/>';
    expect(colourOf(cell(SHD('25397D'), para('marked', yellow)), 'marked')).toBe('FFFFFF');
    expect(colourOf(cell(SHD('25397D'), para('both', yellow + SHD('FFFFFF'))), 'both')).toBe(
      '000000',
    );
    // Black on a black highlight, as Word leaves it.
    expect(colourOf(para('pen', '<w:highlight w:val="black"/>'), 'pen')).toBe('000000');
  });

  it('is judged by what a patterned shading paints, and `nil` paints nothing', () => {
    expect(colourOf(cell(SHD('FFFFFF', 'solid', '25397D'), para('solid')), 'solid')).toBe('FFFFFF');
    expect(colourOf(cell(SHD('FFFFFF', 'pct90', '000000'), para('pct90')), 'pct90')).toBe('FFFFFF');
    expect(colourOf(cell(SHD('FFFFFF', 'pct75', '000000'), para('pct75')), 'pct75')).toBe('000000');
    expect(colourOf(cell(SHD('25397D', 'nil'), para('nil')), 'nil')).toBe('000000');
  });

  it('follows the shading a table style gives the cell', () => {
    const styles =
      '<w:style w:type="table" w:styleId="Dark"><w:name w:val="Dark"/><w:tblPr/>' +
      `<w:tblStylePr w:type="firstRow"><w:tcPr>${SHD('25397D')}</w:tcPr></w:tblStylePr></w:style>`;
    const body =
      '<w:tbl><w:tblPr><w:tblStyle w:val="Dark"/><w:tblW w:w="0" w:type="auto"/>' +
      '<w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc>${para('header')}</w:tc></w:tr><w:tr><w:tc>${para('body')}</w:tc></w:tr>` +
      '</w:tbl><w:p/>';
    const flow = Ream.parse(buildDocxFromBody(body, { stylesXml: styles })).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const colours = new Map<string, string>();
    for (const c of laid.pages[0]!.commands) {
      if (c.type !== 'line') continue;
      for (const t of c.line.tokens)
        if (t.kind === 'text') colours.set(t.text, t.resolvedRun.colorHex);
    }
    expect(colours.get('header')).toBe('FFFFFF');
    expect(colours.get('body')).toBe('000000');
  });

  it('turns a list item’s marker white with its text', () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const item = para('item', '', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const colours = drawnColours(cell(SHD('25397D'), item), numbering);
    expect(colours.get('item')).toBe('FFFFFF');
    expect(colours.get('1.')).toBe('FFFFFF');
  });

  it('is drawn white on a text box filled dark', () => {
    const box = `<w:p><w:r><w:drawing>
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
          <wp:extent cx="2540000" cy="635000"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>
          <wp:wrapNone/><wp:docPr id="1" name="Box"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="635000"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:solidFill><a:srgbClr val="25397D"/></a:solidFill></wps:spPr>
                <wps:txbx><w:txbxContent>${para('boxed')}</w:txbxContent></wps:txbx>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing></w:r></w:p>`;
    expect(colourOf(box, 'boxed')).toBe('FFFFFF');
  });

  it('is left alone in a document Word did not set', () => {
    const flow = Ream.parse(buildDocxFromBody(cell(SHD('25397D'), para('plain')))).flow;
    const { typesetBy: _word, ...flat } = flowRenderOptions(flow);
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flat,
    });
    const token = laid.pages[0]!.commands.flatMap((c) =>
      c.type === 'line' ? c.line.tokens : [],
    ).find((t) => t.kind === 'text' && t.text === 'plain') as TextToken;
    expect(token.resolvedRun.colorHex).toBe('000000');
  });
});

describe('automatic text in the other targets', () => {
  it('is white in HTML too, over the shading HTML now paints', async () => {
    const html = new TextDecoder().decode(
      await Ream.parse(
        buildDocxFromBody(
          cell(SHD('25397D'), para('cellText')) +
            para('bandText', '', SHD('25397D')) +
            para('runText', SHD('000000')),
        ),
      ).convert('html'),
    );
    const span = (text: string) =>
      new RegExp(`<span style="[^"]*">${text}</span>`, 'u').exec(html)?.[0];
    expect(span('cellText')).toContain('color:#FFFFFF');
    expect(span('bandText')).toContain('color:#FFFFFF');
    expect(span('runText')).toContain('color:#FFFFFF');
    expect(span('runText')).toContain('background-color:#000000');
    expect(html).toMatch(/<p style="[^"]*background-color:#25397D[^"]*">/u);
  });

  it('stays automatic in a written DOCX, a stated black stays black, and a highlight a highlight', async () => {
    const docx = await Ream.parse(
      buildDocxFromBody(
        cell(SHD('25397D'), para('auto') + para('black', '<w:color w:val="000000"/>')) +
          para('pen', '<w:highlight w:val="yellow"/>'),
      ),
    ).convert('docx');
    // Read back and laid out, the words fall as they did in the source.
    const flow = Ream.parse(docx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const colours = new Map<string, string>();
    for (const c of laid.pages[0]!.commands) {
      if (c.type !== 'line') continue;
      for (const t of c.line.tokens) {
        if (t.kind === 'text') colours.set(t.text, t.resolvedRun.colorHex);
      }
    }
    expect(colours.get('auto')).toBe('FFFFFF');
    expect(colours.get('black')).toBe('000000');
    const xml = new TextDecoder().decode(unzipSync(docx)['word/document.xml']);
    expect(xml).not.toContain('w:val="FFFFFF"');
    expect(xml).toContain('<w:color w:val="000000"/>');
    expect(xml).toContain('<w:highlight w:val="yellow"/>');
    const pen = readDocx(docx)
      .doc.body.flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []))
      .find((r) => r.text === 'pen')!;
    expect(pen.properties.highlightColorHex).toBe('FFFF00');
    expect(pen.properties.shadingColorHex).toBeUndefined();
  });
});

describe('a cell’s shading (§17.4.33)', () => {
  const shadingOf = (tcPr: string) => {
    const el = readDocx(buildDocxFromBody(cell(tcPr, para('x')))).doc.body[0]!;
    if (el.kind !== 'table') throw new Error('expected a table');
    return el.table.rows[0]!.cells[0]!.properties.shading?.colorHex;
  };

  it('is the colour its pattern paints, and nothing for `nil`', () => {
    expect(shadingOf(SHD('25397D'))).toBe('25397D');
    expect(shadingOf(SHD('FFFFFF', 'solid', '25397D'))).toBe('25397D');
    expect(shadingOf(SHD('FFFFFF', 'pct50', '000000'))).toBe('808080');
    expect(shadingOf('<w:shd w:val="pct25"/>')).toBe('BFBFBF');
    expect(shadingOf(SHD('25397D', 'nil'))).toBeUndefined();
    expect(shadingOf(SHD('auto'))).toBeUndefined();
  });
});
