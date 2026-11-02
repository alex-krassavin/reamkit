// §17.3.2.26 — a run names a face for each kind of character: `w:ascii` for
// ASCII, `w:hAnsi` for the rest above it, `w:eastAsia` and `w:cs` for East
// Asian and complex-script text. Measured in Word for Mac, on a .doc and on a
// .docx whose run has Verdana for ASCII and Courier New for the other
// characters: "split" reports Verdana, "Кириллица" Courier New. The layout drew
// a run in its ascii face alone and asked w:hAnsi for nothing but the line.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc } from './fixtures/build-doc';
import { buildDocxFromBody } from './fixtures/build-docx';
import type { FamilyKey } from '@/core/fonts';
import type { RunProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { TextLineItem, TextToken } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { convertDocxToPdfSync } from '@/core/converter';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry, createFontMeasure } from '@/core/font';
import { familiesInFlow } from '@/core/fonts/families';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDoc } from '@/word/doc/doc-reader';
import { detectDocxFamilyKeys } from '@/word/docx-to-pdf';

const font = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(`tests/fixtures/fonts/${name}.ttf`));
// A stand-in per substitute, each a face of its own, so the face a token is
// drawn in says which family drew it: Verdana is set in arimo's, Courier New in
// cousine's, Times New Roman in tinos's.
const STAND_INS = {
  arimo: font('Roboto-Regular'),
  cousine: font('Roboto-Bold'),
  tinos: font('Roboto-Italic'),
} as const;
const registries = (keys: ReadonlyArray<keyof typeof STAND_INS>): Map<FamilyKey, FontRegistry> =>
  new Map(keys.map((k) => [k, FontRegistry.fromBytes({ regular: STAND_INS[k] })]));
const ALL = registries(['arimo', 'cousine', 'tinos']);

/** The family whose stand-in a token is drawn in. */
function familyOf(t: TextToken, regs: ReadonlyMap<FamilyKey, FontRegistry>): string {
  for (const [key, reg] of regs)
    if (reg.resolveByStyle(false, false).parsed === t.font.parsed) return key;
  return '?';
}

/** Every word-bearing text token of the laid-out document. */
function textTokens(flow: FlowDoc, regs: ReadonlyMap<FamilyKey, FontRegistry>): Array<TextToken> {
  const laid = layoutStyledDocument(flow.body, {
    registry: regs.get('arimo')!,
    registriesByFamily: regs,
    ...flowRenderOptions(flow),
  });
  return laid.pages
    .flatMap((page) => page.commands.filter((c): c is TextLineItem => c.type === 'line'))
    .flatMap((l) => l.line.tokens)
    .filter((t): t is TextToken => t.kind === 'text' && !t.isSpace && t.text !== '');
}

/** Each token's text and the family drawing it. */
const faces = (flow: FlowDoc, regs = ALL): Array<[string, string]> =>
  textTokens(flow, regs).map((t) => [t.text, familyOf(t, regs)]);

/** A one-paragraph .docx: `text` in a run of `rFonts` (and `extra` run properties). */
const docx = (rFonts: string, text: string, extra = ''): Uint8Array =>
  buildDocxFromBody(
    `<w:p><w:r><w:rPr><w:rFonts ${rFonts}/>${extra}</w:rPr>` +
      `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
  );

const VERDANA_COURIER = 'w:ascii="Verdana" w:hAnsi="Courier New"';

describe('a run’s characters are set in the face their slot of w:rFonts names', () => {
  it('sets the letters above ASCII in the hAnsi face, the ASCII ones in the ascii face', () => {
    const flow = Ream.parse(docx(VERDANA_COURIER, 'split Кириллица.')).flow;
    expect(faces(flow)).toEqual([
      ['split', 'arimo'],
      ['Кириллица', 'cousine'],
      ['.', 'arimo'],
    ]);
  });

  it('measures them in it', () => {
    const flow = Ream.parse(docx(VERDANA_COURIER, 'Кириллица')).flow;
    const [word] = textTokens(flow, ALL);
    const sizePt = word!.fontSizePt;
    const inCourier = createFontMeasure(
      ALL.get('cousine')!.resolveByStyle(false, false).parsed,
    ).textWidthPt('Кириллица', sizePt);
    const inVerdana = createFontMeasure(
      ALL.get('arimo')!.resolveByStyle(false, false).parsed,
    ).textWidthPt('Кириллица', sizePt);
    expect(inCourier).not.toBeCloseTo(inVerdana, 1);
    expect(word!.widthPt).toBeCloseTo(inCourier, 6);
  });

  it('draws them in it, and only where the run names a face of its own for them', () => {
    const baseFonts = (bytes: Uint8Array): Set<string> =>
      new Set(
        [
          ...new TextDecoder('latin1')
            .decode(bytes)
            .matchAll(/\/BaseFont \/[A-Z]{6}\+([A-Za-z-]+)/g),
        ].map((m) => m[1]!),
      );
    const options = { fonts: { regular: STAND_INS.arimo }, registriesByFamily: ALL };
    const split = convertDocxToPdfSync(docx(VERDANA_COURIER, 'split Кириллица'), options);
    expect([...baseFonts(split)].sort()).toEqual(['Roboto-Bold', 'Roboto-Regular']);
    const whole = convertDocxToPdfSync(docx('w:ascii="Verdana"', 'split Кириллица'), options);
    expect([...baseFonts(whole)]).toEqual(['Roboto-Regular']);
  });

  it('reads a .doc’s other font the same way', () => {
    // [MS-DOC] sprmCRgFtc0 names the ASCII font and sprmCRgFtc2 the other one.
    const flow = readDoc(
      buildDoc([{ text: 'split Кириллица\r', compressed: false }], {
        fonts: ['Verdana', 'Courier New'],
        formatRuns: [{ length: 16, fontAscii: 0, fontOther: 1 }],
      }),
    ).doc;
    expect(faces(flow)).toEqual([
      ['split', 'arimo'],
      ['Кириллица', 'cousine'],
    ]);
  });

  it('sets a right-to-left run in its complex-script face whole', () => {
    // §17.3.2.30 — a w:rtl run is complex script whatever its characters are.
    // Word: ten i's in a w:rtl run whose cs font is Courier New put the word
    // after them where ten Courier New letters do, 147.95 pt; without w:rtl,
    // where ten Verdana ones do, 109.2 pt.
    const flow = Ream.parse(
      docx('w:ascii="Verdana" w:hAnsi="Verdana" w:cs="Times New Roman"', 'abc', '<w:rtl/>'),
    ).flow;
    expect(faces(flow)).toEqual([['abc', 'tinos']]);
  });

  it('takes a slot the run does not name from the styles under it', () => {
    // Word for Mac sets the Cyrillic of a run that names only `w:ascii` in the
    // hAnsi font of the defaults (the theme's body font there), not in Verdana.
    const flow = Ream.parse(
      buildDocxFromBody(
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Verdana"/></w:rPr><w:t>split Кириллица</w:t></w:r></w:p>',
        {
          stylesXml:
            '<w:docDefaults><w:rPrDefault><w:rPr>' +
            '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Courier New"/>' +
            '</w:rPr></w:rPrDefault></w:docDefaults>',
        },
      ),
    ).flow;
    expect(faces(flow)).toEqual([
      ['split', 'arimo'],
      ['Кириллица', 'cousine'],
    ]);
  });

  it('reads a DrawingML a:latin as the font of every slot it covers', () => {
    // §21.1.2.3.7 — `a:latin` sets everything neither East Asian nor complex
    // script, so the Cyrillic of a .docx's own shape is not left to the
    // document's hAnsi: tdf143476_lockedCanvas_position.docx's WordArt.
    // A member of a group states its words in an `a:txBody`, as the canvas's do.
    const shape =
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Group 1"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"/></wps:spPr>' +
      '<a:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="Courier New"/></a:rPr>' +
      '<a:t>split Кириллица</a:t></a:r></a:p></a:txBody>' +
      '<wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const flow = Ream.parse(
      buildDocxFromBody(shape, {
        stylesXml:
          '<w:docDefaults><w:rPrDefault><w:rPr>' +
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' +
          '</w:rPr></w:rPrDefault></w:docDefaults>',
      }),
    ).flow;
    // The shape's own text, wherever the reader keeps it.
    const runs: Array<{ text: string; properties: RunProperties }> = [];
    const seen = new Set<object>();
    const find = (node: unknown): void => {
      if (typeof node !== 'object' || node === null || seen.has(node)) return;
      seen.add(node);
      const run = node as { text?: unknown; properties?: RunProperties };
      if (run.text === 'split Кириллица' && run.properties) {
        runs.push({ text: run.text, properties: run.properties });
      }
      for (const value of Object.values(node)) find(value);
    };
    find(flow.body);
    expect(runs.map((r) => r.properties.fontFamily)).toEqual([
      { ascii: 'Courier New', hAnsi: 'Courier New' },
    ]);
  });

  it('keeps a tab in a run set in two faces', () => {
    const flow = Ream.parse(docx(VERDANA_COURIER, 'Ж1.\tX')).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: ALL.get('arimo')!,
      registriesByFamily: ALL,
      ...flowRenderOptions(flow),
    });
    const tokens = laid.pages
      .flatMap((page) => page.commands.filter((c): c is TextLineItem => c.type === 'line'))
      .flatMap((l) => l.line.tokens)
      .filter((t): t is TextToken => t.kind === 'text');
    expect(tokens.map((t) => (t.tab ? 'TAB' : t.text))).toEqual(['Ж', '1.', 'TAB', 'X']);
  });

  it('leaves a slot to the ascii face where nothing was fetched for its family', () => {
    const regs = registries(['arimo', 'tinos']);
    const flow = Ream.parse(docx(VERDANA_COURIER, 'split Кириллица')).flow;
    expect(faces(flow, regs)).toEqual([
      ['split', 'arimo'],
      ['Кириллица', 'arimo'],
    ]);
  });
});

describe('the families fetched for a run’s slots', () => {
  const families = (bytes: Uint8Array): Array<string> =>
    [...familiesInFlow(Ream.parse(bytes).flow)].sort();

  it('fetches the hAnsi family where the text has characters it draws', () => {
    expect(families(docx(VERDANA_COURIER, 'Кириллица'))).toEqual(['arimo', 'cousine']);
    expect(families(docx(VERDANA_COURIER, 'split'))).toEqual(['arimo']);
  });

  it('fetches the complex-script family for a right-to-left run', () => {
    const rtl = 'w:ascii="Verdana" w:cs="Times New Roman"';
    expect(families(docx(rtl, 'abc', '<w:rtl/>'))).toEqual(['arimo', 'tinos']);
    expect(families(docx(rtl, 'abc'))).toEqual(['arimo']);
  });

  it('reads the hAnsi family off a .docx’s XML too', () => {
    expect([...detectDocxFamilyKeys(docx(VERDANA_COURIER, 'Кириллица'))].sort()).toEqual([
      'arimo',
      'cousine',
    ]);
  });
});
