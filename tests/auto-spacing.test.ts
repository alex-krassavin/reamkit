// ECMA-376 Part 1 §17.3.1.1/§17.3.1.3 — `w:beforeAutospacing` and
// `w:afterAutospacing` ask for HTML's automatic 14pt, in every compatibility
// mode. Measured in Word for Mac by the positions of 11pt Calibri paragraphs.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const LINE = (11 * 2500) / 2048;
const ASCENT = (11 * 1950) / 2048;
const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>';
const AUTO = 'w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"';

const para = (text: string, spacing: string): string =>
  `<w:p><w:pPr><w:spacing ${spacing} w:line="240" w:lineRule="auto"/>${RPR}</w:pPr>` +
  `<w:r>${RPR}<w:t>${text}</w:t></w:r></w:p>`;

const SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

const mode = (n: number): string =>
  '<w:compat><w:compatSetting w:name="compatibilityMode" ' +
  `w:uri="http://schemas.microsoft.com/office/word" w:val="${String(n)}"/></w:compat>`;

function firstPage(body: string, settingsXml?: string): ReadonlyArray<PageItem> {
  const docx = buildDocxFromBody(body + SECTION, settingsXml === undefined ? {} : { settingsXml });
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  }).pages[0]!.commands;
}

function baseline(items: ReadonlyArray<PageItem>, text: string): number {
  const hit = items.find(
    (c): c is TextLineItem =>
      c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === text),
  );
  if (!hit) throw new Error(`no line "${text}"`);
  return hit.baselineY;
}

describe('automatic paragraph spacing (§17.3.1.1/§17.3.1.3)', () => {
  it('is 14pt whatever the mode, and two such paragraphs share it', () => {
    // Word: 14.02 apart in modes 14, 15 and none — the stated 5pt ignored.
    for (const settings of [mode(14), mode(15), undefined]) {
      const page = firstPage(
        para('A', 'w:before="0" w:after="0"') + para('B', AUTO) + para('C', AUTO),
        settings,
      );
      expect(baseline(page, 'C') - baseline(page, 'B') - LINE).toBeCloseTo(14, 3);
    }
  });

  it('is not given to the document’s first paragraph, as a stated space is', () => {
    // Word: 72.00 for an automatic space before, 95.45 for a stated 24pt.
    expect(baseline(firstPage(para('A', AUTO), mode(15)), 'A') - ASCENT).toBeCloseTo(72, 3);
    expect(
      baseline(firstPage(para('A', 'w:before="480" w:after="0"'), mode(15)), 'A') - ASCENT,
    ).toBeCloseTo(72 + 24, 3);
  });
});
