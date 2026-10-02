// §21.2.2.216 — a chart's text, role by role, as Excel sets it. Excel's own PDF
// of probe workbooks (2026-10-02): every text in the theme's minor font at 10pt
// in the theme's text colour, the chart title and the axis titles bold, the
// chart's `c:chartSpace/c:txPr` over that and each element's own over the
// chart's — never the workbook's Normal font. Without a theme part: Calibri
// when the Normal font is Calibri, Aptos Narrow when it is anything else.

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { Chart } from '@/core/document-model';
import type { ChartTextDefaults } from '@/core/drawingml/chart-text';
import type { FamilyKey } from '@/core/fonts/remote-fonts';
import { flowRenderOptions } from '@/core/converter/project';
import { Ream } from '@/core/converter/ream';
import { CHART_LABEL_PT, CHART_TITLE_PT, buildChartScene } from '@/core/drawingml/chart-geometry';
import { parseChart } from '@/core/drawingml/chart-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';

const here = dirname(fileURLToPath(import.meta.url));
const font = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(here, `fixtures/fonts/${name}.ttf`)));

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** `c:txPr` with one paragraph's defaults. */
const txPr = (attrs: string, inner = ''): string =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${attrs}>${inner}</a:defRPr></a:pPr>` +
  '<a:endParaRPr lang="en-US"/></a:p></c:txPr>';
const latin = (face: string): string => `<a:latin typeface="${face}"/>`;

interface Parts {
  readonly space?: string;
  readonly titleDef?: string;
  readonly titleRun?: string;
  readonly catTx?: string;
  readonly legendTx?: string;
}

/** A column chart with a title, two axes and a legend — each role's text in place. */
const chartXml = (p: Parts = {}): string =>
  `<c:chartSpace ${C_NS}><c:chart>` +
  `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>${p.titleDef ?? ''}</a:defRPr></a:pPr>` +
  `<a:r><a:rPr lang="en-US"${p.titleRun ?? ''}/><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
  '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
  '<c:tx><c:v>Series</c:v></c:tx>' +
  '<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Alpha</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
  '<c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>' +
  '</c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
  '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' +
  `<c:tickLblPos val="nextTo"/>${p.catTx ?? ''}<c:crossAx val="2"/></c:catAx>` +
  '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>' +
  '<c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>' +
  `</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/>${p.legendTx ?? ''}</c:legend></c:chart>` +
  `${p.space ?? ''}</c:chartSpace>`;

const EXCEL: ChartTextDefaults = {
  themeFonts: { major: { latin: 'Trebuchet MS' }, minor: { latin: 'Georgia' } },
  family: 'Georgia',
};

const parse = (xml: string): Chart =>
  parseChart(new TextEncoder().encode(xml), defaultColorResolver, undefined, EXCEL)!;
/** …as a host that resolves no chart text reads it (Word, PowerPoint). */
const parseBare = (xml: string): Chart =>
  parseChart(new TextEncoder().encode(xml), defaultColorResolver)!;

/** A minimal theme naming its two fonts. */
const theme = (minor: string, major: string): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>' +
  '<a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
  `<a:fontScheme name="T"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  '<a:fmtScheme name="T"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>' +
  '</a:themeElements></a:theme>';

/** styles.xml content whose Normal font is `face`. */
const normal = (face: string): string =>
  `<fonts count="1"><font><sz val="11"/><name val="${face}"/></font></fonts>` +
  '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>';

const chartOf = (xlsx: Uint8Array): Chart => {
  const charts = Ream.parse(xlsx).flow.charts;
  const chart = charts ? [...charts.values()][0] : undefined;
  if (!chart) throw new Error('no chart');
  return chart;
};

describe('chart text as Excel sets it (§21.2.2.216)', () => {
  it("sets every text in the theme's minor font at 10pt, the titles bold, when the chart says nothing", () => {
    const text = parse(chartXml()).text!;
    for (const role of ['legend', 'catAxis', 'valAxis', 'dataLabels'] as const) {
      expect(text[role]).toEqual({ family: 'Georgia', sizePt: 10, colorHex: '000000' });
    }
    for (const role of ['title', 'catAxisTitle', 'valAxisTitle'] as const) {
      expect(text[role]).toEqual({ family: 'Georgia', sizePt: 10, colorHex: '000000', bold: true });
    }
  });

  it("takes the chart's own c:txPr over the defaults, and an element's over the chart's", () => {
    const text = parse(
      chartXml({
        space: txPr(' sz="1400"', latin('Arial')),
        catTx: txPr('', latin('Times New Roman')),
        legendTx: txPr(' sz="800"', latin('Courier New')),
      }),
    ).text!;
    expect(text.valAxis).toMatchObject({ family: 'Arial', sizePt: 14 });
    expect(text.title).toMatchObject({ family: 'Arial', sizePt: 14, bold: true });
    expect(text.catAxis).toMatchObject({ family: 'Times New Roman', sizePt: 14 });
    expect(text.legend).toMatchObject({ family: 'Courier New', sizePt: 8 });
  });

  it('resolves +mn-lt and +mj-lt through the theme', () => {
    const text = parse(
      chartXml({ space: txPr('', latin('+mn-lt')), titleDef: latin('+mj-lt') }),
    ).text!;
    expect(text.valAxis?.family).toBe('Georgia');
    expect(text.title?.family).toBe('Trebuchet MS');
  });

  it("reads a title run's own size and weight over the title's defaults", () => {
    const text = parse(chartXml({ titleRun: ' sz="2000" b="0"' })).text!;
    expect(text.title).toMatchObject({ sizePt: 20, bold: false });
  });

  it("colours every text the chart's c:txPr colours, the titles included", () => {
    const text = parse(
      chartXml({ space: txPr('', '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>') }),
    ).text!;
    expect(text.title?.colorHex).toBe('FF0000');
    expect(text.catAxis?.colorHex).toBe('FF0000');
  });

  it('resolves no text for a host that gives no defaults', () => {
    expect(parseBare(chartXml()).text).toBeUndefined();
  });
});

describe('a workbook chart takes its face from the theme, never the Normal font', () => {
  it("is set in the workbook theme's minor font", () => {
    const xlsx = buildXlsx({
      rows: [['x']],
      themeXml: theme('Georgia', 'Trebuchet MS'),
      stylesXml: normal('Verdana'),
      sheetChart: { chartXml: chartXml() },
    });
    expect(chartOf(xlsx).text?.catAxis?.family).toBe('Georgia');
  });

  it('is set in Calibri without a theme when the Normal font is Calibri, else in Aptos Narrow', () => {
    const face = (normalFace: string): string | undefined =>
      chartOf(
        buildXlsx({
          rows: [['x']],
          stylesXml: normal(normalFace),
          sheetChart: { chartXml: chartXml() },
        }),
      ).text?.catAxis?.family;
    expect(face('Calibri')).toBe('Calibri');
    expect(face('Verdana')).toBe('Aptos Narrow');
  });
});

describe('the chart scene sets each label in its role', () => {
  it("carries the chart's face, size and colour onto the labels it measures in that face", () => {
    const chart = parse(chartXml({ catTx: txPr('', latin('Times New Roman')) }));
    const asked = new Set<string>();
    const scene = buildChartScene(chart, 400, 300, (text, sizePt, f) => {
      asked.add(`${f?.family ?? '-'}|${f?.bold === true}`);
      return text.length * sizePt * 0.5;
    })!;
    const label = (text: string) => scene.labels.find((l) => l.text === text)!;
    expect(label('Quarterly Sales')).toMatchObject({ family: 'Georgia', sizePt: 10, bold: true });
    expect(label('Alpha')).toMatchObject({ family: 'Times New Roman', sizePt: 10 });
    expect(label('Alpha').colorHex).toBe('000000');
    expect(label('Series')).toMatchObject({ family: 'Georgia', sizePt: 10 });
    // Measured in the face each is drawn in.
    expect(asked).toContain('Times New Roman|false');
    expect(asked).toContain('Georgia|true');
  });

  it('keeps its own sizes and greys for a chart whose reader resolved no text', () => {
    const chart = parseBare(chartXml());
    const scene = buildChartScene(chart, 400, 300, (text, sizePt) => text.length * sizePt * 0.5)!;
    const title = scene.labels.find((l) => l.text === 'Quarterly Sales')!;
    const cat = scene.labels.find((l) => l.text === 'Alpha')!;
    expect(title).toMatchObject({ sizePt: CHART_TITLE_PT, colorHex: '404040' });
    expect(title.family).toBeUndefined();
    expect(cat).toMatchObject({ sizePt: CHART_LABEL_PT, colorHex: '595959' });
    expect(cat.bold).toBeUndefined();
  });
});

describe('a chart drawn in faces of its own', () => {
  it('draws, in its own face, when several families are loaded and the base face is never asked for', () => {
    // Every text of the chart names its face, so the subset asks for no base
    // face — and with several families loaded the base face then has no
    // resource of its own. Looked up by it regardless, the whole chart went
    // undrawn, and a chart sheet's page with it.
    const serif = FontRegistry.fromBytes({ regular: font('Roboto-Bold') });
    const sans = FontRegistry.fromBytes({ regular: font('Roboto-Regular') });
    const registriesByFamily: ReadonlyMap<FamilyKey, FontRegistry> = new Map([
      ['arimo', sans],
      ['tinos', serif],
    ]);
    const xlsx = buildXlsx({
      rows: [['x']],
      themeXml: theme('Times New Roman', 'Times New Roman'),
      sheetChart: { chartXml: chartXml() },
    });
    const flow = Ream.parse(xlsx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: sans,
      registriesByFamily,
      ...flowRenderOptions(flow),
    });
    const commands = laid.pages.flatMap((p) => p.commands);
    const alpha = commands.find(
      (c) =>
        c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === 'Alpha'),
    );
    expect(alpha?.type).toBe('line');
    if (alpha?.type !== 'line') return;
    const token = alpha.line.tokens.find((t) => t.kind === 'text');
    expect(token?.kind === 'text' ? token.font.parsed : undefined).toBe(
      serif.resolveByStyle(false, false).parsed,
    );
  });
});
