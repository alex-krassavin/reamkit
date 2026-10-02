// §21.2.2.226 — a scatter's two value axes are told apart by where they sit:
// the upright one is the values', the lying one the x's. Excel writes the x
// axis first, and read as "the first c:valAx" its ends, step and number format
// landed on the y axis (DataTableCities.xlsx: longitude −180…180 by 60 in
// `0"°"`, latitude −90…90 by 30).

import { describe, expect, it } from 'vitest';

import { buildChartScene } from '@/core/drawingml/chart-geometry';
import { parseChart } from '@/core/drawingml/chart-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const numCache = (vals: ReadonlyArray<number>): string =>
  `<c:numRef><c:numCache><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  '</c:numCache></c:numRef>';

const title = (text: string): string =>
  `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;

/** A value axis: where it sits, what it is fixed to, its title and format. */
const valAx = (
  id: number,
  cross: number,
  pos: string,
  min: number,
  max: number,
  unit: number,
  name: string,
): string =>
  `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${max}"/><c:min val="${min}"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/>${title(name)}` +
  '<c:numFmt formatCode="0&quot;°&quot;" sourceLinked="0"/><c:tickLblPos val="nextTo"/>' +
  `<c:crossAx val="${cross}"/><c:crossBetween val="midCat"/><c:majorUnit val="${unit}"/></c:valAx>`;

const SCATTER =
  `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
  '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
  `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:xVal>${numCache([-150, 30, 120])}</c:xVal>` +
  `<c:yVal>${numCache([60, -20, 35])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
  valAx(1, 2, 'b', -180, 180, 60, 'Longitude') +
  valAx(2, 1, 'l', -90, 90, 30, 'Latitude') +
  '</c:plotArea></c:chart></c:chartSpace>';

const chart = parseChart(new TextEncoder().encode(SCATTER), defaultColorResolver)!;

describe('a scatter tells its axes apart by where they sit (§21.2.2.226)', () => {
  it('reads the upright axis as the values and the lying one as the x', () => {
    expect(chart).toMatchObject({
      valAxisMin: -90,
      valAxisMax: 90,
      valAxisMajorUnit: 30,
      valAxisTitle: 'Latitude',
      xAxisMin: -180,
      xAxisMax: 180,
      xAxisMajorUnit: 60,
      catAxisTitle: 'Longitude',
      numberFormat: '0"°"',
      xNumberFormat: '0"°"',
    });
  });

  it('labels each axis with its own ends, step and format, its title beside it', () => {
    const scene = buildChartScene(chart, 420, 300, (text, sizePt) => text.length * sizePt * 0.5)!;
    const texts = scene.labels.map((l) => l.text);
    for (const x of ['-180°', '-120°', '-60°', '0°', '60°', '120°', '180°'])
      expect(texts).toContain(x);
    for (const y of ['-90°', '-60°', '-30°', '30°', '90°']) expect(texts).toContain(y);
    expect(texts).not.toContain('-200');
    const lat = scene.labels.find((l) => l.text === 'Latitude');
    expect(lat?.rotationDeg).toBe(90);
    expect(scene.labels.find((l) => l.text === 'Longitude')?.rotationDeg).toBeUndefined();
  });
});
