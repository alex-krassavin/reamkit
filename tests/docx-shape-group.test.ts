// §20.5.2.17 `wpg:wgp` — a drawing group is written as one: its members at
// their offsets in it, pictures among them, groups inside it. The writer wrote
// a group as a single empty shape, and the paths and labels of a PDF's figure
// went with it. And §20.1.2.1.1 `a:bodyPr @wrap="none"` — a text box told not
// to wrap keeps its line whole.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import type { BodyElement, ShapeBlock } from '@/core/document-model';
import type { TextLineItem } from '@/layout/page-doc';
import type { FlowDoc } from '@/core/ir/flow';
import { FontRegistry } from '@/core/font';
import { ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';
import { writeDocx } from '@/word/docx-writer';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const EMU = 12700;

/** A text box holding one line, `noWrap` as asked. */
const label = (text: string, width: number, noWrap: boolean): ShapeBlock => ({
  width: pt(width),
  height: pt(12),
  geometry: { kind: 'preset', preset: 'rect' },
  fill: { kind: 'none' },
  text: {
    content: [
      { kind: 'paragraph', paragraph: { properties: {}, runs: [{ text, properties: {} }] } },
    ],
    insetLeft: pt(0),
    insetTop: pt(0),
    insetRight: pt(0),
    insetBottom: pt(0),
    ...(noWrap ? { noWrap: true } : {}),
  },
  paragraphProperties: {},
});

/** A group: a stroked path, a picture, a group holding a box, and a label. */
function groupDoc(): { doc: FlowDoc; group: ShapeBlock } {
  const resources = new ResourceStore();
  const picture = resources.put(buildTinyPng(4, 4, [0, 0, 0, 128]));
  const path: ShapeBlock = {
    width: pt(40),
    height: pt(20),
    geometry: {
      kind: 'custom',
      custom: {
        pathWidth: 40,
        pathHeight: 20,
        commands: [
          { cmd: 'move', x: 0, y: 20 },
          { cmd: 'cubic', x1: 10, y1: 0, x2: 30, y2: 0, x: 40, y: 20 },
        ],
      },
    },
    fill: { kind: 'none' },
    line: { width: pt(1), colorHex: '000000', fill: 'solid' },
    paragraphProperties: {},
  };
  const inner: ShapeBlock = {
    width: pt(30),
    height: pt(10),
    geometry: { kind: 'custom', custom: { pathWidth: 0, pathHeight: 0, commands: [] } },
    fill: { kind: 'none' },
    children: [
      {
        shape: {
          width: pt(30),
          height: pt(10),
          geometry: { kind: 'preset', preset: 'rect' },
          fill: { kind: 'solid', colorHex: 'CCCCCC' },
          paragraphProperties: {},
        },
        xPt: pt(0),
        yPt: pt(0),
      },
    ],
    paragraphProperties: {},
  };
  const group: ShapeBlock = {
    width: pt(120),
    height: pt(60),
    geometry: { kind: 'custom', custom: { pathWidth: 0, pathHeight: 0, commands: [] } },
    fill: { kind: 'none' },
    children: [
      { shape: path, xPt: pt(5), yPt: pt(10) },
      {
        shape: {
          width: pt(20),
          height: pt(20),
          geometry: { kind: 'preset', preset: 'rect' },
          fill: { kind: 'picture', imageResource: picture },
          paragraphProperties: {},
        },
        xPt: pt(60),
        yPt: pt(5),
      },
      { shape: inner, xPt: pt(80), yPt: pt(40) },
      { shape: label('Monitor', 30, true), xPt: pt(10), yPt: pt(40) },
    ],
    paragraphProperties: {},
  };
  const { doc: base } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
  const body: ReadonlyArray<BodyElement> = [{ kind: 'shape', shape: group }];
  return { doc: { ...base, body, resources }, group };
}

describe('a drawing group in a .docx (§20.5.2.17)', () => {
  it('is written as a group, its members placed in a child space the size of its box', () => {
    const xml = new TextDecoder().decode(
      OpcPackage.open(writeDocx(groupDoc().doc).bytes).getMainDocument().data,
    );
    expect(xml).toContain('<wp:inline');
    expect(xml).toContain(
      '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
        `<a:ext cx="${String(120 * EMU)}" cy="${String(60 * EMU)}"/>` +
        `<a:chOff x="0" y="0"/><a:chExt cx="${String(120 * EMU)}" cy="${String(60 * EMU)}"/>`,
    );
    // Each member at its offset, a picture as a picture, a group as a group.
    expect(xml).toContain(`<a:off x="${String(5 * EMU)}" y="${String(10 * EMU)}"/>`);
    expect(xml).toMatch(/<pic:pic [^>]*>.*<a:blip r:embed="rId\d+">/u);
    expect(xml).toContain('<wpg:grpSp>');
    expect(xml).toContain('<wps:bodyPr wrap="none"');
  });

  it('is read back as the group it was', () => {
    const { doc, group } = groupDoc();
    const { doc: again } = readDocx(writeDocx(doc).bytes);
    const read = again.body.find((el) => el.kind === 'shape');
    const members = read?.kind === 'shape' ? (read.shape.children ?? []) : [];
    expect(members).toHaveLength(group.children!.length);
    members.forEach((m, k) => {
      expect(m.xPt).toBeCloseTo(group.children![k]!.xPt, 2);
      expect(m.yPt).toBeCloseTo(group.children![k]!.yPt, 2);
    });
    expect(members[1]!.shape.fill.kind).toBe('picture');
    expect(members[2]!.shape.children).toHaveLength(1);
  });
});

describe('a text box that does not wrap (§20.1.2.1.1)', () => {
  const linesOf = (box: ShapeBlock): number => {
    const { doc } = groupDoc();
    const laid = layoutStyledDocument([{ kind: 'shape', shape: box }], {
      registry: FontRegistry.fromBytes(FONTS),
      styles: doc.styles,
    });
    return laid.pages[0]!.commands.filter(
      (c): c is TextLineItem =>
        c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text.length > 0),
    ).length;
  };

  it('keeps its line whole, however narrow the box', () => {
    const text = 'side exit to existing trace';
    expect(linesOf(label(text, 20, false))).toBeGreaterThan(1);
    expect(linesOf(label(text, 20, true))).toBe(1);
  });
});
