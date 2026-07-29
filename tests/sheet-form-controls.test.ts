// E-SHEET W8 — form controls. Checkboxes, option buttons, spinners etc. are
// declared on the worksheet (the x14 extLst <controls>) and point through a
// relationship at a ctrlProp part carrying their objectType + state. The reader
// resolves them and the projection lists each in a "Form controls" section after
// the grid with a type-appropriate affordance. Render-only — not written back.
// ActiveX (OLE) controls are a documented graceful loss.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { BodyElement } from '@/core/document-model';
import { parseFormControlProps } from '@/excel/form-control-parser';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

function paragraphTexts(body: ReadonlyArray<BodyElement>): Array<string> {
  const out: Array<string> = [];
  for (const el of body) {
    if (el.kind === 'paragraph') out.push(el.paragraph.runs.map((r) => r.text).join(''));
  }
  return out;
}

describe('ctrlProp parser (E-SHEET W8)', () => {
  it('reads objectType, checked state and a value', () => {
    expect(
      parseFormControlProps(enc('<formControlPr objectType="CheckBox" checked="Checked"/>')),
    ).toEqual({ objectType: 'CheckBox', checked: true });
    expect(
      parseFormControlProps(enc('<formControlPr objectType="Spin" val="7" min="0" max="10"/>')),
    ).toEqual({ objectType: 'Spin', value: 7 });
    expect(
      parseFormControlProps(enc('<formControlPr objectType="CheckBox" checked="Unchecked"/>')),
    ).toEqual({
      objectType: 'CheckBox',
      checked: false,
    });
  });
});

describe('form controls — end to end (E-SHEET W8)', () => {
  it('lists controls in a Form controls section with type affordances + state', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        formControls: [
          { name: 'Agree', objectType: 'CheckBox', checked: true },
          { name: 'No thanks', objectType: 'CheckBox', checked: false },
          { name: 'Option A', objectType: 'Radio', checked: true },
          { name: 'Quantity', objectType: 'Spin', value: 5 },
          { name: 'Run', objectType: 'Buttons' },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('Form controls');
    expect(texts).toContain('[x] Agree');
    expect(texts).toContain('[ ] No thanks');
    expect(texts).toContain('(o) Option A');
    expect(texts).toContain('Quantity (value 5)');
    expect(texts).toContain('[ Run ]');
  });

  it('draws a control that knows where it goes, instead of listing it', () => {
    // The listing was a stand-in for having no geometry. tdf111980's eleven
    // controls carry theirs in the legacy VML, and drawn from it they land
    // where LibreOffice draws them — the group box 209×88pt at 441pt across
    // the sheet, not a line of ASCII at the origin 18cm away.
    const flow = Ream.parse(
      new Uint8Array(readFileSync('tests/fixtures/real/tdf111980_radioButtons.xlsx')),
    ).flow;
    const shapes = flow.body.flatMap((e) => (e.kind === 'shape' ? [e.shape] : []));
    const at = (x: number, y: number) =>
      shapes.filter(
        (s) =>
          Math.abs((s.float?.posH?.offsetPt ?? -1) - x) < 0.01 &&
          Math.abs((s.float?.posV?.offsetPt ?? -1) - y) < 0.01,
      );
    // The group box: its frame, and its caption over it.
    const group = at(441, 7.5);
    expect(group).toHaveLength(2);
    expect(group[0]?.width).toBeCloseTo(209.25, 2);
    expect(group[0]?.height).toBeCloseTo(87.75, 2);

    // A checked option button draws three things — ring, dot, caption; an
    // unchecked one draws two. The ring is centred vertically in the control's
    // box, the dot inset inside the ring, the caption across the whole box.
    expect(at(280.5, 11.25 + (17.25 - 8.4) / 2)).toHaveLength(1); // ring
    expect(at(281.7, 16.875)).toHaveLength(1); // dot — this one is checked
    expect(at(280.5, 11.25)).toHaveLength(1); // caption
    expect(at(282, 33 + (21 - 8.4) / 2)).toHaveLength(1); // ring
    expect(at(283.2, 39.3 + 1.2)).toHaveLength(0); // unchecked: no dot
    expect(at(282, 33)).toHaveLength(1); // caption

    // An ActiveX control's box comes from the VML shape sharing its shapeId.
    expect(at(564.75, 35.25)).toHaveLength(1);

    // And nothing is listed after the grid any more.
    const texts = paragraphTexts(flow.body);
    expect(texts).not.toContain('Form controls');
    expect(texts).not.toContain('ActiveX controls');
  });

  it('adds no section to a sheet without controls (byte-zero)', () => {
    const flow = Ream.parse(buildXlsx({ rows: [['data']] })).flow;
    expect(paragraphTexts(flow.body)).not.toContain('Form controls');
  });

  it('renders a sheet with controls to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(
      buildXlsx({
        rows: [['x']],
        formControls: [{ name: 'Agree', objectType: 'CheckBox', checked: true }],
      }),
      {
        fonts: {
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
        },
      },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
