// E-SHEET W10 — ActiveX controls. An embedded ActiveX control's visible state
// lives in its xl/activeX/activeX#.xml property bag (<ax:ocxPr name value>),
// reached through a worksheet <oleObject> relationship. The reader resolves the
// progId → control type and the property bag → caption/value, and the projection
// lists them in an "ActiveX controls" section after the grid, like form controls.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { BodyElement } from '@/core/document-model';
import { activeXType, parseActiveX } from '@/excel/activex-parser';
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

describe('activeX property-bag parser (E-SHEET W10)', () => {
  it('reads caption / value / groupName from <ax:ocxPr>', () => {
    const xml = `<ax:ocx ax:classid="{x}" ax:persistence="persistPropertyBag"
      xmlns:ax="http://schemas.microsoft.com/office/2006/activeX">
      <ax:ocxPr ax:name="Caption" ax:value="I agree"/>
      <ax:ocxPr ax:name="Value" ax:value="1"/>
      <ax:ocxPr ax:name="GroupName" ax:value="g1"/>
    </ax:ocx>`;
    expect(parseActiveX(enc(xml))).toEqual({ caption: 'I agree', value: '1', groupName: 'g1' });
  });

  it('maps the oleObject progId to a control type', () => {
    expect(activeXType('Forms.CheckBox.1')).toBe('checkbox');
    expect(activeXType('Forms.CommandButton.1')).toBe('button');
    expect(activeXType('Forms.OptionButton.1')).toBe('option');
    expect(activeXType('Forms.TextBox.1')).toBe('textbox');
    expect(activeXType('Forms.SpinButton.1')).toBe('spin');
    expect(activeXType(undefined)).toBe('control');
  });
});

describe('ActiveX controls — end to end (E-SHEET W10)', () => {
  it('lists controls with type-appropriate affordances and their state', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        oleObjects: [
          { progId: 'Forms.CheckBox.1', caption: 'Agree', value: '1' },
          { progId: 'Forms.OptionButton.1', caption: 'Choice', value: '0' },
          { progId: 'Forms.CommandButton.1', caption: 'OK' },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('ActiveX controls');
    expect(texts).toContain('[x] Agree');
    expect(texts).toContain('( ) Choice');
    expect(texts).toContain('[ OK ]');
  });

  it('exposes the resolved controls on the SheetDoc', () => {
    const sheet = Ream.parse(
      buildXlsx({
        rows: [['x']],
        oleObjects: [{ progId: 'Forms.CheckBox.1', caption: 'C', value: '1' }],
      }),
    ).sheet;
    expect(sheet?.sheets[0]?.activeXControls).toEqual([
      { type: 'checkbox', caption: 'C', value: '1' },
    ]);
  });

  it('adds no ActiveX section to a sheet without controls (byte-zero)', () => {
    const flow = Ream.parse(buildXlsx({ rows: [['data']] })).flow;
    expect(paragraphTexts(flow.body)).not.toContain('ActiveX controls');
  });

  it('renders a sheet with an ActiveX control to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(
      buildXlsx({
        rows: [['x']],
        oleObjects: [{ progId: 'Forms.CommandButton.1', caption: 'Go' }],
      }),
      {
        fonts: {
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
        },
      },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
