// E-SHEET W8 — form controls. Checkboxes, option buttons, spinners etc. are
// declared on the worksheet (the x14 extLst <controls>) and point through a
// relationship at a ctrlProp part carrying their objectType + state. The reader
// resolves them and the projection lists each in a "Form controls" section after
// the grid with a type-appropriate affordance. Render-only — not written back.
// ActiveX (OLE) controls are a documented graceful loss.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { sheetDrawings } from './fixtures/sheet-drawings';
import type { BodyElement } from '@/core/document-model';
import { parseFormControlProps } from '@/excel/form-control-parser';
import { parseVmlDrawing } from '@/excel/vml-drawing';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
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

describe('legacy VML "Print object" (E-SHEET W8)', () => {
  const vml = (printObject: string): Uint8Array =>
    enc(
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel">
         <v:shape id="_x0000_s1025" type="#_x0000_t201"
            style='position:absolute;margin-left:68.25pt;margin-top:48pt;width:106.5pt;height:58.5pt'>
           <v:textbox><div><font>Button 1</font></div></v:textbox>
           <x:ClientData ObjectType="Button">${printObject}</x:ClientData>
         </v:shape>
       </xml>`,
    );

  it('keeps a shape that says nothing about printing', () => {
    const drawing = parseVmlDrawing(vml(''));
    expect(drawing.controls.map((c) => c.caption)).toEqual(['Button 1']);
    expect(drawing.nonPrinting.size).toBe(0);
    expect(drawing.boxes.get('1025')?.widthPt).toBeCloseTo(106.5, 2);
  });

  it('reads a length with a long run of blanks inside in time linear in it', () => {
    // The two `\s*` round an empty unit split the run every way between them.
    const xml = new TextDecoder()
      .decode(vml(''))
      .replace('width:106.5pt', `width:0${' '.repeat(100_000)}!`);
    const start = performance.now();
    const drawing = parseVmlDrawing(enc(xml));
    expect(performance.now() - start).toBeLessThan(1000);
    expect(drawing.boxes.get('1025')).toBeUndefined();
  });

  it('drops one that clears it, and names the shape so its <control> goes too', () => {
    const drawing = parseVmlDrawing(vml('<x:PrintObject>False</x:PrintObject>'));
    expect(drawing.controls).toHaveLength(0);
    expect(drawing.nonPrinting.has('1025')).toBe(true);
  });
});

describe('a VML group (E-SHEET W8)', () => {
  // 45540_form_Footer.xlsx's shape: a group in points, its children in the
  // group's own units — 494 across 416.25pt, 183 down 137.25pt.
  const grouped = enc(
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel">
       <v:group id="_x0000_s2138" style='position:absolute;margin-left:7.5pt;margin-top:729pt;width:416.25pt;height:137.25pt'
          coordorigin="9,1206" coordsize="494,183">
         <v:shape id="CheckBox44" o:spid="_x0000_s2101" style='position:absolute;left:9;top:1206;width:22;height:23'>
           <x:ClientData ObjectType="Pict"/></v:shape>
         <v:shape id="CheckBox54" o:spid="_x0000_s2111" style='position:absolute;left:271;top:1226;width:22;height:23'>
           <x:ClientData ObjectType="Pict"/></v:shape>
       </v:group>
     </xml>`,
  );

  it("places the shapes it holds in the group's own space", () => {
    const boxes = parseVmlDrawing(grouped).boxes;
    const first = boxes.get('2101')!;
    expect(first.xPt).toBeCloseTo(7.5, 5);
    expect(first.yPt).toBeCloseTo(729, 5);
    expect(first.widthPt).toBeCloseTo((22 * 416.25) / 494, 5);
    expect(first.heightPt).toBeCloseTo(17.25, 5);
    const second = boxes.get('2111')!;
    expect(second.xPt).toBeCloseTo(7.5 + (262 * 416.25) / 494, 5);
    expect(second.yPt).toBeCloseTo(744, 5);
  });

  it('stands a box placed by points on the rows Excel counted it by', () => {
    // Twenty rows of 20.1pt, which Excel draws 26px tall: a control it put 3px
    // into the eleventh row it wrote at 263px, 197.25pt. Our rows keep their
    // fraction, so that row starts at 201pt, and the control 3px below it.
    const flow = Ream.parse(
      buildXlsx({
        rows: Array.from({ length: 20 }, (_, i) => [`row ${i + 1}`]),
        rowHeights: Array.from({ length: 20 }, (_, i) => ({ row: i, heightPt: 20.1 })),
        legacyVmlXml:
          `<v:group style='position:absolute;margin-left:0;margin-top:197.25pt;width:150pt;height:30pt' coordorigin="0,0" coordsize="200,40">` +
          `<v:shape id="_x0000_s1025" style='position:absolute;left:10;top:0;width:100;height:23'>` +
          '<v:textbox><div><font>Pick me</font></div></v:textbox>' +
          '<x:ClientData ObjectType="Checkbox"/></v:shape></v:group>',
      }),
    ).flow;
    const caption = sheetDrawings(flow.body).flatMap((e) =>
      e.kind === 'shape' && e.shape.text ? [e.shape] : [],
    );
    expect(caption).toHaveLength(1);
    expect(caption[0]!.float?.posV?.offsetPt).toBeCloseTo(10 * 20.1 + 2.25, 2);
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
    // This one sits past the printable width, so it rides the second band —
    // shifted back to the left margin by exactly one page of it (487pt). The
    // page break between the bands is the projection's only paragraph.
    expect(at(564.75 - 487, 35.25)).toHaveLength(1);
    expect(flow.body.filter((e) => e.kind === 'paragraph')).toHaveLength(1);

    // The group box straddles the boundary, so its FRAME continues into the
    // second band — at a negative offset, since its left edge is back in the
    // first. Its caption does not come with it: it was printed where the
    // drawing starts, and printing it again puts the same label on two pages.
    const carried = at(441 - 487, 7.5);
    expect(carried).toHaveLength(1);
    expect(carried[0]?.text).toBeUndefined();
    expect(carried[0]?.line?.colorHex).toBe('808080');

    // And nothing is listed after the grid any more.
    const texts = paragraphTexts(flow.body);
    expect(texts).not.toContain('Form controls');
    expect(texts).not.toContain('ActiveX controls');
  });

  it('leaves off a control whose "Print object" is cleared (§18.3.1.20)', () => {
    // Excel's Print object checkbox is on by default; a control that clears it
    // is on screen only. button-form-control.xlsx says so twice — `print="0"`
    // on the controlPr and `<x:PrintObject>False</x:PrintObject>` in the VML —
    // and LibreOffice prints it as a blank page while we drew the button.
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        formControls: [
          { name: 'Shown', objectType: 'Button' },
          { name: 'Screen only', objectType: 'Button', print: false },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('Shown (Button)');
    expect(texts.some((t) => t.includes('Screen only'))).toBe(false);
  });

  it('ticks a check box with a cross, not a filled square', () => {
    // Excel and Calc both draw ☒. Filling the square the way an option button
    // fills its ring makes checked-vs-unchecked a difference in the amount of
    // black rather than a mark.
    const shapes = Ream.parse(
      new Uint8Array(readFileSync('tests/fixtures/real/singlecontrol.xlsx')),
    ).flow.body.flatMap((e) => (e.kind === 'shape' ? [e.shape] : []));
    const diagonals = shapes.filter(
      (s) => s.geometry.kind === 'preset' && s.geometry.preset === 'line',
    );
    expect(diagonals).toHaveLength(2);
    // One each way — the second is the first mirrored.
    expect(diagonals.map((s) => s.transform?.flipV ?? false)).toEqual([false, true]);
    expect(diagonals.every((s) => s.fill.kind === 'none')).toBe(true);
  });

  it('paginates a drawing anchored below the first page onto its own band', () => {
    // singlecontrol.xlsx has no cells at all and one check box 7331pt down —
    // nine pages past the only page its empty grid produces. Banded across but
    // not down, every trace of it fell off the document.
    const flow = Ream.parse(
      new Uint8Array(readFileSync('tests/fixtures/real/singlecontrol.xlsx')),
    ).flow;
    const caption = flow.body.flatMap((e) => (e.kind === 'shape' && e.shape.text ? [e.shape] : []));
    expect(caption).toHaveLength(1);
    // Its band is 9 pages down, so what is left is the remainder of the printable
    // height (785.2pt for A4 with this sheet's 1cm margins). It stands on the
    // cells its `<x:Anchor>` names — C517, 6px across and 12px down — measured
    // as our grid measures them: 7337.25pt down and 122.25pt across. The
    // shape's `style` says 7331.25 and 122.25, the writer's own measure of the
    // rows and columns it wrote the file from — the columns now agree.
    expect(caption[0]?.float?.posV?.offsetPt).toBeCloseTo(7337.25 - 9 * 785.2, 1);
    expect(caption[0]?.float?.posH?.offsetPt).toBeCloseTo(122.25, 2);
  });

  it("draws a control's caption, never its name", () => {
    // §18.3.1.19 `<control name>` is the shape's IDENTIFIER — Excel shows it in
    // the name box, never on the page. Falling back to it printed "CheckBox28"
    // across 45540_form_Header.xlsx's own text once per captionless check box,
    // forty times over.
    const bytes = new Uint8Array(readFileSync('tests/fixtures/real/45540_form_Header.xlsx'));
    const controls = readXlsxToSheetDoc(bytes).sheets[0]?.activeXControls ?? [];
    expect(controls.length).toBeGreaterThan(30);
    expect(controls.every((c) => c.caption === undefined)).toBe(true);
    expect(controls.some((c) => (c.name ?? '').startsWith('CheckBox'))).toBe(true);

    const drawn = Ream.parse(bytes).flow.body.flatMap((e) =>
      e.kind === 'shape' && e.shape.text
        ? e.shape.text.content.flatMap((b) =>
            b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : [],
          )
        : [],
    );
    expect(drawn.filter((t) => t.startsWith('CheckBox'))).toEqual([]);
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
