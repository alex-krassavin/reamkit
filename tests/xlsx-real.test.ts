// Real .xlsx documents, from real producers (tests/fixtures/real/NOTICE.md).
//
// Everything else in the xlsx suite is built by tests/fixtures/build-xlsx.ts,
// which emits exactly the dialect our parsers expect — so those tests prove the
// parser can read our own writer. These do not: each file below came from an
// upstream bug report and carries a dialect, a malformation or a scale that we
// could not honestly synthesize, and each one broke us before it was adopted.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import type { FontBytesByVariant } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { convertXlsxToPdfSync } from '@/core/converter';
import { Ream } from '@/core/converter/ream';
import { readXlsx, readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

const here = dirname(fileURLToPath(import.meta.url));
const load = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(here, 'fixtures/real', name)));

const FONTS: { fonts: FontBytesByVariant } = {
  fonts: {
    regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
    bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
    italic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Italic.ttf'))),
    boldItalic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-BoldItalic.ttf'))),
  },
};

const pageCount = (pdf: Uint8Array): number =>
  (new TextDecoder('latin1').decode(pdf).match(/\/MediaBox/g) ?? []).length;

/** Cells carrying visible text in the projected document. */
function textCells(bytes: Uint8Array): Array<string> {
  const { doc } = readXlsx(bytes);
  const out: Array<string> = [];
  for (const element of doc.body) {
    if (element.kind !== 'table') continue;
    for (const row of element.table.rows) {
      for (const cell of row.cells) {
        const text = cell.content
          .map((c) => (c.kind === 'paragraph' ? c.paragraph.runs.map((r) => r.text).join('') : ''))
          .join('')
          .trim();
        if (text.length > 0) out.push(text);
      }
    }
  }
  return out;
}

describe('real documents: package-level tolerance', () => {
  it('tdf76115.xlsx — backslash ZIP separators, worksheet outside xl/worksheets/', () => {
    const pkg = OpcPackage.open(load('tdf76115.xlsx'));
    expect(pkg.listParts()).toContain('_rels/.rels');
    expect(pkg.getMainDocumentPath()).toBe('xl/workbook.xml');
    expect(textCells(load('tdf76115.xlsx')).length).toBeGreaterThan(1000);
  });

  it('tdf76115.xlsx — the format sniffer recognises it too', () => {
    // The sniffs scan the raw ZIP bytes for a part name without unzipping, so
    // they were blind to the backslash spelling the OPC layer now normalizes:
    // the reader could read this document but the public entry point refused
    // to dispatch to it.
    expect(() => Ream.parse(load('tdf76115.xlsx'))).not.toThrow();
    expect(Ream.parse(load('tdf76115.xlsx')).format).toBe('xlsx');
  });

  it('tdf82984_zip64XLSXImport.xlsx — zip64 size sentinels are not real sizes', () => {
    // Every entry in this 4.7 KB archive declares originalSize 0xFFFFFFFF, the
    // zip64 "see the extra field" sentinel. Read as a literal 4 GiB it trips
    // the per-entry bomb guard and the whole document is refused.
    expect(() => OpcPackage.open(load('tdf82984_zip64XLSXImport.xlsx'))).not.toThrow();
    expect(textCells(load('tdf82984_zip64XLSXImport.xlsx')).length).toBeGreaterThan(0);
  });
});

describe('real documents: SpreadsheetML dialects', () => {
  it('tdf122336.xlsx — x: prefix, GUID r:id, unparseable cell refs', () => {
    const cells = textCells(load('tdf122336.xlsx'));
    // Blank before the fixes: the prefix hid nothing (the parsers strip it),
    // but r="11_2" dropped every cell on the floor.
    expect(cells.length).toBeGreaterThanOrEqual(19);
    expect(cells).toContain('Van Rompaey Marcus');

    // The refs look like <column>_<row> and the row half even agrees, but the
    // columns they name are wrong: the header row is 11 labels and the data row
    // 11 values, and only document order files each value under its own
    // heading. Read as columns, the start time landed under "Klantnaam" — and
    // then vanished off the right of the page.
    //
    // The sheet is wider than the page, so it comes out as one table per
    // column band; row i of the sheet is row i of each of them.
    const { doc } = readXlsx(load('tdf122336.xlsx'));
    const tables = doc.body.filter((e) => e.kind === 'table');
    // A band whose trailing rows draw nothing is trimmed, so later bands can be
    // shorter than the first — ask each for the row only if it has one.
    const row = (i: number): Array<string> =>
      tables.flatMap((t) =>
        (t.table.rows[i]?.cells ?? []).map((c) =>
          c.content
            .map((b) =>
              b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text).join('') : '',
            )
            .join(''),
        ),
      );
    // The projection clips off a per-character estimate whose buckets
    // over-charge, so it keeps a tenth of the column in hand and leaves the
    // exact cut to the layout: the heading survives even here.
    expect(row(0).slice(0, 3)).toEqual(['Uitvoeringsdatum', 'Starttijd', 'Eindtijd']);

    // …but that is the reader with no render font, where a column is measured
    // in Excel's own 7px digit. Told the font it will actually be drawn in, the
    // columns grow — and the TEXT must not be re-scaled with them. charWidthUnits
    // already reports each character's width in Excel's unit, so measuring the
    // text in the render font's digit too applied 123.54/105 twice and cut
    // "Uitvoeringsdatum" one glyph short inside a column 5% WIDER than the
    // reference's, with 7pt to spare.
    const drawn = projectSheetDoc(readXlsxToSheetDoc(load('tdf122336.xlsx')), {
      digitWidthPt: 6.18,
    });
    const firstRow = drawn.body.flatMap((e) =>
      e.kind === 'table'
        ? (e.table.rows[0]?.cells ?? []).map((c) =>
            c.content
              .map((b) =>
                b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text).join('') : '',
              )
              .join(''),
          )
        : [],
    );
    expect(firstRow[0]).toBe('Uitvoeringsdatum');
    // "Bevestigd via App?" in a column that holds about two thirds of it: the
    // estimate keeps a couple of characters more than fit, and the layout's own
    // measurement takes them back — the drawn page ends at "Bevestigd vi",
    // exactly where LibreOffice ends it.
    expect(firstRow[6]).toBe('Bevestigd via');
    expect(row(1).slice(0, 3)).toEqual(['12/25/2018', '11:30', '14:30']);

    // `<font/><font><b/></font>`: the empty element parses to a string rather
    // than an object, and skipping it moved the bold font to index 0 — so
    // `fontId="1"` resolved to nothing and the header row lost its weight.
    const sd = readXlsxToSheetDoc(load('tdf122336.xlsx'));
    expect(sd.styles.fonts).toHaveLength(2);
    expect(sd.styles.fonts[1]?.bold).toBe(true);
    expect(tables[0]!.table.rows[0]!.cells[0]!.content).toMatchObject([
      { paragraph: { runs: [{ properties: { bold: true } }] } },
    ]);
  });

  it('tdf111980_radioButtons.xlsx — <control> reaches ActiveX, not just form controls', () => {
    // §18.3.1.19 <control> resolves to BOTH kinds: a form control's ctrlProps
    // part and an ActiveX control's activeX#.xml. Only the relationship target
    // says which, and reading an ocx part as a ctrlProps one yields nothing —
    // so this sheet of option buttons came out as five bare names with their
    // captions and states sitting unread in the .bin beside them.
    const sheet = readXlsxToSheetDoc(load('tdf111980_radioButtons.xlsx')).sheets[0]!;
    const controls = sheet.activeXControls ?? [];
    expect(controls).toHaveLength(5);
    // Typed from the class id — <control> carries no progId to type it by.
    expect(controls.every((c) => c.type === 'option')).toBe(true);
    expect(controls.map((c) => c.caption)).toEqual([
      'ActiveX 3',
      'ActiveX nogroup2',
      'ActiveX nogroup2',
      'ActiveX button2',
      'ActiveX button1',
    ]);
    // Exactly one of the group is selected, and the group name came through.
    expect(controls.filter((c) => c.value === '1').map((c) => c.groupName)).toEqual(['Sheet1']);

    // Six more controls sit beside them, declared ONLY in the legacy VML: five
    // Forms-toolbar radio buttons and the group box around two of them. They
    // have no `<control>` entry and no ctrlProps part, so reading just the
    // `<controls>` list lost them without a word — LibreOffice draws all six.
    // Only actual control types: a cell comment is a VML shape too
    // (`ObjectType="Note"`) and must not be listed as a control.
    // The VML textbox is the control's CAPTION; `name` keeps it too, for the
    // listing a control without geometry falls back to. What is drawn is the
    // caption alone — `<control name>` is an identifier, never a label.
    const shown = (name: string) => ({ name, caption: name });
    expect(sheet.formControls?.map((c) => ({ ...c, box: undefined }))).toEqual([
      {
        objectType: 'Radio',
        ...shown('Form button1'),
        checked: true,
        fontSizePt: 8,
        box: undefined,
      },
      { objectType: 'Radio', ...shown('Form button2'), fontSizePt: 8, box: undefined },
      { objectType: 'GBox', ...shown('Group Box 7'), fontSizePt: 8, box: undefined },
      { objectType: 'Radio', ...shown('Form groupbox1'), fontSizePt: 8, box: undefined },
      {
        objectType: 'Radio',
        ...shown('Form groupbox2'),
        checked: true,
        fontSizePt: 8,
        box: undefined,
      },
      { objectType: 'Radio', ...shown('Form outside groupbox3'), fontSizePt: 8, box: undefined },
    ]);

    // Every one of them knows where it goes. The box is what turns the listing
    // into a drawing: the group box is 209×88pt at 441pt across the sheet, and
    // without it all eleven controls collapsed to a text list at the origin.
    expect(sheet.formControls?.[2]?.box).toEqual({
      xPt: 441,
      yPt: 7.5,
      widthPt: 209.25,
      heightPt: 87.75,
    });
    // An ActiveX control's box lives in the `Pict` shape sharing its shapeId —
    // its own part carries a class id and nothing else.
    expect(sheet.activeXControls?.map((c) => c.box?.xPt)).toEqual([564.75, 129, 131.25, 0, 2.25]);
  });

  it('AverageTaxRates.xlsx — hidden column and rows stay out of the render', () => {
    // The sheet hides a currency column between two visible ones and seven rows
    // in the middle of its table. Rendering them showed a column no other
    // reader shows, an extra data row, a second copy of the header band, and
    // broke the table's borders around them.
    const sheet = readXlsxToSheetDoc(load('AverageTaxRates.xlsx')).sheets[0]!;
    expect(sheet.grid.columns.filter((c) => c.hidden)).toHaveLength(1);
    expect(sheet.grid.rowHeights.filter((r) => r.hidden)).toHaveLength(7);

    const { doc } = readXlsx(load('AverageTaxRates.xlsx'));
    const table = doc.body.find((e) => e.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    // Twelve printed columns: the label plus 1997..2007. Thirteen would mean the
    // hidden currency column came back.
    expect(table.table.grid).toHaveLength(12);
    // The hidden rows carry these; none may appear.
    const text = textCells(load('AverageTaxRates.xlsx'));
    expect(text).not.toContain('4,726');
    expect(text).not.toContain('6,000');
  });

  it('tdf58243.xlsx — three things a page-count check cannot see', () => {
    const sd = readXlsxToSheetDoc(load('tdf58243.xlsx'));

    // §18.8.3 <color indexed="10"> — the legacy palette, still what Excel writes
    // for a colour picked from the classic dropdown. Ignoring the attribute
    // rendered these headers black where every other reader shows red.
    expect(sd.styles.fonts.some((f) => f.colorHex === 'FF0000')).toBe(true);

    // The dropdown arrow of a data-validation `list` cell is a selection
    // affordance; neither Excel nor LibreOffice prints it, and drawing it also
    // reserved a gutter and a minimum height that cost a whole page. The flag
    // stays on the cell — the HTML writer renders an interactive view and wants
    // it — but the paginated layout must not paint it.
    expect(sd.sheets[0]!.grid.dataValidations?.some((v) => v.type === 'list')).toBe(true);
    const { doc } = readXlsx(load('tdf58243.xlsx'));
    const table = doc.body.find((e) => e.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows.some((r) => r.cells.some((c) => c.properties.dropdown))).toBe(true);
    // Two pages, the same as LibreOffice. The sheet asks to fit 2×2 pages and
    // the closed-form scale gave 52 %, which spills the last three rows onto a
    // third page; fit-to-page now paginates for real and lands on 50 %.
    expect(pageCount(convertXlsxToPdfSync(load('tdf58243.xlsx'), FONTS))).toBe(2);

    // A literal CR LF inside the centre header region is a line break, not text
    // — drawn as text it came out as a missing-glyph box mid-title.
    // (The regions themselves share one line, on tab stops, so the break shows
    // as a second LINE of the band rather than a second centred paragraph.)
    const bands = [...(doc.headersFooters?.values() ?? [])].flat();
    const deepest = Math.max(
      ...[...(doc.headersFooters?.values() ?? [])].map(
        (band) => band.filter((b) => b.kind === 'paragraph').length,
      ),
    );
    expect(deepest).toBeGreaterThanOrEqual(2);
    for (const b of bands) {
      if (b.kind !== 'paragraph') continue;
      for (const run of b.paragraph.runs) expect(run.text).not.toMatch(/[\r\n]/);
    }
  });

  it('open-as-read-only.xlsx — text overflows past the end of the used range', () => {
    // `<dimension ref="A1"/>`: one cell, one column, and a sentence far wider
    // than it. Excel and LibreOffice run the text across the empty grid to the
    // right — there is nothing there to block it — and print one line. With no
    // columns to run into, the cell kept its own width and the layout wrapped
    // the sentence into nine stacked lines.
    const { doc } = readXlsx(load('open-as-read-only.xlsx'));
    const table = doc.body.find((e) => e.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.grid.length).toBeGreaterThan(1);
    expect(table.table.rows[0]!.cells[0]!.properties.colSpan).toBe(table.table.grid.length);
    // The widened grid must still fit the page, or it is split into bands and
    // the document gains a blank second page LibreOffice does not print.
    expect(pageCount(convertXlsxToPdfSync(load('open-as-read-only.xlsx'), FONTS))).toBe(1);
  });

  it('tdf171828 — overflow runs across a decorated but empty neighbour', () => {
    // The labels sit in a filled, top-ruled block: every neighbour to the right
    // is empty but carries the band's fill and rule. Treating any decoration as
    // a blocker clipped them to their own narrow column — "Kre" for
    // "Kreditsumme" — where every other reader prints the label in full.
    const cells = textCells(load('tdf171828_fail_to_import_file.xlsx'));
    expect(cells).toContain('Kreditsumme');
    expect(cells).toContain('Zahlungsbeginn');

    // §18.2.19: the third sheet is state="hidden" — a lookup table nobody
    // prints. Excel and LibreOffice leave it out; we printed two pages of
    // working data at the end of the document.
    const sd0 = readXlsxToSheetDoc(load('tdf171828_fail_to_import_file.xlsx'));
    expect(sd0.sheets.map((s) => s.hidden ?? false)).toEqual([false, false, true]);
    expect(cells).not.toContain('Hitab');

    // The block is styled entirely from the workbook theme — `theme="2"
    // tint="-0.5"` and friends. Unresolved, those colours parse to nothing and
    // a solid fill with no foreground paints nothing: the whole header block
    // came out white where every other reader shows it khaki.
    const sd = readXlsxToSheetDoc(load('tdf171828_fail_to_import_file.xlsx'));
    expect(sd.styles.fills.some((f) => f.fgColorHex === '948A54')).toBe(true);
  });

  it('duplicate-filename.xlsx — t="inlineStr" written into <v>', () => {
    expect(textCells(load('duplicate-filename.xlsx'))).toContain('v2');
  });
});

describe('real documents: scale and amplification', () => {
  it('53105.xlsx — all 16 384 columns render, none clipped', () => {
    // Two rows across the full column space. A 1024-column cap of ours cut this
    // to 103 pages where Excel and LibreOffice print 1639; the sheet is not
    // pathological, only wide, and the memory bound is the total-cell budget
    // that a 2-row sheet comes nowhere near.
    const { doc, losses } = readXlsx(load('53105.xlsx'));
    let columns = 0;
    for (const element of doc.body) {
      if (element.kind === 'table') columns += element.table.grid.length;
    }
    expect(columns).toBe(16_384);
    expect(losses.filter((l) => /column/i.test(l.detail))).toEqual([]);
  });

  it('too-many-cols-rows.xlsx — A1:XFE16777217 from a 5 KB file stays bounded', () => {
    const { doc, losses } = readXlsx(load('too-many-cols-rows.xlsx'));
    let cells = 0;
    for (const element of doc.body) {
      if (element.kind === 'table') for (const row of element.table.rows) cells += row.cells.length;
    }
    expect(cells).toBeLessThanOrEqual(1_000_000);
    expect(losses.filter((l) => l.severity === 'dropped').length).toBeGreaterThanOrEqual(2);

    // And it says WHY. This sheet asks for more than SpreadsheetML has — a cell
    // at XFE is past column XFD, the last one the format defines, so no reader
    // can put it anywhere the file names. Reporting that as a memory guard
    // would tell the reader to buy RAM for a file that is malformed.
    const detail = losses.map((l) => l.detail).join('\n');
    expect(detail).toContain('past column XFD');
    expect(detail).toContain('past row 1048576');
    expect(detail).not.toContain('memory guard');
  });
});
