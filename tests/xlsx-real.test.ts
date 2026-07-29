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
  });

  it('tdf111980_radioButtons.xlsx — <control> reaches ActiveX, not just form controls', () => {
    // §18.3.1.19 <control> resolves to BOTH kinds: a form control's ctrlProps
    // part and an ActiveX control's activeX#.xml. Only the relationship target
    // says which, and reading an ocx part as a ctrlProps one yields nothing —
    // so this sheet of option buttons came out as five bare names with their
    // captions and states sitting unread in the .bin beside them.
    const sheet = readXlsxToSheetDoc(load('tdf111980_radioButtons.xlsx')).sheets[0]!;
    expect(sheet.formControls ?? []).toHaveLength(0);
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
    // Pinned, not argued: removing the button did not by itself close the
    // page-count gap to LibreOffice's 2, so the remaining difference is
    // something else and this records where we stand.
    expect(pageCount(convertXlsxToPdfSync(load('tdf58243.xlsx'), FONTS))).toBe(3);

    // A literal CR LF inside the centre header region is a line break, not text
    // — drawn as text it came out as a missing-glyph box mid-title.
    const bands = [...(doc.headersFooters?.values() ?? [])].flat();
    const centre = bands.filter(
      (b) => b.kind === 'paragraph' && b.paragraph.properties.alignment === 'center',
    );
    expect(centre.length).toBeGreaterThanOrEqual(2);
    for (const b of bands) {
      if (b.kind !== 'paragraph') continue;
      for (const run of b.paragraph.runs) expect(run.text).not.toMatch(/[\r\n]/);
    }
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
  });
});
