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

import { OpcPackage } from '@/core/opc';
import { Ream } from '@/core/converter/ream';
import { readXlsx } from '@/excel/xlsx-reader';

const here = dirname(fileURLToPath(import.meta.url));
const load = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(here, 'fixtures/real', name)));

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

  it('duplicate-filename.xlsx — t="inlineStr" written into <v>', () => {
    expect(textCells(load('duplicate-filename.xlsx'))).toContain('v2');
  });
});

describe('real documents: scale and amplification', () => {
  it('53105.xlsx — 16 384 declared columns clip, and say so', () => {
    const { doc, losses } = readXlsx(load('53105.xlsx'));
    expect(doc.body.length).toBeGreaterThan(0);
    const clip = losses.find((l) => /columns/.test(l.detail));
    expect(clip).toBeDefined();
    expect(clip!.severity).toBe('dropped');
    expect(clip!.detail).toContain('16384');
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
