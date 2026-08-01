// Rich values — the value a cell really holds when `<v>` holds a stand-in.
// §18.3.1.4's `vm` walks xl/metadata.xml → futureMetadata → rdrichvalue.xml →
// rdrichvaluestructure.xml. Only the `_error` structure resolves to anything;
// everything else leaves the legacy value alone.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { BodyElement } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';
import { parseRichValueText } from '@/excel/rich-value';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

// The three parts as Spill.xlsx writes them, trimmed to what the walk reads.
const METADATA = `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="2">
    <metadataType name="XLDAPR" minSupportedVersion="120000"/>
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
  <valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata>
</metadata>`;

const STRUCTURES = `<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_error">
    <k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/>
  </s>
</rvStructures>`;

const values = (errorType: number): string =>
  `<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
     <rv s="0"><v>0</v><v>${errorType}</v><v>3</v><v>1</v></rv>
   </rvData>`;

function texts(body: ReadonlyArray<BodyElement>): Array<string> {
  const out: Array<string> = [];
  for (const el of body) {
    if (el.kind !== 'table') continue;
    for (const row of el.table.rows)
      for (const cell of row.cells)
        for (const block of cell.content)
          if (block.kind === 'paragraph')
            out.push(block.paragraph.runs.map((r) => r.text).join(''));
  }
  return out;
}

describe('rich values (§18.3.1.4 vm)', () => {
  it('walks metadata → rich value → structure to name the error', () => {
    // errorType 8 with an offset to the cell in the way: a blocked spill.
    const map = parseRichValueText(enc(METADATA), enc(STRUCTURES), enc(values(8)));
    // The cell's `vm` is 1-based over <valueMetadata>.
    expect(map.get(1)).toBe('#SPILL!');
    expect(map.get(0)).toBeUndefined();
  });

  it('leaves an error code it cannot vouch for alone', () => {
    // The rest of the enumeration is not written down anywhere this code can
    // check, and a wrong error reads worse than the stand-in already in <v>.
    expect(parseRichValueText(enc(METADATA), enc(STRUCTURES), enc(values(3))).size).toBe(0);
  });

  it('resolves nothing when a part is missing (byte-zero)', () => {
    expect(parseRichValueText(undefined, enc(STRUCTURES), enc(values(8))).size).toBe(0);
    expect(parseRichValueText(enc(METADATA), undefined, enc(values(8))).size).toBe(0);
    expect(parseRichValueText(enc(METADATA), enc(STRUCTURES), undefined).size).toBe(0);
  });

  it('ignores a value metadata record pointing at another type', () => {
    // `rc t="1"` is XLDAPR — a dynamic-array marker, not a rich value.
    const other = METADATA.replace(
      '<rc t="2" v="0"/></bk></valueMetadata>',
      '<rc t="1" v="0"/></bk></valueMetadata>',
    );
    expect(parseRichValueText(enc(other), enc(STRUCTURES), enc(values(8))).size).toBe(0);
  });
});

describe('a blocked spill end to end (Spill.xlsx)', () => {
  const bytes = new Uint8Array(readFileSync('tests/fixtures/real/Spill.xlsx'));

  it('prints the error Excel prints, not the one written for readers that predate it', () => {
    // D2, E2 and G2 each hold a dynamic array whose spill range runs into the
    // text in row 5, and each caches `#VALUE!` for a reader that knows nothing
    // of rich values. The rich value records errorType 8 three rows down —
    // which is exactly where the blocking text sits.
    const grid = readXlsxToSheetDoc(bytes).sheets[0]!.grid;
    const errors = grid.cells.filter((c) => c.type === 'e');
    expect(errors).toHaveLength(3);
    expect(errors.map((c) => c.rawValue)).toEqual(['#SPILL!', '#SPILL!', '#SPILL!']);
    expect(errors.map((c) => c.valueMetadataIndex)).toEqual([1, 1, 1]);

    const rendered = texts(Ream.parse(bytes).flow.body);
    expect(rendered).toContain('#SPILL!');
    expect(rendered).not.toContain('#VALUE!');
  });

  it('leaves the cells that spilled successfully alone', () => {
    // C, F and H hold the same arrays with nothing in the way, over the source
    // column A: four real cached values that must not be touched.
    const grid = readXlsxToSheetDoc(bytes).sheets[0]!.grid;
    const row2 = grid.cells.filter((c) => c.row === 1 && c.type !== 'e');
    expect(row2.map((c) => c.rawValue)).toEqual(['10', '10', '10', '10']);
    expect(row2.map((c) => c.valueMetadataIndex)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
  });
});
