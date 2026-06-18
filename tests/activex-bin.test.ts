// E-SHEET W10 tail — the MS-OFORMS MorphDataControl binary (activeX#.bin) parse.
// A control saved with ax:persistence="persistStreamInit" keeps no <ax:ocxPr>;
// its caption / value / group name live in the binary stream. The bytes below are
// the REAL stream from LibreOffice's `activex_checkbox.xlsx` fixture (the OOXML
// ActiveX format Excel also writes), so this validates the layout against a real
// Office file — not a self-consistent fixture.

import { describe, expect, it } from 'vitest';

import { activeXBinRelId, parseActiveXBin } from '@/excel/activex-parser';

// hex → bytes (whitespace ignored).
function hex(s: string): Uint8Array {
  const clean = s.replace(/\s+/g, '');
  const out = new Uint8Array(clean.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(clean.slice(i * 2, i * 2 + 2), 16);
  return out;
}

// The 116-byte activeX1.bin from activex_checkbox.xlsx: a Forms.CheckBox
// (classid {8BD21D40-…}) with Caption "Custom Caption", Value "1" (checked) and
// GroupName "Sheet1".
const CHECKBOX_BIN =
  '401dd28b42ecce119e0d00aa006002f3000244004601c080010000000d0000800a0000800400' +
  '0000010000800e00008006000080d91000000707000031000000437573746f6d204361707469' +
  '6f6e00005368656574312043000218003500000007000080e1000000ee02000043616c696272' +
  '6900';

describe('ActiveX .bin (MS-OFORMS MorphDataControl) — W10 tail', () => {
  it('recovers caption / value / group name from a real Office checkbox stream', () => {
    expect(parseActiveXBin(hex(CHECKBOX_BIN))).toEqual({
      caption: 'Custom Caption',
      value: '1',
      groupName: 'Sheet1',
    });
  });

  it('tolerates a stream without the 16-byte classid prefix', () => {
    // Drop the leading GUID — the control still parses from the version onward.
    expect(parseActiveXBin(hex(CHECKBOX_BIN).slice(16))).toEqual({
      caption: 'Custom Caption',
      value: '1',
      groupName: 'Sheet1',
    });
  });

  it('returns nothing for a CFB-storage .bin (the unhandled variant)', () => {
    expect(parseActiveXBin(hex('d0cf11e0a1b11ae1' + '00'.repeat(60)))).toEqual({});
  });

  it('returns nothing for a structurally implausible blob (never a wrong caption)', () => {
    expect(parseActiveXBin(hex('deadbeef'))).toEqual({});
    expect(parseActiveXBin(new Uint8Array(0))).toEqual({});
  });

  it('reads the bin relationship id only for a stream/storage-persisted control', () => {
    const enc = new TextEncoder();
    const streamInit = enc.encode(
      '<ax:ocx ax:persistence="persistStreamInit" r:id="rId7" xmlns:ax="x" xmlns:r="y"/>',
    );
    expect(activeXBinRelId(streamInit)).toBe('rId7');
    // A property-bag control keeps its state inline — no .bin to fetch.
    const bag = enc.encode(
      '<ax:ocx ax:persistence="persistPropertyBag" r:id="rId7" xmlns:ax="x" xmlns:r="y"/>',
    );
    expect(activeXBinRelId(bag)).toBeUndefined();
  });
});
