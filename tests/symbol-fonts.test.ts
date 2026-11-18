import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { fromSymbolFont } from '@/core/metafile/symbol-fonts';
import { readDocx } from '@/word/docx-reader';

// The private-use spelling Word writes a symbol-font code in: the low byte
// plus 0xF000.
const pua = (low: number): string => String.fromCharCode(0xf000 + low);

describe('Symbol codes are the signs SymbolMT draws', () => {
  it('reads 0x2D as the minus sign, in either spelling', () => {
    expect(fromSymbolFont(pua(0x2d), 'Symbol')).toBe('−');
    expect(fromSymbolFont('-', 'Symbol')).toBe('−');
  });

  it('reads 0xB4 as the times sign and 0xD7 as the dot operator', () => {
    expect(fromSymbolFont(pua(0xb4), 'Symbol')).toBe('×');
    expect(fromSymbolFont(pua(0xd7), 'Symbol')).toBe('⋅');
  });

  it('leaves a hyphen in an ordinary font alone', () => {
    expect(fromSymbolFont('-', 'Times New Roman')).toBe('-');
  });

  it('reads a w:sym minus as the minus sign', () => {
    const body =
      '<w:p><w:r><w:t>a</w:t><w:sym w:font="Symbol" w:char="F02D"/><w:t>b</w:t></w:r></w:p>';
    const { doc } = readDocx(buildDocxFromBody(body));
    const text = doc.body.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : [],
    );
    expect(text.join('')).toBe('a−b');
  });
});
