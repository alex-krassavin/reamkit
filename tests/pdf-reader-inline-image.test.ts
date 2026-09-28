// §8.9.7 — an inline image's bytes, between `ID` and `EI`, read to where the
// image's own encoding ends rather than to the first `EI` its bytes happen to
// spell.

import { describe, expect, it } from 'vitest';

import { interpretContent } from '@/pdf-reader/content';

/** A content stream as the bytes it is, one byte per character. */
const bytesOf = (content: string): Uint8Array =>
  Uint8Array.from([...content].map((c) => c.charCodeAt(0)));

/** The data of the one inline image a content stream draws, as a byte string. */
function inlineData(content: string): string | undefined {
  const { images } = interpretContent(bytesOf(content), new Map());
  const data = images[0]?.inline?.data;
  return data === undefined ? undefined : String.fromCharCode(...data);
}

/**
 * The smallest JPEG the walk has to step through: a comment segment that
 * happens to spell ` EI `, a frame, a scan with its coded bytes, the end.
 */
const JPEG =
  '\xff\xd8' +
  '\xff\xfe\x00\x06 EI ' +
  '\xff\xc0\x00\x0b\x08\x00\x01\x00\x01\x01\x01\x11\x00' +
  '\xff\xda\x00\x08\x01\x01\x00\x00\x3f\x00' +
  '\xd2\xcf\xff\x00\x20' +
  '\xff\xd9';

describe('an inline image read to where its own encoding ends (§8.9.7)', () => {
  it('reads a JPEG to its end-of-image marker, past an EI its bytes spell', () => {
    // A JPEG is dense with EI: searched for, the end fell inside the picture
    // and what was left would not decode.
    const content = `q BI /W 1 /H 1 /CS /G /BPC 8 /F /DCT ID ${JPEG} EI Q`;
    expect(inlineData(content)).toBe(JPEG);
  });

  it('reads base-85 text to its ~>, past an EI it spells', () => {
    // Four grey samples whose base-85 is "!5EIC".
    const content = 'BI /W 2 /H 2 /CS /G /BPC 8 /F /A85 ID !5EIC~> EI';
    expect(inlineData(content)).toBe('!5EIC~>');
  });

  it('ends anything else at an EI that stands as a word of its own', () => {
    const content = 'BI /W 1 /H 1 /CS /G /BPC 8 /F /Fl ID abEIcd EI';
    expect(inlineData(content)?.startsWith('abEIcd')).toBe(true);
  });

  it('takes CR LF after ID as the one end of line it is', () => {
    // bug1065245.pdf writes `ID` CR LF, and read as one byte and data the LF
    // went in front of each JPEG's start marker: none of its banners decoded.
    const content = `BI /W 1 /H 1 /CS /G /BPC 8 /F /DCT ID\r\n${JPEG}\r\nEI`;
    expect(inlineData(content)).toBe(JPEG);
  });

  it('keeps the one byte after ID for an image it measures', () => {
    // Unfiltered, the data is exactly W·H bytes and may itself start with LF.
    const content = 'BI /W 2 /H 1 /CS /G /BPC 8 ID\r\n\nx EI';
    expect(inlineData(content)).toBe('\n\n');
  });
});
