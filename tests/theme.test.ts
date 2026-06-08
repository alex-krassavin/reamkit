import { describe, expect, it } from 'vitest';

import {
  DEFAULT_THEME_PALETTE,
  makeColorResolver,
  resolveSchemeName,
} from '@/ooxml/drawingml/colors';
import { parseTheme } from '@/ooxml/drawingml/theme-parser';

const enc = new TextEncoder();

const THEME_XML = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="ff0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;

describe('parseTheme', () => {
  it('reads srgbClr and sysClr (lastClr) scheme slots', () => {
    const theme = parseTheme(enc.encode(THEME_XML));
    expect(theme.get('accent1')).toBe('FF0000'); // upper-cased
    expect(theme.get('accent2')).toBe('00FF00');
    expect(theme.get('dk1')).toBe('000000'); // sysClr lastClr
    expect(theme.get('lt1')).toBe('FFFFFF');
    expect(theme.get('dk2')).toBe('1F3864');
    expect(theme.get('accent3')).toBeUndefined();
  });

  it('returns an empty map when there is no clrScheme', () => {
    const empty = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`;
    expect(parseTheme(enc.encode(empty)).size).toBe(0);
  });
});

describe('colour resolver', () => {
  it('passes srgb through and resolves scheme names via the palette', () => {
    const resolve = makeColorResolver(DEFAULT_THEME_PALETTE);
    expect(resolve({ srgb: '4472c4' })).toBe('4472C4');
    expect(resolve({ scheme: 'accent1' })).toBe('4472C4');
    expect(resolve({ scheme: 'accent2' })).toBe('ED7D31');
  });

  it('aliases tx1/bg1/tx2/bg2 onto dk/lt slots', () => {
    expect(resolveSchemeName('tx1')).toBe('dk1');
    expect(resolveSchemeName('bg1')).toBe('lt1');
    expect(resolveSchemeName('tx2')).toBe('dk2');
    expect(resolveSchemeName('bg2')).toBe('lt2');
    const resolve = makeColorResolver(DEFAULT_THEME_PALETTE);
    expect(resolve({ scheme: 'tx1' })).toBe('000000'); // dk1
    expect(resolve({ scheme: 'bg1' })).toBe('FFFFFF'); // lt1
  });

  it('returns undefined for an unknown scheme name', () => {
    expect(makeColorResolver(DEFAULT_THEME_PALETTE)({ scheme: 'phClr' })).toBeUndefined();
  });
});
