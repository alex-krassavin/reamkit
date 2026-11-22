import { describe, expect, it } from 'vitest';

import { headingLevelOf } from '@/core/outline';
import { EMPTY_STYLE_SHEET, resolveParagraphProperties } from '@/core/style-cascade';

const level = (p: Parameters<typeof resolveParagraphProperties>[0]) =>
  headingLevelOf(resolveParagraphProperties(p, EMPTY_STYLE_SHEET));

describe('headingLevelOf', () => {
  it('maps outline levels 0–8 onto headings 1–6', () => {
    expect(level({ outlineLevel: 0 })).toBe(1);
    expect(level({ outlineLevel: 4 })).toBe(5);
    expect(level({ outlineLevel: 8 })).toBe(6);
  });

  it('falls back to a Heading N style id when no level is stated', () => {
    expect(level({ styleId: 'Heading3' })).toBe(3);
    expect(level({ styleId: 'Title' })).toBe(1);
  });

  it('takes a stated body-text level 9 at its word, whatever the style id', () => {
    expect(level({ styleId: 'Heading3', outlineLevel: 9 })).toBeUndefined();
  });
});
