import { describe, expect, it } from 'vitest';

import type { Style, StyleSheet } from '@/document-model';
import { resolveParagraphProperties, resolveRunProperties } from '@/style-cascade';

function sheet(styles: Array<Style>, defaults: Partial<StyleSheet> = {}): StyleSheet {
  return {
    defaultRunProperties: {},
    defaultParagraphProperties: {},
    ...defaults,
    styles: new Map(styles.map((s) => [s.id, s])),
  };
}

const STYLE_BOLD: Style = {
  id: 'Bold',
  type: 'character',
  isDefault: false,
  runProperties: { bold: true },
  paragraphProperties: {},
};

describe('resolveRunProperties', () => {
  it('falls back to hardcoded defaults when nothing is set', () => {
    const r = resolveRunProperties({}, {}, sheet([]));
    expect(r.bold).toBe(false);
    expect(r.italic).toBe(false);
    expect(r.fontSizeHalfPoints).toBe(22);
    expect(r.colorHex).toBe('000000');
  });

  it('applies docDefault run properties', () => {
    const r = resolveRunProperties(
      {},
      {},
      sheet([], { defaultRunProperties: { fontSizeHalfPoints: 28, colorHex: 'FF0000' } }),
    );
    expect(r.fontSizeHalfPoints).toBe(28);
    expect(r.colorHex).toBe('FF0000');
  });

  it('paragraph style rPr applies as default for runs in that paragraph', () => {
    const r = resolveRunProperties(
      {},
      { styleId: 'Heading1' },
      sheet([
        {
          id: 'Heading1',
          type: 'paragraph',
          isDefault: false,
          runProperties: { fontSizeHalfPoints: 32, bold: true },
          paragraphProperties: {},
        },
      ]),
    );
    expect(r.bold).toBe(true);
    expect(r.fontSizeHalfPoints).toBe(32);
  });

  it('character style overrides paragraph style', () => {
    const r = resolveRunProperties(
      { styleId: 'Bold' },
      { styleId: 'Heading1' },
      sheet([
        {
          id: 'Heading1',
          type: 'paragraph',
          isDefault: false,
          runProperties: { bold: false, fontSizeHalfPoints: 32 },
          paragraphProperties: {},
        },
        STYLE_BOLD,
      ]),
    );
    expect(r.bold).toBe(true);
    expect(r.fontSizeHalfPoints).toBe(32);
  });

  it('direct run properties take highest priority', () => {
    const r = resolveRunProperties({ styleId: 'Bold', bold: false }, {}, sheet([STYLE_BOLD]));
    expect(r.bold).toBe(false);
  });

  it('resolves a basedOn chain', () => {
    const r = resolveRunProperties(
      {},
      { styleId: 'Sub' },
      sheet([
        {
          id: 'Root',
          type: 'paragraph',
          isDefault: false,
          runProperties: { fontSizeHalfPoints: 22, italic: true },
          paragraphProperties: {},
        },
        {
          id: 'Sub',
          type: 'paragraph',
          basedOn: 'Root',
          isDefault: false,
          runProperties: { bold: true },
          paragraphProperties: {},
        },
      ]),
    );
    expect(r.italic).toBe(true);
    expect(r.bold).toBe(true);
    expect(r.fontSizeHalfPoints).toBe(22);
  });

  it('breaks a cycle through basedOn without infinite loop', () => {
    const r = resolveRunProperties(
      {},
      { styleId: 'A' },
      sheet([
        {
          id: 'A',
          type: 'paragraph',
          basedOn: 'B',
          isDefault: false,
          runProperties: { bold: true },
          paragraphProperties: {},
        },
        {
          id: 'B',
          type: 'paragraph',
          basedOn: 'A',
          isDefault: false,
          runProperties: { italic: true },
          paragraphProperties: {},
        },
      ]),
    );
    expect(r.bold).toBe(true);
    expect(r.italic).toBe(true);
  });
});

describe('resolveParagraphProperties', () => {
  it('returns hardcoded defaults when no style is referenced', () => {
    const p = resolveParagraphProperties({}, sheet([]));
    expect(p.alignment).toBe('left');
    expect(p.spacingBeforeTwips).toBe(0);
  });

  it('direct alignment overrides paragraph style alignment', () => {
    const p = resolveParagraphProperties(
      { styleId: 'Centered', alignment: 'right' },
      sheet([
        {
          id: 'Centered',
          type: 'paragraph',
          isDefault: false,
          runProperties: {},
          paragraphProperties: { alignment: 'center' },
        },
      ]),
    );
    expect(p.alignment).toBe('right');
  });
});
