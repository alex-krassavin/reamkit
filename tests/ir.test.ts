import { describe, expect, it } from 'vitest';

import {
  ConversionLossError,
  FEATURES,
  ResourceStore,
  eighthPtToPt,
  emuToPt,
  featureWithin,
  formatLoss,
  halfPtToPt,
  inchToPt,
  mmToPt,
  pt,
  pxToPt,
  twipsToPt,
} from '@/core/ir';

describe('IR units — canonical Pt', () => {
  it('converts format-native units to points', () => {
    expect(twipsToPt(20)).toBe(1);
    expect(twipsToPt(11906)).toBeCloseTo(595.3, 5); // A4 width in dxa
    expect(halfPtToPt(24)).toBe(12); // w:sz 24 → 12 pt font
    expect(eighthPtToPt(8)).toBe(1); // border w:sz 8 → 1 pt
    expect(emuToPt(914400)).toBe(72); // 1 inch
    expect(pxToPt(96)).toBe(72); // 1 inch of CSS px
    expect(inchToPt(1)).toBe(72);
    expect(mmToPt(25.4)).toBeCloseTo(72, 10);
  });

  it('pt() brands without changing the value', () => {
    const v = pt(12.5);
    expect(v).toBe(12.5);
    // Pt is assignable to number — arithmetic reads freely.
    expect(v * 2).toBe(25);
  });
});

describe('IR ResourceStore — content-addressed', () => {
  it('round-trips bytes by id', () => {
    const store = new ResourceStore();
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const id = store.put(bytes);
    expect(store.get(id)).toBe(bytes);
    expect(store.has(id)).toBe(true);
    expect(store.size).toBe(1);
  });

  it('deduplicates identical bytes (same id, single entry)', () => {
    const store = new ResourceStore();
    const a = store.put(new Uint8Array([9, 9, 9]));
    const b = store.put(new Uint8Array([9, 9, 9]));
    expect(b).toBe(a);
    expect(store.size).toBe(1);
    expect(store.ids()).toEqual([a]);
  });

  it('gives different bytes different ids', () => {
    const store = new ResourceStore();
    const a = store.put(new Uint8Array([1]));
    const b = store.put(new Uint8Array([2]));
    expect(a).not.toBe(b);
    expect(store.size).toBe(2);
  });

  it('ids are deterministic across stores (content-addressed)', () => {
    const s1 = new ResourceStore();
    const s2 = new ResourceStore();
    const bytes = new Uint8Array([5, 6, 7, 8, 9]);
    expect(s1.put(bytes)).toBe(s2.put(new Uint8Array([5, 6, 7, 8, 9])));
  });

  it('unknown id → undefined', () => {
    const store = new ResourceStore();
    expect(store.get('r-deadbeef-0' as never)).toBeUndefined();
  });
});

describe('IR features', () => {
  it('registry exposes hierarchical dot-names', () => {
    expect(FEATURES.tablesNested).toBe('tables.nested');
    expect(FEATURES.fontsSubstitution).toBe('fonts.substitution');
  });

  it('featureWithin matches exact and nested, not prefixes of words', () => {
    expect(featureWithin('tables.nested', 'tables')).toBe(true);
    expect(featureWithin('tables', 'tables')).toBe(true);
    expect(featureWithin('tablesNested', 'tables')).toBe(false); // no dot → different name
    expect(featureWithin('text', 'tables')).toBe(false);
  });
});

describe('IR loss reporting', () => {
  it('formats a loss with and without location', () => {
    expect(
      formatLoss({
        severity: 'substituted',
        feature: 'fonts.substitution',
        detail: 'Calibri → Roboto',
      }),
    ).toBe('[substituted] fonts.substitution: Calibri → Roboto');
    expect(
      formatLoss({
        severity: 'dropped',
        feature: 'math',
        detail: 'OMML equation',
        where: 'page 3',
      }),
    ).toBe('[dropped] math at page 3: OMML equation');
  });

  it('ConversionLossError carries the loss and a strict-mode message', () => {
    const err = new ConversionLossError({
      severity: 'dropped',
      feature: 'math',
      detail: 'OMML equation',
    });
    expect(err.name).toBe('ConversionLossError');
    expect(err.loss.feature).toBe('math');
    expect(err.message).toContain('strict conversion failed');
    expect(err.message).toContain('[dropped] math');
  });
});
