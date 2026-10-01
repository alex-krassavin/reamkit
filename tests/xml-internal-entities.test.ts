// XML 1.0 §4.2.2 — the internal DTD subset. fast-xml-parser never registers
// these declarations, so every `&name;` used to reach the page as its own
// spelling: poc-xmlbomb-empty.xlsx drew "test123&a5;&a5;&a5;…" where the
// entities expand to nothing at all.

import { describe, expect, it } from 'vitest';

import { resolveInternalEntities } from '@/core/opc/xml-entities';

const BOMB = `<?xml version="1.0"?>
<!DOCTYPE foo[
<!ENTITY a0 "">
<!ENTITY a1 "&a0;&a0;&a0;&a0;">
<!ENTITY a2 "&a1;&a1;&a1;&a1;">
]>
<sst><si><t>test123&a2;&a2;</t></si></sst>`;

describe('internal DTD subset (§4.2.2)', () => {
  it('leaves a part with no DOCTYPE exactly as it was read', () => {
    const xml = '<sst><si><t>Tom &amp; Jerry</t></si></sst>';
    expect(resolveInternalEntities(xml)).toBe(xml);
  });

  it('expands a declared entity and drops the declaration with it', () => {
    const out = resolveInternalEntities(
      '<!DOCTYPE r[<!ENTITY who "Acme">]><r><t>&who; Ltd</t></r>',
    );
    expect(out).toBe('<r><t>Acme Ltd</t></r>');
  });

  it('resolves the nested chain of an empty bomb to nothing', () => {
    const out = resolveInternalEntities(BOMB);
    expect(out).not.toContain('&a2;');
    expect(out).not.toContain('<!DOCTYPE');
    expect(out).toContain('<t>test123</t>');
  });

  it('refuses an expansion that runs away, rather than building it', () => {
    // Four levels of ×8 over a 64-char seed is past the budget; the reference
    // resolves to nothing instead of a megabyte of text.
    const seed = 'x'.repeat(64);
    const decls = [`<!ENTITY b0 "${seed}">`];
    for (let i = 1; i <= 4; i += 1) {
      decls.push(`<!ENTITY b${String(i)} "${`&b${String(i - 1)};`.repeat(8)}">`);
    }
    const out = resolveInternalEntities(`<!DOCTYPE r[${decls.join('')}]><r>&b4;</r>`);
    expect(out).toBe('<r></r>');
  });

  it('survives an entity that references itself', () => {
    const out = resolveInternalEntities('<!DOCTYPE r[<!ENTITY loop "&loop;">]><r>&loop;</r>');
    expect(out).toBe('<r></r>');
  });

  it('drops a reference that names nothing', () => {
    expect(resolveInternalEntities('<!DOCTYPE r[<!ENTITY a "x">]><r>&a;&ghost;</r>')).toBe(
      '<r>x</r>',
    );
  });

  it('keeps the predefined and numeric references for the parser', () => {
    const out = resolveInternalEntities('<!DOCTYPE r[<!ENTITY a "x">]><r>&amp;&#10;&a;</r>');
    expect(out).toBe('<r>&amp;&#10;x</r>');
  });

  it('escapes markup an entity carries instead of splicing it in', () => {
    const out = resolveInternalEntities('<!DOCTYPE r[<!ENTITY evil "<b>&amp;</b>">]><r>&evil;</r>');
    expect(out).toBe('<r>&lt;b&gt;&amp;amp;&lt;/b&gt;</r>');
  });

  it('takes the first declaration that closes, past one that never does', () => {
    // The `[` of the first opens a subset nothing closes; the second ends at its `>`.
    const out = resolveInternalEntities('<!DOCTYPE x [ <!DOCTYPE r><r>&amp;</r>');
    expect(out).toBe('<!DOCTYPE x [ <r>&amp;</r>');
  });

  // The declaration is looked for before the parser sees the part, in raw
  // text the file decides. Searched for by one expression, a part of
  // `<!DOCTYPE ` written over and over was read to its end from each of them.
  it('gives up on declarations that never close in time linear in them', () => {
    for (const unclosed of [
      `${'<!DOCTYPE '.repeat(40_000)}r`,
      `${'<!DOCTYPE [ '.repeat(80_000)}<r/>`,
    ]) {
      const start = performance.now();
      expect(resolveInternalEntities(unclosed)).toBe(unclosed);
      expect(performance.now() - start).toBeLessThan(1000);
    }
  });
});
