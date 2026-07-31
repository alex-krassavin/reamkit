// ISO/IEC 29500-3 — Markup Compatibility. `<mc:AlternateContent>` offers the
// same content twice: an `<mc:Choice Requires="…">` written against a namespace
// the producer's own application understands, and an `<mc:Fallback>` in plain
// OOXML for everyone else. A reader that does not implement the required
// namespace takes the Fallback.
//
// Resolved as TEXT, before the XML parser sees the part, because the parser
// groups children by element name and loses the order they were written in.
// That order is the whole meaning of a style collection: §18.8.23 indexes fonts
// by position, so a `<fonts count="14">` holding eight plain `<font>` elements
// and six wrapped ones parses as eight fonts, and every `fontId` past the first
// wrapper names the wrong one. style-alternate-content.xlsx is such a file in
// both its fonts and its cellXfs — 29 declared cell formats, 22 read — and it
// lost its title's face, its table's borders and its alignment together.

/** Deepest nesting of one AlternateContent inside another that we will unwrap. */
const MAX_DEPTH = 8;

const OPEN = /<(?:[A-Za-z_][\w.-]*:)?AlternateContent(?=[\s/>])/g;

/**
 * Find the element starting at `from` (the index of its `<`) and return the
 * span of its content plus the index just past its closing tag. Returns
 * undefined for an unclosed element — malformed markup we leave alone.
 */
function elementAt(
  xml: string,
  from: number,
  local: string,
): { readonly inner: string; readonly end: number } | undefined {
  const nameEnd = xml.indexOf('>', from);
  if (nameEnd < 0) return undefined;
  if (xml[nameEnd - 1] === '/') return { inner: '', end: nameEnd + 1 };
  const open = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${local}(?=[\\s/>])`, 'g');
  const close = new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${local}\\s*>`, 'g');
  let depth = 1;
  let cursor = nameEnd + 1;
  while (depth > 0) {
    close.lastIndex = cursor;
    const shut = close.exec(xml);
    if (!shut) return undefined;
    open.lastIndex = cursor;
    for (let o = open.exec(xml); o && o.index < shut.index; o = open.exec(xml)) {
      // A self-closing `<x/>` opens nothing.
      if (xml[xml.indexOf('>', o.index) - 1] !== '/') depth++;
    }
    depth--;
    cursor = shut.index + shut[0].length;
    if (depth === 0) return { inner: xml.slice(nameEnd + 1, shut.index), end: cursor };
  }
  return undefined;
}

/** The content of the first `<mc:Fallback>` directly inside `inner`, or ''. */
function fallbackOf(inner: string): string {
  const at = new RegExp('<(?:[A-Za-z_][\\w.-]*:)?Fallback(?=[\\s/>])').exec(inner);
  if (!at) return '';
  return elementAt(inner, at.index, 'Fallback')?.inner ?? '';
}

/**
 * Replace every `<mc:AlternateContent>` with the markup its `<mc:Fallback>`
 * holds, so the part reaching the XML parser is plain OOXML in its original
 * element order. A block with no Fallback resolves to nothing — which is what
 * §10.2 of the Markup Compatibility spec asks of a reader that understands
 * none of the offered choices.
 *
 * Text with no AlternateContent is returned unchanged, which is most files.
 *
 * @param xml The decoded part, as read from the package.
 * @returns The part with its compatibility blocks resolved.
 */
export function resolveAlternateContent(xml: string): string {
  let out = xml;
  for (let pass = 0; pass < MAX_DEPTH; pass++) {
    if (!out.includes('AlternateContent')) return out;
    let cursor = 0;
    let next = '';
    OPEN.lastIndex = 0;
    for (let found = OPEN.exec(out); found; found = OPEN.exec(out)) {
      const el = elementAt(out, found.index, 'AlternateContent');
      // Unclosed: leave this one and everything after it as written.
      if (!el) break;
      next += out.slice(cursor, found.index) + fallbackOf(el.inner);
      cursor = el.end;
      OPEN.lastIndex = cursor;
    }
    out = next + out.slice(cursor);
  }
  return out;
}
