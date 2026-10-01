// Searching a raw XML part for a tag with a regular expression, in linear time.
//
// A few readers take one attribute out of a part without parsing it — a
// font's name, an embedded font's key, the document's language — with an
// expression of the shape `<w:name\b[^>]*…`. Searching, the engine tries it at
// every `<w:name`, and from each it reads the rest of the tag again: a part
// the file wrote with many such starts and no `>` makes that quadratic. Every
// start inside a tag sees a part of what the first one saw, though, so where
// the first fails they all fail; tried at the first start of each tag only,
// the expression reads each tag once.

/** Whether the code unit is one `\w` matches. NaN, past the end, is not. */
export function isWordChar(c: number): boolean {
  return (
    (c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a) || c === 0x5f
  );
}

/**
 * The matches a global search with `re` finds in `xml` from `from`, in order.
 *
 * @param xml      The part's text.
 * @param open     The tag's opening, `<w:font`, which `re` begins with.
 * @param boundary Whether `re` puts `\b` right after `open`.
 * @param re       The expression, sticky (`y`): `open`, the `\b` if any, then
 *                 `[^>]*` (greedy or lazy) before what it looks for in the tag.
 * @param from     Where the search starts.
 */
export function* tagMatches(
  xml: string,
  open: string,
  boundary: boolean,
  re: RegExp,
  from = 0,
): Generator<RegExpExecArray> {
  let at = xml.indexOf(open, from);
  while (at >= 0) {
    if (boundary && isWordChar(xml.charCodeAt(at + open.length))) {
      at = xml.indexOf(open, at + 1);
      continue;
    }
    re.lastIndex = at;
    const found = re.exec(xml);
    if (found) {
      const next = Math.max(re.lastIndex, at + 1);
      yield found;
      at = xml.indexOf(open, next);
      continue;
    }
    // No start before the tag's `>` sees anything this one did not.
    const end = xml.indexOf('>', at + open.length);
    if (end < 0) return;
    at = xml.indexOf(open, end + 1);
  }
}

/** The first of {@link tagMatches}, or undefined when there is none. */
export function firstTagMatch(
  xml: string,
  open: string,
  boundary: boolean,
  re: RegExp,
  from = 0,
): RegExpExecArray | undefined {
  for (const found of tagMatches(xml, open, boundary, re, from)) return found;
  return undefined;
}
