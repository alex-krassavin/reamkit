// Hyperlink scheme allowlist, shared by every writer that emits something
// clickable (the HTML writer's <a href>, the PDF writer's /URI action).
//
// Documents are UNTRUSTED input: a crafted w:hyperlink target of
// `javascript:…` (or data:/file:/vbscript:) would turn a converted document
// into an execution vector the moment a browser or viewer follows it. Only
// plain web/mail targets pass through; everything else renders as plain text
// and the writer reports a degraded-hyperlink loss.

const ALLOWED_SCHEMES = new Set(['http', 'https', 'mailto']);

/**
 * Returns the trimmed URL when its scheme is on the allowlist, undefined
 * otherwise (including scheme-less/relative targets — those are not external
 * links for a converted document).
 */
export function sanitizeHref(href: string): string | undefined {
  const url = href.trim();
  const m = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(url);
  if (!m) return undefined;
  return ALLOWED_SCHEMES.has(m[1]!.toLowerCase()) ? url : undefined;
}

// Any UTF-16 code unit above U+007F, surrogate halves included.
const NON_ASCII = /[\x80-\uffff]/;
const UTF8 = new TextEncoder();

/**
 * Maps an IRI (RFC 3987) to the URI it stands for, for a writer whose link
 * target must be 7-bit ASCII — the PDF writer's `/URI` (ISO 32000-1
 * §12.6.4.7). The HTML, SVG, Markdown and DOCX writers keep the IRI itself —
 * an `href` and an OPC relationship's `Target` (xsd:anyURI) may carry one —
 * and a browser following the HTML link maps it to this same URI.
 *
 * The mapping is RFC 3987 §3.1's: every non-ASCII character in the userinfo,
 * path, query and fragment becomes the `%XX` of its UTF-8 octets (a lone
 * surrogate as U+FFFD's, the way a browser encodes one). A non-ASCII host
 * becomes IDNA A-labels (`пример.рф` → `xn--e1afmkfd.xn--p1ai`), the form DNS
 * resolves; a host that is no valid domain name is percent-encoded instead.
 * Nothing ASCII is touched, so an ASCII target, `%XX` sequences included, is
 * returned as it is, and mapping a URI again changes nothing.
 *
 * @param iri A hyperlink target, typically one {@link sanitizeHref} let through.
 * @returns The URI, in 7-bit ASCII.
 */
export function iriToUri(iri: string): string {
  if (!NON_ASCII.test(iri)) return iri;
  // RFC 3986 §3.2: an authority follows "//" and runs to the first "/", "?" or
  // "#". The userinfo ends at its last "@" — where a browser ends it too.
  const m = /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\/[^/?#]*/.exec(iri);
  if (!m) return percentEncodeNonAscii(iri); // mailto: has no authority
  const scheme = m[0].indexOf('//') + 2;
  const authority = m[0].slice(scheme);
  const at = authority.lastIndexOf('@') + 1;
  // §3.2.3: the port is the digits after the host's last ":".
  const [, host, port] = /^(.*?)(:[0-9]*)?$/.exec(authority.slice(at))!;
  return (
    iri.slice(0, scheme) +
    percentEncodeNonAscii(authority.slice(0, at)) +
    asciiHost(host!) +
    (port ?? '') +
    percentEncodeNonAscii(iri.slice(m[0].length))
  );
}

// RFC 3986 §3.2.2: a name meant for DNS goes to the resolver in the IDNA
// encoding, and producers should write it so "rather than a percent-encoding"
// for legacy resolvers. (Handed the host as raw UTF-8, macOS PDFKit gave the
// annotation no URL at all.) The conversion is the URL host parser's (UTS #46
// ToASCII), the one a browser runs on the HTML writer's href, so the PDF link
// lands on the host the HTML link does.
function asciiHost(host: string): string {
  if (!NON_ASCII.test(host)) return host;
  try {
    const url = new URL(`http://${host}/`);
    // The parser reads "\" as a path delimiter: take its host only when it
    // kept all of ours (nothing went to the path).
    if (url.pathname === '/' && url.hostname !== '' && !NON_ASCII.test(url.hostname)) {
      return url.hostname;
    }
  } catch {
    // Not a valid domain name (a space, a malformed "xn--" label): fall through.
  }
  // RFC 3986 §3.2.2 still lets a reg-name be percent-encoded; it stays ASCII.
  return percentEncodeNonAscii(host);
}

// RFC 3987 §3.1 step 2: a non-ASCII character becomes the %XX of its UTF-8
// octets, upper-case hex (RFC 3986 §2.1); ASCII passes through untouched.
function percentEncodeNonAscii(s: string): string {
  const hex = (b: number) => `%${b.toString(16).toUpperCase().padStart(2, '0')}`;
  return s.replace(/[\x80-\uffff]+/g, (run) => Array.from(UTF8.encode(run), hex).join(''));
}
