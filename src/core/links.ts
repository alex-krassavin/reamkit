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
