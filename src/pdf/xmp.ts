// XMP metadata packet (ISO 16684-1 / Adobe XMP) for the document /Metadata
// stream. PDF/A-1 requires document-level XMP carrying the PDF/A identifier
// (pdfaid:part + pdfaid:conformance) and that the standard properties agree
// with the /Info dictionary.

export interface XmpInput {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string; // application that created the doc (xmp:CreatorTool)
  readonly producer?: string;
  readonly createDate?: Date;
  readonly modifyDate?: Date;
  // PDF/A identifier: part 1 (ISO 19005-1) / 2 / 3; conformance level
  // 'A' (tagged) / 'B' (visual) / 'U' (Unicode — parts 2/3 only).
  readonly pdfaPart?: '1' | '2' | '3';
  readonly pdfaConformance?: 'A' | 'B' | 'U';
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// XMP dates are ISO 8601. We emit UTC with a 'Z' suffix for determinism.
function xmpDate(d: Date): string {
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  return (
    `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`
  );
}

export function buildXmpPacket(input: XmpInput): Uint8Array {
  const props: Array<string> = [];

  // pdfaid (PDF/A identification) — required for PDF/A.
  if (input.pdfaPart) {
    props.push(
      `      <rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">`,
      `        <pdfaid:part>${input.pdfaPart}</pdfaid:part>`,
      `        <pdfaid:conformance>${input.pdfaConformance ?? 'B'}</pdfaid:conformance>`,
      `      </rdf:Description>`,
    );
  }

  // Dublin Core.
  const dc: Array<string> = [];
  if (input.title) {
    dc.push(
      `        <dc:title><rdf:Alt><rdf:li xml:lang="x-default">${escapeXml(input.title)}</rdf:li></rdf:Alt></dc:title>`,
    );
  }
  if (input.author) {
    dc.push(
      `        <dc:creator><rdf:Seq><rdf:li>${escapeXml(input.author)}</rdf:li></rdf:Seq></dc:creator>`,
    );
  }
  if (input.subject) {
    dc.push(
      `        <dc:description><rdf:Alt><rdf:li xml:lang="x-default">${escapeXml(input.subject)}</rdf:li></rdf:Alt></dc:description>`,
    );
  }
  if (dc.length > 0) {
    props.push(
      `      <rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/">`,
      ...dc,
      `      </rdf:Description>`,
    );
  }

  // XMP basic.
  const xmp: Array<string> = [];
  if (input.creator)
    xmp.push(`        <xmp:CreatorTool>${escapeXml(input.creator)}</xmp:CreatorTool>`);
  if (input.createDate)
    xmp.push(`        <xmp:CreateDate>${xmpDate(input.createDate)}</xmp:CreateDate>`);
  if (input.modifyDate)
    xmp.push(`        <xmp:ModifyDate>${xmpDate(input.modifyDate)}</xmp:ModifyDate>`);
  if (xmp.length > 0) {
    props.push(
      `      <rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/">`,
      ...xmp,
      `      </rdf:Description>`,
    );
  }

  // Adobe PDF schema (Producer + Keywords).
  const pdf: Array<string> = [];
  if (input.producer) pdf.push(`        <pdf:Producer>${escapeXml(input.producer)}</pdf:Producer>`);
  if (input.keywords) pdf.push(`        <pdf:Keywords>${escapeXml(input.keywords)}</pdf:Keywords>`);
  if (pdf.length > 0) {
    props.push(
      `      <rdf:Description rdf:about="" xmlns:pdf="http://ns.adobe.com/pdf/1.3/">`,
      ...pdf,
      `      </rdf:Description>`,
    );
  }

  const body = [
    `<?xpacket begin="\uFEFF" id="W5M0MpCehiHzreSzNTczkc9d"?>`,
    `<x:xmpmeta xmlns:x="adobe:ns:meta/">`,
    `  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">`,
    ...props,
    `  </rdf:RDF>`,
    `</x:xmpmeta>`,
    `<?xpacket end="w"?>`,
  ].join('\n');

  return new TextEncoder().encode(body);
}
