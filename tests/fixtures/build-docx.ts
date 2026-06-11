// Build a minimal valid WordprocessingML docx in memory for tests.
// Produces only what ECMA-376 Part 2 requires for a recognizable package:
// [Content_Types].xml, _rels/.rels, word/document.xml, plus optional
// word/numbering.xml, word/header1.xml, word/footer1.xml with the
// corresponding part relationships.

import { zipSync } from 'fflate';

const encoder = new TextEncoder();

const CONTENT_TYPES_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`;

const NUMBERING_TYPE_OVERRIDE =
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';
const STYLES_TYPE_OVERRIDE =
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';

const SETTINGS_TYPE_OVERRIDE =
  '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
const THEME_TYPE_OVERRIDE =
  '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const chartOverride = (n: number): string =>
  `<Override PartName="/word/charts/chart${n}.xml" ContentType="${CHART_CT}"/>`;
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const HEADER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const headerOverride = (n: number): string =>
  `<Override PartName="/word/header${n}.xml" ContentType="${HEADER_CT}"/>`;
const footerOverride = (n: number): string =>
  `<Override PartName="/word/footer${n}.xml" ContentType="${FOOTER_CT}"/>`;

const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export interface FixtureImage {
  readonly contentType: 'image/png' | 'image/jpeg';
  readonly bytes: Uint8Array;
  readonly extension: 'png' | 'jpg' | 'jpeg';
}

export interface BuildDocxOptions {
  readonly numberingXml?: string;
  /** Inner XML of word/styles.xml (the <w:style>/<w:docDefaults> elements). */
  readonly stylesXml?: string;
  // Default header/footer (legacy fields, rId10 / rId11).
  readonly headerXml?: string;
  readonly footerXml?: string;
  // Optional first-page-only / even-page-only variants.
  readonly firstHeaderXml?: string; // rId12 → word/header2.xml
  readonly firstFooterXml?: string; // rId13 → word/footer2.xml
  readonly evenHeaderXml?: string; // rId14 → word/header3.xml
  readonly evenFooterXml?: string; // rId15 → word/footer3.xml
  // Inner XML of word/settings.xml (e.g. "<w:evenAndOddHeaders/>"). Wrapped in
  // <w:settings> here.
  readonly settingsXml?: string;
  // Full word/theme/theme1.xml content (an <a:theme>…</a:theme> document).
  readonly themeXml?: string;
  readonly images?: Readonly<Record<string, FixtureImage>>;
  // Images owned by the default header (word/_rels/header1.xml.rels) — for
  // exercising per-part relationship resolution.
  readonly headerImages?: Readonly<Record<string, FixtureImage>>;
  // Chart parts keyed by the relationship id a c:chart @r:id references. Each
  // value is a full <c:chartSpace>…</c:chartSpace> document.
  readonly charts?: Readonly<Record<string, string>>;
  /** External hyperlink rels for document.xml: rId → URL (TargetMode External). */
  readonly hyperlinks?: Readonly<Record<string, string>>;
}

export function buildDocx(paragraphs: ReadonlyArray<string>): Uint8Array {
  const body = paragraphs
    .map((p) => `  <w:p><w:r><w:t xml:space="preserve">${escapeXml(p)}</w:t></w:r></w:p>`)
    .join('\n');
  return buildDocxFromBody(body);
}

export function buildDocxFromBody(
  bodyInnerXml: string,
  options: BuildDocxOptions = {},
): Uint8Array {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
${bodyInnerXml}
  </w:body>
</w:document>`;

  const imageDefaults = new Set<string>();
  for (const img of Object.values({ ...options.images, ...options.headerImages })) {
    imageDefaults.add(`<Default Extension="${img.extension}" ContentType="${img.contentType}"/>`);
  }
  const headerSlots: Array<{ idx: number; xml: string; rId: string }> = [];
  const footerSlots: Array<{ idx: number; xml: string; rId: string }> = [];
  if (options.headerXml) headerSlots.push({ idx: 1, xml: options.headerXml, rId: 'rId10' });
  if (options.firstHeaderXml)
    headerSlots.push({ idx: 2, xml: options.firstHeaderXml, rId: 'rId12' });
  if (options.evenHeaderXml) headerSlots.push({ idx: 3, xml: options.evenHeaderXml, rId: 'rId14' });
  if (options.footerXml) footerSlots.push({ idx: 1, xml: options.footerXml, rId: 'rId11' });
  if (options.firstFooterXml)
    footerSlots.push({ idx: 2, xml: options.firstFooterXml, rId: 'rId13' });
  if (options.evenFooterXml) footerSlots.push({ idx: 3, xml: options.evenFooterXml, rId: 'rId15' });

  const chartSlots: Array<{ idx: number; xml: string; rId: string }> = [];
  if (options.charts) {
    let i = 1;
    for (const [rId, xml] of Object.entries(options.charts)) {
      chartSlots.push({ idx: i, xml, rId });
      i++;
    }
  }

  const contentTypes =
    CONTENT_TYPES_HEADER +
    [...imageDefaults].join('') +
    (options.numberingXml ? NUMBERING_TYPE_OVERRIDE : '') +
    (options.stylesXml ? STYLES_TYPE_OVERRIDE : '') +
    (options.settingsXml ? SETTINGS_TYPE_OVERRIDE : '') +
    (options.themeXml ? THEME_TYPE_OVERRIDE : '') +
    headerSlots.map((s) => headerOverride(s.idx)).join('') +
    footerSlots.map((s) => footerOverride(s.idx)).join('') +
    chartSlots.map((s) => chartOverride(s.idx)).join('') +
    '\n</Types>';

  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(contentTypes),
    '_rels/.rels': encoder.encode(ROOT_RELS),
    'word/document.xml': encoder.encode(documentXml),
  };

  if (options.stylesXml) {
    entries['word/styles.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
${options.stylesXml}
</w:styles>`,
    );
  }

  if (options.numberingXml) {
    entries['word/numbering.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
${options.numberingXml}
</w:numbering>`,
    );
  }

  if (options.settingsXml) {
    entries['word/settings.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
${options.settingsXml}
</w:settings>`,
    );
  }

  if (options.themeXml) {
    entries['word/theme/theme1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${options.themeXml}`,
    );
  }

  for (const c of chartSlots) {
    entries[`word/charts/chart${c.idx}.xml`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${c.xml}`,
    );
  }

  for (const h of headerSlots) {
    entries[`word/header${h.idx}.xml`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
${h.xml}
</w:hdr>`,
    );
  }
  for (const f of footerSlots) {
    entries[`word/footer${f.idx}.xml`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
${f.xml}
</w:ftr>`,
    );
  }

  const docRels: Array<string> = [];
  for (const h of headerSlots) {
    docRels.push(`<Relationship Id="${h.rId}" Type="${REL_HEADER}" Target="header${h.idx}.xml"/>`);
  }
  for (const f of footerSlots) {
    docRels.push(`<Relationship Id="${f.rId}" Type="${REL_FOOTER}" Target="footer${f.idx}.xml"/>`);
  }
  if (options.themeXml) {
    docRels.push(`<Relationship Id="rIdTheme" Type="${REL_THEME}" Target="theme/theme1.xml"/>`);
  }
  for (const c of chartSlots) {
    docRels.push(
      `<Relationship Id="${c.rId}" Type="${REL_CHART}" Target="charts/chart${c.idx}.xml"/>`,
    );
  }
  if (options.hyperlinks) {
    for (const [rId, url] of Object.entries(options.hyperlinks)) {
      docRels.push(
        `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url.replaceAll('&', '&amp;')}" TargetMode="External"/>`,
      );
    }
  }
  if (options.images) {
    let imgIdx = 1;
    for (const [rId, img] of Object.entries(options.images)) {
      const target = `media/image${imgIdx}.${img.extension}`;
      entries[`word/${target}`] = img.bytes;
      docRels.push(`<Relationship Id="${rId}" Type="${REL_IMAGE}" Target="${target}"/>`);
      imgIdx++;
    }
  }
  if (options.headerImages) {
    const hdrRels: Array<string> = [];
    let imgIdx = 1;
    for (const [rId, img] of Object.entries(options.headerImages)) {
      const target = `media/hImage${imgIdx}.${img.extension}`;
      entries[`word/${target}`] = img.bytes;
      hdrRels.push(`<Relationship Id="${rId}" Type="${REL_IMAGE}" Target="${target}"/>`);
      imgIdx++;
    }
    entries['word/_rels/header1.xml.rels'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${hdrRels.join('\n')}
</Relationships>`,
    );
  }
  if (docRels.length > 0) {
    entries['word/_rels/document.xml.rels'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${docRels.join('\n')}
</Relationships>`,
    );
  }

  return zipSync(entries);
}

export interface FixtureRun {
  readonly text: string;
  readonly rPrXml?: string;
}

export interface FixtureParagraph {
  readonly pPrXml?: string;
  readonly runs: ReadonlyArray<FixtureRun>;
}

export function buildRichDocx(paragraphs: ReadonlyArray<FixtureParagraph>): Uint8Array {
  const body = paragraphs.map(renderParagraph).join('\n');
  return buildDocxFromBody(body);
}

function renderParagraph(p: FixtureParagraph): string {
  const pPr = p.pPrXml ? `    ${p.pPrXml}\n` : '';
  const runs = p.runs.map(renderRun).join('\n');
  return `  <w:p>\n${pPr}${runs}\n  </w:p>`;
}

function renderRun(r: FixtureRun): string {
  const rPr = r.rPrXml ? `      ${r.rPrXml}\n` : '';
  return `    <w:r>\n${rPr}      <w:t xml:space="preserve">${escapeXml(r.text)}</w:t>\n    </w:r>`;
}
