// ECMA-376 Part 2 — OPC package writer (the inverse of OpcPackage.open).
// Assembles parts + relationships into a deterministic ZIP: the same input
// always produces the same bytes (fixed ZIP mtime, stable part order), so a
// writer's output can sit under the byte-gate like everything else. The first
// step of the docx-writer epic (E-DOCX D1) and reusable by any future
// OOXML-emitting writer.

import { zipSync } from 'fflate';

import type { Relationship } from '@/core/opc/relationships';

/** One part to write into the package: its path, bytes, and content type. */
export interface OpcPart {
  /** Package path, e.g. `'word/document.xml'`. No leading slash. */
  readonly path: string;
  readonly data: Uint8Array;
  /** Content type for the `[Content_Types].xml` Override (XML parts). */
  readonly contentType: string;
}

/**
 * A part's relationships (id/type/target[/External]). The reader returns the
 * same {@link Relationship} shape, so a writer round-trips through
 * `parseRelationships`.
 */
export interface OpcPartRelationships {
  /** The owning part, e.g. `'word/document.xml'` (or `''` for the package root). */
  readonly sourcePart: string;
  readonly relationships: ReadonlyArray<Relationship>;
}

/** The full input to {@link buildOpcPackage}: parts plus their relationships. */
export interface OpcWriteOptions {
  /** The content parts (excluding `[Content_Types].xml` and `.rels`, which are generated). */
  readonly parts: ReadonlyArray<OpcPart>;
  /** Package-level relationships (`_rels/.rels`). */
  readonly rootRelationships: ReadonlyArray<Relationship>;
  /** Per-part relationships (word/_rels/<part>.rels). */
  readonly partRelationships?: ReadonlyArray<OpcPartRelationships>;
  /**
   * Default content types by lowercase extension, for parts not given an
   * Override (typically binaries: png, jpeg…). `rels` and `xml` are always
   * present and need not be repeated.
   */
  readonly defaultsByExtension?: Readonly<Record<string, string>>;
}

const encoder = new TextEncoder();
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

// A fixed ZIP timestamp keeps output byte-stable (fflate stamps the current
// time otherwise). DOS time floors at 1980-01-01.
const FIXED_MTIME = Date.UTC(1980, 0, 1, 0, 0, 0);

/**
 * Assemble parts + relationships into a deterministic OPC ZIP: it synthesizes
 * `[Content_Types].xml` and the `.rels` parts, then zips everything with a fixed
 * timestamp so the same input always yields byte-identical output.
 *
 * @returns The packaged `.docx`/`.xlsx`/… bytes.
 */
export function buildOpcPackage(options: OpcWriteOptions): Uint8Array {
  const entries: Record<string, [Uint8Array, { mtime: number }]> = {};
  const add = (path: string, data: Uint8Array): void => {
    entries[path] = [data, { mtime: FIXED_MTIME }];
  };

  add('[Content_Types].xml', encoder.encode(contentTypesXml(options)));
  add('_rels/.rels', encoder.encode(relationshipsXml(options.rootRelationships)));

  for (const part of options.parts) add(part.path, part.data);

  for (const pr of options.partRelationships ?? []) {
    if (pr.relationships.length === 0) continue;
    add(relsPathFor(pr.sourcePart), encoder.encode(relationshipsXml(pr.relationships)));
  }

  return zipSync(entries);
}

/**
 * The relationships-part path for a part: `word/document.xml` →
 * `word/_rels/document.xml.rels`. (`''` (root) is handled separately as
 * `_rels/.rels`.)
 */
export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash >= 0 ? partPath.slice(0, slash) : '';
  const name = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

// §10.1.2.2 [Content_Types].xml: Default elements per extension (rels + xml
// always, plus binaries), then an Override for each part whose content type
// differs from its extension's default.
function contentTypesXml(options: OpcWriteOptions): string {
  const defaults: Record<string, string> = {
    rels: 'application/vnd.openxmlformats-package.relationships+xml',
    xml: 'application/xml',
    ...(options.defaultsByExtension ?? {}),
  };
  const lines: Array<string> = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Types xmlns="${CT_NS}">`,
  ];
  for (const ext of Object.keys(defaults).sort()) {
    lines.push(`<Default Extension="${ext}" ContentType="${escapeAttr(defaults[ext]!)}"/>`);
  }
  for (const part of options.parts) {
    const ext = extensionOf(part.path);
    if (ext !== undefined && defaults[ext] === part.contentType) continue; // covered by Default
    lines.push(
      `<Override PartName="/${escapeAttr(part.path)}" ContentType="${escapeAttr(part.contentType)}"/>`,
    );
  }
  lines.push('</Types>');
  return lines.join('');
}

function relationshipsXml(rels: ReadonlyArray<Relationship>): string {
  const lines: Array<string> = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${RELS_NS}">`,
  ];
  for (const r of rels) {
    const mode = r.targetMode === 'External' ? ' TargetMode="External"' : '';
    lines.push(
      `<Relationship Id="${escapeAttr(r.id)}" Type="${escapeAttr(r.type)}"` +
        ` Target="${escapeAttr(r.target)}"${mode}/>`,
    );
  }
  lines.push('</Relationships>');
  return lines.join('');
}

function extensionOf(path: string): string | undefined {
  const dot = path.lastIndexOf('.');
  return dot >= 0 ? path.slice(dot + 1).toLowerCase() : undefined;
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
