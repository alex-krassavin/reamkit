// ECMA-376 Part 2 — Open Packaging Conventions.
// A package is a ZIP archive. Top-level "_rels/.rels" describes package-level
// relationships; the officeDocument relationship points to the main document
// part (word/document.xml for WordprocessingML).

import { unzipSync } from 'fflate';

import type { Relationship } from '@/core/opc/relationships';
import { parseRelationships } from '@/core/opc/relationships';
import { isOoxmlRel } from '@/core/opc/relationship-types';

const ROOT_RELS_PATH = '_rels/.rels';

const MIB = 1024 * 1024;

/**
 * Resource caps applied while unzipping a package — a defence against
 * decompression ("zip") bombs and pathological archives. Generous defaults that
 * never reject a legitimate office document; tighten for untrusted input.
 */
export interface OpcOpenOptions {
  /** Reject an archive whose raw (compressed) size exceeds this (default 128 MiB). */
  readonly maxArchiveBytes?: number;
  /** Reject any single entry declaring an uncompressed size over this (default 256 MiB). */
  readonly maxEntryBytes?: number;
  /** Reject if the total declared uncompressed size exceeds this (default 512 MiB). */
  readonly maxTotalBytes?: number;
  /** Reject archives with more than this many entries (default 65 536). */
  readonly maxEntries?: number;
}

/**
 * An opened OPC package (ECMA-376 Part 2): a ZIP archive's parts plus its
 * package-level relationships, with helpers to look up parts, resolve part
 * relationships, and find the main document part.
 */
export class OpcPackage {
  /**
   * @param parts             The package parts keyed by ZIP-convention path (no leading slash).
   * @param rootRelationships The package-level relationships from `_rels/.rels`.
   */
  private constructor(
    private readonly parts: ReadonlyMap<string, Uint8Array>,
    private readonly rootRelationships: ReadonlyArray<Relationship>,
  ) {}

  private readonly relsCache = new Map<string, ReadonlyArray<Relationship>>();

  /**
   * Unzip and validate a package's bytes. Rejects an OLE compound file (an
   * encrypted OOXML container or a legacy binary `.doc`/`.xls`) and enforces the
   * zip-bomb caps from `options`.
   *
   * @throws Error when the bytes are not a ZIP package, a cap is exceeded, or
   *         `_rels/.rels` is missing.
   */
  static open(buffer: Uint8Array, options: OpcOpenOptions = {}): OpcPackage {
    // OLE CFB magic (D0 CF 11 E0): a password-protected (encrypted) OOXML
    // container or a legacy binary .doc/.xls — either way, not a ZIP package.
    if (
      buffer.length >= 4 &&
      buffer[0] === 0xd0 &&
      buffer[1] === 0xcf &&
      buffer[2] === 0x11 &&
      buffer[3] === 0xe0
    ) {
      throw new Error(
        'OLE compound file: a password-protected (encrypted) OOXML document or a legacy binary .doc/.xls — re-save as an unencrypted .docx/.xlsx',
      );
    }
    const maxArchive = options.maxArchiveBytes ?? 128 * MIB;
    const maxEntry = options.maxEntryBytes ?? 256 * MIB;
    const maxTotal = options.maxTotalBytes ?? 512 * MIB;
    const maxEntries = options.maxEntries ?? 65_536;

    if (buffer.byteLength > maxArchive) {
      throw new Error(
        `OPC archive too large: ${buffer.byteLength} bytes (limit ${maxArchive}); refusing to unzip`,
      );
    }

    // Zip-bomb guard. fflate's filter runs per entry BEFORE decompression, so
    // returning false skips a (potentially huge) entry entirely rather than
    // expanding it. Declared sizes that under-report (forged) are still bounded
    // by maxArchiveBytes (the compressed input) and, for untrusted input,
    // operationally by running the parse in a memory-limited process.
    let total = 0;
    let count = 0;
    let violation: string | undefined;
    const entries = unzipSync(buffer, {
      filter: (info) => {
        if (++count > maxEntries) {
          violation ??= `more than ${maxEntries} entries`;
          return false;
        }
        if (info.originalSize > maxEntry) {
          violation ??= `entry "${info.name}" declares ${info.originalSize} bytes (limit ${maxEntry})`;
          return false;
        }
        total += info.originalSize;
        if (total > maxTotal) {
          violation ??= `total uncompressed size exceeds ${maxTotal} bytes`;
          return false;
        }
        return true;
      },
    });
    if (violation) {
      throw new Error(`OPC archive rejected (zip-bomb guard): ${violation}`);
    }

    const parts = new Map<string, Uint8Array>();
    for (const [path, data] of Object.entries(entries)) {
      parts.set(normalizePath(path), data);
    }
    const relsBytes = parts.get(ROOT_RELS_PATH);
    if (!relsBytes) {
      throw new Error(`OPC package missing ${ROOT_RELS_PATH}`);
    }
    return new OpcPackage(parts, parseRelationships(relsBytes));
  }

  /** The bytes of the part at `path`, or undefined when absent. */
  getPart(path: string): Uint8Array | undefined {
    return this.parts.get(normalizePath(path));
  }

  /**
   * The bytes of the part at `path`.
   *
   * @throws Error when the part is absent.
   */
  requirePart(path: string): Uint8Array {
    const data = this.getPart(path);
    if (!data) throw new Error(`OPC part not found: ${path}`);
    return data;
  }

  /** Every part path in the package (ZIP convention, no leading slash). */
  listParts(): Array<string> {
    return [...this.parts.keys()];
  }

  /**
   * The relationships of the part at `partPath` (ECMA-376 Part 2 §9.3.4). For a
   * part at `dir/name.ext` they live at `dir/_rels/name.ext.rels`. Returns `[]`
   * if the rels part is absent. Parsed rels are cached per part.
   */
  getPartRelationships(partPath: string): ReadonlyArray<Relationship> {
    const normalized = normalizePath(partPath);
    // One conversion asks for the main part's rels several times (images,
    // headers/footers, charts, embedded fonts) — parse each .rels once.
    const cached = this.relsCache.get(normalized);
    if (cached) return cached;
    const slash = normalized.lastIndexOf('/');
    const dir = slash >= 0 ? normalized.substring(0, slash) : '';
    const base = slash >= 0 ? normalized.substring(slash + 1) : normalized;
    const relsPath = dir.length > 0 ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
    const data = this.parts.get(relsPath);
    const rels = data ? parseRelationships(data) : [];
    this.relsCache.set(normalized, rels);
    return rels;
  }

  /**
   * Resolve a relationship against its source part and return the related part's
   * resolved path + data (Internal relationships only). External relationships
   * (e.g. http hyperlinks) and unresolvable targets return undefined.
   */
  resolveRelatedPart(
    sourcePartPath: string,
    relationship: Relationship,
  ): { readonly path: string; readonly data: Uint8Array } | undefined {
    if (relationship.targetMode !== 'Internal') return undefined;
    const sourceNormalized = normalizePath(sourcePartPath);
    const resolved = resolveTarget(`/${sourceNormalized}`, relationship.target);
    const data = this.getPart(resolved);
    if (!data) return undefined;
    return { path: resolved, data };
  }

  /**
   * The path of the main document part, from the package's single
   * `officeDocument` relationship (ECMA-376 Part 2 §11.1).
   *
   * @throws Error when there is no such relationship, or more than one.
   */
  getMainDocumentPath(): string {
    const candidates = this.rootRelationships.filter(
      (r) => isOoxmlRel(r.type, 'officeDocument') && r.targetMode === 'Internal',
    );
    if (candidates.length === 0) {
      throw new Error('OPC package has no officeDocument relationship');
    }
    if (candidates.length > 1) {
      throw new Error(
        `OPC package has multiple officeDocument relationships (${candidates.length})`,
      );
    }
    return resolveTarget('/', candidates[0]!.target);
  }

  /**
   * The main document part's path and bytes.
   *
   * @throws Error when the officeDocument relationship is missing/ambiguous or
   *         the part it points at is absent.
   */
  getMainDocument(): { path: string; data: Uint8Array } {
    const path = this.getMainDocumentPath();
    return { path, data: this.requirePart(path) };
  }
}

// Paths inside a ZIP have no leading slash; OPC PartNames are conceptually
// absolute with a leading slash. Normalize to the ZIP convention.
function normalizePath(p: string): string {
  return p.startsWith('/') ? p.slice(1) : p;
}

// ECMA-376 Part 2 §9.3.2 — Target resolution.
// A relative Target is resolved against the source part's location; an
// absolute Target starts with '/'.
function resolveTarget(sourcePath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target);
  const sourceDir = sourcePath.endsWith('/')
    ? sourcePath
    : sourcePath.substring(0, sourcePath.lastIndexOf('/') + 1);
  const combined = `${sourceDir}${target}`;
  return normalizePath(collapseDotSegments(combined));
}

function collapseDotSegments(p: string): string {
  const segments = p.split('/');
  const out: Array<string> = [];
  for (const s of segments) {
    if (s === '.' || s === '') {
      if (out.length === 0) out.push('');
      continue;
    }
    if (s === '..') {
      if (out.length > 1) out.pop();
      continue;
    }
    out.push(s);
  }
  return out.join('/');
}
