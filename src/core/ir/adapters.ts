// Adapter contracts (ir-design §7) — @experimental until they stabilize
// against three-plus adapters (handoff v1 §3a). Readers and writers are SYNC
// by contract (the bytes are already in memory; all async — font fetching,
// resource I/O — lives in the conversion facade).

import type { Feature } from '@/core/ir/features';
import type { FlowDoc } from '@/core/ir/flow';
import type { LossReport } from '@/core/ir/loss';

/** Reader-call options. Open-ended; per-reader limits live in the readers. */
export interface ReadOptions {
  /** Reserved for reader-specific knobs (limits live in the readers). */
  readonly [key: string]: unknown;
}

/** The result of {@link DocumentReader.read}: the parsed tree plus any losses. */
export interface ReadResult<TDoc> {
  /** The parsed document tree. */
  readonly doc: TDoc;
  /** Losses recorded while reading. */
  readonly losses: LossReport;
}

/**
 * A synchronous reader for one source format (ir-design §7). All async I/O —
 * font fetching, resource loading — lives in the conversion facade, not here.
 *
 * @experimental The adapter contracts may change until they stabilize against
 * three or more adapters.
 */
export interface DocumentReader<TDoc = FlowDoc> {
  /** Format id: 'docx', 'xlsx', 'pdf', … */
  readonly id: string;
  /** Which tree this reader produces ('sheet' projects to flow at the boundary). */
  readonly produces: 'flow' | 'page' | 'sheet';
  /** Feature vocabulary this reader understands (capability matrix source). */
  readonly supports: ReadonlySet<Feature>;
  /** Cheap format detection (magic bytes / container markers). */
  sniff: (bytes: Uint8Array) => boolean;
  read: (bytes: Uint8Array, opts?: ReadOptions) => ReadResult<TDoc>;
}

/** Writer-call options. Open-ended; per-writer knobs live in the writers. */
export interface WriteOptions {
  readonly [key: string]: unknown;
}

/** The result of {@link DocumentWriter.write}: the encoded bytes plus any losses. */
export interface WriteResult {
  /** The encoded output bytes. */
  readonly bytes: Uint8Array;
  /** Losses recorded while writing. */
  readonly losses: LossReport;
}

/**
 * A synchronous writer for one output format (ir-design §7).
 *
 * @experimental The adapter contracts may change until they stabilize against
 * three or more adapters.
 */
export interface DocumentWriter<TDoc> {
  /** Output format id: 'pdf', 'svg', 'html', 'docx', 'xlsx', … */
  readonly id: string;
  /** Which tree this writer consumes ('sheet' for the native grid writer). */
  readonly consumes: 'flow' | 'page' | 'sheet';
  /** Feature vocabulary this writer understands. */
  readonly supports: ReadonlySet<Feature>;
  /** Encode `doc` to the output bytes, collecting any losses. */
  write: (doc: TDoc, opts?: WriteOptions) => WriteResult;
}
