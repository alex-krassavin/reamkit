// IR core — conversion loss reporting (ir-design.md §9).
//
// Conversions never silently eat content: every place the pipeline drops,
// degrades or substitutes something it records a Loss. The default mode keeps
// converting and returns the report alongside the output; `strict: true`
// (a caller decision — e.g. legal/archival projects) turns the FIRST loss into
// a thrown ConversionLossError instead.

import type { Feature } from '@/core/ir/features';

/**
 * - `dropped`     — content is absent from the output (e.g. math → markdown).
 * - `degraded`    — content survives with reduced fidelity (e.g. gradient → flat fill).
 * - `substituted` — content was replaced by an equivalent (e.g. font substitution).
 */
export type LossSeverity = 'dropped' | 'degraded' | 'substituted';

/** A single conversion loss: something the pipeline dropped, degraded or substituted. */
export interface Loss {
  /** How badly the content was affected. */
  readonly severity: LossSeverity;
  /** The affected {@link Feature} (drives the capability/loss matrix). */
  readonly feature: Feature;
  /** Human-readable specifics: what exactly happened to what. */
  readonly detail: string;
  /** Optional location hint (page/paragraph/sheet reference). */
  readonly where?: string;
}

/** An ordered, read-only list of {@link Loss} entries, returned alongside the output. */
export type LossReport = ReadonlyArray<Loss>;

/**
 * Format a {@link Loss} as one human-readable line, e.g.
 * `[degraded] charts at sheet "Q3": gradient flattened`.
 *
 * @param loss The loss to format.
 * @returns The one-line description.
 */
export function formatLoss(loss: Loss): string {
  const where = loss.where ? ` at ${loss.where}` : '';
  return `[${loss.severity}] ${loss.feature}${where}: ${loss.detail}`;
}

/** Thrown by strict-mode conversions on the first recorded loss. */
export class ConversionLossError extends Error {
  /** @param loss The first loss that tripped strict mode (retained on the error). */
  constructor(readonly loss: Loss) {
    super(`strict conversion failed — ${formatLoss(loss)}`);
    this.name = 'ConversionLossError';
  }
}
