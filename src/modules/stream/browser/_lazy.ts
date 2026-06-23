/**
 * Lazy Duplex.from() registry
 *
 * Breaks the circular import between Readable and Duplex:
 *   readable.ts → duplex.ts → readable.ts (cycle!)
 *
 * This file has ZERO imports from the browser stream class files, so it
 * can be safely imported by readable.ts without creating a cycle.
 * The actual Duplex.from binding is registered at module load time by
 * index.browser.ts.
 *
 * @internal
 */

import type { IDuplex } from "@stream/types";

type DuplexFromFn = (source: unknown) => IDuplex<unknown, unknown>;

let _duplexFrom: DuplexFromFn | null = null;

/**
 * Get the registered Duplex.from() factory.
 * Returns null if not yet registered (callers should handle gracefully).
 */
export function getDuplexFrom(): DuplexFromFn | null {
  return _duplexFrom;
}

/**
 * Register the Duplex.from() factory.
 * Called once from index.browser.ts after all stream classes are loaded.
 *
 * @internal
 */
export function registerDuplexFrom(factory: DuplexFromFn): void {
  _duplexFrom = factory;
}
