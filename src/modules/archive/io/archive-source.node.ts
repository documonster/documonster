import { Readable } from "node:stream";

import type { ArchiveSource } from "@archive/io/archive-source";
import { toAsyncIterable } from "@archive/io/archive-source";

/**
 * Convert an ArchiveSource into a Node.js Readable stream.
 *
 * Useful when a consumer expects Node stream semantics (e.g. `.pipe()`).
 */
export function toNodeReadable(
  source: ArchiveSource,
  options: { signal?: AbortSignal } = {}
): Readable {
  return Readable.from(toAsyncIterable(source, options) as AsyncIterable<Uint8Array>);
}
