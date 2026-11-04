import { Readable } from "node:stream";

import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import { checkSourceArgument, toAsyncIterable } from "@archive/io/archive-source";

/**
 * Convert an ArchiveSource into a Node.js Readable stream.
 *
 * Useful when a consumer expects Node stream semantics (e.g. `.pipe()`).
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function toNodeReadable(
  source: ArchiveSource | ArchiveBytes,
  options: { signal?: AbortSignal } = {}
): Readable {
  checkSourceArgument("Archive.toNodeReadable", source);
  return Readable.from(toAsyncIterable(source, options) as AsyncIterable<Uint8Array>);
}
