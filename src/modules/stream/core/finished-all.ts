/**
 * Stream common - finishedAll
 *
 * Shared implementation for waiting on multiple streams to finish.
 */

import { StreamTypeError } from "@stream/errors";
import type { PipelineStreamLike } from "@stream/types";

/**
 * Create a `finishedAll` function bound to a platform-specific `finished`.
 */
export function createFinishedAll(
  finished: (stream: PipelineStreamLike) => Promise<void>
): (streams: ReadonlyArray<PipelineStreamLike>) => Promise<void> {
  return async function finishedAll(streams: ReadonlyArray<PipelineStreamLike>): Promise<void> {
    if (!Array.isArray(streams)) {
      throw new StreamTypeError("an array of streams", streams === null ? "null" : typeof streams, {
        field: "streams",
        where: "Streams.waitForFinished"
      });
    }
    for (let i = 0; i < streams.length; i++) {
      const stream = streams[i] as {
        on?: unknown;
        getReader?: unknown;
        getWriter?: unknown;
      } | null;
      if (
        typeof stream !== "object" ||
        stream === null ||
        (typeof stream.on !== "function" &&
          typeof stream.getReader !== "function" &&
          typeof stream.getWriter !== "function")
      ) {
        throw new StreamTypeError("a stream", stream === null ? "null" : typeof stream, {
          field: `streams[${i}]`,
          where: "Streams.waitForFinished"
        });
      }
    }
    const len = streams.length;
    if (len === 0) {
      return;
    }
    if (len === 1) {
      await finished(streams[0]);
      return;
    }
    const promises = new Array<Promise<void>>(len);
    for (let i = 0; i < len; i++) {
      promises[i] = finished(streams[i]);
    }
    await Promise.all(promises);
  };
}
