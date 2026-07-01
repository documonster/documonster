/**
 * Stream common - finishedAll
 *
 * Shared implementation for waiting on multiple streams to finish.
 */

import type { PipelineStreamLike } from "@stream/types";

/**
 * Create a `finishedAll` function bound to a platform-specific `finished`.
 */
export function createFinishedAll(
  finished: (stream: PipelineStreamLike) => Promise<void>
): (streams: ReadonlyArray<PipelineStreamLike>) => Promise<void> {
  return async function finishedAll(streams: ReadonlyArray<PipelineStreamLike>): Promise<void> {
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
