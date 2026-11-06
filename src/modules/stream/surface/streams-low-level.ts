/**
 * `Streams.LowLevel` — building blocks for adapting event-emitter-shaped objects and
 * arbitrary sinks. The library's own writers and readers are built on these; most
 * callers want the `Streams` members (`bytes`, `text`, `collect`, `fromWebReadable`, …)
 * instead.
 *
 * `import { Streams } from "documonster/stream"` →
 *   `new Streams.LowLevel.SinkAdapter(sink)`,
 *   `Streams.LowLevel.eventsToAsyncIterableNoDestroy(emitter)`,
 *   `Streams.LowLevel.onceEvent(emitter, "finish")`,
 *   `Streams.LowLevel.fromCallback(cb => fs.close(fd, cb))`.
 *
 * Identical on both platforms.
 */
export { eventedReadableToAsyncIterableNoDestroy as eventsToAsyncIterableNoDestroy } from "@stream/core/evented-readable-to-async-iterable";
export { SinkAdapter } from "@stream/core/sink-adapter";
export { onceEvent } from "@stream/core/event-utils";
export { promisify as fromCallback } from "@stream/core/utils";
