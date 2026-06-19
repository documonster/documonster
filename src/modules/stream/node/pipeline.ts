/**
 * Node.js Stream - Pipeline & Finished
 *
 * Pipeline, stream normalization, and stream completion for Node.js.
 */

import {
  Readable,
  Writable as NodeWritable,
  Transform,
  Duplex,
  finished as nodeFinished
} from "stream";
import { pipeline as nodePipeline } from "stream/promises";

import { createFinishedAll } from "@stream/core/finished-all";
import type { PipelineOptions, PipelineCallback, FinishedOptions } from "@stream/core/options";
import { isPipelineOptions } from "@stream/core/options";
import { isReadableStream, isTransformStream, isWritableStream } from "@stream/core/type-guards";
import type { PipelineStreamLike } from "@stream/types";

// Re-export for consumers
export type { PipelineOptions, FinishedOptions } from "@stream/core/options";
export { isPipelineOptions } from "@stream/core/options";

// =============================================================================
// Pipeline
// =============================================================================

type PipelineStream = PipelineStreamLike;

export const toNodePipelineStream = (stream: PipelineStream): unknown => {
  // Node native streams (Readable/Transform/Duplex/Writable) are already compatible.
  if (
    stream instanceof Readable ||
    stream instanceof Transform ||
    stream instanceof Duplex ||
    stream instanceof NodeWritable
  ) {
    return stream;
  }

  if (isTransformStream(stream)) {
    return (Transform as any).fromWeb(stream as any);
  }
  if (isReadableStream(stream)) {
    return (Readable as any).fromWeb(stream as any);
  }
  if (isWritableStream(stream)) {
    return (NodeWritable as any).fromWeb(stream as any);
  }

  return stream;
};

/**
 * Pipeline streams together with proper error handling and cleanup.
 * Node.js compatible with support for options and callbacks.
 */
export function pipeline(
  ...args: [...PipelineStream[], PipelineOptions | PipelineCallback] | PipelineStream[]
): Promise<void> {
  let streams: PipelineStream[];
  let options: PipelineOptions | undefined;
  let callback: PipelineCallback | undefined;

  const lastArg = args[args.length - 1] as unknown;

  if (typeof lastArg === "function") {
    callback = lastArg as PipelineCallback;
    // Check for combined style: pipeline(s1, s2, ..., options, callback)
    const secondToLast = args[args.length - 2];
    if (isPipelineOptions(secondToLast)) {
      options = secondToLast;
      streams = args.slice(0, -2) as PipelineStream[];
    } else {
      // Callback only: pipeline(s1, s2, ..., callback)
      streams = args.slice(0, -1) as PipelineStream[];
    }
  } else if (isPipelineOptions(lastArg)) {
    options = lastArg;
    streams = args.slice(0, -1) as PipelineStream[];
  } else {
    streams = args as PipelineStream[];
  }

  const normalizedStreams = streams.map(toNodePipelineStream);

  if (streams.length < 2) {
    return Promise.reject(new Error("Pipeline requires at least 2 streams"));
  }

  const promise: Promise<void> = options
    ? (nodePipeline as any)(...normalizedStreams, options)
    : (nodePipeline as any)(...normalizedStreams);

  if (callback) {
    promise.then(() => callback!()).catch(err => callback!(err));
  }

  return promise;
}

// =============================================================================
// Finished
// =============================================================================

/**
 * Wait for a stream to finish, close, or error.
 * Node.js compatible with support for options and callbacks.
 */
export function finished(
  stream: PipelineStreamLike,
  optionsOrCallback?: FinishedOptions | PipelineCallback,
  callback?: PipelineCallback
): Promise<void> {
  let options: FinishedOptions | undefined;
  let cb: PipelineCallback | undefined;

  if (typeof optionsOrCallback === "function") {
    cb = optionsOrCallback;
  } else {
    options = optionsOrCallback;
    cb = callback;
  }

  const promise = new Promise<void>((resolve, reject) => {
    const normalizedStream = toNodePipelineStream(stream);
    (nodeFinished as any)(normalizedStream, options, (err: Error | null) => {
      // Node.js native finished() already handles options.error internally.
      // With error:false it still passes the error through the close handler
      // (via stream.errored check), so we must NOT filter it here.
      if (err) {
        reject(err);
        return;
      }
      resolve();
    });
  });

  if (cb) {
    promise.then(() => cb!()).catch(err => cb!(err));
  }

  return promise;
}

/**
 * Wait for multiple streams to finish
 */
export const finishedAll = createFinishedAll(finished);
