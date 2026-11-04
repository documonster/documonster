/**
 * Node.js Stream - pipeline() and finished()
 *
 * The module's calling convention over `node:stream`; see `pipeline` for which
 * implementation runs which kind of pipeline.
 */

import { Duplex, PassThrough, Readable, Writable, finished as nodeFinished } from "stream";
import { pipeline as nodePipeline } from "stream/promises";

import { createFinishedAll } from "@stream/core/finished-all";
import type { FinishedOptions, PipelineCallback, PipelineOptions } from "@stream/core/options";
import { createPipeline } from "@stream/core/pipeline-api";
import { isWebOnlyPipeline, pipeWebStreams } from "@stream/core/web-pipeline";
import { assertDistinctStages, classifyStages, pipelineImpl } from "@stream/kernel/pipeline-impl";
import type { StreamKernel } from "@stream/kernel/state";
import { isWebStream } from "@stream/kernel/stream-utils";
import type { PipelineStageLike, PipelineStreamLike } from "@stream/types";

export type { PipelineOptions, FinishedOptions } from "@stream/core/options";
export { isPipelineOptions } from "@stream/core/options";

/** Node's own classes, for the module's pipeline on a mixed Node/Web pipeline. */
const nativeKernel: StreamKernel = { Readable, Writable, Duplex, PassThrough };

/**
 * Which implementation runs a pipeline:
 *
 * - Node streams, with an iterable source at most → Node's own
 *   `stream.pipeline`. Every stage is checked first, so nothing can fail after
 *   the first connections are made.
 * - only Web streams → the platform's `pipeTo`.
 * - anything else (a Web stream among Node ones, a generator stage) → the
 *   module's pipeline over Node's classes. Node's pumps a Web stream directly,
 *   so an abort never reaches a read waiting for data and a Web writable left
 *   open by `end: false` stays locked; and a generator whose return value is
 *   invalid is found only after earlier stages are connected, which Node's
 *   leaves running.
 */
const runPipeline = createPipeline((stages, options) => {
  if (isWebOnlyPipeline(stages)) {
    return pipeWebStreams(stages, options);
  }
  assertDistinctStages(stages);
  const checked = classifyStages(nativeKernel, stages);
  if (checked.some(stage => typeof stage === "function" || isWebStream(stage))) {
    return new Promise<void>((resolve, reject) => {
      pipelineImpl(nativeKernel, checked, err => (err ? reject(err) : resolve()), options);
    });
  }
  const run = nodePipeline as (...pipelineArgs: unknown[]) => Promise<unknown>;
  return (options ? run(...checked, options) : run(...checked)).then(() => undefined);
});

/**
 * Pipe `streams` together. With a trailing callback it is called once every stage is done
 * (or with the first error) and the last stage is returned, as `node:stream`'s `pipeline`
 * does; without one, a promise settles instead.
 */
export function pipeline<D extends PipelineStageLike>(
  ...args:
    | [...PipelineStageLike[], D, PipelineCallback]
    | [...PipelineStageLike[], D, PipelineOptions, PipelineCallback]
): D;
export function pipeline(
  ...args: [...PipelineStageLike[], PipelineOptions] | PipelineStageLike[]
): Promise<void>;
export function pipeline(
  ...args: [...PipelineStageLike[], PipelineOptions | PipelineCallback] | PipelineStageLike[]
): Promise<void> | PipelineStageLike {
  return runPipeline(...args);
}

/**
 * Wait for `stream` to finish, error or close early. With a callback it is called then and
 * a function that removes the listeners is returned, as `node:stream`'s `finished` does;
 * without one, a promise settles instead.
 */
export function finished(stream: PipelineStreamLike, callback: PipelineCallback): () => void;
export function finished(
  stream: PipelineStreamLike,
  options: FinishedOptions | undefined,
  callback: PipelineCallback
): () => void;
export function finished(stream: PipelineStreamLike, options?: FinishedOptions): Promise<void>;
export function finished(
  stream: PipelineStreamLike,
  optionsOrCallback?: FinishedOptions | PipelineCallback,
  callback?: PipelineCallback
): Promise<void> | (() => void) {
  let options: FinishedOptions | undefined;
  let cb: PipelineCallback | undefined;
  if (typeof optionsOrCallback === "function") {
    cb = optionsOrCallback;
  } else {
    options = optionsOrCallback;
    cb = callback;
  }

  const run = nodeFinished as (
    target: unknown,
    opts: FinishedOptions,
    done: (err?: Error | null) => void
  ) => () => void;
  if (cb) {
    const done = cb;
    return run(stream, options ?? {}, err => (err ? done(err) : done()));
  }
  return new Promise<void>((resolve, reject) => {
    run(stream, options ?? {}, err => (err ? reject(err) : resolve()));
  });
}

export const finishedAll = createFinishedAll(finished);
