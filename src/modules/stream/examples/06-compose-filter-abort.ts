/**
 * Example: compose / filter / Streams.waitForFinished / addAbortSignal
 *
 * Demonstrates composition + functional helpers:
 * - compose()               — fuse several transforms into one transform stream
 * - Streams.filter()        — functional predicate transform (drops non-matching chunks)
 * - Streams.waitForFinished()   — await completion of multiple streams at once
 * - addAbortSignal()        — wire an AbortSignal to destroy a stream on abort
 *
 * Usage: pnpm example --filter 06-compose-filter-abort
 */

import { addAbortSignal, compose, finished, pipeline, Streams } from "@stream";

/** compose(): chain transforms into a single composite transform. */
export async function exampleCompose(): Promise<void> {
  const double = Streams.createTransform<number, number>(n => n * 2, { objectMode: true });
  const addOne = Streams.createTransform<number, number>(n => n + 1, { objectMode: true });

  // compose -> a single transform: n => (n * 2) + 1
  const composed = compose<number, number>(double, addOne);

  const source = Streams.createReadableFromArray([1, 2, 3], { objectMode: true });
  const out = source.pipe(composed);

  const results = await Streams.collect<number>(out);
  console.log("compose results:", results);
}

/** Streams.filter(): keep only chunks matching the predicate. */
export async function exampleFilter(): Promise<void> {
  const source = Streams.createReadableFromArray([1, 2, 3, 4, 5, 6], { objectMode: true });
  const evens = Streams.filter<number>(n => n % 2 === 0);

  const results = await Streams.collect<number>(source.pipe(evens));
  console.log("filter (evens):", results);
}

/** Streams.waitForFinished(): resolve once every stream has finished. */
export async function exampleFinishedAll(): Promise<void> {
  const a = Streams.createReadableFromArray(["a1", "a2"], { objectMode: true });
  const b = Streams.createReadableFromArray(["b1", "b2", "b3"], { objectMode: true });

  // Streams must be consumed to reach their end.
  a.resume();
  b.resume();

  await Streams.waitForFinished([a, b]);
  console.log("finishedAll: both streams finished");
}

/** addAbortSignal(): destroy a stream when the signal aborts. */
export async function exampleAddAbortSignal(): Promise<void> {
  const controller = new AbortController();
  const source = Streams.createReadableFromArray(["x", "y", "z"], { objectMode: true });

  addAbortSignal(controller.signal, source);

  // Abort before consuming — the stream is destroyed with an AbortError.
  controller.abort();

  try {
    await finished(source);
    console.log("addAbortSignal: finished without abort (unexpected)");
  } catch (err) {
    console.log("addAbortSignal: stream aborted ->", (err as Error).name);
  }
}

export async function exampleComposeFilterAbort(): Promise<void> {
  await exampleCompose();
  await exampleFilter();
  await exampleFinishedAll();
  await exampleAddAbortSignal();
}

// pipeline is imported to show it pairs with these helpers; keep a tiny demo.
export async function examplePipelineWithFilter(): Promise<void> {
  const source = Streams.createReadableFromArray([10, 15, 20, 25], { objectMode: true });
  const big = Streams.filter<number>(n => n >= 20);
  const collected: number[] = [];
  const sink = Streams.createTransform<number, number>(
    n => {
      collected.push(n);
      return n;
    },
    { objectMode: true }
  );
  sink.resume();
  await pipeline(source, big, sink);
  console.log("pipeline + filter (>=20):", collected);
}

await exampleComposeFilterAbort();
await examplePipelineWithFilter();
