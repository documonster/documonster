/**
 * Example: readable/writable factory family + chunk constructors
 *
 * Demonstrates the remaining stream creation helpers:
 * - Streams.createReadableFromGenerator()    — drive a readable from an async generator
 * - Streams.createReadableFromAsyncIterable() — wrap any async iterable as a readable
 * - Streams.createReadableFromPromise()      — emit a single value once a promise settles
 * - Streams.createEmptyReadable()            — a readable that ends immediately
 * - Streams.createNullWritable()             — a sink that discards everything (/dev/null)
 * - Streams.createStringChunk() / Streams.createByteChunk() — build StreamDataChunk objects from
 *   strings / bytes for use with BufferedStream-style buffering
 *
 * Usage: pnpm example --filter 08-readable-factories
 */

import { finished, Streams } from "@stream";

const decoder = new TextDecoder();

/** Streams.createReadableFromGenerator(): pull values from an async generator. */
export async function exampleFromGenerator(): Promise<void> {
  async function* countUp(): AsyncGenerator<number, void, unknown> {
    for (let i = 1; i <= 4; i++) {
      yield i * 10;
    }
  }
  const source = Streams.createReadableFromGenerator(countUp);
  const values = await Streams.collect<number>(source);
  console.log("createReadableFromGenerator:", values);
}

/** Streams.createReadableFromAsyncIterable(): wrap an arbitrary async iterable. */
export async function exampleFromAsyncIterable(): Promise<void> {
  const iterable: AsyncIterable<string> = {
    async *[Symbol.asyncIterator]() {
      yield "alpha";
      yield "beta";
      yield "gamma";
    }
  };
  const source = Streams.createReadableFromAsyncIterable(iterable);
  const values = await Streams.collect<string>(source);
  console.log("createReadableFromAsyncIterable:", values);
}

/** Streams.createReadableFromPromise(): emit one resolved value, then end. */
export async function exampleFromPromise(): Promise<void> {
  const source = Streams.createReadableFromPromise(Promise.resolve({ ready: true, id: 7 }));
  const values = await Streams.collect<{ ready: boolean; id: number }>(source);
  console.log("createReadableFromPromise:", values);
}

/** Streams.createEmptyReadable(): a readable that yields nothing and ends. */
export async function exampleEmptyReadable(): Promise<void> {
  const source = Streams.createEmptyReadable<Uint8Array>();
  const text = await Streams.text(source);
  console.log("createEmptyReadable produced length:", text.length);
}

/** Streams.createNullWritable(): discard every chunk written to it. */
export async function exampleNullWritable(): Promise<void> {
  const sink = Streams.createNullWritable();
  const encoder = new TextEncoder();
  sink.write(encoder.encode("this data goes nowhere"));
  sink.write(encoder.encode(" and so does this"));
  sink.end();
  await finished(sink);
  console.log("createNullWritable discarded all writes (writableEnded:", sink.writableEnded, ")");
}

/** Streams.createStringChunk() / Streams.createByteChunk(): build StreamDataChunk objects. */
export function exampleChunkConstructors(): void {
  const strChunk = Streams.createStringChunk("héllo");
  console.log("createStringChunk length (bytes):", strChunk.length);
  console.log("createStringChunk decoded:", decoder.decode(strChunk.toUint8Array()));

  const byteChunk = Streams.createByteChunk(new Uint8Array([0x42, 0x59, 0x54, 0x45]));
  console.log("createByteChunk length:", byteChunk.length);
  console.log("createByteChunk decoded:", decoder.decode(byteChunk.toUint8Array()));

  // copy() blits a slice of the chunk into a caller-provided target buffer.
  const target = new Uint8Array(2);
  const copied = byteChunk.copy(target, 0, 0, 2);
  console.log("createByteChunk copy() copied bytes:", copied, "->", Array.from(target));
}

export async function exampleReadableFactories(): Promise<void> {
  await exampleFromGenerator();
  await exampleFromAsyncIterable();
  await exampleFromPromise();
  await exampleEmptyReadable();
  await exampleNullWritable();
  exampleChunkConstructors();
}

await exampleReadableFactories();
