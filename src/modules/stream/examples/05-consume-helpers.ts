/**
 * Example: consume helpers (collect / json / bytes / streamTo* / drain / copy)
 *
 * Demonstrates the full family of consumption utilities:
 * - Streams.collect()        — gather every chunk of an async-iterable stream into an array
 * - Streams.json() / Streams.fromJson()   — round-trip a JS value through a byte stream
 * - Streams.bytes() / Streams.fromBytes() — round-trip raw bytes through a stream
 * - Streams.text() / Streams.bytes() — drain to a string or a Uint8Array
 *   (`Streams.text(s, { encoding })` for non-UTF-8 text)
 * - Streams.discard()    — consume and discard everything (run side effects to end)
 * - pipeline()     — pipe a readable into a writable
 * - Streams.LowLevel.eventsToAsyncIterableNoDestroy() / new Streams.LowLevel.SinkAdapter() — adapt emitters and sinks
 *
 * Usage: pnpm example --filter 05-consume-helpers
 */

import { pipeline, Streams } from "@stream";

const decoder = new TextDecoder();

/** Streams.collect(): gather all object-mode chunks into a plain array. */
export async function exampleCollect(): Promise<void> {
  const source = Streams.createReadableFromArray([1, 2, 3, 4], { objectMode: true });
  const values = await Streams.collect<number>(source);
  console.log("collect:", values);
}

/** Streams.json()/Streams.fromJson(): serialize a value to a byte stream and read it back. */
export async function exampleJson(): Promise<void> {
  const payload = { name: "stream", count: 42, tags: ["a", "b"] };
  const stream = Streams.fromJson(payload);
  const parsed = await Streams.json<typeof payload>(stream);
  console.log("json round-trip:", parsed);
}

/** Streams.bytes()/fromBytes(): round-trip raw bytes. */
export async function exampleBytes(): Promise<void> {
  const original = new Uint8Array([0xde, 0xad, 0xbe, 0xef]);
  const collected = await Streams.bytes(Streams.fromBytes(original));
  console.log("bytes round-trip:", Array.from(collected));
}

/** Streams.text / Streams.bytes: drain a readable. */
export async function exampleStreamTo(): Promise<void> {
  const text = "The quick brown fox";

  // Each helper consumes a fresh stream (streams are one-shot).
  const asString = await Streams.text(Streams.fromBytes(new TextEncoder().encode(text)));
  console.log("text:", asString);

  const asU8 = await Streams.bytes(Streams.fromBytes(new TextEncoder().encode(text)));
  console.log("bytes length:", asU8.length, "head:", decoder.decode(asU8.subarray(0, 3)));

  const latin1 = await Streams.text(Streams.fromBytes(new Uint8Array([0x63, 0x61, 0x66, 0xe9])), {
    encoding: "latin1"
  });
  console.log("latin1 text:", latin1);
}

/** Streams.discard(): consume every chunk without keeping it. */
export async function exampleDrain(): Promise<void> {
  let count = 0;
  const source = Streams.createReadableFromArray(["a", "b", "c"], { objectMode: true });
  source.on("data", () => count++);
  await Streams.discard(source);
  console.log("drain consumed chunks:", count);
}

/** pipeline(): pipe a readable into a writable. */
export async function exampleCopy(): Promise<void> {
  const source = Streams.createReadableFromArray(["x", "y", "z"], { objectMode: true });
  const sink = Streams.createCollector<string>();
  await pipeline(source, sink);
  console.log("pipeline collected:", sink.chunks);
}

/**
 * Streams.LowLevel.eventsToAsyncIterableNoDestroy(): iterate an evented readable and stop early without
 * destroying it; new Streams.LowLevel.SinkAdapter(): one awaited `write`/`end` over any sink.
 */
export async function exampleEventsAndSinks(): Promise<void> {
  const source = Streams.createReadableFromArray(["first", "second"], { objectMode: true });
  for await (const chunk of Streams.LowLevel.eventsToAsyncIterableNoDestroy<string>(source)) {
    console.log("first chunk:", chunk);
    break;
  }
  console.log("source destroyed after break:", source.destroyed);

  const collector = Streams.createCollector();
  const sink = new Streams.LowLevel.SinkAdapter(collector);
  await sink.write(new TextEncoder().encode("via SinkAdapter"));
  await sink.end();
  console.log("sink received:", decoder.decode(collector.toUint8Array()));
}

export async function exampleConsumeHelpers(): Promise<void> {
  await exampleCollect();
  await exampleJson();
  await exampleBytes();
  await exampleStreamTo();
  await exampleDrain();
  await exampleCopy();
  await exampleEventsAndSinks();
}

await exampleConsumeHelpers();
