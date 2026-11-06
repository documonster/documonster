import { BufferedStream } from "@stream/buffered-stream";
import { ChunkedBuilder, TransactionalChunkedBuilder } from "@stream/chunked-builder";
import { SinkAdapter } from "@stream/core/sink-adapter";
import * as entry from "@stream/index";
import { Streams } from "@stream/index";
import * as base from "@stream/index.base";
import { Collector } from "@stream/node/collector";
import { PullStream } from "@stream/pull-stream";
import { describe, expect, it } from "vitest";

describe("Streams namespace", () => {
  it("bytes/text/json consume an async iterable and a Web ReadableStream alike", async () => {
    const bytes = new TextEncoder().encode('{"a":1}');
    expect(await Streams.bytes(Streams.fromBytes(bytes))).toEqual(bytes);
    expect(await Streams.text(new Blob([bytes]).stream())).toBe('{"a":1}');
    expect(await Streams.json(Streams.fromString('{"a":1}'))).toEqual({ a: 1 });
    const latin1 = Streams.fromBytes(new Uint8Array([0xe9]));
    expect(await Streams.text(latin1, { encoding: "latin1" })).toBe("é");
  });

  it("fromJson builds a readable of the JSON text", async () => {
    expect(await Streams.json(Streams.fromJson({ x: [1, 2] }))).toEqual({ x: [1, 2] });
  });

  it("publishes the stream classes, built with their constructors", () => {
    expect(Streams.ChunkedBuilder).toBe(ChunkedBuilder);
    expect(Streams.TransactionalChunkedBuilder).toBe(TransactionalChunkedBuilder);
    expect(new Streams.Collector()).toBeInstanceOf(Collector);
    expect(new Streams.PullStream()).toBeInstanceOf(PullStream);
    expect(new Streams.BufferedStream()).toBeInstanceOf(BufferedStream);
  });

  it("keeps the deprecated constructor-forwarding factories", async () => {
    expect(Streams.createCollector()).toBeInstanceOf(Streams.Collector);
    expect(Streams.createPullStream()).toBeInstanceOf(Streams.PullStream);
    expect(Streams.createBufferedStream()).toBeInstanceOf(Streams.BufferedStream);
    expect(entry.createCollector).toBe(Streams.createCollector);
    expect(entry.createPullStream).toBe(Streams.createPullStream);
    expect(entry.createBufferedStream).toBe(Streams.createBufferedStream);
    // The constructor performs the checks the factory used to name.
    expect(() => Streams.createPullStream({ highWaterMark: -1 })).toThrow(/^Streams\.PullStream: /);
    expect(() => Streams.createBufferedStream({ batchSize: 0 })).toThrow(
      /^Streams\.BufferedStream: /
    );
    const collector = Streams.createCollector<string>({ objectMode: true });
    await new Promise<void>(resolve => collector.end("x", () => resolve()));
    expect(collector.chunks).toEqual(["x"]);
  });

  it("keeps UnsupportedStreamTypeError as a deprecated name of StreamUnsupportedTypeError", () => {
    expect(entry.UnsupportedStreamTypeError).toBe(entry.StreamUnsupportedTypeError);
    expect(new entry.UnsupportedStreamTypeError("op", "type").name).toBe(
      "StreamUnsupportedTypeError"
    );
  });

  it("keeps StreamTypeError as a deprecated name of StreamArgumentError", () => {
    expect(entry.StreamTypeError).toBe(entry.StreamArgumentError);
    const error = new entry.StreamTypeError("a string", "number", { field: "x", where: "W" });
    expect(error).toBeInstanceOf(entry.StreamArgumentError);
    expect(error).toBeInstanceOf(entry.StreamError);
    expect(error.field).toBe("x");
    expect(error.message).toBe("W: expected a string, got number");
  });

  it("keeps every flat helper as a deprecated forward", () => {
    expect(entry.createTransform).toBe(Streams.createTransform);
    expect(entry.createReadableFromArray).toBe(Streams.createReadableFromArray);
    expect(entry.fromJSON).toBe(Streams.fromJson);
    expect(entry.streamToUint8Array).toBe(Streams.bytes);
    expect(entry.drainStream).toBe(Streams.discard);
    expect(entry.finishedAll).toBe(Streams.waitForFinished);
    expect(entry.readableFromWeb).toBe(Streams.fromWebReadable);
    expect(entry.toWritable).toBe(Streams.toWritable);
    expect(entry.isTransform).toBe(Streams.isTransform);
    expect(entry.copyStream).toBe(entry.pipeline);
    expect(entry.streamToPromise).toBe(entry.finished);
    expect(entry.Collector).toBe(Collector);
    expect(entry.PullStream).toBe(PullStream);
    expect(base.ChunkedBuilder).toBe(ChunkedBuilder);
    expect(base.isReadableStreamLike).toBe(Streams.isReadableStream);
    expect(base.onceEvent).toBe(Streams.LowLevel.onceEvent);
    expect(base.collect).toBe(Streams.collect);
    expect(typeof base.SinkAdapter).toBe("function");
    expect(typeof base.eventedReadableToAsyncIterableNoDestroy).toBe("function");
  });

  it("streamToString forwards keep the positional encoding", async () => {
    const s = Streams.fromBytes(new Uint8Array([0xe9]));
    expect(await entry.streamToString(s, "latin1")).toBe("é");
  });

  it("publishes the stream classes as values alongside their factories", () => {
    expect(Streams.ChunkedBuilder).toBe(ChunkedBuilder);
    expect(Streams.TransactionalChunkedBuilder).toBe(TransactionalChunkedBuilder);
    expect(Streams.Collector).toBe(Collector);
    expect(Streams.PullStream).toBe(PullStream);
    expect(Streams.BufferedStream).toBe(BufferedStream);
    expect(Streams.LowLevel.SinkAdapter).toBe(SinkAdapter);
    expect(base.SinkAdapter).toBe(SinkAdapter);
    class Sub extends Streams.ChunkedBuilder {}
    expect(new Sub()).toBeInstanceOf(Streams.ChunkedBuilder);
  });

  it("eventsToAsyncIterable is the old NoDestroy iterator: early exit leaves the source alive", async () => {
    expect(Streams.LowLevel.eventsToAsyncIterableNoDestroy).toBe(
      base.eventedReadableToAsyncIterableNoDestroy
    );
    const source = Streams.createReadableFromArray(["a", "b", "c"], { objectMode: true });
    for await (const chunk of Streams.LowLevel.eventsToAsyncIterableNoDestroy<string>(source)) {
      expect(chunk).toBe("a");
      break;
    }
    expect(source.destroyed).toBe(false);
  });

  it("LowLevel.SinkAdapter writes through to the sink", async () => {
    const collector = new Streams.Collector();
    const sink = new Streams.LowLevel.SinkAdapter(collector);
    expect(sink).toBeInstanceOf(SinkAdapter);
    await sink.write(new TextEncoder().encode("hi"));
    await sink.end();
    expect(new TextDecoder().decode(collector.toUint8Array())).toBe("hi");
  });

  it("LowLevel.fromCallback runs the operation once; the deprecated promisify is the same function", async () => {
    expect(base.promisify).toBe(Streams.LowLevel.fromCallback);
    let calls = 0;
    const value = await Streams.LowLevel.fromCallback<number>(cb => {
      calls++;
      cb(null, 42);
    });
    expect(value).toBe(42);
    expect(calls).toBe(1);
  });

  it("discard reads a stream to its end without keeping it", async () => {
    const source = Streams.createReadableFromArray(["a", "b"], { objectMode: true });
    await Streams.discard(source);
    expect(source.readableEnded).toBe(true);
  });

  it("streamToBuffer resolves to a Uint8Array, not a Buffer, like Streams.bytes", async () => {
    const out = await entry.streamToBuffer(Streams.fromBytes(new Uint8Array([1])));
    expect(out).toBeInstanceOf(Uint8Array);
    expect(Buffer.isBuffer(out)).toBe(false);
  });

  it("the collapsed duplicates behave identically to their replacements", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    expect(await entry.streamToBuffer(Streams.fromBytes(bytes))).toEqual(
      await Streams.bytes(Streams.fromBytes(bytes))
    );
    expect(await base.bytes(Streams.fromBytes(bytes))).toEqual(bytes);
    expect(await base.text(Streams.fromString("x"))).toBe(
      await Streams.text(Streams.fromString("x"))
    );
  });
});
