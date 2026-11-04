/**
 * Contracts of the `Streams` helpers: a value a stream cannot carry is refused rather than
 * silently ending the stream, arguments are checked where they are given, and a builder's
 * readers do not change it.
 */
import { AbortError as ArchiveAbortError } from "@archive/index";
import {
  AbortError,
  StreamTypeError,
  Streams,
  UnsupportedStreamTypeError,
  addAbortSignal,
  isStreamError
} from "@stream/index";
import { describe, expect, it } from "vitest";

describe("Streams.createReadableFromArray", () => {
  it("refuses a null entry, which would end the stream and drop what follows", () => {
    expect(() => Streams.createReadableFromArray(["a", null, "b"])).toThrow(StreamTypeError);
    expect(() => Streams.createReadableFromArray([1, null])).toThrow("index 1");
  });

  it("still skips undefined entries and keeps falsy values", async () => {
    expect(
      await Streams.collect(Streams.createReadableFromArray([0, undefined, "", false]))
    ).toEqual([0, "", false]);
  });

  it("refuses something that is not an array", () => {
    expect(() => Streams.createReadableFromArray("abc" as never)).toThrow(StreamTypeError);
  });
});

describe("Streams argument checks", () => {
  it.each([
    ["collect", () => Streams.collect(null as never)],
    ["waitForFinished", () => Streams.waitForFinished(null as never)],
    ["fromString", () => Streams.fromString(42 as never)],
    ["fromBytes", () => Streams.fromBytes("abc" as never)],
    ["fromBytes", () => Streams.fromBytes(null as never)]
  ])("%s refuses a wrong argument with a stream error", async (_name, call) => {
    let error: unknown;
    try {
      await call();
    } catch (caught) {
      error = caught;
    }
    expect(isStreamError(error)).toBe(true);
  });

  it("toWritable refuses something that is not a writable", () => {
    expect(() => Streams.toWritable(null as never)).toThrow(UnsupportedStreamTypeError);
    expect(() => Streams.toWritable({} as never)).toThrow(UnsupportedStreamTypeError);
  });

  it("SinkAdapter refuses something that is not a sink", () => {
    expect(() => new Streams.LowLevel.SinkAdapter(null as never)).toThrow(
      UnsupportedStreamTypeError
    );
  });
});

describe("ChunkedBuilder readers", () => {
  it("toString() leaves cursor and length as they were", () => {
    const builder = new Streams.ChunkedBuilder({ chunkSize: 2 });
    builder.push("a");
    builder.push("b");
    builder.push("c");
    const before = [builder.cursor, builder.length];
    expect(builder.toString()).toBe("abc");
    expect([builder.cursor, builder.length]).toEqual(before);
    builder.push("d");
    expect(builder.cursor).toBe(4);
    expect(builder.toString()).toBe("abcd");
  });

  it("a rollback after toString() inside a snapshot restores the snapshot", () => {
    const builder = new Streams.TransactionalChunkedBuilder({ chunkSize: 1 });
    builder.push("x");
    builder.snapshot();
    builder.push("y");
    expect(builder.toString()).toBe("xy");
    builder.rollback();
    expect(builder.toString()).toBe("x");
    expect(builder.stringLength).toBe(1);
  });

  it("refuses a piece that is not a string", () => {
    const builder = new Streams.ChunkedBuilder();
    expect(() => builder.push(42 as never)).toThrow(StreamTypeError);
    expect(() => builder.pushAll(["a", 1] as never)).toThrow(StreamTypeError);
    expect(builder.toString()).toBe("");
    expect(builder.stringLength).toBe(0);
  });
});

describe("BufferedStream.readAll", () => {
  it("takes everything buffered, as the deprecated toUint8Array did", () => {
    const stream = Streams.createBufferedStream();
    stream.write("hello");
    stream.write(new TextEncoder().encode(" world"));
    expect(new TextDecoder().decode(stream.readAll())).toBe("hello world");
    expect(stream.bufferedLength).toBe(0);
    expect(stream.readAll()).toEqual(new Uint8Array(0));
    stream.write("again");
    expect(new TextDecoder().decode(stream.toUint8Array())).toBe("again");
  });
});

describe("PullStream.pullUntil", () => {
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const stream = () => {
    const pull = Streams.createPullStream();
    pull.write(encoder.encode("a\nb\nc"));
    pull.end();
    return pull;
  };

  it("takes { includePattern }, as the deprecated boolean did", async () => {
    const pull = stream();
    expect(decoder.decode(await pull.pullUntil(encoder.encode("\n")))).toBe("a");
    expect(
      decoder.decode(await pull.pullUntil(encoder.encode("\n"), { includePattern: true }))
    ).toBe("b\n");
    const legacy = stream();
    expect(decoder.decode(await legacy.pullUntil(encoder.encode("\n"), true))).toBe("a\n");
    expect(decoder.decode(await legacy.pull(encoder.encode("\n"), false))).toBe("b");
  });
});

describe("the remaining Streams entry points check their arguments", () => {
  const refuses = async (call: () => unknown) => {
    let error: unknown;
    try {
      await Promise.resolve(call());
    } catch (caught) {
      error = caught;
    }
    return isStreamError(error);
  };

  it.each<[string, () => unknown]>([
    ["BufferedStream.write(1)", () => Streams.createBufferedStream().write(1 as never)],
    ["PullStream.write('abc')", () => Streams.createPullStream().write("abc" as never)],
    ["PullStream.pullUntil('x')", () => Streams.createPullStream().pullUntil("x" as never)],
    ["PullStream.pullUntil(empty)", () => Streams.createPullStream().pullUntil(new Uint8Array(0))],
    ["createTransform(null)", () => Streams.createTransform(null as never)],
    ["filter(null)", () => Streams.filter(null as never)],
    ["createReadableFromGenerator(null)", () => Streams.createReadableFromGenerator(null as never)],
    [
      "createReadableFromAsyncIterable(null)",
      () => Streams.createReadableFromAsyncIterable(null as never)
    ],
    ["createReadableFromPromise(42)", () => Streams.createReadableFromPromise(42 as never)],
    ["fromWebReadable(null)", () => Streams.fromWebReadable(null as never)],
    ["toWebReadable(null)", () => Streams.toWebReadable(null as never)],
    ["fromWebWritable(null)", () => Streams.fromWebWritable(null as never)],
    ["toWebWritable(null)", () => Streams.toWebWritable(null as never)],
    ["fromJson(() => 1)", () => Streams.fromJson(() => 1)],
    ["fromJson(1n)", () => Streams.fromJson(1n)],
    ["fromJson(undefined)", () => Streams.fromJson(undefined)],
    ["createStringChunk(1)", () => Streams.createStringChunk(1 as never)],
    ["createByteChunk('a')", () => Streams.createByteChunk("a" as never)],
    [
      "readableStreamToAsyncIterable(null)",
      () => Streams.readableStreamToAsyncIterable(null as never).next()
    ],
    [
      "createPullStream({ highWaterMark: -1 })",
      () => Streams.createPullStream({ highWaterMark: -1 })
    ],
    [
      "createBufferedStream({ batchSize: 0 })",
      () => Streams.createBufferedStream({ batchSize: 0 } as never)
    ]
  ])("%s is refused with a stream error", async (_name, call) => {
    expect(await refuses(call)).toBe(true);
  });

  it("fromJson names what it could not serialise", () => {
    expect(() => Streams.fromJson(() => 1)).toThrow(/function/);
    expect(() => Streams.fromJson(1n)).toThrow(/bigint|BigInt/);
  });

  it("a refused BufferedStream.write keeps what was written before and after", () => {
    const buffered = Streams.createBufferedStream();
    buffered.write("a");
    expect(() => buffered.write(1 as never)).toThrow(StreamTypeError);
    buffered.write("b");
    expect(new TextDecoder().decode(buffered.readAll())).toBe("ab");
  });
});

describe("argument checks name the field", () => {
  const fieldOf = async (call: () => unknown): Promise<unknown> => {
    try {
      await Promise.resolve(call());
    } catch (error) {
      expect(isStreamError(error)).toBe(true);
      return (error as { field?: string }).field;
    }
    return "did not throw";
  };

  it.each<[string, () => unknown]>([
    // The path from the parameter name: an option is `options.<name>`.
    ["options.highWaterMark", () => Streams.createPullStream({ highWaterMark: -1 })],
    ["chunk", () => Streams.createPullStream().write("x" as never)],
    ["pattern", () => Streams.createPullStream().pullUntil("x" as never)],
    ["streams", () => Streams.waitForFinished(null as never)],
    ["iterable", () => Streams.createReadableFromAsyncIterable(null as never)],
    ["data", () => Streams.createReadableFromArray("x" as never)],
    ["promise", () => Streams.createReadableFromPromise(42 as never)],
    ["generator", () => Streams.createReadableFromGenerator(null as never)],
    ["transformFn", () => Streams.createTransform(null as never)],
    ["predicate", () => Streams.filter(null as never)],
    ["data", () => Streams.fromBytes("x" as never)],
    ["str", () => Streams.fromString(1 as never)],
    ["data", () => Streams.fromJson(() => 1)],
    ["stream", () => Streams.collect(null as never)],
    ["stream", () => Streams.toWritable(null as never)],
    ["stream", () => Streams.fromWebReadable(null as never)],
    ["stream", () => Streams.toWebWritable(null as never)],
    ["sink", () => new Streams.LowLevel.SinkAdapter(null as never)],
    ["piece", () => new Streams.ChunkedBuilder().push(1 as never)],
    ["data", () => Streams.createStringChunk(1 as never)],
    ["options.batchSize", () => Streams.createBufferedStream({ batchSize: 0 })],
    ["chunk", () => Streams.createBufferedStream().write(1 as never)]
  ])("%s", async (field, call) => {
    expect(await fieldOf(call)).toBe(field);
  });

  it.each<[string, () => unknown]>([
    ["Streams.bytes", () => Streams.bytes(42 as never)],
    ["Streams.discard", () => Streams.discard(42 as never)],
    ["Streams.text", () => Streams.text(42 as never)],
    ["Streams.json", () => Streams.json(42 as never)],
    ["Streams.collect", () => Streams.collect(42 as never)],
    ["Streams.createTransform", () => Streams.createTransform(null as never)],
    ["Streams.fromString", () => Streams.fromString(1 as never)],
    ["Streams.PullStream#pullUntil", () => Streams.createPullStream().pullUntil("x" as never)],
    ["Streams.BufferedStream#write", () => Streams.createBufferedStream().write(1 as never)],
    ["Streams.toWebWritable", () => Streams.toWebWritable(null as never)]
  ])("starts the message with the public call %s", async (where, call) => {
    const error = await Promise.resolve()
      .then(call)
      .then(
        () => undefined,
        (e: unknown) => e
      );
    expect(isStreamError(error)).toBe(true);
    expect((error as Error).message.startsWith(`${where}: `), (error as Error).message).toBe(true);
  });

  it("is absent on an error about the stream's contents", async () => {
    const error = await Streams.text(Streams.createReadableFromArray([1, 2]) as never).catch(
      e => e
    );
    expect(isStreamError(error)).toBe(true);
    expect((error as { field?: string }).field).toBeUndefined();
  });

  it("is carried by the constructors", () => {
    expect(new StreamTypeError("a string", "number", { field: "x" }).field).toBe("x");
    expect(new UnsupportedStreamTypeError("collect", "null", { field: "stream" }).field).toBe(
      "stream"
    );
    expect(new StreamTypeError("a string", "number").field).toBeUndefined();
  });
});

describe("Streams handles", () => {
  it.each<[string, () => unknown]>([
    ["streams[0]", () => Streams.waitForFinished([42 as never])],
    ["emitter", () => Streams.LowLevel.onceEvent(null as never, "finish")],
    ["stream", () => Streams.LowLevel.eventsToAsyncIterableNoDestroy(null as never)],
    ["fn", () => Streams.LowLevel.fromCallback(null as never)]
  ])("%s is checked", async (field, call) => {
    let error: unknown;
    try {
      await call();
    } catch (caught) {
      error = caught;
    }
    expect(isStreamError(error)).toBe(true);
    expect((error as { field?: string }).field).toBe(field);
  });
});

describe("AbortError", () => {
  it("is the class documonster/archive exports", () => {
    expect(AbortError).toBe(ArchiveAbortError);
  });

  it("is what an aborted stream is destroyed with", async () => {
    const controller = new AbortController();
    const stream = addAbortSignal(controller.signal, Streams.createReadableFromArray([1, 2, 3]));
    const failed = new Promise<unknown>(resolve => stream.on("error", resolve));
    controller.abort();
    const error = await failed;
    expect(error).toBeInstanceOf(AbortError);
    expect(isStreamError(error)).toBe(false);
  });
});
