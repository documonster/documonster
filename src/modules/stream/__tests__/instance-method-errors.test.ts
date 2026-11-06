/**
 * The stateful stream classes report a refusal as this module's error: a wrong argument as a
 * `StreamArgumentError` naming it, an operation on a destroyed stream as a `StreamStateError`.
 */
import { Streams, StreamArgumentError, StreamStateError, isStreamError } from "@stream/index";
import { describe, expect, it } from "vitest";

describe("Streams.PullStream", () => {
  it("refuses a size that is not a non-negative integer", async () => {
    const stream = new Streams.PullStream();
    await expect(stream.pull(-1)).rejects.toMatchObject({
      name: "StreamArgumentError",
      field: "size"
    });
  });

  it("reports a write to a destroyed stream as a state error", async () => {
    const stream = new Streams.PullStream();
    const reported = new Promise<unknown>(resolve => stream.on("error", resolve));
    stream.destroy();
    stream.write(new Uint8Array([1]));
    const error = await reported;
    expect(error).toBeInstanceOf(StreamStateError);
    expect(isStreamError(error)).toBe(true);
  });

  it("rejects a pending pull with a state error when destroyed", async () => {
    const stream = new Streams.PullStream();
    const pending = stream.pull(4);
    const until = stream.pullUntil(new Uint8Array([9]));
    stream.destroy();
    await expect(pending).rejects.toBeInstanceOf(StreamStateError);
    await expect(until).rejects.toBeInstanceOf(StreamStateError);
  });
});

describe("Streams.LowLevel.SinkAdapter", () => {
  it("refuses a write after end with a state error", async () => {
    const adapter = new Streams.LowLevel.SinkAdapter(new WritableStream<Uint8Array>());
    await adapter.end();
    await expect(adapter.write(new Uint8Array([1]))).rejects.toBeInstanceOf(StreamStateError);
  });
});

describe("Streams.ChunkedBuilder#pushAll", () => {
  it("refuses something that is not an array", () => {
    const builder = new Streams.ChunkedBuilder();
    expect(() => builder.pushAll(null as never)).toThrow(StreamArgumentError);
    expect(() => builder.pushAll("ab" as never)).toThrow(
      expect.objectContaining({ field: "pieces" })
    );
    expect(builder.toString()).toBe("");
  });
});

describe("Streams.ChunkedBuilder#pushAll items", () => {
  it("names the item that is not a string", () => {
    const builder = new Streams.ChunkedBuilder();
    expect(() => builder.pushAll(["a", 1 as never])).toThrow(
      expect.objectContaining({ field: "pieces[1]" })
    );
    expect(builder.toString()).toBe("");
  });
});

describe("Streams.ChunkedBuilder#pushAll holes", () => {
  it("refuses a sparse array and leaves the builder as it was", () => {
    const builder = new Streams.ChunkedBuilder();
    builder.push("kept");
    // oxlint-disable-next-line no-sparse-arrays
    expect(() => builder.pushAll(["a", , "b"] as never)).toThrow(
      expect.objectContaining({ name: "StreamArgumentError", field: "pieces[1]" })
    );
    expect(builder.toString()).toBe("kept");
    expect(builder.stringLength).toBe(4);
  });
});
