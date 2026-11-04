/**
 * With a callback, `pipeline` and `finished` return what `node:stream`'s return — the last
 * stage and a cleanup function — so code written against Node's API works unchanged. Without
 * one they return a promise, as before.
 */
import { Streams, finished, pipeline } from "@stream/index";
import { describe, expect, it } from "vitest";

const source = () => Streams.createReadableFromArray(["a", "b"]);

describe("pipeline with a callback", () => {
  it("returns the last stage and calls back once it is done", async () => {
    const sink = Streams.createCollector<string>({ objectMode: true });
    const done = new Promise<Error | null | undefined>(resolve => {
      const returned = pipeline(source(), sink, err => resolve(err));
      expect(returned).toBe(sink);
    });
    expect(await done).toBeFalsy();
    expect(sink.chunks).toEqual(["a", "b"]);
  });

  it("returns the last stage when options come before the callback", async () => {
    const sink = Streams.createCollector<string>({ objectMode: true });
    const done = new Promise<void>(resolve => {
      expect(pipeline(source(), sink, { end: true }, () => resolve())).toBe(sink);
    });
    await done;
  });

  it("still returns a promise without a callback", async () => {
    const sink = Streams.createCollector<string>({ objectMode: true });
    const returned = pipeline(source(), sink);
    expect(returned).toBeInstanceOf(Promise);
    await returned;
    expect(sink.chunks).toEqual(["a", "b"]);
  });
});

describe("finished with a callback", () => {
  it("returns a cleanup function and calls back once the stream is done", async () => {
    const stream = source();
    const done = new Promise<void>(resolve => {
      const cleanup = finished(stream, () => resolve());
      expect(typeof cleanup).toBe("function");
    });
    stream.resume();
    await done;
  });

  it("the cleanup function stops the callback", async () => {
    const stream = source();
    let called = false;
    const cleanup = finished(stream, {}, () => {
      called = true;
    });
    cleanup();
    stream.resume();
    await Streams.waitForFinished([stream]);
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(called).toBe(false);
  });

  it("still returns a promise without a callback", async () => {
    const stream = source();
    const returned = finished(stream);
    expect(returned).toBeInstanceOf(Promise);
    stream.resume();
    await returned;
  });
});
