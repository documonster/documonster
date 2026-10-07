import { StreamBuf } from "@excel/utils/stream-buf";
import { describe, expect, it } from "vitest";

/**
 * In batch mode a `data` listener receives strings joined into chunks of about `bufSize` characters rather than one
 * chunk per write. What it receives overall — the bytes and their order — must be exactly what an unbatched stream
 * delivers.
 */
describe("StreamBuf batching to a data listener", () => {
  const decoder = new TextDecoder();

  function collect(stream: StreamBuf): { chunks: Uint8Array[]; text: () => string } {
    const chunks: Uint8Array[] = [];
    stream.on("data", (chunk: Uint8Array) => chunks.push(chunk));
    return {
      chunks,
      text: () => chunks.map(chunk => decoder.decode(chunk)).join("")
    };
  }

  it("joins small string writes into chunks of about bufSize, and emits the rest at end()", () => {
    const stream = new StreamBuf({ bufSize: 100, batch: true });
    const out = collect(stream);
    const rows = Array.from({ length: 50 }, (_, i) => `<row r="${i + 1}"/>`);
    for (const row of rows) {
      void stream.write(row);
    }
    const beforeEnd = out.chunks.length;
    stream.end();
    expect(out.text()).toBe(rows.join(""));
    expect(beforeEnd).toBeGreaterThan(0);
    expect(beforeEnd).toBeLessThan(rows.length / 2);
    expect(out.chunks.length).toBeGreaterThan(beforeEnd);
  });

  it("keeps byte writes in order with the strings around them", () => {
    const stream = new StreamBuf({ bufSize: 1000, batch: true });
    const out = collect(stream);
    void stream.write("<a>");
    void stream.write(new TextEncoder().encode("<b/>"));
    void stream.write("</a>");
    stream.end();
    expect(out.text()).toBe("<a><b/></a>");
  });

  it("emits pending strings before finish", () => {
    const stream = new StreamBuf({ bufSize: 1000, batch: true });
    const out = collect(stream);
    let atFinish = "";
    stream.on("finish", () => (atFinish = out.text()));
    void stream.write("one");
    stream.end("two");
    expect(atFinish).toBe("onetwo");
  });

  it("flushes pending strings when paused, so nothing is lost or reordered", () => {
    const stream = new StreamBuf({ bufSize: 1000, batch: true });
    const out = collect(stream);
    void stream.write("before");
    stream.pause();
    expect(out.text()).toBe("before");
  });

  it("encodes non-ASCII text the same as an unbatched stream", () => {
    const text = ["日本語", "é", "😀", "plain"];
    const batched = new StreamBuf({ bufSize: 4, batch: true });
    const plain = new StreamBuf();
    const a = collect(batched);
    const b = collect(plain);
    for (const piece of text) {
      void batched.write(piece);
      void plain.write(piece);
    }
    batched.end();
    plain.end();
    expect(a.text()).toBe(b.text());
    expect(a.text()).toBe(text.join(""));
  });

  it("is unchanged without batch: one chunk per write", () => {
    const stream = new StreamBuf();
    const out = collect(stream);
    void stream.write("a");
    void stream.write("b");
    expect(out.chunks.length).toBe(2);
  });
});

describe("StreamBuf batching keeps every write's own encoding", () => {
  function bytesOf(stream: StreamBuf, writes: string[]): number[] {
    const out: number[] = [];
    stream.on("data", (chunk: Uint8Array) => out.push(...chunk));
    for (const w of writes) {
      void stream.write(w);
    }
    stream.end();
    return out;
  }

  it.each([
    [["a\uD83D", "\uDE00b"]],
    [["\uD83D", "\uDE00"]],
    [["x", "\uDE00", "y\uD83D"]],
    [["ok😀", "fine"]]
  ])("%j encodes as it does unbatched", writes => {
    expect(bytesOf(new StreamBuf({ bufSize: 1000, batch: true }), writes)).toEqual(
      bytesOf(new StreamBuf(), writes)
    );
  });
});

describe("StreamBuf batching across listener and pipe changes", () => {
  const decoder = new TextDecoder();

  it("delivers batched text to its listener before the listener is removed", () => {
    for (const remove of [
      (stream: StreamBuf, listener: (chunk: Uint8Array) => void) =>
        stream.removeListener("data", listener),
      (stream: StreamBuf, listener: (chunk: Uint8Array) => void) => stream.off("data", listener),
      (stream: StreamBuf) => stream.removeAllListeners("data"),
      (stream: StreamBuf) => stream.removeAllListeners()
    ]) {
      const stream = new StreamBuf({ batch: true, bufSize: 1000 });
      const got: string[] = [];
      const listener = (chunk: Uint8Array) => got.push(decoder.decode(chunk));
      stream.on("data", listener);
      void stream.write("kept");
      remove(stream, listener);
      stream.end();
      expect(got).toEqual(["kept"]);
    }
  });

  it("delivers text batched for a listener before a pipe takes over", () => {
    const stream = new StreamBuf({ batch: true, bufSize: 1000 });
    const got: string[] = [];
    stream.on("data", (chunk: Uint8Array) => got.push(`listener:${decoder.decode(chunk)}`));
    void stream.write("one");
    stream.pipe({
      write: (chunk: Uint8Array, callback?: () => void) => {
        got.push(`pipe:${decoder.decode(chunk)}`);
        callback?.();
      }
    });
    void stream.write("two");
    stream.end();
    expect(got).toEqual(["listener:one", "pipe:two"]);
  });
});
