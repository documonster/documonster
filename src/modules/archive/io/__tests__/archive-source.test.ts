import { toAsyncIterable } from "@archive/io/archive-source";
import { describe, it, expect } from "vitest";

async function collect(iter: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const c of iter) {
    chunks.push(c);
    total += c.length;
  }
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

describe("archive-source", () => {
  it("toAsyncIterable() should stream Blob correctly", async () => {
    const data = new Uint8Array(256 * 1024);
    for (let i = 0; i < data.length; i++) {
      data[i] = i & 0xff;
    }

    const blob = new Blob([data]);
    const out = await collect(toAsyncIterable(blob));

    expect(out.length).toBe(data.length);
    expect(out.slice(0, 64)).toEqual(data.slice(0, 64));
    expect(out.slice(out.length - 64)).toEqual(data.slice(data.length - 64));
  });
});
