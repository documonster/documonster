import { resolveArchiveSourceToBuffer, toReadableStream } from "@archive/io/archive-source";
import { toNodeReadable } from "@archive/io/archive-source.node";
import { describe, it, expect } from "vitest";

async function* makeChunks(): AsyncIterable<Uint8Array> {
  yield new Uint8Array([1, 2, 3]);
  yield new Uint8Array([4, 5]);
  yield new Uint8Array([6]);
}

describe("archive-stream adapters", () => {
  it("toReadableStream() should adapt AsyncIterable to ReadableStream", async () => {
    const stream = toReadableStream(makeChunks());
    const out = await resolveArchiveSourceToBuffer(stream);
    expect(Array.from(out)).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it("toNodeReadable() should adapt AsyncIterable to Node Readable", async () => {
    const readable = toNodeReadable(makeChunks());
    const chunks: Buffer[] = [];

    await new Promise<void>((resolve, reject) => {
      readable.on("data", (c: Buffer) => chunks.push(c));
      readable.on("end", () => resolve());
      readable.on("error", err => reject(err));
    });

    const out = Buffer.concat(chunks);
    expect(Array.from(out)).toEqual([1, 2, 3, 4, 5, 6]);
  });
});
