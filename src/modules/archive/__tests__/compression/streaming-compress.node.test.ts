import { createDeflateStream } from "@archive/compression/streaming-compress";
import { describe, it, expect } from "vitest";

describe("streaming-compress (Node)", () => {
  it("createDeflateStream output should be inflateRaw compatible", async () => {
    const { inflateRawSync } = await import("zlib");

    const deflate = createDeflateStream({ level: 6 });
    const chunks: Buffer[] = [];

    deflate.on("data", (chunk: Buffer) => chunks.push(chunk));

    const originalData = Buffer.from("Hello World! ".repeat(1000));
    deflate.write(originalData);

    await new Promise<void>(resolve => {
      deflate.end(() => resolve());
    });

    const compressed = Buffer.concat(chunks);
    const decompressed = inflateRawSync(compressed);

    expect(decompressed.toString()).toBe(originalData.toString());
  });
});
