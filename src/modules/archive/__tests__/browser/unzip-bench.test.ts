/**
 * Browser unzip parsing (lightweight correctness test)
 *
 * Keep this test small and deterministic (avoid timing assertions/logging)
 * so it stays stable across browsers/CI.
 */

import { zip } from "@archive/create-archive";
import { createParse, type ZipEntry } from "@archive/unzip/stream.browser";
import { describe, it, expect } from "vitest";

function createTestZip(fileCount: number, bytesPerFile: number): Uint8Array {
  const encoder = new TextEncoder();
  const content = "a".repeat(bytesPerFile);
  const z = zip({ level: 6 });
  for (let i = 0; i < fileCount; i++) {
    z.add(`file_${i}.txt`, encoder.encode(content));
  }
  return z.bytesSync();
}

describe("browser unzip parsing", () => {
  it("parses and buffers entries", async () => {
    // Keep this small to avoid flakiness in CI.
    const fileCount = 50;
    const bytesPerFile = 8 * 1024;

    const zipData = createTestZip(fileCount, bytesPerFile);

    const parse = createParse({
      // Exercise backpressure paths a bit.
      inputHighWaterMarkBytes: 256 * 1024,
      inputLowWaterMarkBytes: 64 * 1024
    });

    const buffers: Promise<Uint8Array>[] = [];

    const done = new Promise<void>((resolve, reject) => {
      parse.on("entry", (entry: ZipEntry) => {
        buffers.push(entry.buffer());
      });
      parse.on("close", () => resolve());
      parse.on("error", reject);
    });

    // Feed in chunks to simulate real streaming input.
    const chunkSize = 32 * 1024;
    for (let i = 0; i < zipData.length; i += chunkSize) {
      const ok = parse.write(zipData.subarray(i, Math.min(zipData.length, i + chunkSize)));
      // If the writable side signals backpressure, yield to let parser drain.
      if (!ok) {
        await new Promise<void>(resolve => setTimeout(resolve, 0));
      }
    }
    parse.end();

    await done;
    const results = await Promise.all(buffers);

    expect(results.length).toBe(fileCount);
    for (const b of results) {
      expect(b.length).toBe(bytesPerFile);
    }
  });
});
