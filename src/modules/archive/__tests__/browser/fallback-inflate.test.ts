import { Parse } from "@archive/unzip/stream.browser";
import { StreamingZip, ZipDeflateFile } from "@archive/zip/stream";
import { concatUint8Arrays } from "@utils/binary";
/**
 * Test that Parse.browser.ts (with FallbackInflateRaw) can parse
 * ZIP files created by StreamingZip
 */
import { describe, it, expect } from "vitest";

// Type helper for browser Parse which has different methods than Node version
type BrowserParse = Parse & {
  on(event: string, listener: (...args: any[]) => void): void;
  write(chunk: Uint8Array): void;
  end(): void;
};

async function createZipBytesWithStreamingZip(
  entries: Array<{ name: string; content: Uint8Array; level?: number }>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];

  let resolveFinish: (() => void) | null = null;
  const finishPromise = new Promise<void>(resolve => {
    resolveFinish = resolve;
  });

  const zip = new StreamingZip((err: Error | null, data: Uint8Array, final: boolean) => {
    if (err) {
      throw err;
    }
    if (data && data.length > 0) {
      chunks.push(data);
    }
    if (final) {
      resolveFinish?.();
    }
  });

  for (const entry of entries) {
    const file = new ZipDeflateFile(entry.name, { level: entry.level ?? 6 });
    zip.add(file);
    await file.push(entry.content, true);
    await file.complete();
  }

  zip.end();
  await finishPromise;
  return concatUint8Arrays(chunks);
}

describe("FallbackInflateRaw", () => {
  it("should fall back to non-streaming DEFLATE when CompressionStream is unavailable", async () => {
    const originalCompressionStream = globalThis.CompressionStream;
    const originalDecompressionStream = globalThis.DecompressionStream;

    // Simulate an older browser without CompressionStream/DecompressionStream.
    globalThis.CompressionStream = undefined as any;
    globalThis.DecompressionStream = undefined as any;

    try {
      const payload = new TextEncoder().encode("stream-me");
      const fullZip = await createZipBytesWithStreamingZip([
        { name: "test.txt", content: payload, level: 6 }
      ]);

      const parser = new Parse() as BrowserParse;
      const entries: Array<{ name: string; content: string }> = [];

      const entryPromise = new Promise<void>((resolve, reject) => {
        parser.on("entry", (entry: any) => {
          if (entry.type === "File") {
            const readChunks: Uint8Array[] = [];
            entry.on("data", (chunk: Uint8Array) => readChunks.push(chunk));
            entry.on("end", () => {
              const combined = concatUint8Arrays(readChunks);
              entries.push({
                name: entry.path,
                content: new TextDecoder().decode(combined)
              });
              resolve();
            });
            entry.on("error", (err: Error) => reject(err));
          } else {
            entry.autodrain();
          }
        });
        parser.on("error", (err: Error) => reject(err));
      });

      parser.write(fullZip);
      parser.end();
      await entryPromise;

      expect(entries.length).toBe(1);
      expect(entries[0]!.name).toBe("test.txt");
      expect(entries[0]!.content).toBe("stream-me");
    } finally {
      globalThis.CompressionStream = originalCompressionStream;
      globalThis.DecompressionStream = originalDecompressionStream;
    }
  }, 30000);

  it("should parse ZIP created by StreamingZip", async () => {
    const fullZip = await createZipBytesWithStreamingZip([
      { name: "test.txt", content: new TextEncoder().encode("Hello, World!"), level: 6 }
    ]);

    const parser = new Parse() as BrowserParse;
    const entries: Array<{ name: string; content: string }> = [];

    const entryPromise = new Promise<void>((resolve, reject) => {
      parser.on("entry", (entry: any) => {
        if (entry.type === "File") {
          const readChunks: Uint8Array[] = [];
          entry.on("data", (chunk: Uint8Array) => readChunks.push(chunk));
          entry.on("end", () => {
            const combined = concatUint8Arrays(readChunks);
            entries.push({
              name: entry.path,
              content: new TextDecoder().decode(combined)
            });
            resolve();
          });
          entry.on("error", (err: Error) => reject(err));
        } else {
          entry.autodrain();
        }
      });

      parser.on("error", (err: Error) => reject(err));
    });

    parser.write(fullZip);
    parser.end();
    await entryPromise;

    expect(entries.length).toBe(1);
    expect(entries[0]!.name).toBe("test.txt");
    expect(entries[0]!.content).toBe("Hello, World!");
  }, 30000);

  it("should parse ZIP with multiple files", async () => {
    const encoder = new TextEncoder();
    const fullZip = await createZipBytesWithStreamingZip([
      { name: "file1.txt", content: encoder.encode("File 1 content"), level: 6 },
      { name: "file2.txt", content: encoder.encode("File 2 content"), level: 6 },
      { name: "dir/file3.txt", content: encoder.encode("File 3 content"), level: 6 }
    ]);

    const parser = new Parse() as BrowserParse;
    const entries: Array<{ name: string; content: string }> = [];

    const allEntriesPromise = new Promise<void>((resolve, reject) => {
      parser.on("entry", (entry: any) => {
        if (entry.type === "File") {
          const readChunks: Uint8Array[] = [];
          entry.on("data", (chunk: Uint8Array) => readChunks.push(chunk));
          entry.on("end", () => {
            const combined = concatUint8Arrays(readChunks);
            entries.push({
              name: entry.path,
              content: new TextDecoder().decode(combined)
            });
            if (entries.length === 3) {
              resolve();
            }
          });
          entry.on("error", (err: Error) => reject(err));
        } else {
          entry.autodrain();
        }
      });

      parser.on("error", (err: Error) => reject(err));
    });

    parser.write(fullZip);
    parser.end();
    await allEntriesPromise;

    expect(entries.length).toBe(3);
    expect(entries.find(e => e.name === "file1.txt")?.content).toBe("File 1 content");
    expect(entries.find(e => e.name === "file2.txt")?.content).toBe("File 2 content");
    expect(entries.find(e => e.name === "dir/file3.txt")?.content).toBe("File 3 content");
  }, 30000);
});
