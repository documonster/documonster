import { hasDeflateRawDecompressionStream } from "@archive/compression/compress.base";
import { zip } from "@archive/create-archive";
import { createParse, type ZipEntry } from "@archive/unzip/stream.browser";
import { describe, it, expect } from "vitest";

describe("parse.browser - worker inflate (optional)", () => {
  function createTestZip(files: Array<{ name: string; content: string }>): Uint8Array {
    const encoder = new TextEncoder();
    const z = zip({ level: 6 });
    for (const f of files) {
      z.add(f.name, encoder.encode(f.content));
    }
    return z.bytesSync();
  }

  it("should parse entries when worker inflate is enabled (with fallback)", async () => {
    const testContent = "Hello from worker";
    const zipData = createTestZip([{ name: "test.txt", content: testContent }]);

    const parse = createParse({ useWorkerInflate: true });
    const results: Array<{ path: string; content: string }> = [];
    const bufferPromises: Promise<void>[] = [];

    await new Promise<void>((resolve, reject) => {
      parse.on("entry", (entry: ZipEntry) => {
        bufferPromises.push(
          entry.buffer().then(buf => {
            results.push({ path: entry.path, content: new TextDecoder().decode(buf) });
          })
        );
      });

      parse.on("close", async () => {
        try {
          await Promise.all(bufferPromises);
          expect(results.length).toBe(1);
          expect(results[0]!.path).toBe("test.txt");
          expect(results[0]!.content).toBe(testContent);
          resolve();
        } catch (e) {
          reject(e);
        }
      });

      parse.on("error", reject);
      parse.end(zipData);
    });
  }, 20_000);

  it("should parse entries with an injected workerInflateUrl", async () => {
    if (typeof Worker === "undefined" || !hasDeflateRawDecompressionStream()) {
      return;
    }

    const workerCode = `
let ds = new DecompressionStream('deflate-raw');
let writer = ds.writable.getWriter();
let reader = ds.readable.getReader();

async function pump() {
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        postMessage({ t: 'end' });
        return;
      }
      postMessage({ t: 'data', chunk: value }, [value.buffer]);
    }
  } catch (e) {
    const m = e && e.message ? e.message : String(e);
    postMessage({ t: 'error', message: m });
  }
}

onmessage = async (ev) => {
  const msg = ev.data;
  if (!msg || typeof msg.t !== 'string') return;

  if (msg.t === 'write') {
    try {
      await writer.write(msg.chunk);
      postMessage({ t: 'ack', id: msg.id });
    } catch (e) {
      const m = e && e.message ? e.message : String(e);
      postMessage({ t: 'error', message: m, id: msg.id });
    }
    return;
  }

  if (msg.t === 'close') {
    try { await writer.close(); } catch (_) {}
    pump();
    return;
  }

  if (msg.t === 'abort') {
    try { await writer.abort(); } catch (_) {}
    postMessage({ t: 'aborted' });
    return;
  }
};
`;

    const url = URL.createObjectURL(new Blob([workerCode], { type: "text/javascript" }));

    try {
      const testContent = "Hello injected worker";
      const zipData = createTestZip([{ name: "test.txt", content: testContent }]);

      const parse = createParse({ useWorkerInflate: true, workerInflateUrl: url });
      const results: Array<{ path: string; content: string }> = [];
      const bufferPromises: Promise<void>[] = [];

      await new Promise<void>((resolve, reject) => {
        parse.on("entry", (entry: ZipEntry) => {
          bufferPromises.push(
            entry.buffer().then(buf => {
              results.push({ path: entry.path, content: new TextDecoder().decode(buf) });
            })
          );
        });

        parse.on("close", async () => {
          try {
            await Promise.all(bufferPromises);
            expect(results.length).toBe(1);
            expect(results[0]!.path).toBe("test.txt");
            expect(results[0]!.content).toBe(testContent);
            resolve();
          } catch (e) {
            reject(e);
          }
        });

        parse.on("error", reject);
        parse.end(zipData);
      });
    } finally {
      URL.revokeObjectURL(url);
    }
  }, 20_000);
});
