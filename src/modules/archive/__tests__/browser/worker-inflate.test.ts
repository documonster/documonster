import { hasDeflateRawDecompressionStream } from "@archive/compression/compress.base";
import { zip } from "@archive/create-archive";
import type { ZipEntry } from "@archive/unzip/stream.browser";
import { createParse } from "@archive/unzip/stream.browser";
import { describe, it, expect } from "vitest";

describe("parse.browser - worker inflate (optional)", () => {
  function createTestZip(files: Array<{ name: string; content: string }>): Uint8Array {
    const encoder = new TextEncoder();
    const z = zip({ level: 6 });
    for (const f of files) {
      z.add(f.name, encoder.encode(f.content));
    }
    return z.toBufferSync();
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

  // A worker refused by a Content Security Policy, or a `workerInflateUrl` that does not boot,
  // reports it only after the first chunk was posted; the entry completes on this thread.
  it("completes entries on this thread when the inflate worker fails to load", async () => {
    if (typeof Worker === "undefined" || !hasDeflateRawDecompressionStream()) {
      return;
    }
    const url = URL.createObjectURL(
      new Blob(["throw new Error('boot failure');"], { type: "text/javascript" })
    );
    try {
      // Several entries, each large enough to arrive in more than one chunk
      const files = ["a", "b", "c"].map(name => ({
        name: `${name}.txt`,
        content: `${name} — worker load failure. `.repeat(20_000)
      }));
      const parse = createParse({ useWorkerInflate: true, workerInflateUrl: url });
      const results = new Map<string, string>();
      await new Promise<void>((resolve, reject) => {
        const pending: Promise<void>[] = [];
        parse.on("entry", (entry: ZipEntry) => {
          pending.push(
            entry.buffer().then(buf => {
              results.set(entry.path, new TextDecoder().decode(buf));
            })
          );
        });
        parse.on("close", () => Promise.all(pending).then(() => resolve(), reject));
        parse.on("error", reject);
        const zipData = createTestZip(files);
        // In pieces, so the parser writes to the inflater while the worker is still loading
        for (let i = 0; i < zipData.length; i += 16_384) {
          parse.write(zipData.subarray(i, i + 16_384));
        }
        parse.end();
      });
      expect(results).toEqual(new Map(files.map(f => [f.name, f.content])));
    } finally {
      URL.revokeObjectURL(url);
    }
  }, 20_000);
});
