/**
 * In a browser, damaged compressed data meets `DecompressionStream` or the pure-JS inflater rather than zlib, and is
 * still an `ArchiveDecompressionError` — the one class a document reader takes for "the bytes are damaged".
 */
import {
  Archive,
  ArchiveDecompressionError,
  ArchiveError,
  Compression
} from "@archive/index.browser";
import { afterEach, describe, expect, it } from "vitest";

const BAD = new Uint8Array([0xff, 0xff, 0xff, 0xff, 1, 2, 3, 4, 5, 6, 7, 8]);

describe("ArchiveDecompressionError in a browser", () => {
  it("is what the one-shot decompressors fail with", async () => {
    await expect(Compression.inflate(BAD)).rejects.toBeInstanceOf(ArchiveDecompressionError);
    await expect(Compression.gunzip(BAD)).rejects.toBeInstanceOf(ArchiveDecompressionError);
    await expect(Compression.unzlib(BAD)).rejects.toBeInstanceOf(ArchiveDecompressionError);
  });

  // Data that asks for what no inflater here implements — a zlib preset dictionary (FDICT set in a valid header), a gzip
  // method other than DEFLATE — is reported as data that does not decompress, by the native codec and the JS one alike.
  it("is what data asking for an unimplemented feature fails with", async () => {
    const presetDictionary = new Uint8Array([0x78, 0xbb, 0, 0, 0, 1, 0x4b, 0x04, 0, 0, 0, 0, 0, 0]);
    const otherMethod = (await Compression.gzip(new TextEncoder().encode("floor"))).slice();
    otherMethod[2] = 7;
    await expect(Compression.unzlib(presetDictionary)).rejects.toBeInstanceOf(
      ArchiveDecompressionError
    );
    await expect(Compression.gunzip(otherMethod)).rejects.toBeInstanceOf(ArchiveDecompressionError);
    for (const [make, data] of [
      [Compression.createUnzlibStream, presetDictionary],
      [Compression.createGunzipStream, otherMethod]
    ] as const) {
      const stream = make() as unknown as {
        on(event: string, listener: (value?: unknown) => void): void;
        write(chunk: Uint8Array): void;
        end(): void;
      };
      const error = await new Promise<unknown>(resolve => {
        stream.on("error", resolve);
        stream.on("data", () => {});
        stream.on("end", () => resolve(undefined));
        stream.write(data);
        stream.end();
      });
      expect(error).toBeInstanceOf(ArchiveDecompressionError);
    }
  });

  // The class says the bytes are damaged, so an error a listener throws — the caller's own — must keep its identity.
  it("is not what a listener that throws fails the stream with", async () => {
    const input = new TextEncoder().encode("valid input");
    const cases: [string, () => unknown, Uint8Array][] = [
      ["createInflateStream", Compression.createInflateStream, await Compression.deflate(input)],
      ["createGunzipStream", Compression.createGunzipStream, await Compression.gzip(input)],
      ["createUnzlibStream", Compression.createUnzlibStream, await Compression.zlib(input)]
    ];
    for (const [name, make, data] of cases) {
      for (const event of ["data", "end"]) {
        const stream = make() as {
          on(event: string, listener: (value?: unknown) => void): void;
          write(chunk: Uint8Array): void;
          end(): void;
        };
        const thrown = new Error(`${event} listener failed`);
        const error = await new Promise<unknown>(resolve => {
          stream.on("error", resolve);
          if (event === "end") {
            stream.on("data", () => {});
          }
          stream.on(event, () => {
            throw thrown;
          });
          stream.write(data);
          stream.end();
        });
        expect(error, `${name}, ${event}`).toBe(thrown);
      }
    }
  });

  it("is what the streaming decompressors fail with", async () => {
    for (const make of [
      Compression.createInflateStream,
      Compression.createGunzipStream,
      Compression.createUnzlibStream
    ]) {
      // The browser codec is an event emitter with a stream's methods; its declared type names the writable side only.
      const stream = make() as unknown as {
        on(event: string, listener: (value?: unknown) => void): void;
        write(chunk: Uint8Array): void;
        end(): void;
      };
      const error = await new Promise<unknown>(resolve => {
        stream.on("error", resolve);
        stream.on("data", () => {});
        stream.on("end", () => resolve(undefined));
        stream.write(BAD);
        stream.end();
      });
      expect(error).toBeInstanceOf(ArchiveDecompressionError);
    }
  });

  describe.skipIf(!Compression.hasWorkerSupport())("in a worker", () => {
    afterEach(() => Compression.Workers.terminateDefaultPool());

    // An Error loses its class through postMessage; the worker says the data was refused and the
    // pool rebuilds the class, so damage found on another thread reads as damage found on this one.
    it("is what the pool's inflate fails with", async () => {
      await expect(Compression.Workers.inflate(BAD)).rejects.toBeInstanceOf(
        ArchiveDecompressionError
      );
    });

    // With allowTransfer the bytes are gone from this thread, so the worker's verdict cannot be checked here
    // again — it is still the verdict on the data, not a failure of the worker.
    it("is what inflate fails with when allowTransfer gave the worker the input", async () => {
      const bad = BAD.slice();
      await expect(
        Compression.inflate(bad, { useWorker: true, allowTransfer: true })
      ).rejects.toBeInstanceOf(ArchiveDecompressionError);
      expect(bad.byteLength).toBe(0);
    });

    it("is what a streaming decompressor on a worker fails with", async () => {
      const stream = Compression.createInflateStream({ useWorker: true }) as unknown as {
        on(event: string, listener: (value?: unknown) => void): void;
        write(chunk: Uint8Array): void;
        end(): void;
      };
      const error = await new Promise<unknown>(resolve => {
        stream.on("error", resolve);
        stream.on("data", () => {});
        stream.on("end", () => resolve(undefined));
        stream.write(BAD);
        stream.end();
      });
      expect(error).toBeInstanceOf(ArchiveDecompressionError);
    });

    it("leaves a failure of the pool itself a plain ArchiveError", async () => {
      const pool = new Compression.Workers.Pool();
      pool.terminate();
      const error = await pool.execute("inflate", BAD).catch((e: unknown) => e);
      expect(error).toBeInstanceOf(ArchiveError);
      expect(error).not.toBeInstanceOf(ArchiveDecompressionError);
    });
  });

  it("is what a ZIP entry whose compressed data is broken fails with", async () => {
    const text = new TextEncoder().encode("hello ".repeat(200));
    const zip = await Archive.packZip([["a.txt", text]], { level: 9 });
    const nameLength = zip[26]! | (zip[27]! << 8);
    const extraLength = zip[28]! | (zip[29]! << 8);
    const broken = zip.slice();
    broken[30 + nameLength + extraLength + 1]! ^= 0xff;
    await expect(Archive.unpackZip(broken)).rejects.toBeInstanceOf(ArchiveDecompressionError);
  });
});
