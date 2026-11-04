import { Archive, ArchiveError, Compression } from "@archive";
import { parseTarGzStream as deprecatedParseTarGzStream } from "@archive/tar/tar-gzip";
import { describe, expect, it } from "vitest";

const decoder = new TextDecoder();

/** Incompressible bytes, so the compressed archive is as large as the payload. */
function randomBytes(length: number): Uint8Array {
  const out = new Uint8Array(length);
  for (let i = 0; i < length; i += 65536) {
    crypto.getRandomValues(out.subarray(i, i + 65536));
  }
  return out;
}

describe("tar.gz reading is incremental", () => {
  it("parseTarGzStream yields the first entry before the compressed source is exhausted", async () => {
    const tgz = await Archive.packTarGz([
      ["first.txt", "first"],
      ["big.bin", randomBytes(256 * 1024)]
    ]);
    let sourceDone = false;
    async function* source(): AsyncIterable<Uint8Array> {
      for (let i = 0; i < tgz.length; i += 1024) {
        yield tgz.subarray(i, i + 1024);
      }
      sourceDone = true;
    }

    const seen: string[] = [];
    let doneAtFirst: boolean | undefined;
    for await (const entry of Archive.parseTarGzStream(source())) {
      doneAtFirst ??= sourceDone;
      seen.push(entry.info.path);
      await entry.data();
    }
    expect(seen).toEqual(["first.txt", "big.bin"]);
    expect(doneAtFirst).toBe(false);
  });

  it("openTarGz supports getEntry / list and repeated iteration", async () => {
    const tgz = await Archive.packTarGz([
      ["a.txt", "alpha"],
      ["b.txt", "beta"]
    ]);
    const reader = Archive.openTarGz(tgz);
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
    expect(decoder.decode((await reader.bytes("b.txt"))!)).toBe("beta");
    expect(await reader.get("missing")).toBeNull();

    const names: string[] = [];
    for (let pass = 0; pass < 2; pass++) {
      for await (const entry of reader.entries()) {
        names.push(entry.path);
      }
    }
    expect(names).toEqual(["a.txt", "b.txt", "a.txt", "b.txt"]);
  });

  it("TarReader over a buffer does not duplicate entries after repeated iteration", async () => {
    const reader = Archive.openTar(await Archive.packTar([["a.txt", "alpha"]]));
    for (let pass = 0; pass < 2; pass++) {
      for await (const _ of reader.entries()) {
        // drain
      }
    }
    expect(await reader.list()).toEqual(["a.txt"]);
  });

  it("TarReader over a stream supports list()", async () => {
    const tar = await Archive.packTar([["a.txt", "alpha"]]);
    const reader = Archive.openTar(
      (async function* () {
        yield tar;
      })()
    );
    expect(await reader.list()).toEqual(["a.txt"]);
  });
});

/** A one-shot `ReadableStream` over `bytes`, like a `fetch` body. */
function oneShot(bytes: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      for (let i = 0; i < bytes.length; i += 97) {
        controller.enqueue(bytes.slice(i, i + 97));
      }
      controller.close();
    }
  });
}

async function names(reader: ReturnType<typeof Archive.openTarGz>): Promise<string[]> {
  const out: string[] = [];
  for await (const entry of reader.entries()) {
    out.push(entry.path);
  }
  return out;
}

describe("openTarGz over a one-shot source", () => {
  const files: Array<[string, string]> = [
    ["a.txt", "alpha"],
    ["b.txt", "beta"]
  ];

  it("list → entries → get", async () => {
    const reader = Archive.openTarGz(oneShot(await Archive.packTarGz(files)));
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
    expect(await names(reader)).toEqual(["a.txt", "b.txt"]);
    expect(await (await reader.get("b.txt"))!.text()).toBe("beta");
  });

  it("entries → entries → list", async () => {
    const reader = Archive.openTarGz(oneShot(await Archive.packTarGz(files)));
    expect(await names(reader)).toEqual(["a.txt", "b.txt"]);
    expect(await names(reader)).toEqual(["a.txt", "b.txt"]);
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
  });

  it("get → entries after an early break", async () => {
    const reader = Archive.openTarGz(oneShot(await Archive.packTarGz(files)));
    for await (const _ of reader.entries()) {
      break;
    }
    expect(decoder.decode((await reader.bytes("a.txt"))!)).toBe("alpha");
    expect(await names(reader)).toEqual(["a.txt", "b.txt"]);
  });
});

/** A tar.gz whose TAR is cut off inside the last entry's data. */
async function truncatedTgz(): Promise<Uint8Array> {
  const tar = await Archive.packTar([
    ["a.txt", "alpha"],
    ["b.txt", "b".repeat(1000)]
  ]);
  return Compression.gzip(tar.subarray(0, 512 * 3 + 100));
}

describe("a truncated tar.gz", () => {
  // These pinned the old result — the cut entry returned as far as it went, data shorter
  // than its header says — which is silent data loss. Every reader now refuses it.
  it("is refused by every reader, buffered or incremental", async () => {
    const tgz = await truncatedTgz();
    const drain = async (entries: AsyncIterable<unknown>) => {
      for await (const _ of entries) {
        // drain
      }
    };
    await expect(Archive.parseTarGz(tgz)).rejects.toThrow(/Unexpected end/);
    await expect(drain(Archive.parseTarGzStream(tgz))).rejects.toThrow(/Unexpected end/);
    await expect(drain(deprecatedParseTarGzStream(tgz))).rejects.toThrow(/Unexpected end/);
    await expect(Archive.openTarGz(oneShot(tgz)).list()).rejects.toThrow(/Unexpected end/);
    await expect(Archive.openTar(await Compression.gunzip(tgz)).list()).rejects.toThrow(
      /Unexpected end/
    );
  });

  it("the deprecated flat parseTarGzStream fails on corrupt gzip before yielding any entry", async () => {
    const tgz = await Archive.packTarGz([
      ["a.txt", "alpha"],
      ["b.txt", "x".repeat(100000)]
    ]);
    const corrupt = tgz.slice(0, tgz.length - 20);
    const seen: string[] = [];
    await expect(
      (async () => {
        for await (const entry of deprecatedParseTarGzStream(corrupt)) {
          seen.push(entry.info.path);
        }
      })()
    ).rejects.toThrow();
    expect(seen).toEqual([]);
  });
});

describe("parseTarGzStream on a damaged gzip stream", () => {
  /** A deterministic payload big enough to span several deflate blocks. */
  const payload = new Uint8Array(70000).map((_, i) => (i * 2654435761) >>> 24);

  async function damaged(): Promise<Record<string, Uint8Array>> {
    const tgz = await Archive.packTarGz([
      ["a.txt", "hello"],
      ["b.bin", payload]
    ]);
    return {
      // Cut inside the deflate stream, after both entries' data
      tail: tgz.subarray(0, tgz.length - 20),
      // Cut inside the second entry
      middle: tgz.subarray(0, 200),
      // Only the gzip trailer is missing: the TAR ends cleanly
      trailer: tgz.subarray(0, tgz.length - 8),
      empty: new Uint8Array(0)
    };
  }

  const sources: Record<
    string,
    (bytes: Uint8Array) => Parameters<typeof Archive.parseTarGzStream>[0]
  > = {
    bytes: bytes => bytes,
    asyncIterable: bytes =>
      (async function* () {
        for (let i = 0; i < bytes.length; i += 1000) {
          yield bytes.subarray(i, i + 1000);
        }
      })(),
    readableStream: bytes => oneShot(bytes)
  };

  /** Give a late zlib `error` event the chance to surface as an unhandled error. */
  const settle = (): Promise<void> => new Promise(resolve => setTimeout(resolve, 20));

  for (const [sourceName, toSource] of Object.entries(sources)) {
    for (const damage of ["tail", "middle", "trailer", "empty"]) {
      it(`rejects with an ArchiveError (${damage}, ${sourceName})`, async () => {
        const bytes = (await damaged())[damage]!;
        const iteration = (async () => {
          for await (const entry of Archive.parseTarGzStream(toSource(bytes))) {
            await entry.data();
          }
        })();
        await expect(iteration).rejects.toBeInstanceOf(ArchiveError);
        await settle();
      });
    }
  }

  it("the buffered readers report a damaged gzip stream as an ArchiveError too", async () => {
    const { tail, empty } = await damaged();
    for (const bytes of [tail, empty]) {
      await expect(Archive.unpackTarGz(bytes)).rejects.toBeInstanceOf(ArchiveError);
      await expect(Archive.parseTarGz(oneShot(bytes))).rejects.toBeInstanceOf(ArchiveError);
      await expect(Archive.openTarGz(bytes).list()).rejects.toBeInstanceOf(ArchiveError);
      expect(() => Archive.unpackTarGzSync(bytes)).toThrow(ArchiveError);
    }
  });

  it("does not raise an unhandled error when the consumer stops early", async () => {
    const { tail } = await damaged();
    for await (const _ of Archive.parseTarGzStream(tail)) {
      break;
    }
    await settle();
  });

  it("closes the source when the consumer stops early", async () => {
    const tgz = await Archive.packTarGz([
      ["a.txt", "alpha"],
      ["b.bin", randomBytes(64 * 1024)]
    ]);
    let closed = false;
    async function* source(): AsyncIterable<Uint8Array> {
      try {
        for (let i = 0; i < tgz.length; i += 512) {
          yield tgz.subarray(i, i + 512);
        }
      } finally {
        closed = true;
      }
    }
    for await (const _ of Archive.parseTarGzStream(source())) {
      break;
    }
    expect(closed).toBe(true);
  });

  it("still reads an intact archive in full", async () => {
    const tgz = await Archive.packTarGz([
      ["a.txt", "hello"],
      ["b.bin", payload]
    ]);
    for (const toSource of Object.values(sources)) {
      const seen: Array<[string, number]> = [];
      for await (const entry of Archive.parseTarGzStream(toSource(tgz))) {
        seen.push([entry.info.path, (await entry.data()).length]);
      }
      expect(seen).toEqual([
        ["a.txt", 5],
        ["b.bin", payload.length]
      ]);
    }
  });
});

/** A one-shot stream that hands out one small chunk per macrotask, so a read stays pending. */
function slowStream(bytes: Uint8Array): ReadableStream<Uint8Array> {
  let offset = 0;
  return new ReadableStream({
    async pull(controller) {
      await new Promise(resolve => setTimeout(resolve, 1));
      if (offset >= bytes.length) {
        controller.close();
        return;
      }
      controller.enqueue(bytes.slice(offset, offset + 64));
      offset += 64;
    }
  });
}

/** A one-shot stream that fails after its first chunk. */
function failingStream(bytes: Uint8Array, error: Error): ReadableStream<Uint8Array> {
  let sent = false;
  return new ReadableStream({
    pull(controller) {
      if (sent) {
        controller.error(error);
        return;
      }
      sent = true;
      controller.enqueue(bytes.slice(0, 16));
    }
  });
}

describe("openTarGz shared decompression", () => {
  const files: Array<[string, string]> = [
    ["a.txt", "alpha"],
    ["b.txt", "beta"]
  ];

  it("one caller aborting does not fail another waiting on the same read", async () => {
    const reader = Archive.openTarGz(slowStream(await Archive.packTarGz(files)));
    const controller = new AbortController();
    const aborted = (async () => {
      for await (const _ of reader.entries({ signal: controller.signal })) {
        // drain
      }
    })();
    const listed = reader.list();
    controller.abort();
    await expect(aborted).rejects.toThrow(/abort/i);
    expect(await listed).toEqual(["a.txt", "b.txt"]);
  });

  it("a one-shot source that failed rethrows the original error on retry", async () => {
    const error = new Error("network down");
    const reader = Archive.openTarGz(failingStream(await Archive.packTarGz(files), error));
    await expect(reader.list()).rejects.toBe(error);
    await expect(reader.list()).rejects.toBe(error);
  });

  it("a one-shot source whose every caller aborted is cancelled for later calls", async () => {
    const reader = Archive.openTarGz(slowStream(await Archive.packTarGz(files)));
    const controller = new AbortController();
    const first = reader.entries({ signal: controller.signal })[Symbol.asyncIterator]().next();
    controller.abort();
    // Joined in the same tick: must not attach to the cancelled read
    const later = reader.list();
    await expect(first).rejects.toThrow(/abort/i);
    await expect(later).rejects.toThrow(/cancelled by an earlier caller/);
    await expect(reader.list()).rejects.toThrow(/cancelled by an earlier caller/);
  });

  it("aborting the only caller as decompression completes keeps the decompressed archive", async () => {
    const reader = Archive.openTarGz(slowStream(await Archive.packTarGz(files)));
    // Start the shared load before any caller, so this callback runs ahead of the caller's
    const internal = reader as unknown as {
      _load: { result: Promise<Uint8Array> } | undefined;
      _startLoad(): { result: Promise<Uint8Array> };
    };
    internal._load = internal._startLoad();
    const controller = new AbortController();
    void internal._load.result.then(() => controller.abort());
    const first = reader.entries({ signal: controller.signal })[Symbol.asyncIterator]().next();
    await expect(first).rejects.toThrow(/abort/i);
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
  });

  it("an in-memory source is not cancelled when its only caller aborts", async () => {
    const reader = Archive.openTarGz(await Archive.packTarGz(files));
    const controller = new AbortController();
    const first = reader.entries({ signal: controller.signal })[Symbol.asyncIterator]().next();
    controller.abort();
    const later = reader.list();
    await expect(first).rejects.toThrow(/abort/i);
    expect(await later).toEqual(["a.txt", "b.txt"]);
  });

  it("close() cancels a one-shot read once its last caller aborts", async () => {
    const bytes = await Archive.packTarGz(files);
    let cancelled = false;
    let offset = 0;
    const source = new ReadableStream<Uint8Array>({
      async pull(c) {
        await new Promise(r => setTimeout(r, 5));
        if (offset >= bytes.length) {
          c.close();
          return;
        }
        c.enqueue(bytes.slice(offset, offset + 16));
        offset += 16;
      },
      cancel() {
        cancelled = true;
      }
    });
    const reader = Archive.openTarGz(source);
    const controller = new AbortController();
    const first = reader.entries({ signal: controller.signal })[Symbol.asyncIterator]().next();
    await reader.close();
    controller.abort();
    await expect(first).rejects.toThrow(/abort/i);
    await new Promise(r => setTimeout(r, 20));
    expect(cancelled).toBe(true);
  });

  it("close() during a one-shot read lets the waiting caller finish", async () => {
    const reader = Archive.openTarGz(slowStream(await Archive.packTarGz(files)));
    const listed = reader.list();
    await reader.close();
    expect(await listed).toEqual(["a.txt", "b.txt"]);
    // The decompressed copy is dropped on close, and a one-shot source cannot be read again
    await expect(reader.list()).rejects.toThrow(/closed/);
  });

  it("an in-memory source is retried after a caller aborted", async () => {
    const reader = Archive.openTarGz(await Archive.packTarGz(files));
    await expect(
      reader.entries({ signal: AbortSignal.abort() })[Symbol.asyncIterator]().next()
    ).rejects.toThrow(/abort/i);
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
  });

  it("close() releases the decompressed archive", async () => {
    const memory = Archive.openTarGz(await Archive.packTarGz(files));
    expect(await memory.list()).toEqual(["a.txt", "b.txt"]);
    await memory.close();
    // An in-memory source is decompressed again, like openTar re-parses after close()
    expect(await memory.list()).toEqual(["a.txt", "b.txt"]);

    const stream = Archive.openTarGz(oneShot(await Archive.packTarGz(files)));
    expect(await stream.list()).toEqual(["a.txt", "b.txt"]);
    await stream.close();
    // A one-shot source cannot be read again once its decompressed copy is dropped
    await expect(stream.list()).rejects.toThrow(/closed/);
  });
});
