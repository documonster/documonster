import * as entry from "@archive";
import { Archive, ArchiveError, ArchiveWorkerUnsupportedError, Compression, Fs } from "@archive";
import { compress, decompress, detectCompressionFormat } from "@archive/compression/compress";
import { zip } from "@archive/create-archive";
import { encodePng } from "@archive/png";
import { unzip } from "@archive/read-archive";
import {
  TarArchive,
  TarReader,
  createTarArchive,
  createTarReader,
  tar,
  tarSync
} from "@archive/tar/tar-archive";
import { TarGzArchive, targz, untargz } from "@archive/tar/tar-gzip";
import { untar } from "@archive/tar/tar-parser";
import { ZipArchive } from "@archive/zip/zip-archive";
import { glob, safeStats } from "@utils/fs";
import { describe, expect, it } from "vitest";

describe("Archive namespace", () => {
  it("keeps the format-dispatching zip/unzip as deprecated forwards", () => {
    expect(Archive.zip).toBe(zip);
    expect(Archive.unzip).toBe(unzip);
    expect(Archive.zip()).toBeInstanceOf(ZipArchive);
    expect(Archive.zip({ format: "tar" })).toBeInstanceOf(TarArchive);
    expect(Archive.unzip(new Uint8Array(), { format: "tar" })).toBeInstanceOf(TarReader);
  });

  it("keeps the renamed TAR members as deprecated forwards", () => {
    expect(Archive.createTarArchive).toBe(createTarArchive);
    expect(Archive.createTarReader).toBe(createTarReader);
    expect(Archive.tar).toBe(tar);
    expect(Archive.tarSync).toBe(tarSync);
    expect(Archive.untar).toBe(untar);
    expect(Archive.TarArchive).toBe(TarArchive);
    expect(Archive.TarReader).toBe(TarReader);
  });

  it("packTarGz takes the same entries as packTar, tuples included", async () => {
    const entries: [string, string][] = [["a.txt", "hello"]];
    const tgz = await Archive.packTarGz(entries);
    const files = await Archive.unpackTarGz(tgz);
    expect(new TextDecoder().decode(files.get("a.txt")!.data)).toBe("hello");
    const plain = await Archive.unpackTar(await Archive.packTar(entries));
    expect([...plain.keys()]).toEqual([...files.keys()]);
  });

  it("publishes crc32 and encodePng, keeping the flat forwards", () => {
    expect(Archive.crc32(new TextEncoder().encode("abc"))).toBe(0x352441c2);
    expect(Archive.encodePng).toBe(encodePng);
    expect(entry.encodePng).toBe(encodePng);
    // `Archive.crc32` checks its argument; the deprecated flat form keeps the old function.
    expect(entry.crc32(new TextEncoder().encode("abc"))).toBe(0x352441c2);
  });

  it("keeps the Node-only flat forwards", () => {
    expect(entry.targz).toBe(targz);
    expect(entry.untargz).toBe(untargz);
    expect(entry.TarGzArchive).toBe(TarGzArchive);
    expect(entry.ArchiveFile).toBe(Archive.ArchiveFile);
    expect(entry.toNodeReadable).toBe(Archive.toNodeReadable);
  });
});

describe("Compression namespace", () => {
  it("names raw DEFLATE deflate/inflate, like its streams and pool", async () => {
    // The namespace members check their arguments (see compression-surface-checks.test.ts),
    // so they wrap the codecs rather than being them; the output is the codecs'.
    const data = new TextEncoder().encode("hello hello hello");
    expect(await Compression.deflate(data)).toEqual(await compress(data));
    expect(await Compression.inflate(await compress(data))).toEqual(
      await decompress(await compress(data))
    );
    expect(Compression.detectFormat(data)).toBe(detectCompressionFormat(data));
    expect(await Compression.inflate(await Compression.deflate(data))).toEqual(data);
  });

  it("keeps the flat forwards", async () => {
    const data = new TextEncoder().encode("hello hello hello");
    expect(entry.compress).toBe(compress);
    expect(entry.decompress).toBe(decompress);
    expect(await entry.gzip(data)).toEqual(await Compression.gzip(data));
    expect(typeof entry.createDeflateStream).toBe("function");
    expect(entry.WorkerPool).toBe(Compression.Workers.WorkerPool);
    expect(entry.deflateWithPool).toBe(Compression.Workers.deflate);
    expect(entry.inflateBatchWithPool).toBe(Compression.Workers.inflateBatch);
    expect(entry.getDefaultWorkerPool).toBe(Compression.Workers.getDefaultPool);
  });

  it("Workers.WorkerPool cannot be constructed in Node", () => {
    expect(Compression.hasWorkerSupport()).toBe(false);
    expect(() => new Compression.Workers.WorkerPool()).toThrow(ArchiveWorkerUnsupportedError);
    expect(() => Compression.Workers.terminateDefaultPool()).not.toThrow();
  });

  it("Workers members have the browser signatures on Node, and reject", async () => {
    // Type-checked: code guarded by hasWorkerSupport() compiles against the Node types.
    const data = new Uint8Array([1, 2, 3]);
    const rejected = await Compression.Workers.deflate(data, { level: 6 }).catch((e: unknown) => e);
    expect(rejected).toBeInstanceOf(ArchiveWorkerUnsupportedError);
    expect(rejected).toBeInstanceOf(ArchiveError);
    expect((rejected as ArchiveWorkerUnsupportedError).operation).toBe(
      "Compression.Workers.deflate"
    );
    expect((rejected as Error).message).toMatch(
      /^Compression\.Workers\.deflate: .*not available in Node/
    );
    await expect(Compression.Workers.inflateBatch([{ data }])).rejects.toThrow(
      ArchiveWorkerUnsupportedError
    );
    expect(() => Compression.Workers.getDefaultPool({ maxWorkers: 2 })).toThrow(
      ArchiveWorkerUnsupportedError
    );
  });
});

describe("Archive.Crypto", () => {
  const password = "secret";
  const plain = new TextEncoder().encode("hello crypto");
  const crc = Archive.crc32(plain);

  it("keeps the flat Archive members as deprecated forwards", () => {
    expect(Archive.aesEncrypt).toBe(Archive.Crypto.encryptAes);
    expect(Archive.AES_KEY_LENGTH).toBe(Archive.Crypto.AES_KEY_LENGTH);
    expect("randomBytes" in Archive.Crypto).toBe(false);
    expect(Archive.randomBytes(16)).toHaveLength(16);
  });

  it("zipCryptoDecrypt throws on a wrong password; the deprecated Archive form returns null", () => {
    const encrypted = Archive.Crypto.encryptZipCrypto(plain, password, crc, n =>
      crypto.getRandomValues(new Uint8Array(n))
    );
    expect(Archive.Crypto.decryptZipCrypto(encrypted, password, crc)).toEqual(plain);
    expect(() => Archive.Crypto.decryptZipCrypto(encrypted, "wrong", crc)).toThrow(ArchiveError);
    expect(() => Archive.Crypto.decryptZipCrypto(new Uint8Array(3), password, crc)).toThrow(
      ArchiveError
    );
    expect(Archive.zipCryptoDecrypt(encrypted, password, crc)).toEqual(plain);
    expect(Archive.zipCryptoDecrypt(encrypted, "wrong", crc)).toBeNull();
  });
});

describe("Archive reader lookups", () => {
  const files = new Map([["a.txt", "alpha"]]);

  it("ZipReader.get/bytes return null for a missing entry", async () => {
    const zip = await Archive.packZip(files);
    const reader = Archive.openZip(zip);
    expect(new TextDecoder().decode((await reader.bytes("a.txt"))!)).toBe("alpha");
    expect((await reader.get("a.txt"))?.path).toBe("a.txt");
    expect(await reader.get("missing")).toBeNull();
    expect(await reader.bytes("missing")).toBeNull();
  });

  it("TarReader.get/bytes return null for a missing entry", async () => {
    const reader = Archive.openTar(await Archive.packTar(files));
    expect(new TextDecoder().decode((await reader.bytes("a.txt"))!)).toBe("alpha");
    expect(await reader.get("missing")).toBeNull();
    expect(await reader.bytes("missing")).toBeNull();
  });

  it("RemoteZipReader.fromReader reads through any RandomAccessReader; extract returns null when missing", async () => {
    const zip = await Archive.packZip(files);
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip));
    expect(new TextDecoder().decode((await remote.extract("a.txt"))!)).toBe("alpha");
    expect(await remote.extract("missing")).toBeNull();
    await remote.close();
  });

  it("HttpRangeReader.open reads over Range requests", async () => {
    const body = new Uint8Array([1, 2, 3, 4]);
    const fetch = (async (_url: string, init?: RequestInit) => {
      const range = new Headers(init?.headers).get("Range");
      if (init?.method === "HEAD") {
        return new Response(null, {
          headers: { "Content-Length": "4", "Accept-Ranges": "bytes" }
        });
      }
      const [, start, end] = /bytes=(\d+)-(\d+)/.exec(range ?? "")!;
      return new Response(body.slice(Number(start), Number(end) + 1), { status: 206 });
    }) as typeof globalThis.fetch;
    const reader = await Archive.HttpRangeReader.open("https://example.test/a.zip", {
      fetch,
      validateRangeSupport: false
    });
    expect(reader).toBeInstanceOf(Archive.HttpRangeReader);
    expect(reader.size).toBe(4);
    expect(await reader.read(1, 3)).toEqual(new Uint8Array([2, 3]));
  });
});

describe("Archive builders", () => {
  it("toBuffer/toBufferSync build the archive; bytes/bytesSync stay as deprecated aliases", async () => {
    const zip = Archive.createZip();
    zip.add("a.txt", "alpha");
    const built = zip.toBufferSync();
    expect([...Archive.unpackZipSync(built).keys()]).toEqual(["a.txt"]);

    const tar = Archive.createTar();
    tar.add("a.txt", "alpha");
    expect(await tar.toBuffer()).toEqual(tar.bytesSync());

    const tgz = Archive.createTarGz();
    tgz.add("a.txt", "alpha");
    const files = await Archive.unpackTarGz(tgz.toBufferSync());
    expect(new TextDecoder().decode(files.get("a.txt")!.data)).toBe("alpha");
  });

  it("ZipEditPlan.toJson/fromJson round-trip; the JSON-cased names forward", async () => {
    const plan = new Archive.ZipEditPlan().set("a.txt", new TextEncoder().encode("x"));
    const json = await plan.toJson();
    expect(await Archive.ZipEditPlan.fromJson(json).toJson()).toBe(json);
    expect(await Archive.ZipEditPlan.fromJSON(json).toJSON()).toBe(json);
  });
});

describe("Fs namespace", () => {
  it("holds the file-system helpers; the flat forwards keep the unchecked v0.21.1 functions", () => {
    expect(typeof Fs.glob).toBe("function");
    expect(entry.glob).toBe(glob);
    expect(entry.safeStats).toBe(safeStats);
    // Fs checks its arguments first, so it wraps rather than re-exports them.
    expect(Fs.glob).not.toBe(glob);
  });

  it("safeStats reports a missing path as null", async () => {
    const missing = "/definitely/not/here/documonster";
    expect(await Fs.safeStats(missing)).toBeNull();
    expect(Fs.safeStatsSync(missing)).toBeNull();
    expect(await entry.safeStats(missing)).toBeNull();
    expect((await Fs.safeStats(import.meta.dirname))?.isDirectory()).toBe(true);
  });
});
