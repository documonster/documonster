/**
 * A `Uint8Array` made in another realm — a worker, an iframe, a test DOM such as jsdom, whose
 * `TextEncoder` returns the host's arrays — is bytes all the same. The argument checks used to
 * test `instanceof Uint8Array`, so every one of these refused such bytes ("data must be a
 * Uint8Array, got a Uint8Array") although the operation itself handles them.
 */
import { runInNewContext } from "node:vm";

import { Archive, Compression, Fs } from "@archive";
import { describe, expect, it } from "vitest";

const local = new TextEncoder().encode("documonster bytes from another realm");
const foreign = (): Uint8Array =>
  runInNewContext(`new Uint8Array([${[...local].join(",")}])`) as Uint8Array;

describe("bytes from another realm", () => {
  it("is not an instance of this realm's Uint8Array", () => {
    expect(foreign() instanceof Uint8Array).toBe(false);
  });

  it.each([
    ["Compression.deflateSync", (data: Uint8Array) => Compression.deflateSync(data)],
    ["Compression.gzipSync", (data: Uint8Array) => Compression.gzipSync(data)],
    ["Compression.zlibSync", (data: Uint8Array) => Compression.zlibSync(data)],
    ["Compression.isGzipData", (data: Uint8Array) => Compression.isGzipData(data)],
    ["Archive.crc32", (data: Uint8Array) => Archive.crc32(data)]
  ])("%s gives what it gives for local bytes", (_, run) => {
    expect(run(foreign())).toEqual(run(local));
  });

  it("Compression.deflate gives what it gives for local bytes", async () => {
    expect(await Compression.deflate(foreign())).toEqual(await Compression.deflate(local));
  });

  it("Archive.Crypto.createZipCryptoKeys takes a password from another realm", () => {
    expect(Archive.Crypto.createZipCryptoKeys(foreign())).toEqual(
      Archive.Crypto.createZipCryptoKeys(local)
    );
  });

  it("Fs.writeFileBytesSync writes bytes from another realm", () => {
    const dir = Fs.createTempDirSync("documonster-realm-");
    try {
      const path = `${dir}/bytes.bin`;
      Fs.writeFileBytesSync(path, foreign());
      expect(Fs.readFileBytesSync(path)).toEqual(local);
    } finally {
      Fs.removeSync(dir);
    }
  });

  it("Archive.Crypto.encryptZipCrypto takes random bytes from another realm", () => {
    const random = (): Uint8Array => new Uint8Array(11).fill(7);
    const crc = Archive.crc32(local);
    expect(
      Archive.Crypto.encryptZipCrypto(local, "pw", crc, () =>
        runInNewContext(`new Uint8Array(11).fill(7)`)
      )
    ).toEqual(Archive.Crypto.encryptZipCrypto(local, "pw", crc, random));
  });
});
