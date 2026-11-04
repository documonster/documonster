/**
 * The low-level `WordCrypto.Agile` / `WordCrypto.Cfb` primitives check their
 * arguments as `WordCrypto.encrypt` / `decrypt` do. Before, `Agile.deriveKey`
 * derived a key for an unknown hash (silently SHA-512), zero rounds or zero
 * bits, `Cfb.read` called a string "too small", and `Agile.decryptPackage`
 * took its size cap positionally while `WordCrypto.decrypt` takes it in
 * `DecryptOptions`.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError, DocxDecryptionError, WordCrypto } from "../crypto";
import { Document, Io } from "../index";

async function encrypted(): Promise<{ plain: Uint8Array; enc: Uint8Array }> {
  const d = Document.create();
  Document.addParagraph(d, "secret");
  const plain = await Io.toBuffer(d);
  return { plain, enc: await WordCrypto.encrypt(plain, "pw", { spinCount: 1 }) };
}

function streams(enc: Uint8Array) {
  const entries = WordCrypto.Cfb.read(enc);
  const infoBytes = entries.find(e => e.name === "EncryptionInfo")!.data;
  const info = WordCrypto.Agile.parseEncryptionInfo(
    new TextDecoder().decode(infoBytes.subarray(8))
  );
  const pkg = entries.find(e => e.name === "EncryptedPackage")!.data;
  return { info, pkg };
}

const goodInfo = {
  keySalt: new Uint8Array(16),
  spinCount: 1,
  hashAlgorithm: "SHA512",
  keyBits: 256
};

describe("WordCrypto.Agile.deriveKey", () => {
  it("derives a key from valid arguments", async () => {
    const key = await WordCrypto.Agile.deriveKey("pw", goodInfo, new Uint8Array(8));
    expect(key.length).toBe(32);
  });

  it("refuses what would derive a wrong key", async () => {
    const block = new Uint8Array(8);
    const bad: [string, () => Promise<unknown>][] = [
      ["password", () => WordCrypto.Agile.deriveKey(42 as never, goodInfo, block)],
      ["info", () => WordCrypto.Agile.deriveKey("pw", 42 as never, block)],
      [
        "info.keySalt",
        () => WordCrypto.Agile.deriveKey("pw", { ...goodInfo, keySalt: "x" as never }, block)
      ],
      [
        "info.spinCount",
        () => WordCrypto.Agile.deriveKey("pw", { ...goodInfo, spinCount: Number.NaN }, block)
      ],
      [
        "info.spinCount",
        () => WordCrypto.Agile.deriveKey("pw", { ...goodInfo, spinCount: -1 }, block)
      ],
      [
        "info.hashAlgorithm",
        () => WordCrypto.Agile.deriveKey("pw", { ...goodInfo, hashAlgorithm: "bogus" }, block)
      ],
      ["info.keyBits", () => WordCrypto.Agile.deriveKey("pw", { ...goodInfo, keyBits: 7 }, block)],
      ["blockKey", () => WordCrypto.Agile.deriveKey("pw", goodInfo, "x" as never)]
    ];
    for (const [field, call] of bad) {
      const error = await call().catch((e: unknown) => e);
      expect(error, field).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, field).toBe(field);
      expect((error as Error).message, field).toMatch(/^WordCrypto\.Agile\.deriveKey: /);
    }
  });
});

describe("WordCrypto.Agile.decryptPackage", () => {
  it("takes the size cap in DecryptOptions, as WordCrypto.decrypt does", async () => {
    const { plain, enc } = await encrypted();
    const { info, pkg } = streams(enc);
    expect(await WordCrypto.Agile.decryptPackage(pkg, info, "pw")).toEqual(plain);
    await expect(
      WordCrypto.Agile.decryptPackage(pkg, info, "pw", { maxDecryptedSize: 10 })
    ).rejects.toBeInstanceOf(DocxDecryptionError);
    await expect(
      WordCrypto.Agile.decryptPackage(pkg, info, "pw", { maxDecryptedSize: Number.NaN })
    ).rejects.toBeInstanceOf(DocxArgumentError);
    await expect(
      WordCrypto.Agile.decryptPackage(pkg, info, "pw", 42 as never)
    ).rejects.toBeInstanceOf(DocxArgumentError);
  });

  it("checks the package bytes and the info", async () => {
    const { enc } = await encrypted();
    const { info } = streams(enc);
    await expect(WordCrypto.Agile.decryptPackage("x" as never, info, "pw")).rejects.toBeInstanceOf(
      DocxArgumentError
    );
    await expect(
      WordCrypto.Agile.decryptPackage(new Uint8Array(16), 42 as never, "pw")
    ).rejects.toBeInstanceOf(DocxArgumentError);
  });
});

describe("WordCrypto.Cfb", () => {
  it("read refuses non-bytes with DocxArgumentError", () => {
    for (const value of ["x", 42, null, {}]) {
      expect(() => WordCrypto.Cfb.read(value as never)).toThrow(DocxArgumentError);
    }
  });

  it("read accepts an ArrayBuffer", async () => {
    const { enc } = await encrypted();
    const ab = enc.buffer.slice(enc.byteOffset, enc.byteOffset + enc.byteLength) as ArrayBuffer;
    expect(WordCrypto.Cfb.read(ab).map(e => e.name)).toEqual(
      WordCrypto.Cfb.read(enc).map(e => e.name)
    );
  });

  it("write refuses entries that are not { name, data }", () => {
    const bad: unknown[] = [42, [42], [{ name: 1, data: new Uint8Array(1) }], [{ name: "a" }]];
    for (const value of bad) {
      expect(() => WordCrypto.Cfb.write(value as never)).toThrow(DocxArgumentError);
    }
    const written = WordCrypto.Cfb.write([{ name: "a", data: new Uint8Array([1, 2, 3]) }]);
    expect(WordCrypto.Cfb.read(written).find(e => e.name === "a")?.data).toEqual(
      new Uint8Array([1, 2, 3])
    );
  });
});
