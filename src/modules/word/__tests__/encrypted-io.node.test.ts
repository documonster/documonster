import { readFileSync } from "node:fs";
/**
 * Word reads and writes encrypted packages through one implementation shared with Excel: `Io.read` /
 * `Stream.DocumentReader` / `WordCrypto.decrypt` open what Word 16 saved with "Encrypt with Password"
 * (`data/office-encrypted/word-encrypted.docx`, password `Word-pw1`), and `Io.toBuffer` / `toBase64` / `writeFile`
 * encrypt with `{ password }`.
 */
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { WordCrypto } from "../crypto";
import {
  Document,
  DocxArgumentError,
  DocxDecryptionError,
  DocxEncryptedError,
  Io,
  Query,
  Stream
} from "../index";
import type { DocxDocument } from "../index";

const OFFICE = new Uint8Array(
  readFileSync(new URL("data/office-encrypted/word-encrypted.docx", import.meta.url))
);

function text(doc: DocxDocument): string {
  return Query.getText(doc);
}

async function rejection(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => undefined,
    (e: unknown) => e
  );
}

describe("a document Word encrypted", () => {
  it("Io.read opens it with its password", async () => {
    expect(text(await Io.read(OFFICE, { password: "Word-pw1" }))).toContain("encrypted by Word");
  });

  it("WordCrypto.decrypt returns its package", async () => {
    // It returned garbage: the package segments were keyed with the password salt, not the key data's.
    const zip = await WordCrypto.decrypt(OFFICE, "Word-pw1");
    expect([...zip.subarray(0, 2)]).toEqual([0x50, 0x4b]);
    expect(text(await Io.read(zip))).toContain("encrypted by Word");
  });

  it("WordCrypto.Agile decrypts it from parsed info", async () => {
    const entries = WordCrypto.Cfb.read(OFFICE);
    const infoStream = entries.find(e => e.name === "EncryptionInfo")!.data;
    const info = WordCrypto.Agile.parseEncryptionInfo(
      new TextDecoder().decode(infoStream.subarray(8))
    );
    expect(info.keyData).toBeDefined();
    const pkg = entries.find(e => e.name === "EncryptedPackage")!.data;
    const zip = await WordCrypto.Agile.decryptPackage(pkg, info, "Word-pw1");
    expect([...zip.subarray(0, 2)]).toEqual([0x50, 0x4b]);
  });

  it("Stream.DocumentReader opens it with its password", async () => {
    const body: unknown[] = [];
    for await (const item of new Stream.DocumentReader(OFFICE, { password: "Word-pw1" })) {
      body.push(item);
    }
    expect(body).toEqual((await Io.read(OFFICE, { password: "Word-pw1" })).body);
  });

  it("without a password is DocxEncryptedError; with a wrong one DocxDecryptionError", async () => {
    expect(await rejection(Io.read(OFFICE))).toBeInstanceOf(DocxEncryptedError);
    const wrong = await rejection(Io.read(OFFICE, { password: "nope" }));
    expect(wrong).toBeInstanceOf(DocxDecryptionError);
    expect((wrong as Error).message).toBe("Incorrect password");
    expect((wrong as Error).cause).toBeInstanceOf(Error);
    expect(await WordCrypto.verifyPassword(OFFICE, "Word-pw1")).toBe(true);
    expect(await WordCrypto.verifyPassword(OFFICE, "nope")).toBe(false);
  });
});

describe("Io writes an encrypted package with { password }", () => {
  function sample(): DocxDocument {
    const doc = Document.create();
    Document.addParagraph(doc, "secret text");
    return doc;
  }
  const fast = { spinCount: 1000 } as const;

  it("Io.toBuffer encrypts; Io.read with the password reads it back", async () => {
    const bytes = await Io.toBuffer(sample(), { password: "pw", encryption: fast });
    expect(WordCrypto.isEncrypted(bytes)).toBe(true);
    expect(text(await Io.read(bytes, { password: "pw" }))).toContain("secret text");
    expect(await rejection(Io.read(bytes, { password: "wrong" }))).toBeInstanceOf(
      DocxDecryptionError
    );
    expect(await rejection(Io.read(bytes))).toBeInstanceOf(DocxEncryptedError);
  });

  it("uses Office's defaults when no encryption options are given", async () => {
    const bytes = await Io.toBuffer(sample(), { password: "pw" });
    const info = WordCrypto.Cfb.read(bytes).find(e => e.name === "EncryptionInfo")!.data;
    const xml = new TextDecoder().decode(info.subarray(8));
    expect(xml).toContain('spinCount="100000"');
    expect(xml).toContain('keyBits="256"');
    expect(xml).toContain('hashAlgorithm="SHA512"');
  });

  it("Io.toBase64 and Io.writeFile encrypt too", async () => {
    const base64 = await Io.toBase64(sample(), { password: "pw", encryption: fast });
    expect(text(await Io.read(base64, { base64: true, password: "pw" }))).toContain("secret text");
    const path = join(await mkdtemp(join(tmpdir(), "dm-word-enc-")), "secret.docx");
    await Io.writeFile(sample(), path, { password: "pw", encryption: fast });
    expect(text(await Io.readFile(path, { password: "pw" }))).toContain("secret text");
  });

  it("writes an unencrypted package without a password, as before", async () => {
    const bytes = await Io.toBuffer(sample());
    expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b]);
  });

  it.each([
    [{ password: "" }, "options.password"],
    [{ password: 42 }, "options.password"],
    [{ password: "x".repeat(256) }, "options.password"],
    [{ encryption: fast }, "options.encryption"],
    [{ password: "pw", encryption: { keyBits: 100 } }, "options.encryption.keyBits"],
    [{ password: "pw", encryption: { hashAlgorithm: "MD5" } }, "options.encryption.hashAlgorithm"],
    [{ password: "pw", encryption: { spinCount: -1 } }, "options.encryption.spinCount"]
  ])("refuses %j naming %s", async (options, field) => {
    for (const write of [
      () => Io.toBuffer(sample(), options as never),
      () => Io.toBase64(sample(), options as never),
      () => Io.writeFile(sample(), join(tmpdir(), "never-written.docx"), options as never)
    ]) {
      const error = await rejection(write());
      expect(error).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field).toBe(field);
    }
  });
});
