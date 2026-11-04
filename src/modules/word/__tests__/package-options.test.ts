/**
 * `Package.*` reads `.docx` bytes as `Io.read` reads them: under the caller's
 * `securityPolicy`, decrypting with `password`, and refusing an encrypted file
 * without one with `DocxEncryptedError`. Before, every `Package` function read
 * under the default policy whatever it was given — `{ securityPolicy }` was
 * accepted and had no effect — and `listParts` / `readPart` / `edit` called an
 * encrypted file a broken ZIP. The writing ones take `Io.toBuffer`'s `modTime`
 * and `securityPolicy` beside `compressionLevel`.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import {
  Document,
  DocxEncryptedError,
  DocxLimitExceededError,
  Io,
  Package,
  Query,
  Security
} from "../index";
import type { DocxArgumentError } from "../index";

async function sample(): Promise<Uint8Array> {
  const d = Document.create();
  Document.addParagraph(d, "Hello {{name}}");
  return Io.toBuffer(d);
}

const tight = { maxPartCount: 1 };

describe("Package.* read options", () => {
  it("apply securityPolicy limits on every read", async () => {
    const bytes = await sample();
    const calls: [string, () => Promise<unknown>][] = [
      ["fillTemplate", () => Package.fillTemplate(bytes, { name: "x" }, { securityPolicy: tight })],
      ["patch", () => Package.patch(bytes, [], { securityPolicy: tight })],
      ["edit", () => Package.edit(bytes, [], { securityPolicy: tight })],
      ["compileTemplate", () => Package.compileTemplate(bytes, { securityPolicy: tight })],
      ["listParts", () => Package.listParts(bytes, { securityPolicy: tight })],
      ["readPart", () => Package.readPart(bytes, "word/document.xml", { securityPolicy: tight })]
    ];
    for (const [label, call] of calls) {
      await expect(call(), label).rejects.toBeInstanceOf(DocxLimitExceededError);
    }
  });

  it("limit each part's size where its bytes are read", async () => {
    const bytes = await sample();
    const policy = { maxPartSize: 10 };
    await expect(
      Package.readPart(bytes, "word/document.xml", { securityPolicy: policy })
    ).rejects.toBeInstanceOf(DocxLimitExceededError);
    await expect(Package.edit(bytes, [], { securityPolicy: policy })).rejects.toBeInstanceOf(
      DocxLimitExceededError
    );
    // `listParts` decompresses nothing, so no part's size is read.
    await expect(Package.listParts(bytes, { securityPolicy: policy })).resolves.toContain(
      "word/document.xml"
    );
  });

  it("report a bad policy as the call's own option", async () => {
    const bytes = await sample();
    const error = await Package.listParts(bytes, {
      securityPolicy: { maxEntries: 1 } as never
    }).catch((e: unknown) => e);
    expect((error as DocxArgumentError).field).toBe("options.securityPolicy.maxEntries");
    expect((error as Error).message).toMatch(/^Package\.listParts: /);
    const write = await Package.patch(bytes, [], {
      securityPolicy: { maxPartSize: Number.NaN }
    }).catch((e: unknown) => e);
    expect((write as DocxArgumentError).field).toBe("options.securityPolicy.maxPartSize");
    expect((write as Error).message).toMatch(/^Package\.patch: /);
  });

  it("apply STRICT_SECURITY_POLICY", async () => {
    const bytes = await sample();
    await expect(
      Package.fillTemplate(
        bytes,
        { name: "x" },
        { securityPolicy: Security.STRICT_SECURITY_POLICY }
      )
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  it("refuse an encrypted package without a password, as Io.read does", async () => {
    const enc = await WordCrypto.encrypt(await sample(), "pw", { spinCount: 1 });
    const calls: [string, () => Promise<unknown>][] = [
      ["fillTemplate", () => Package.fillTemplate(enc, {})],
      ["patch", () => Package.patch(enc, [])],
      ["edit", () => Package.edit(enc, [])],
      ["compileTemplate", () => Package.compileTemplate(enc)],
      ["listParts", () => Package.listParts(enc)],
      ["readPart", () => Package.readPart(enc, "word/document.xml")]
    ];
    for (const [label, call] of calls) {
      await expect(call(), label).rejects.toBeInstanceOf(DocxEncryptedError);
    }
  });

  it("decrypt an encrypted package with { password }", async () => {
    const plain = await sample();
    const enc = await WordCrypto.encrypt(plain, "pw", { spinCount: 1 });
    expect(await Package.listParts(enc, { password: "pw" })).toEqual(
      await Package.listParts(plain)
    );
    expect(await Package.readPart(enc, "word/document.xml", { password: "pw" })).toEqual(
      await Package.readPart(plain, "word/document.xml")
    );
    const filled = await Package.fillTemplate(enc, { name: "Ann" }, { password: "pw" });
    expect(Document.getContentCount(await Io.read(filled))).toBe(1);
    const compiled = await Package.compileTemplate(enc, { password: "pw" });
    await expect(Package.patchTemplate(compiled, [])).resolves.toBeInstanceOf(Uint8Array);
    await expect(Package.patch(enc, [], { password: "pw" })).resolves.toBeInstanceOf(Uint8Array);
    await expect(Package.edit(enc, [], { password: "pw" })).resolves.toBeInstanceOf(Uint8Array);
  });

  it("read a Flat OPC package as Io.read does", async () => {
    const d = Document.create();
    Document.addParagraph(d, "flat");
    const flat = new TextEncoder().encode(await Io.toFlatOpc(d));
    const zipped = await Io.toBuffer(d);
    expect((await Package.listParts(flat)).sort()).toEqual(
      (await Package.listParts(zipped)).sort()
    );
    const xml = await Package.readPart(flat, "word/document.xml");
    expect(new TextDecoder().decode(xml)).toContain("flat");
    const edited = await Package.edit(flat, []);
    expect(Query.getText(await Io.read(edited))).toContain("flat");
  });
});

describe("Package.* write options", () => {
  it("honour modTime, so the same input and time give the same bytes", async () => {
    const bytes = await sample();
    const modTime = new Date(Date.UTC(2020, 0, 2, 3, 4, 6));
    const other = new Date(Date.UTC(2021, 5, 6, 7, 8, 10));
    const runs: [string, (t: Date) => Promise<Uint8Array>][] = [
      ["edit", t => Package.edit(bytes, [], { modTime: t })],
      ["patch", t => Package.patch(bytes, [], { modTime: t })],
      ["fillTemplate", t => Package.fillTemplate(bytes, { name: "x" }, { modTime: t })]
    ];
    for (const [label, run] of runs) {
      expect(await run(modTime), label).toEqual(await run(modTime));
      expect(await run(modTime), label).not.toEqual(await run(other));
    }
  });
});
