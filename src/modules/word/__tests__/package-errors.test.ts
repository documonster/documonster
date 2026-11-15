/**
 * Bytes that are not a Word package this library can read are a `DocxPackageError` whose `reason` says what was found
 * — the same reason from `Document.read`, `Stream.DocumentReader` and the `Package` functions — so a caller can tell a
 * legacy `.doc` from a damaged file, from something that is not a Word document, without reading the message.
 */
import { createZip } from "@archive/create-archive";
import { encryptIfRequested } from "@word/security/package-encryption";
import { describe, expect, it } from "vitest";

import {
  Document,
  DocxMissingPartError,
  DocxPackageError,
  DocxParseError,
  isDocxError,
  Package,
  Stream
} from "../index";

type Reason = DocxPackageError["reason"];

/** A compound file with no encrypted package in it: what a legacy `.doc` looks like to the reader. */
function compoundFile(): Uint8Array {
  const bytes = new Uint8Array(512);
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return bytes;
}

/** `bytes` with the first part's compressed data broken, so it does not inflate. */
function brokenDeflate(bytes: Uint8Array): Uint8Array {
  const out = bytes.slice();
  const nameLength = out[26]! | (out[27]! << 8);
  const extraLength = out[28]! | (out[29]! << 8);
  out[30 + nameLength + extraLength + 2]! ^= 0xff;
  return out;
}

async function zipOf(parts: Record<string, string>): Promise<Uint8Array> {
  const archive = createZip();
  for (const [path, text] of Object.entries(parts)) {
    archive.add(path, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

async function failure(read: () => Promise<unknown>): Promise<unknown> {
  try {
    await read();
  } catch (error) {
    return error;
  }
  throw new Error("expected the read to fail");
}

async function streamAll(bytes: Uint8Array): Promise<void> {
  for await (const _item of new Stream.DocumentReader(bytes)) {
    // read through
  }
}

function expectReason(error: unknown, reason: Reason): void {
  expect(error).toBeInstanceOf(DocxPackageError);
  expect(error).toBeInstanceOf(DocxParseError);
  expect(isDocxError(error)).toBe(true);
  expect((error as DocxPackageError).reason).toBe(reason);
}

/** Every reader that reads a package's body, given `bytes`. */
async function everyReader(bytes: Uint8Array): Promise<unknown[]> {
  return [await failure(() => Document.read(bytes)), await failure(() => streamAll(bytes))];
}

describe("DocxPackageError", () => {
  it('reports a compound file that is not encrypted as "legacy-format"', async () => {
    for (const error of [
      ...(await everyReader(compoundFile())),
      await failure(() => Package.listParts(compoundFile()))
    ]) {
      expectReason(error, "legacy-format");
      expect((error as Error).message).toContain("legacy .doc");
    }
  });

  it('reports bytes that are not a ZIP package as "not-a-document"', async () => {
    const text = new TextEncoder().encode("name,amount\nwidget,3\n");
    for (const error of [
      ...(await everyReader(text)),
      await failure(() => Package.listParts(text))
    ]) {
      expectReason(error, "not-a-document");
    }
  });

  it('reports a truncated package as "damaged"', async () => {
    const bytes = await Document.toBuffer(Document.create());
    const truncated = bytes.slice(0, bytes.length >> 1);
    for (const error of [
      ...(await everyReader(truncated)),
      await failure(() => Package.listParts(truncated))
    ]) {
      expectReason(error, "damaged");
    }
  });

  it('reports compressed data that does not inflate as "damaged"', async () => {
    const bytes = brokenDeflate(await Document.toBuffer(Document.create()));
    expectReason(await failure(() => Document.read(bytes)), "damaged");
  });

  // Damaged or not-a-ZIP is told by the first bytes of what the ZIP reader was given — for an encrypted
  // document, the decrypted package, not the compound file that held it.
  it('reports an encrypted package that decrypts to a damaged one as "damaged"', async () => {
    const bytes = await Document.toBuffer(Document.create());
    const encrypted = await encryptIfRequested(bytes.slice(0, bytes.length >> 1), {
      password: "secret"
    });
    const options = { password: "secret" };
    for (const error of [
      await failure(() => Document.read(encrypted, options)),
      await failure(async () => {
        for await (const _item of new Stream.DocumentReader(encrypted, options)) {
          // read through
        }
      }),
      await failure(() => Package.listParts(encrypted, options))
    ]) {
      expectReason(error, "damaged");
    }
  });

  it('reports a package whose main part is not a Word document as "not-a-document"', async () => {
    const workbook = await zipOf({
      "_rels/.rels":
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        "</Relationships>",
      "xl/workbook.xml":
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>'
    });
    for (const error of await everyReader(workbook)) {
      expectReason(error, "not-a-document");
      expect((error as Error).message).toContain("<workbook>");
    }
  });

  it('reports a package with no main part as "not-a-document"', async () => {
    const empty = await zipOf({ "docProps/app.xml": "<Properties/>" });
    for (const error of await everyReader(empty)) {
      expectReason(error, "not-a-document");
      expect(error).toBeInstanceOf(DocxMissingPartError);
    }
  });

  it("leaves a Word document with malformed content a plain DocxParseError", async () => {
    const bytes = await zipOf({
      "word/document.xml":
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    });
    const error = await failure(() => Document.read(bytes));
    expect(error).toBeInstanceOf(DocxParseError);
    expect(error).not.toBeInstanceOf(DocxPackageError);
  });

  // Flip a byte or cut the file anywhere and the read either succeeds or fails with a Word error.
  it("answers every damaged variant of a package with a Word error", async () => {
    const docx = await Document.toBuffer(Document.create());
    let seed = 7;
    const random = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const variants: Uint8Array[] = [];
    for (let i = 0; i < 40; i++) {
      const copy = docx.slice();
      copy[Math.floor(random() * copy.length)]! ^= 1 + Math.floor(random() * 254);
      variants.push(copy);
    }
    for (let i = 1; i < 20; i++) {
      variants.push(docx.slice(0, Math.floor((docx.length * i) / 20)));
    }
    for (const bytes of variants) {
      for (const read of [() => Document.read(bytes), () => streamAll(bytes)]) {
        const error = await read().then(
          () => undefined,
          (e: unknown) => e
        );
        if (error !== undefined) {
          expect(isDocxError(error), String(error)).toBe(true);
        }
      }
    }
  });
});
