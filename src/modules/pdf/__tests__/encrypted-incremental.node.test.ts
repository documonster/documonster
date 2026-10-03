/**
 * Incremental saves of encrypted PDFs written by another producer (qpdf
 * 12.4): RC4 40/128-bit, AES-128 and AES-256 revision 6 — the handler
 * Acrobat and qpdf use by default. The appended objects must be encrypted
 * with the file's own handler; the same outputs were also checked with
 * `qpdf --check`, pypdf and macOS PDFKit.
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";

import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

const DIR = join(__dirname, "data", "encrypted");
const latin1 = (bytes: Uint8Array) => new TextDecoder("latin1").decode(bytes);

describe.each(["rc4-40", "rc4-128", "aes-128", "aes-256-r6"])("%s", name => {
  const original = new Uint8Array(readFileSync(join(DIR, `${name}.pdf`)));

  it("opens with the user and the owner password, and refuses a wrong one", async () => {
    for (const password of ["user", "owner"]) {
      const read = await Pdf.read(original, { password });
      expect(read.pages[0].text).toContain("Base text");
    }
    await expect(Pdf.read(original, { password: "wrong" })).rejects.toThrow();
  });

  it("appends an encrypted update that reads back", async () => {
    const editor = Pdf.Editor.load(original, { password: "user" });
    editor.getPage(0).drawText("Appended (é) text", { x: 72, y: 450, fontSize: 12 });
    const updated = await editor.saveIncremental();
    expect(updated.subarray(0, original.length)).toEqual(original);
    expect(latin1(updated.subarray(original.length))).not.toContain("Appended");
    const read = await Pdf.read(updated, { password: "user" });
    expect(read.pages[0].text).toContain("Base text");
    expect(read.pages[0].text).toContain("Appended (é) text");
  });
});

/**
 * Crypt filters (ISO 32000-1 §7.6.5): `/StrF` and `/StmF` are resolved
 * separately, `/Identity` passes data through in both directions, and with
 * `/EncryptMetadata false` the metadata stream is stored in the clear.
 *
 * `aes-128-identity-*.pdf` are written by `make-identity-fixtures.py` (qpdf
 * cannot produce an Identity filter); `aes-128-cleartext-metadata.pdf` by
 * `qpdf --encrypt user owner 128 --use-aes=y --cleartext-metadata`, and
 * `aes-128-object-streams.pdf` by the same with `--object-streams=generate`.
 * qpdf 12.4 `--check` and pypdf 6.11 read all four, and the incremental
 * outputs below, with the user password.
 */
describe("crypt filters", () => {
  const load = (name: string) => new Uint8Array(readFileSync(join(DIR, `${name}.pdf`)));

  async function roundTrip(original: Uint8Array): Promise<Uint8Array> {
    const editor = Pdf.Editor.load(original, { password: "user" });
    editor.getPage(0).drawText("Appended text", { x: 72, y: 450, fontSize: 12 });
    const updated = await editor.saveIncremental();
    const read = await Pdf.read(updated, { password: "user" });
    expect(read.pages[0].text).toContain("Base text");
    expect(read.pages[0].text).toContain("Appended text");
    return updated;
  }

  it.each(["aes-128-identity-streams", "aes-128-identity-strings"])(
    "%s: reads strings and streams through their own filter",
    async name => {
      const read = await Pdf.read(load(name), { password: "user" });
      expect(read.pages[0].text).toContain("Base text");
      expect(read.metadata.title).toBe("Secret title");
    }
  );

  it("leaves stream data in the clear when /StmF is /Identity", async () => {
    const original = load("aes-128-identity-streams");
    // The fixture's content stream is stored plain …
    expect(latin1(original)).toContain("(Base text) Tj");
    // … and the update's must be plain too: its page dict still decrypts.
    const updated = await roundTrip(original);
    const appended = latin1(updated.subarray(original.length));
    expect(appended).toMatch(/stream\n(?:BT|q|[0-9.\s]+)/);
  });

  it("leaves strings in the clear when /StrF is /Identity", async () => {
    await roundTrip(load("aes-128-identity-strings"));
  });

  it("reads a cleartext metadata stream when /EncryptMetadata is false", async () => {
    const original = load("aes-128-cleartext-metadata");
    expect(latin1(original)).toContain("Cleartext XMP");
    const read = await Pdf.read(original, { password: "user" });
    expect(read.metadata.xmpXml).toContain("Cleartext XMP");
    expect(read.pages[0].text).toContain("Base text");
    await roundTrip(original);
  });

  it("does not decrypt strings of objects inside an object stream twice", async () => {
    const original = load("aes-128-object-streams");
    const read = await Pdf.read(original, { password: "user" });
    expect(read.metadata.title).toBe("Secret title");
    await roundTrip(original);
  });
});
