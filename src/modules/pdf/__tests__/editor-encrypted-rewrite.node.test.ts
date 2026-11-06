/**
 * Full rewrites of encrypted PDFs. A rewrite rebuilds every object, so it has to encrypt them
 * again — with the file's own security handler, so the passwords that opened the original
 * still open the output and its permissions are unchanged. Writing the result in the clear
 * takes an explicit `{ encryption: "remove" }` — except through the deprecated `save()` (and
 * `saveIncremental()`'s fallback to it), which was documented to write a decrypted copy and
 * still does.
 *
 * The fixtures are the qpdf 12.4 files of `encrypted-incremental.node.test.ts`: RC4 40/128,
 * AES-128 and AES-256 R6, plus the crypt-filter variants. The same outputs were also checked
 * with `qpdf --check` and `qpdf --show-encryption` under both passwords.
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";

import { generateSelfSignedCertificate } from "@pdf/examples/utils/self-signed-certificate";
import { Pdf, PdfEncryptionError, PdfIncrementalUpdateError } from "@pdf/index";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { describe, expect, it } from "vitest";

import { buildRawPdf } from "./raw-pdf.helpers";

const DIR = join(__dirname, "data", "encrypted");
const load = (name: string) => new Uint8Array(readFileSync(join(DIR, `${name}.pdf`)));
const latin1 = (bytes: Uint8Array) => new TextDecoder("latin1").decode(bytes);

type Editor = ReturnType<typeof Pdf.Editor.read>;

/** The security handler parameters a reader derives the file key from. */
function securityOf(bytes: Uint8Array) {
  const doc = new PdfDocument(bytes);
  const encrypt = doc.derefDict(doc.trailer.get("Encrypt") ?? null);
  const id = doc.trailer.get("ID");
  return { encrypt, firstId: Array.isArray(id) ? id[0] : undefined };
}

async function expectProtectedLike(output: Uint8Array, original: Uint8Array): Promise<void> {
  await expect(Pdf.extract(output)).rejects.toThrow();
  await expect(Pdf.extract(output, { password: "wrong" })).rejects.toThrow();
  for (const password of ["user", "owner"]) {
    const read = await Pdf.extract(output, { password });
    expect(read.pages[0].text).toContain("Base text");
  }
  expect(securityOf(output)).toEqual(securityOf(original));
}

const handlers = ["rc4-40", "rc4-128", "aes-128", "aes-256-r6"];
const variants = ["aes-128-identity-streams", "aes-128-identity-strings"];
const variantsWithMetadata = ["aes-128-cleartext-metadata", "aes-128-object-streams"];

const rewrites: Array<[string, (editor: Editor) => Promise<Uint8Array>]> = [
  ["toBuffer()", editor => editor.toBuffer()],
  ['save({ encryption: "keep" })', editor => editor.save({ encryption: "keep" })],
  [
    "toBuffer({ incremental: true }) after a structural edit",
    editor => editor.rotatePage(0, 90).toBuffer({ incremental: true })
  ],
  [
    'saveIncremental({ encryption: "keep" }) after a structural edit',
    editor => editor.addPage() && editor.saveIncremental({ encryption: "keep" })
  ]
];

describe.each([...handlers, ...variants, ...variantsWithMetadata])("%s", name => {
  const original = load(name);

  describe.each(["user", "owner"])("opened with the %s password", password => {
    it.each(rewrites)("%s stays encrypted with the original handler", async (_, write) => {
      const editor = Pdf.Editor.read(original, { password });
      editor.getPage(0).drawText("Overlay (é) text", { x: 72, y: 450, fontSize: 12 });
      const output = await write(editor);
      expect(output.subarray(0, original.length)).not.toEqual(original);
      expect(latin1(output)).not.toContain("Overlay");
      await expectProtectedLike(output, original);
      const read = await Pdf.extract(output, { password: "user" });
      expect(read.pages[0].text).toContain("Overlay (é) text");
    });
  });
});

describe("full rewrite of an encrypted file", () => {
  it("the deprecated save() still writes a decrypted copy, as it was documented to", async () => {
    for (const name of handlers) {
      const editor = Pdf.Editor.read(load(name), { password: "user" });
      const output = await editor.save();
      expect(securityOf(output).encrypt).toBeNull();
      expect((await Pdf.extract(output)).pages[0].text).toContain("Base text");
    }
  });

  it("the deprecated saveIncremental() appends encrypted, and falls back to save()'s decrypted copy", async () => {
    const original = load("aes-128");
    const appending = Pdf.Editor.read(original, { password: "user" });
    appending.getPage(0).drawText("Appended", { x: 72, y: 450, fontSize: 12 });
    const appended = await appending.saveIncremental();
    expect(appended.subarray(0, original.length)).toEqual(original);
    await expectProtectedLike(appended, original);

    const falling = Pdf.Editor.read(original, { password: "user" }).rotatePage(0, 90);
    const rewritten = await falling.saveIncremental();
    expect(securityOf(rewritten).encrypt).toBeNull();
    expect((await Pdf.extract(rewritten)).pages[0].text).toContain("Base text");
  });

  it("keeps the strings and metadata readable", async () => {
    for (const name of variantsWithMetadata) {
      const original = load(name);
      const before = await Pdf.extract(original, { password: "user" });
      const output = await Pdf.Editor.read(original, { password: "user" }).toBuffer();
      const after = await Pdf.extract(output, { password: "owner" });
      expect(after.metadata.title).toBe(before.metadata.title);
      expect(after.metadata.xmpXml).toBe(before.metadata.xmpXml);
    }
  });

  it("leaves the metadata stream in the clear when /EncryptMetadata is false", async () => {
    const output = await Pdf.Editor.read(load("aes-128-cleartext-metadata"), {
      password: "user"
    }).toBuffer();
    // Read without any security handler installed: only cleartext data decodes.
    const doc = new PdfDocument(output);
    const metadata = doc.derefStreamWithObjNum(doc.getCatalog().get("Metadata") ?? null);
    expect(metadata).not.toBeNull();
    expect(latin1(doc.getStreamData(metadata!.stream))).toContain("Cleartext XMP");
  });

  it("keeps restricted permissions and both passwords of a file this library encrypted", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().drawText("Base text", { x: 72, y: 700, fontSize: 12 });
    builder.setEncryption({
      userPassword: "user",
      ownerPassword: "owner",
      permissions: { print: true, copy: false, modify: false }
    });
    const original = await builder.toBuffer();
    for (const password of ["user", "owner"]) {
      const editor = Pdf.Editor.read(original, { password });
      editor.removePage(0).addPage().drawText("Base text", { x: 72, y: 700, fontSize: 12 });
      await expectProtectedLike(await editor.toBuffer(), original);
    }
  });

  it("copies pages into an encrypted file encrypted", async () => {
    const source = await Pdf.fromTable([["Copied cell"]]);
    const editor = Pdf.Editor.read(load("aes-256-r6"), { password: "owner" });
    const output = await editor.copyPagesFrom(source).toBuffer();
    expect(latin1(output)).not.toContain("Copied");
    await expectProtectedLike(output, load("aes-256-r6"));
    expect((await Pdf.extract(output, { password: "user" })).text).toContain("Copied cell");
  });

  it("splitPages() keeps each part encrypted", async () => {
    const original = load("aes-128");
    const [part] = await Pdf.Editor.read(original, { password: "user" }).splitPages();
    await expectProtectedLike(part, original);
  });

  it("sign() keeps the encryption, and the signature verifies", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate("Encrypted");
    const original = load("aes-256-r6");
    const signed = await Pdf.Editor.read(original, { password: "user" }).sign({
      certificate,
      privateKey
    });
    await expectProtectedLike(signed, original);
    const [report] = await Pdf.verifySignatures(signed, { password: "user" });
    expect(report.intact).toBe(true);
  });

  describe('{ encryption: "remove" }', () => {
    it("writes the document decrypted", async () => {
      for (const name of handlers) {
        const editor = Pdf.Editor.read(load(name), { password: "owner" });
        const output = await editor.toBuffer({ encryption: "remove" });
        expect(securityOf(output).encrypt).toBeNull();
        expect((await Pdf.extract(output)).pages[0].text).toContain("Base text");
      }
    });

    it("splitPages() writes decrypted parts with it", async () => {
      const editor = Pdf.Editor.read(load("aes-128"), { password: "user" });
      const [part] = await editor.splitPages({ encryption: "remove" });
      expect(securityOf(part).encrypt).toBeNull();
      expect((await Pdf.extract(part)).pages[0].text).toContain("Base text");
    });

    it("rewrites rather than appending, since an update cannot remove encryption", async () => {
      const original = load("aes-128");
      const unedited = Pdf.Editor.read(original, { password: "user" });
      const output = await unedited.toBuffer({ incremental: true, encryption: "remove" });
      expect(securityOf(output).encrypt).toBeNull();
      const edited = Pdf.Editor.read(original, { password: "user" });
      edited.getPage(0).drawText("Overlay", { x: 72, y: 450, fontSize: 12 });
      const saved = await edited.toBuffer({ incremental: true, encryption: "remove" });
      expect((await Pdf.extract(saved)).text).toContain("Overlay");
      await expect(
        edited.toBuffer({ incremental: "required", encryption: "remove" })
      ).rejects.toThrow(PdfIncrementalUpdateError);
    });

    it("is a no-op for an unencrypted file", async () => {
      const plain = await Pdf.fromTable([["Plain"]]);
      const kept = await Pdf.Editor.read(plain).toBuffer();
      expect(await Pdf.Editor.read(plain).toBuffer({ encryption: "remove" })).toEqual(kept);
      const unchanged = Pdf.Editor.read(plain);
      expect(await unchanged.toBuffer({ incremental: "required", encryption: "remove" })).toBe(
        plain
      );
    });
  });

  it("refuses to rewrite a document whose /Encrypt cannot be read, unless told to remove it", async () => {
    const plain = latin1(
      buildRawPdf([
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>"
      ])
    );
    const broken = Uint8Array.from(plain.replace("/Root 1 0 R", "/Root 1 0 R /Encrypt 9 0 R"), c =>
      c.charCodeAt(0)
    );
    const editor = Pdf.Editor.read(broken);
    await expect(editor.toBuffer()).rejects.toThrow(PdfEncryptionError);
    const removed = await editor.toBuffer({ encryption: "remove" });
    expect(latin1(removed)).not.toContain("/Encrypt");
  });

  it("an incremental update is unaffected", async () => {
    const original = load("rc4-128");
    const editor = Pdf.Editor.read(original, { password: "owner" });
    editor.getPage(0).drawText("Appended", { x: 72, y: 450, fontSize: 12 });
    const updated = await editor.toBuffer({ incremental: "required" });
    expect(updated.subarray(0, original.length)).toEqual(original);
    await expectProtectedLike(updated, original);
  });
});
