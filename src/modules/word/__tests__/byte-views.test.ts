/**
 * Every Word entry point that takes bytes reads any `ArrayBufferView` — a
 * `DataView`, a `Uint16Array`, a Node `Buffer` — as the bytes it views, as
 * `Workbook.read` does. Before, `ArrayBuffer.isView` let them past the check
 * and the code then treated them as a `Uint8Array`: a `DataView` failed deep
 * inside (`content.slice is not a function`) and a wider typed array was
 * read element by element.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import { DocxArgumentError } from "../errors";
import { Convert, Document, Io, Package, Stream, Vba } from "../index";
import type { DocxDocument } from "../index";

const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

/** `bytes` seen through each kind of view, offset into a larger buffer. */
function views(bytes: Uint8Array): [string, ArrayBufferView][] {
  // An even offset and length, so a 16-bit view can cover exactly these bytes.
  const padded = bytes.length % 2 === 0 ? bytes : Uint8Array.from([...bytes, 0]);
  const backing = new Uint8Array(padded.length + 8);
  backing.set(padded, 4);
  return [
    ["DataView", new DataView(backing.buffer, 4, bytes.length)],
    ["Uint8ClampedArray", new Uint8ClampedArray(backing.buffer, 4, bytes.length)],
    ...(bytes.length % 2 === 0
      ? [
          ["Uint16Array", new Uint16Array(backing.buffer, 4, bytes.length / 2)] as [
            string,
            ArrayBufferView
          ]
        ]
      : [])
  ];
}

function doc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "hello");
  return d;
}

describe("bytes given as any ArrayBufferView", () => {
  it("are read by the readers as the bytes they view", async () => {
    const bytes = await Io.toBuffer(doc());
    const parts = await Package.listParts(bytes);
    const flat = new TextEncoder().encode(await Io.toFlatOpc(doc()));
    for (const [label, view] of views(bytes)) {
      expect((await Io.read(view as never)).body.length, label).toBe(1);
      expect(await Package.listParts(view as never), label).toEqual(parts);
      const blocks: string[] = [];
      for await (const block of new Stream.DocumentReader(view as never)) {
        blocks.push(block.type);
      }
      expect(blocks, label).toEqual(["paragraph"]);
      expect(WordCrypto.isEncrypted(view as never), label).toBe(false);
    }
    for (const [label, view] of views(flat)) {
      expect(Convert.isFlatOpc(view as never), label).toBe(true);
      expect(Convert.flatOpcToParts(view as never).size, label).toBeGreaterThan(0);
    }
  });

  it("are taken by the builders and the encryptor as the bytes they view", async () => {
    const bytes = await Io.toBuffer(doc());
    for (const [label, view] of views(bytes)) {
      const enc = await WordCrypto.encrypt(view as never, "pw", { spinCount: 1 });
      expect(await WordCrypto.decrypt(enc, "pw"), label).toEqual(bytes);
    }
    for (const [label, view] of views(PNG)) {
      const d = Document.create();
      Document.addImage(d, view as never, "png", 9525, 9525);
      expect(d.images?.[0]?.data, label).toEqual(PNG);
      await expect(Io.toBuffer(d), label).resolves.toBeInstanceOf(Uint8Array);
    }
    const ole = new Uint8Array(512);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    for (const [label, view] of views(ole)) {
      expect(Vba.add(doc(), view as never).vbaProject, label).toEqual(ole);
    }
  });

  it("an ArrayBuffer is taken wherever a package or file is, as Io.read takes it", async () => {
    const bytes = await Io.toBuffer(doc());
    const ab = (u: Uint8Array): ArrayBuffer =>
      u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
    const buf = ab(bytes);
    const parts = await Package.listParts(bytes);
    expect(await Package.listParts(buf)).toEqual(parts);
    expect(await Package.readPart(buf, "word/document.xml")).toEqual(
      await Package.readPart(bytes, "word/document.xml")
    );
    await expect(Package.edit(buf, [])).resolves.toBeInstanceOf(Uint8Array);
    await expect(Package.patch(buf, [])).resolves.toBeInstanceOf(Uint8Array);
    await expect(Package.compileTemplate(buf)).resolves.toBeDefined();
    await expect(Package.fillTemplate(buf, {})).resolves.toBeInstanceOf(Uint8Array);
    const blocks: string[] = [];
    for await (const block of new Stream.DocumentReader(buf)) {
      blocks.push(block.type);
    }
    expect(blocks).toEqual(["paragraph"]);

    const enc = await WordCrypto.encrypt(buf, "pw", { spinCount: 1 });
    expect(WordCrypto.isEncrypted(ab(enc))).toBe(true);
    expect(await WordCrypto.verifyPassword(ab(enc), "pw")).toBe(true);
    expect(await WordCrypto.decrypt(ab(enc), "pw")).toEqual(bytes);

    const flat = new TextEncoder().encode(await Io.toFlatOpc(doc()));
    expect(Convert.isFlatOpc(ab(flat))).toBe(true);
    expect(Convert.flatOpcToParts(ab(flat)).size).toBeGreaterThan(0);
    const odt = await Convert.toOdt(doc());
    expect(Document.getContentCount(await Convert.fromOdt(ab(odt)))).toBe(1);
  });

  it("anything else is refused with DocxArgumentError", async () => {
    for (const value of [42, "x", [1, 2], { length: 4 }]) {
      await expect(Io.read(value as never)).rejects.toBeInstanceOf(DocxArgumentError);
      expect(() => new Stream.DocumentReader(value as never)).toThrow(DocxArgumentError);
    }
  });
});
