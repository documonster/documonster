/**
 * `Document.setWatermark(doc, undefined)` and `Document.setBackground(doc,
 * undefined)` take the watermark / background off, as Excel's
 * `Watermark.remove` does. Before, both refused `undefined`, so a watermark a
 * file was read with could only be dropped by rebuilding the document outside
 * `Document`.
 */

import { describe, it, expect } from "vitest";

import { Document, DocxArgumentError, Io, Package } from "../index";

async function documentXml(bytes: Uint8Array, path: string): Promise<string> {
  const part = await Package.readPart(bytes, path);
  return part ? new TextDecoder().decode(part) : "";
}

describe("clearing a watermark or background", () => {
  it("setWatermark(doc, undefined) removes a read watermark from every header", async () => {
    const d = Document.create();
    Document.addParagraph(d, "x");
    Document.setWatermark(d, { type: "text", text: "DRAFT" });
    const read = await Io.read(await Io.toBuffer(d));
    expect(read.watermark).toBeDefined();

    Document.setWatermark(read, undefined);
    expect(read.watermark).toBeUndefined();
    const bytes = await Io.toBuffer(read);
    const parts = await Package.listParts(bytes);
    for (const header of parts.filter(p => /^word\/header\d*\.xml$/.test(p))) {
      expect(await documentXml(bytes, header)).not.toContain("DRAFT");
    }
    expect((await Io.read(bytes)).watermark).toBeUndefined();
  });

  it("setBackground(doc, undefined) removes the background", async () => {
    const d = Document.create();
    Document.addParagraph(d, "x");
    Document.setBackground(d, { color: "FF0000" });
    const read = await Io.read(await Io.toBuffer(d));
    expect(read.background).toEqual({ color: "FF0000" });

    Document.setBackground(read, undefined);
    expect("background" in read).toBe(false);
    const bytes = await Io.toBuffer(read);
    expect(await documentXml(bytes, "word/document.xml")).not.toContain("w:background");
    expect((await Io.read(bytes)).background).toBeUndefined();
  });

  it("still refuses a value that is neither a watermark nor undefined", () => {
    const d = Document.create();
    expect(() => Document.setWatermark(d, null as never)).toThrow(DocxArgumentError);
    expect(() => Document.setBackground(d, 42 as never)).toThrow(DocxArgumentError);
  });
});
