/**
 * What an incremental update chains onto with /Prev.
 *
 * - Hybrid-reference files (ISO 32000-1 §7.5.8.4): `data/hybrid/hybrid-reference.pdf`
 *   is written by `data/hybrid/make-hybrid.py`; its page and font live only in an
 *   object stream listed by the trailer's /XRefStm. The update is a classic
 *   section whose /Prev names that trailer. The output was also checked with
 *   `qpdf --check` (12.4) and pypdf.
 * - Recovered cross-reference data: a wrong startxref leaves no reliable
 *   section to chain to, so saveIncremental() becomes a full save().
 *   `data/objstm/objstm.pdf` (written by `data/objstm/make-objstm.py` with
 *   `qpdf --object-streams=generate`) keeps its catalog, page tree and pages in
 *   an object stream and has no classic trailer, so recovery must enumerate
 *   the object stream and take /Root from the xref stream's dictionary.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

import { PdfSignatureInvalidationError } from "@pdf/errors";
import { Pdf } from "@pdf/index";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { describe, expect, it } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";

// Byte-for-byte: TextDecoder("latin1") is really windows-1252 and would not
// round-trip 0x80–0x9F through charCodeAt, corrupting compressed streams.
const latin1 = (bytes: Uint8Array) => Buffer.from(bytes).toString("latin1");
const fromLatin1 = (text: string) => new Uint8Array(Buffer.from(text, "latin1"));
const OBJSTM = new Uint8Array(readFileSync(join(__dirname, "data", "objstm", "objstm.pdf")));
const HYBRID = new Uint8Array(
  readFileSync(join(__dirname, "data", "hybrid", "hybrid-reference.pdf"))
);

function dictName(obj: unknown, key: string): unknown {
  return obj instanceof Map ? obj.get(key) : undefined;
}

/** Point the last startxref at `to`, where no xref section starts. */
function breakStartxref(pdf: Uint8Array, to = "1"): Uint8Array {
  const text = latin1(pdf);
  const at = text.lastIndexOf("startxref");
  const fixed = text.slice(0, at) + text.slice(at).replace(/startxref\s+\d+/, `startxref\n${to}`);
  return fromLatin1(fixed);
}

describe("hybrid-reference file", () => {
  it("resolves objects listed only in /XRefStm", async () => {
    const doc = new PdfDocument(HYBRID);
    expect(doc.xrefRecovered).toBe(false);
    expect(dictName(doc.resolve(3), "Type")).toBe("Page");
    expect(dictName(doc.resolve(5), "BaseFont")).toBe("Helvetica");
    const read = await Pdf.read(HYBRID);
    expect(read.pages[0].text).toContain("Original hybrid text");
  });

  it("appends a classic section chained to the hybrid trailer", async () => {
    const editor = Pdf.Editor.load(HYBRID);
    editor.getPage(0).drawText("Appended text", { x: 72, y: 600, fontSize: 12 });
    const updated = await editor.saveIncremental();
    expect(updated.subarray(0, HYBRID.length)).toEqual(HYBRID);

    const original = new PdfDocument(HYBRID);
    const appended = latin1(updated.subarray(HYBRID.length));
    expect(appended).toMatch(/\ntrailer\n/);
    expect(appended).toContain(`/Prev ${original.startxrefOffset}\n`);

    const doc = new PdfDocument(updated);
    expect(doc.xrefRecovered).toBe(false);
    // The font is still reached through the original trailer's /XRefStm.
    expect(dictName(doc.resolve(5), "BaseFont")).toBe("Helvetica");
    const read = await Pdf.read(updated);
    expect(read.pages[0].text).toContain("Original hybrid text");
    expect(read.pages[0].text).toContain("Appended text");
  });
});

describe("recovered cross-reference data", () => {
  it("is reported by PdfDocument", () => {
    const broken = breakStartxref(HYBRID);
    expect(new PdfDocument(broken).xrefRecovered).toBe(true);
    expect(new PdfDocument(broken).startxrefOffset).toBeNull();
  });

  it("falls back to a full save instead of chaining to a broken section", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().drawText("Base text", { x: 72, y: 700, fontSize: 12 });
    const broken = breakStartxref(await builder.build());
    expect(new PdfDocument(broken).xrefRecovered).toBe(true);

    const editor = Pdf.Editor.load(broken);
    editor.getPage(0).drawText("Appended text", { x: 72, y: 600, fontSize: 12 });
    const out = await editor.saveIncremental();

    // A rewrite, not the broken bytes plus an update.
    expect(latin1(out.subarray(0, broken.length))).not.toBe(latin1(broken));
    const doc = new PdfDocument(out);
    expect(doc.xrefRecovered).toBe(false);
    const read = await Pdf.read(out);
    expect(read.pages[0].text).toContain("Base text");
    expect(read.pages[0].text).toContain("Appended text");
  });

  it("keeps the signature guard on that fallback", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate("Recovered");
    const builder = new Pdf.Builder();
    builder.addPage().drawText("Signed", { x: 72, y: 750, fontSize: 14 });
    const signed = await Pdf.Editor.load(await builder.build()).sign({ certificate, privateKey });
    const broken = breakStartxref(signed);

    const editor = Pdf.Editor.load(broken);
    expect(editor.hasSignatures).toBe(true);
    editor.getPage(0).drawText("Appended", { x: 72, y: 600, fontSize: 12 });
    await expect(editor.saveIncremental()).rejects.toThrow(PdfSignatureInvalidationError);
    await expect(editor.saveIncremental({ invalidateSignatures: true })).resolves.toBeInstanceOf(
      Uint8Array
    );
  });
});

describe("recovery of objects stored in object streams", () => {
  const broken = breakStartxref(OBJSTM, "999999999");

  it("registers object-stream members and finds /Root in the xref stream", async () => {
    const doc = new PdfDocument(broken);
    expect(doc.xrefRecovered).toBe(true);
    expect(doc.getPages()).toHaveLength(3);
    const read = await Pdf.read(broken);
    expect(read.pages.map(p => p.text.trim())).toEqual([
      "Page 1 text",
      "Page 2 text",
      "Page 3 text"
    ]);
  });

  it("lets a direct object win over a compressed copy", () => {
    // Append a direct copy of the first page with a different /MediaBox and no
    // xref section: recovery sees that object number both in the object
    // stream and as `N 0 obj`, and must take the direct one.
    const first = new PdfDocument(broken).getPagesWithObjInfo()[0];
    const page = new PdfDocument(OBJSTM).resolve(first.objNum) as Map<string, unknown>;
    const parent = page.get("Parent") as { objNum: number };
    const contents = page.get("Contents") as { objNum: number };
    const patched =
      latin1(broken) +
      `${first.objNum} 0 obj\n<< /Type /Page /Parent ${parent.objNum} 0 R /MediaBox [0 0 111 222] ` +
      `/Contents ${contents.objNum} 0 R >>\nendobj\n`;
    const doc = new PdfDocument(fromLatin1(patched));
    expect(doc.getPages()).toHaveLength(3);
    expect(doc.resolvePageBox(doc.getPages()[0])).toEqual({ width: 111, height: 222 });
    expect(doc.resolvePageBox(doc.getPages()[1])).toEqual({ width: 300, height: 200 });
  });

  it("falls back to a full save that keeps every page", async () => {
    const editor = Pdf.Editor.load(broken);
    editor.getPage(2).drawText("Appended text", { x: 20, y: 50, fontSize: 12 });
    const out = await editor.saveIncremental();
    expect(latin1(out.subarray(0, broken.length))).not.toBe(latin1(broken));
    const doc = new PdfDocument(out);
    expect(doc.xrefRecovered).toBe(false);
    const read = await Pdf.read(out);
    expect(read.pages).toHaveLength(3);
    expect(read.pages[0].text).toContain("Page 1 text");
    expect(read.pages[2].text).toContain("Page 3 text");
    expect(read.pages[2].text).toContain("Appended text");
    if (process.env.OBJSTM_DUMP) {
      writeFileSync(process.env.OBJSTM_DUMP, out);
      writeFileSync(`${process.env.OBJSTM_DUMP}.broken.pdf`, broken);
    }
  });
});
