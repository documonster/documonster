/**
 * One page-index model for the editor: indices `0 … pageCount - 1` name the pages the document
 * was read with, and the pages added with `addPage` / `copyPagesFrom` follow them, in the order
 * added. Every index call — `getPage`, `removePage`, `rotatePage`, `isPageRemoved`,
 * `splitPages` — takes both ranges. (Added pages used to be splittable but not fetchable,
 * removable or rotatable by index.)
 */
import { describe, expect, it } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import { Pdf, PdfArgumentError, PdfSignatureInvalidationError } from "../index";

async function pages(...labels: string[]): Promise<Uint8Array> {
  const builder = new Pdf.Builder();
  for (const label of labels) {
    builder.addPage().drawText(label, { x: 72, y: 750, fontSize: 12 });
  }
  return builder.toBuffer();
}

async function texts(bytes: Uint8Array): Promise<string[]> {
  return (await Pdf.extract(bytes)).pages.map(page => page.text.trim());
}

async function rotations(bytes: Uint8Array): Promise<number[]> {
  return (await Pdf.extract(bytes, { extractText: false })).pages.map(page => page.rotation);
}

describe("PdfEditor — one index space for read and added pages", () => {
  it("getPage returns a handle for an added page that draws onto it", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    const added = editor.addPage();
    added.drawText("NEW", { x: 72, y: 750, fontSize: 12 });
    const handle = editor.getPage(1);
    expect(handle).toBeInstanceOf(Pdf.EditorPage);
    // The same handle every time, as for a page the document was read with.
    expect(editor.getPage(1)).toBe(handle);
    expect(handle.width).toBe(added.width);
    handle.drawText("MORE", { x: 72, y: 700, fontSize: 12 });
    const [, page] = (await Pdf.extract(await editor.toBuffer())).pages;
    expect(page.text).toContain("NEW");
    expect(page.text).toContain("MORE");
  });

  it("getPage returns a handle for a copied page that overlays its content", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.copyPagesFrom(await pages("S1", "S2"));
    editor.getPage(2).drawText("STAMP", { x: 72, y: 700, fontSize: 12 });
    const out = await editor.toBuffer();
    const result = await texts(out);
    expect(result[0]).toBe("P1");
    expect(result[1]).toBe("S1");
    expect(result[2]).toContain("S2");
    expect(result[2]).toContain("STAMP");
    // Written again, the overlay is still there once.
    expect((await texts(await editor.toBuffer()))[2].match(/STAMP/g)).toHaveLength(1);
  });

  it("removePage drops an added page from the output, and isPageRemoved says so", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    editor.addPage().drawText("NEW", { x: 72, y: 750, fontSize: 12 });
    editor.copyPagesFrom(await pages("S1"));
    expect(editor.outputPageCount).toBe(4);
    editor.removePage(2).removePage(2); // idempotent, as for a read page
    expect(editor.isPageRemoved(2)).toBe(true);
    expect(editor.isPageRemoved(3)).toBe(false);
    expect(editor.pageCount).toBe(2);
    expect(editor.outputPageCount).toBe(3);
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "P2", "S1"]);
    editor.removePage(0);
    expect(editor.outputPageCount).toBe(2);
    expect(await texts(await editor.toBuffer())).toEqual(["P2", "S1"]);
  });

  it("a removed added page refuses to be fetched, rotated, split or drawn on", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    const added = editor.addPage();
    const handle = editor.getPage(1);
    editor.removePage(1);
    for (const call of [
      () => editor.getPage(1),
      () => editor.rotatePage(1, 90),
      () => handle.drawText("x", { x: 1, y: 1 }),
      () => added.drawText("x", { x: 1, y: 1 }),
      () => added.drawRect({ x: 0, y: 0, width: 5, height: 5 })
    ]) {
      expect(call).toThrow(PdfArgumentError);
    }
    expect(added.measureText("x")).toBeGreaterThan(0);
    await expect(editor.splitPages({ pageIndices: [1] })).rejects.toThrow(PdfArgumentError);
    expect(await texts(await editor.toBuffer())).toEqual(["P1"]);
  });

  it("rotatePage sets /Rotate on an added and a copied page", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.addPage();
    editor.copyPagesFrom(await pages("S1"));
    editor.rotatePage(1, 90).rotatePage(2, 270);
    expect(await rotations(await editor.toBuffer())).toEqual([0, 90, 270]);
    editor.rotatePage(2, 0);
    expect(await rotations(await editor.toBuffer())).toEqual([0, 90, 0]);
    // And in a part of splitPages, which writes the same page.
    const [part] = await editor.splitPages({ pageIndices: [1] });
    expect(await rotations(part)).toEqual([90]);
  });

  it("an index past every page throws PdfArgumentError", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    editor.addPage();
    expect(() => editor.getPage(3)).toThrow(PdfArgumentError);
    expect(() => editor.getPage(3)).toThrow(/^Pdf\.Editor#getPage: index 3 names no page \(0-2\)$/);
    expect(() => editor.removePage(-1)).toThrow(
      /^Pdf\.Editor#removePage: index -1 names no page \(0-2\)$/
    );
    expect(() => editor.isPageRemoved(2.5)).toThrow(PdfArgumentError);
  });

  it("removing every page, added ones included, refuses to write a page-less PDF", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.addPage();
    editor.removePage(0).removePage(1);
    expect(editor.outputPageCount).toBe(0);
    await expect(editor.toBuffer()).rejects.toThrow(/no pages/);
  });

  it("form field names on a removed added page are free again", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.addPage().addFormField({ type: "text", name: "f", rect: [0, 0, 10, 10] });
    editor.removePage(1);
    editor.getPage(0).addFormField({ type: "text", name: "f", rect: [0, 0, 10, 10] });
    const out = await Pdf.extract(await editor.toBuffer());
    expect(out.formFields.map(field => field.name)).toEqual(["f"]);
  });
});

describe("PdfEditor — added-page edits and incremental updates", () => {
  it("adding then removing the same page leaves nothing to rewrite: the update appends", async () => {
    const source = await pages("P1");
    const editor = Pdf.Editor.read(source);
    editor.addPage();
    editor.removePage(1);
    editor.getPage(0).drawText("x", { x: 10, y: 10 });
    const out = await editor.toBuffer({ incremental: "required" });
    // An appended update starts with the original bytes.
    expect(out.subarray(0, source.length)).toEqual(source);
  });

  it("rotating an added page is structural: incremental falls back, and a signed file is guarded", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.addPage();
    editor.rotatePage(1, 90);
    await expect(editor.toBuffer({ incremental: "required" })).rejects.toThrow(
      /page additions, removals or rotations/
    );

    const { certificate, privateKey } = await generateSelfSignedCertificate();
    const signer = new Pdf.Builder();
    signer.addPage();
    signer.sign({ certificate, privateKey });
    const signed = Pdf.Editor.read(await signer.toBuffer());
    signed.addPage();
    signed.removePage(1);
    // Nothing structural is left, so the update keeps the signature.
    const kept = await signed.toBuffer({ incremental: true });
    expect((await Pdf.verifySignatures(kept))[0].intact).toBe(true);
    signed.addPage();
    signed.rotatePage(2, 90);
    await expect(signed.toBuffer({ incremental: true })).rejects.toThrow(
      PdfSignatureInvalidationError
    );
  });
});
