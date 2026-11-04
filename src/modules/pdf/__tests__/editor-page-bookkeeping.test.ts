/**
 * Page bookkeeping in the editor. Every index the editor takes — `getPage`, `removePage`,
 * `rotatePage` — names a page of the document as it was read, or (numbered after those, in the
 * order added) a page added with `addPage` or `copyPagesFrom`, and does not shift when another
 * page is removed. See `editor-page-model.test.ts` for the added pages.
 */
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError, PdfStructureError } from "../index";

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

describe("PdfEditor — removing pages", () => {
  it("indices name original pages: removing from the middle repeatedly", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3", "P4", "P5"));
    editor.removePage(1).removePage(3);
    expect(editor.pageCount).toBe(5);
    expect([0, 1, 2, 3, 4].map(i => editor.isPageRemoved(i))).toEqual([
      false,
      true,
      false,
      true,
      false
    ]);
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "P3", "P5"]);
  });

  it("removal is idempotent: removing a removed page again removes it once", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
    editor.removePage(0).removePage(0).removePage(0);
    expect(editor.pageCount).toBe(3);
    expect(editor.isPageRemoved(0)).toBe(true);
    expect(editor.isPageRemoved(1)).toBe(false);
    expect(await texts(await editor.toBuffer())).toEqual(["P2", "P3"]);
  });

  it("a list of pages with duplicates removes each page once", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3", "P4"));
    for (const index of [2, 0, 2, 0]) {
      editor.removePage(index);
    }
    expect(await texts(await editor.toBuffer())).toEqual(["P2", "P4"]);
  });

  it("an index outside the document throws PdfStructureError, as it always did", async () => {
    const editor = Pdf.Editor.read(await pages("P1"));
    for (const index of [1, -1, Infinity, -0.5]) {
      expect(() => editor.isPageRemoved(index)).toThrow(PdfStructureError);
      expect(() => editor.removePage(index)).toThrow(PdfStructureError);
      expect(() => editor.getPage(index)).toThrow(PdfStructureError);
      expect(() => editor.rotatePage(index, 90)).toThrow(PdfStructureError);
    }
  });

  it("a fractional or NaN index throws PdfStructureError instead of matching no page", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    for (const index of [0.5, NaN]) {
      expect(() => editor.isPageRemoved(index)).toThrow(PdfStructureError);
      expect(() => editor.removePage(index)).toThrow(PdfStructureError);
      expect(() => editor.getPage(index)).toThrow(PdfStructureError);
      expect(() => editor.rotatePage(index, 90)).toThrow(PdfStructureError);
    }
    // The rejected calls recorded nothing.
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "P2"]);
  });

  it("the message names what was passed when it is not a number", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
    const bad = (index: unknown) => editor.removePage(index as number);
    const fails = (call: () => unknown, message: string) => {
      expect(call).toThrow(PdfStructureError);
      expect(call).toThrow(message);
    };
    fails(() => bad("1"), 'Page index must be an integer, got string "1"');
    fails(() => bad("9"), 'Page index must be an integer, got string "9"');
    fails(() => bad(null), "Page index must be an integer, got null");
    fails(() => bad(undefined), "Page index must be an integer, got undefined");
    fails(() => bad(1n), "Page index must be an integer, got bigint 1");
    fails(() => bad([0, 2]), "Page index must be an integer, got an array (2 items)");
    fails(() => bad({}), "Page index must be an integer, got object [object Object]");
    fails(() => editor.getPage(1.5), "Page index must be an integer, got 1.5");
    fails(() => editor.getPage(NaN), "Page index must be an integer, got NaN");
    fails(() => editor.isPageRemoved(2.5), "Page index must be an integer, got 2.5");
    fails(() => editor.rotatePage(0.5, 90), "Page index must be an integer, got 0.5");
    fails(() => editor.getPage(3), "Page index 3 out of range (0-2)");
    await expect(editor.splitPages(["1" as unknown as number])).rejects.toThrow(
      new PdfStructureError('Pdf.Editor#splitPages: page index must be an integer, got string "1"')
    );
    await expect(editor.splitPages([0.5])).rejects.toThrow(
      new PdfStructureError("Pdf.Editor#splitPages: page index must be an integer, got 0.5")
    );
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "P2", "P3"]);
  });

  it("an index past the original pages names an added page once one is added", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    expect(() => editor.removePage(2)).toThrow(/^Page index 2 out of range \(0-1\)$/);
    editor.addPage();
    // `pageCount` keeps meaning the pages the document was read with.
    expect(editor.pageCount).toBe(2);
    expect(editor.outputPageCount).toBe(3);
    editor.removePage(2);
    expect(editor.outputPageCount).toBe(2);
    expect(() => editor.removePage(3)).toThrow(/^Page index 3 out of range \(0-2\)$/);
    expect(() => editor.removePage(-1)).toThrow(/^Page index -1 out of range \(0-2\)$/);
  });

  it.each([
    ["toBuffer()", (editor: Pdf.Editor) => editor.toBuffer()],
    [
      "toBuffer({ incremental: true })",
      (editor: Pdf.Editor) => editor.toBuffer({ incremental: true })
    ]
  ])(
    "removing every page and adding none: %s refuses to write a page-less PDF",
    async (_, write) => {
      const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
      editor.removePage(0).removePage(1).removePage(2);
      await expect(write(editor)).rejects.toThrow(PdfStructureError);
      await expect(write(editor)).rejects.toThrow(/no pages/);
    }
  );

  it("removing every page then adding one writes only the new page", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
    editor.removePage(0).removePage(1).removePage(2);
    editor.addPage().drawText("NEW", { x: 72, y: 700, fontSize: 12 });
    expect(await texts(await editor.toBuffer())).toEqual(["NEW"]);
  });

  it("removing every page then copying one in writes only the copy", async () => {
    const source = await pages("S1", "S2");
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    editor
      .removePage(0)
      .removePage(1)
      .copyPagesFrom(source, { pageIndices: [1] });
    expect(await texts(await editor.toBuffer())).toEqual(["S2"]);
  });

  it("an overlay or rotation on a removed page goes nowhere else", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
    // Made before the removal; after it, the page refuses both (editor-call-time-checks).
    editor.getPage(1).drawText("LOST", { x: 72, y: 600, fontSize: 12 });
    editor.rotatePage(1, 90);
    editor.removePage(1);
    expect(() => editor.getPage(1)).toThrow(PdfArgumentError);
    expect(() => editor.rotatePage(1, 90)).toThrow(PdfArgumentError);
    editor.getPage(2).drawText("KEPT", { x: 72, y: 600, fontSize: 12 });
    const out = await editor.toBuffer();
    const result = await texts(out);
    expect(result).toHaveLength(2);
    expect(result[0]).toBe("P1");
    expect(result[1]).toContain("P3");
    expect(result[1]).toContain("KEPT");
    expect(result.join(" ")).not.toContain("LOST");
    expect(new TextDecoder("latin1").decode(out)).not.toContain("/Rotate 90");
  });

  it("a value set on a field of a removed page does not bring the field back", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().addFormField({ type: "text", name: "a", rect: [60, 690, 300, 712] });
    builder.addPage().addFormField({ type: "text", name: "b", rect: [60, 690, 300, 712] });
    const editor = Pdf.Editor.read(await builder.toBuffer());
    editor.removePage(1).setFormFields({ a: "kept", b: "gone" });
    const result = await Pdf.extract(await editor.toBuffer(), {
      extractFormFields: true,
      extractText: false
    });
    expect(result.pages).toHaveLength(1);
    expect(result.formFields.map(field => [field.name, field.value])).toEqual([["a", "kept"]]);
  });
});

describe("PdfEditor — appended pages keep the order they were added in", () => {
  it("addPage then copyPagesFrom", async () => {
    const source = await pages("S1", "S2");
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.addPage().drawText("NEW", { x: 72, y: 700, fontSize: 12 });
    editor.copyPagesFrom(source, { pageIndices: [0] });
    editor.addPage().drawText("LAST", { x: 72, y: 700, fontSize: 12 });
    editor.copyPagesFrom(source, { pageIndices: [1] });
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "NEW", "S1", "LAST", "S2"]);
    expect(await texts(await editor.toBuffer({ incremental: true }))).toEqual([
      "P1",
      "NEW",
      "S1",
      "LAST",
      "S2"
    ]);
  });

  it("the deprecated positional copyPagesFrom skips a source index out of range, as it always did", async () => {
    const source = await pages("S1", "S2");
    const editor = Pdf.Editor.read(await pages("P1"));
    editor.copyPagesFrom(source, [7, 1, -1, 0, 2]);
    expect(await texts(await editor.toBuffer())).toEqual(["P1", "S2", "S1"]);
  });

  it("copyPagesFrom({ pageIndices }) refuses an index that names no source page, copying nothing", async () => {
    const source = await pages("S1", "S2");
    const editor = Pdf.Editor.read(await pages("P1"));
    for (const [bad, at] of [
      [7, 1],
      [-1, 1],
      [2, 1],
      [0.5, 1],
      [NaN, 1],
      [null, 1]
    ] as const) {
      let error: unknown;
      try {
        editor.copyPagesFrom(source, { pageIndices: [0, bad as number] });
      } catch (err) {
        error = err;
      }
      expect(error).toBeInstanceOf(PdfArgumentError);
      expect((error as PdfArgumentError).field).toBe(`options.pageIndices[${at}]`);
      expect((error as Error).message).toMatch(
        /^Pdf\.Editor#copyPagesFrom: options\.pageIndices\[1\] must be the 0-based index of a page of the source \(0 to 1\)/
      );
    }
    expect(await texts(await editor.toBuffer())).toEqual(["P1"]);
  });

  it("the deprecated positional copyPagesFrom rejects an index in range that is not an integer", async () => {
    const source = await pages("S1", "S2");
    const editor = Pdf.Editor.read(await pages("P1"));
    for (const bad of [0.5, NaN, null]) {
      const call = () => editor.copyPagesFrom(source, [0, bad as number]);
      expect(call).toThrow(PdfStructureError);
      expect(call).toThrow(
        `Pdf.Editor#copyPagesFrom: page index must be an integer, got ${String(bad)}`
      );
    }
    expect(await texts(await editor.toBuffer())).toEqual(["P1"]);
  });
});

describe("PdfEditor — a write reflects the pages as they were when it was called", () => {
  it("removals and additions made after a write was called do not reach it", async () => {
    const source = await pages("S1");
    const editor = Pdf.Editor.read(await pages("P1", "P2", "P3"));
    editor.removePage(0);
    const first = editor.toBuffer();
    editor.removePage(2);
    editor.copyPagesFrom(source);
    const second = editor.toBuffer();
    editor.removePage(1);
    editor.addPage().drawText("NEW", { x: 72, y: 700, fontSize: 12 });
    const third = editor.toBuffer();
    expect(await texts(await first)).toEqual(["P2", "P3"]);
    expect(await texts(await second)).toEqual(["P2", "S1"]);
    expect(await texts(await third)).toEqual(["S1", "NEW"]);
  });

  it("a write called after every page was removed fails alone", async () => {
    const editor = Pdf.Editor.read(await pages("P1", "P2"));
    editor.removePage(0);
    const before = editor.toBuffer();
    editor.removePage(1);
    const empty = editor.toBuffer();
    editor.addPage().drawText("NEW", { x: 72, y: 700, fontSize: 12 });
    const after = editor.toBuffer();
    expect(await texts(await before)).toEqual(["P2"]);
    await expect(empty).rejects.toThrow(PdfStructureError);
    expect(await texts(await after)).toEqual(["NEW"]);
  });
});
