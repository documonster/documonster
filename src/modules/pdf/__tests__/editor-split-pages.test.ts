/**
 * `splitPages()` writes each part on its own path rather than as a full rewrite of the whole
 * document with every other page removed. These tests pin that the parts are still exactly
 * that rewrite (see `split-pages.helpers.ts`), across the structures whose handling differs —
 * shared and inherited resources, form fields spanning pages, malformed field trees — and that
 * the work per part no longer grows with the document. The fixture files are covered in
 * `editor-split-pages.node.test.ts`.
 */
import { PdfDocument } from "@pdf/reader/pdf-document";
import { createRng } from "@test/rng";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";

import { Pdf } from "../index";
import { buildRawPdf, twoPageFormBodies } from "./raw-pdf.helpers";
import { expectSplitMatchesReference } from "./split-pages.helpers";
import type { Editor } from "./split-pages.helpers";

beforeAll(() => {
  // A rewrite stamps /CreationDate with the current time when the original has none.
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date("2026-01-02T03:04:05Z"));
});
afterAll(() => {
  vi.useRealTimers();
});

const CONTENT = "<< /Length 0 >>\nstream\n\nendstream";
const page = (extra: string, contents = 99): string =>
  `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents ${contents} 0 R ${extra} >>`;
const widget = (pageObj: number, extra: string): string =>
  `<< /Type /Annot /Subtype /Widget /Rect [60 600 300 622] /P ${pageObj} 0 R /DA (/Helv 12 Tf 0 g) ${extra} >>`;

/** `bodies` numbered from 1; object 99 (an empty content stream) is padded in. */
function raw(bodies: Record<number, string>): Uint8Array {
  const max = Math.max(99, ...Object.keys(bodies).map(Number));
  const list: string[] = [];
  for (let i = 1; i <= max; i++) {
    list.push(i === 99 ? CONTENT : (bodies[i] ?? "null"));
  }
  return buildRawPdf(list);
}

/**
 * `n` pages, each with a text field of its own (a parent field holding one widget, listed in
 * the calculation order `/CO`), one widget of a radio group spanning every page, and a link;
 * the group also has a widgetless kid, and the form a widgetless field.
 */
function formPages(n: number): Uint8Array {
  const bodies: Record<number, string> = {};
  const pageNum = (i: number) => 10 + i * 5;
  const refs = (offset: number) =>
    Array.from({ length: n }, (_, i) => `${pageNum(i) + offset} 0 R`).join(" ");
  bodies[1] = `<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [3 0 R 4 0 R ${refs(4)}] /CO [${refs(4)}] /DA (/Helv 0 Tf 0 g) >> >>`;
  bodies[2] = `<< /Type /Pages /Kids [${refs(0)}] /Count ${n} >>`;
  bodies[3] = `<< /FT /Btn /Ff 49152 /T (group) /V /p0 /Kids [${refs(2)} 5 0 R] >>`;
  bodies[4] = "<< /FT /Tx /T (hidden) /V (kept) >>";
  bodies[5] = "<< /T (detail) /V (no widget) /Parent 3 0 R >>";
  for (let i = 0; i < n; i++) {
    const p = pageNum(i);
    bodies[p] = page(`/Resources << >> /Annots [${p + 1} 0 R ${p + 2} 0 R ${p + 3} 0 R]`);
    bodies[p + 1] = widget(p, `/Parent ${p + 4} 0 R`);
    bodies[p + 2] = widget(p, `/Parent 3 0 R /AS /Off`);
    bodies[p + 3] =
      `<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] /Dest [${pageNum((i + 1) % n)} 0 R /Fit] >>`;
    bodies[p + 4] = `<< /FT /Tx /T (field${i}) /V (value ${i}) /Kids [${p + 1} 0 R] >>`;
  }
  return raw(bodies);
}

describe("splitPages() writes what removing every other page writes", () => {
  it("a form whose fields span pages, with a field no page shows", async () => {
    await expectSplitMatchesReference(buildRawPdf(twoPageFormBodies()));
  });

  it("a radio group spanning every page, links between pages", async () => {
    await expectSplitMatchesReference(formPages(4));
  });

  it("a single page", async () => {
    await expectSplitMatchesReference(formPages(1));
  });

  it("a subset, out of range and repeated indices", async () => {
    const bytes = formPages(4);
    const { parts } = (await expectSplitMatchesReference(bytes, {
      indices: [2, -1, 0, 9, 2]
    })) as { parts: Uint8Array[] };
    expect(parts).toHaveLength(3);
  });

  it("resources inherited from the page tree and shared between pages", async () => {
    await expectSplitMatchesReference(
      raw({
        1: "<< /Type /Catalog /Pages 2 0 R /Metadata 9 0 R >>",
        2: "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 /Resources << /Font << /F1 6 0 R >> /XObject << /X1 7 0 R >> >> /Rotate 90 >>",
        3: page(""),
        4: page("/Resources << /Font << /F1 6 0 R /F2 8 0 R >> >>"),
        5: page("/CropBox [10 10 500 500] /UserUnit 2"),
        6: "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        7: "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /Font << /F1 6 0 R >> >> /Length 0 >>\nstream\n\nendstream",
        8: "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>",
        9: "<< /Type /Metadata /Subtype /XML /Length 5 >>\nstream\n<x/>\n\nendstream"
      })
    );
  });

  it("the same annotation on two pages, a direct annotation, and a missing one", async () => {
    await expectSplitMatchesReference(
      raw({
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>",
        2: "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>",
        3: page(
          "/Annots [6 0 R << /Type /Annot /Subtype /Square /Rect [0 0 5 5] /P 4 0 R >> 50 0 R]"
        ),
        4: page("/Annots [6 0 R]"),
        5: page("/Annots [7 0 R]"),
        6: widget(3, "/FT /Tx /T (shared)"),
        7: "<< /Type /Annot /Subtype /Text /Rect [0 0 5 5] /Popup 8 0 R >>",
        8: "<< /Type /Annot /Subtype /Popup /Parent 7 0 R /Rect [0 0 5 5] >>"
      })
    );
  });

  it("the same page object twice in the page tree, and a page written inline", async () => {
    await expectSplitMatchesReference(
      raw({
        1: "<< /Type /Catalog /Pages 2 0 R >>",
        2: `<< /Type /Pages /Kids [3 0 R 3 0 R ${page("/Rotate 180")} 4 0 R] /Count 4 >>`,
        3: page("/Annots [5 0 R]"),
        4: page(""),
        5: "<< /Type /Annot /Subtype /Square /Rect [0 0 5 5] /P 3 0 R >>"
      })
    );
  });

  describe("field trees that are not trees", () => {
    const pages = {
      2: "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>",
      3: page("/Annots [10 0 R]"),
      4: page("/Annots [11 0 R]"),
      5: page("/Annots [12 0 R]"),
      12: widget(5, "/FT /Tx /T (third)")
    };

    it("a node shared by two parents", async () => {
      const bytes = raw({
        ...pages,
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R 7 0 R 12 0 R] >> >>",
        6: "<< /T (a) /Kids [8 0 R 10 0 R] >>",
        7: "<< /T (b) /Kids [8 0 R] >>",
        8: "<< /T (shared) /Kids [11 0 R] >>",
        10: widget(3, "/Parent 6 0 R"),
        11: widget(4, "/Parent 8 0 R")
      });
      await expectSplitMatchesReference(bytes);
    });

    it("a /Kids cycle", async () => {
      const bytes = raw({
        ...pages,
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R 12 0 R] >> >>",
        6: "<< /T (a) /Kids [7 0 R] >>",
        7: "<< /T (b) /Kids [6 0 R 10 0 R 11 0 R] /Parent 6 0 R >>",
        10: widget(3, "/Parent 7 0 R"),
        11: widget(4, "/Parent 7 0 R")
      });
      await expectSplitMatchesReference(bytes);
    });

    it("a field listed twice", async () => {
      const bytes = raw({
        ...pages,
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [12 0 R 6 0 R 12 0 R] >> >>",
        6: "<< /T (pair) /Kids [10 0 R 11 0 R] >>",
        10: widget(3, "/Parent 6 0 R"),
        11: widget(4, "/Parent 6 0 R")
      });
      await expectSplitMatchesReference(bytes);
    });

    it("a widget with kids of its own", async () => {
      const bytes = raw({
        ...pages,
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [10 0 R 12 0 R] >> >>",
        10: widget(3, "/FT /Tx /T (outer) /Kids [11 0 R 13 0 R]"),
        11: widget(4, "/Parent 10 0 R"),
        13: "<< /T (inner) /Parent 10 0 R /Kids [14 0 R] >>",
        14: "<< /T (leaf) >>"
      });
      await expectSplitMatchesReference(bytes);
    });

    it("an inline field and a field whose /Kids is empty", async () => {
      const bytes = raw({
        ...pages,
        1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [<< /T (inline) /Kids [10 0 R] >> 6 0 R 7 0 R 12 0 R] >> >>",
        6: "<< /T (empty) /Kids [] >>",
        7: "<< /T (deep) /Kids [8 0 R] >>",
        8: "<< /T (mid) /Kids [11 0 R << /T (direct) >>] /Parent 7 0 R >>",
        10: widget(3, "/T (x)"),
        11: widget(4, "/Parent 8 0 R")
      });
      await expectSplitMatchesReference(bytes);
    });
  });

  it("a document built by this library: text, annotations, fields, bookmarks", async () => {
    const builder = new Pdf.Builder();
    builder.setMetadata({ title: "Split", author: "Tests" });
    for (let i = 0; i < 3; i++) {
      const p = builder.addPage();
      p.drawText(`Page ${i + 1} 页`, { x: 72, y: 700, fontSize: 14 });
      p.drawRect({ x: 72, y: 600, width: 100, height: 40, fill: { r: 0.2, g: 0.4, b: 0.6 } });
      p.addAnnotation({ type: "Text", rect: [10, 10, 30, 30], contents: `Note ${i}` });
      p.addFormField({ type: "text", name: `name${i}`, rect: [72, 500, 300, 520], value: "v" });
      p.addFormField({
        type: "radio",
        name: `choice${i}`,
        buttons: [
          { rect: [72, 450, 90, 468], value: "a" },
          { rect: [100, 450, 118, 468], value: "b" }
        ],
        selected: "b"
      });
      builder.addBookmark(`Chapter ${i + 1}`, i);
    }
    await expectSplitMatchesReference(await builder.toBuffer());
  });

  it("an encrypted document built by this library", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().drawText("One", { x: 72, y: 700, fontSize: 12 });
    builder.addPage().drawText("Two", { x: 72, y: 700, fontSize: 12 });
    builder.setEncryption({ userPassword: "user", ownerPassword: "owner" });
    const bytes = await builder.toBuffer();
    await expectSplitMatchesReference(bytes, { password: "user" });
    await expectSplitMatchesReference(bytes, { password: "owner" });
  });

  it("a signed document is refused the same way", async () => {
    const bytes = raw({
      1: "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] /SigFlags 3 >> >>",
      2: "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      3: page("/Annots [5 0 R]"),
      4: page(""),
      5: widget(3, "/FT /Sig /T (sig) /V 6 0 R"),
      6: "<< /Type /Sig /ByteRange [0 0 0 0] /Contents <00> >>"
    });
    const outcome = await expectSplitMatchesReference(bytes);
    expect(outcome).toMatchObject({
      error: expect.stringMatching(/^PdfSignatureInvalidationError/)
    });
    const parts = await Pdf.Editor.read(bytes).splitPages({
      invalidateSignatures: true
    });
    expect(parts).toHaveLength(2);
  });

  // Mutated copies of the fixtures above: whatever each one does — parts or a typed error —
  // the dedicated path must do the same.
  describe("mutated inputs", () => {
    const sources = [buildRawPdf(twoPageFormBodies()), formPages(3)];
    for (const seed of [1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 233]) {
      it(`seed ${seed}`, async () => {
        const rng = createRng(seed);
        const bytes = rng.pick(sources).slice();
        const edits = rng.int(1, 6);
        for (let i = 0; i < edits; i++) {
          const at = rng.int(0, bytes.length - 1);
          bytes[at] = rng.pick([0x20, 0x30, 0x35, 0x52, 0x5b, 0x5d, 0x2f, rng.int(0, 255)]);
        }
        await expectSplitMatchesReference(bytes);
      });
    }
  });
});

describe("splitPages() does work proportional to the document, not to its square", () => {
  /** Calls into the parsed document made while splitting an already-read document. */
  async function lookupsToSplit(bytes: Uint8Array): Promise<number> {
    const editor = Pdf.Editor.read(bytes);
    const spies = [
      vi.spyOn(PdfDocument.prototype, "deref"),
      vi.spyOn(PdfDocument.prototype, "derefDict"),
      vi.spyOn(PdfDocument.prototype, "resolvePageBox"),
      vi.spyOn(PdfDocument.prototype, "resolvePageResources")
    ];
    try {
      const parts = await editor.splitPages();
      expect(parts).toHaveLength(editor.pageCount);
      return spies.reduce((sum, spy) => sum + spy.mock.calls.length, 0);
    } finally {
      for (const spy of spies) {
        spy.mockRestore();
      }
    }
  }

  it("four times the pages costs about four times the lookups", async () => {
    const small = await lookupsToSplit(formPages(30));
    const large = await lookupsToSplit(formPages(120));
    // Linear work gives 4; one full rewrite per part gives about 16.
    expect(large / small).toBeLessThan(5);
  });

  it("so does a document with a bookmark for every page", async () => {
    const bookmarked = async (count: number): Promise<Uint8Array> => {
      const builder = new Pdf.Builder();
      for (let i = 0; i < count; i++) {
        builder.addPage().drawText(`P${i}`, { x: 72, y: 700 });
        builder.addBookmark(`Page ${i}`, i);
      }
      return builder.toBuffer();
    };
    const small = await lookupsToSplit(await bookmarked(30));
    const large = await lookupsToSplit(await bookmarked(120));
    // Reading the outline per part, rather than once, gives about 16.
    expect(large / small).toBeLessThan(5);
  });
});

describe("splitPages() splits the document as it would be written now", () => {
  // The parts used to be the pages of the file as read, whatever had been done to the editor:
  // a removed page came back as a part, and drawing, rotating, filling in a field or adding a
  // page changed none of them.
  const labelled = async (labels: string[]): Promise<Uint8Array> => {
    const builder = new Pdf.Builder();
    for (const label of labels) {
      builder.addPage().drawText(label, { x: 50, y: 700, fontSize: 20 });
    }
    return builder.toBuffer();
  };
  const texts = (parts: Uint8Array[]): Promise<string[]> =>
    Promise.all(
      parts.map(async part => {
        const { pages } = await Pdf.extract(part);
        return pages.map(p => p.text.replace(/\s+/g, "") || "∅").join(",");
      })
    );

  it("leaves out a removed page", async () => {
    const editor = Pdf.Editor.read(await labelled(["P0", "P1", "P2", "P3", "P4"]));
    editor.removePage(0);
    editor.removePage(3);
    expect(await texts(await editor.splitPages())).toEqual(["P1", "P2", "P4"]);
  });

  it("includes added and copied pages, after the original ones, in the order they were added", async () => {
    const editor = Pdf.Editor.read(await labelled(["P0", "P1", "P2"]));
    editor.addPage().drawText("A1", { x: 50, y: 700, fontSize: 20 });
    editor.copyPagesFrom(await labelled(["O0", "O1"]), { pageIndices: [1, 0] });
    editor.addPage().drawText("A2", { x: 50, y: 700, fontSize: 20 });
    editor.removePage(1);
    expect(await texts(await editor.splitPages())).toEqual(["P0", "P2", "A1", "O1", "O0", "A2"]);
  });

  it("indices name pages as read, and added pages after them; removed pages are skipped", async () => {
    const editor = Pdf.Editor.read(await labelled(["P0", "P1", "P2"]));
    editor.addPage().drawText("A1", { x: 50, y: 700, fontSize: 20 });
    editor.copyPagesFrom(await labelled(["O0"]));
    editor.removePage(0);
    expect(await texts(await editor.splitPages([4, 0, 2, 3, 5, -1]))).toEqual(["O0", "P2", "A1"]);
  });

  it("a document whose every page was removed and none added splits into nothing", async () => {
    const editor = Pdf.Editor.read(await labelled(["P0", "P1"]));
    editor.removePage(0).removePage(1);
    expect(await editor.splitPages()).toEqual([]);
  });

  it("applies overlays, rotations and form values: each part is that page of toBuffer()", async () => {
    const edit = (editor: Editor): void => {
      editor.getPage(0).drawText("drawn", { x: 72, y: 72, fontSize: 9 });
      editor
        .getPage(2)
        .drawRect({ x: 10, y: 10, width: 50, height: 20, fill: { r: 1, g: 0, b: 0 } });
      editor.rotatePage(1, 90);
      editor.removePage(3);
      editor.setFormField("field0", "edited");
      editor.setFormField("group", "p1");
    };
    const { parts } = (await expectSplitMatchesReference(formPages(4), { edit })) as {
      parts: Uint8Array[];
    };
    expect(parts).toHaveLength(3);
    expect((await Pdf.extract(parts[0])).pages[0].text).toContain("drawn");
    expect(new TextDecoder("latin1").decode(parts[1])).toMatch(/\/Rotate 90\b/);
  });

  it("a filled-in value is in the part whose page shows the field", async () => {
    const edit = (editor: Editor): void => {
      editor.setFormField("name", "edited");
    };
    const { parts } = (await expectSplitMatchesReference(buildRawPdf(twoPageFormBodies()), {
      edit
    })) as { parts: Uint8Array[] };
    const values = async (part: Uint8Array) =>
      Object.fromEntries(
        ((await Pdf.extract(part)).formFields ?? []).map(field => [field.name, field.value])
      );
    expect(await values(parts[0])).toMatchObject({ name: "edited", hidden: "kept" });
    expect(await values(parts[1])).not.toHaveProperty("name");
  });

  it("writes the edits made before it was called, not after", async () => {
    const editor = Pdf.Editor.read(await labelled(["P0", "P1"]));
    editor.removePage(0);
    const split = editor.splitPages();
    editor.addPage().drawText("late", { x: 50, y: 700, fontSize: 20 });
    expect(await texts(await split)).toEqual(["P1"]);
  });
});
