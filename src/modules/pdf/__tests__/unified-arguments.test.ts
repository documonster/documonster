/**
 * One concept, one behaviour, wherever it appears (r4):
 *
 * - the shared `fonts` option is checked the same way by every entry point (R4P-C1);
 * - `splitPages({ pageIndices })` names pages, so an index that names none is refused
 *   (R4P-C2) — the deprecated positional form keeps skipping it;
 * - `isPageRemoved` takes an added page's index too, as `splitPages` does (R4P-S4);
 * - every PDF byte reader takes any `ArrayBufferView` and an `ArrayBuffer` (R4P-S5);
 * - `keywords` can be written, not only read (R4P-G2).
 */
import { Chart, Row, Workbook } from "@excel/index";
import { Pdf, PdfArgumentError, PdfStructureError } from "@pdf/index";
import { describe, expect, it } from "vitest";

async function rejection(run: () => unknown): Promise<unknown> {
  try {
    await run();
  } catch (err) {
    return err;
  }
  throw new Error("expected the call to throw");
}

async function pages(...labels: string[]): Promise<Uint8Array> {
  const b = new Pdf.Builder();
  for (const label of labels) {
    b.addPage().drawText(label, { x: 10, y: 10 });
  }
  return b.toBuffer();
}

async function texts(bytes: Uint8Array): Promise<string[]> {
  return (await Pdf.extract(bytes)).pages.map(page => page.text.trim());
}

describe("the fonts option is checked the same way everywhere (R4P-C1)", () => {
  const wb = Workbook.create();
  Row.add(Workbook.addWorksheet(wb, "S"), ["a"]);
  const entryPoints: Array<[string, (fonts: unknown) => unknown]> = [
    ["Pdf.fromTable", fonts => Pdf.fromTable([["a"]], { fonts } as never)],
    ["Pdf.fromExcel", fonts => Pdf.fromExcel(wb, { fonts } as never)],
    ["Pdf.fromDocx", fonts => Pdf.fromDocx({ body: [] } as never, { fonts } as never)],
    ["new Pdf.Builder", fonts => new Pdf.Builder({ fonts } as never)]
  ];

  it.each(entryPoints)("%s refuses a font configuration of the wrong shape", async (where, run) => {
    for (const [fonts, field] of [
      [42, "options.fonts"],
      [{}, "options.fonts.default"],
      [{ default: {} }, "options.fonts.default.regular"],
      [{ default: { regular: "x" } }, "options.fonts.default.regular"],
      [
        { default: { regular: { data: new Uint8Array(4), collectionIndex: -1 } } },
        "options.fonts.default.regular.collectionIndex"
      ],
      [{ default: { regular: new Uint8Array(4) }, families: {} }, "options.fonts.families"],
      [
        { default: { regular: new Uint8Array(4) }, families: [{ name: "", faces: {} }] },
        "options.fonts.families[0].name"
      ],
      [
        { default: { regular: new Uint8Array(4) }, fallbackFamilies: ["Nope"] },
        "options.fonts.fallbackFamilies[0]"
      ]
    ] as const) {
      const err = await rejection(() => run(fonts));
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(field);
      expect((err as Error).message.startsWith(`${where}: `)).toBe(true);
    }
  });

  it("names Builder#embedFonts' parameter config", () => {
    let error: unknown;
    try {
      new Pdf.Builder().embedFonts({ default: {} } as never);
    } catch (err) {
      error = err;
    }
    expect(error).toBeInstanceOf(PdfArgumentError);
    expect((error as PdfArgumentError).field).toBe("config.default.regular");
  });

  it("refuses font and fonts together before anything is changed", async () => {
    const err = await rejection(() =>
      Pdf.fromTable([["a"]], {
        font: new Uint8Array(4),
        fonts: { default: { regular: new Uint8Array(4) } }
      })
    );
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect((err as PdfArgumentError).field).toBe("options.fonts");
  });
});

describe("splitPages names pages (R4P-C2)", () => {
  it("refuses an index out of range or naming a removed page, before writing any part", async () => {
    const editor = Pdf.Editor.read(await pages("A", "B", "C"));
    editor.removePage(1);
    for (const [pageIndices, at, pattern] of [
      [[0, 9], 1, /must be the 0-based index of a page of the document \(0 to 2/],
      [[-1], 0, /must be the 0-based index/],
      [[0.5], 0, /must be the 0-based index/],
      [[2, 1], 1, /names page 1, which has been removed/]
    ] as const) {
      const err = await rejection(() => editor.splitPages({ pageIndices }));
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(`options.pageIndices[${at}]`);
      expect((err as Error).message).toMatch(/^Pdf\.Editor#splitPages: /);
      expect((err as Error).message).toMatch(pattern);
    }
    // One part per index, in the order given — so the result can be read against the request.
    const parts = await editor.splitPages({ pageIndices: [2, 0] });
    expect(await Promise.all(parts.map(async part => (await texts(part))[0]))).toEqual(["C", "A"]);
    // Omitting the list splits every page that is still there.
    expect(await editor.splitPages()).toHaveLength(2);
  });

  it("numbers an added page after the pages the document was read with", async () => {
    const editor = Pdf.Editor.read(await pages("A"));
    editor.addPage().drawText("NEW", { x: 10, y: 10 });
    const [part] = await editor.splitPages({ pageIndices: [1] });
    expect(await texts(part)).toEqual(["NEW"]);
  });

  it("the deprecated positional form keeps skipping, as it always did", async () => {
    const editor = Pdf.Editor.read(await pages("A", "B"));
    editor.removePage(1);
    expect(await editor.splitPages([0, 1, 9, -1])).toHaveLength(1);
    await expect(editor.splitPages([0.5])).rejects.toThrow(PdfStructureError);
  });
});

describe("isPageRemoved takes both index ranges (R4P-S4)", () => {
  it("answers for an added page, and every index call takes added pages", async () => {
    const editor = Pdf.Editor.read(await pages("A", "B"));
    expect(() => editor.isPageRemoved(2)).toThrow(PdfArgumentError);
    editor.addPage();
    editor.copyPagesFrom(await pages("C"));
    expect(editor.isPageRemoved(2)).toBe(false);
    expect(editor.isPageRemoved(3)).toBe(false);
    expect(() => editor.isPageRemoved(4)).toThrow(PdfArgumentError);
    expect(() => editor.isPageRemoved(2.5)).toThrow(PdfArgumentError);
    expect(editor.getPage(3)).toBeInstanceOf(Pdf.EditorPage);
    editor.removePage(2);
    expect(editor.isPageRemoved(2)).toBe(true);
  });
});

describe("PDF byte readers take any ArrayBufferView (R4P-S5)", () => {
  it("reads a DataView, a Buffer-like view at an offset, and an ArrayBuffer", async () => {
    const bytes = await pages("hello");
    const padded = new Uint8Array(bytes.length + 8);
    padded.set(bytes, 4);
    const views: unknown[] = [
      new DataView(padded.buffer, 4, bytes.length),
      new Uint8Array(padded.buffer, 4, bytes.length),
      new Int8Array(padded.buffer, 4, bytes.length),
      bytes.slice().buffer
    ];
    for (const view of views) {
      expect((await Pdf.extract(view as never)).text).toContain("hello");
      expect(Pdf.Editor.read(view as never).pageCount).toBe(1);
      expect(await Pdf.verifySignatures(view as never)).toEqual([]);
      const editor = Pdf.Editor.read(bytes);
      editor.copyPagesFrom(view as never);
      expect(await texts(await editor.toBuffer())).toEqual(["hello", "hello"]);
      expect((await Pdf.Signature.verify(view as never, "00", [0, 1, 2, 3])).intact).toBe(false);
    }
    expect(Pdf.Signature.parseAsn1(new DataView(new Uint8Array([0x05, 0x00]).buffer)).tag).toBe(5);
  });

  it("still refuses what is not bytes, naming the parameter", async () => {
    for (const [run, field] of [
      [() => Pdf.extract("x" as never), "data"],
      [() => Pdf.Editor.read([1, 2] as never), "data"],
      [() => Pdf.Signature.parseAsn1({} as never), "data"],
      [() => Pdf.Signature.verify(null as never, "00", [0, 1, 2, 3]), "pdfData"],
      [() => Pdf.Signature.fillPlaceholder(42 as never, {} as never), "pdfBytes"]
    ] as const) {
      const err = await rejection(run);
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(field);
    }
  });
});

describe("Pdf.fromChart refuses a region-map topology it cannot read", () => {
  it("as a PdfArgumentError naming the path, not a silent preview", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["a", 1]);
    const chart = Chart.place(ws, "D1:H10", {
      type: "bar",
      series: [{ values: "S!$B$1:$B$1", categories: "S!$A$1:$A$1" }]
    });
    const err = await rejection(() =>
      Pdf.fromChart(chart, {
        regionMap: { topology: { objects: { countries: {} }, arcs: 3 }, objectName: "countries" }
      })
    );
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect((err as PdfArgumentError).field).toBe("options.regionMap.topology.arcs");
    expect((err as Error).message).toBe(
      "Pdf.fromChart: options.regionMap.topology.arcs must be an array of arcs, got 3"
    );
  });
});

describe("keywords are written, not only read (R4P-G2)", () => {
  it("Builder#setMetadata, Editor#setMetadata and the converters write /Keywords", async () => {
    const b = new Pdf.Builder();
    b.setMetadata({ keywords: "alpha, beta" });
    b.addPage();
    const built = await b.toBuffer();
    expect((await Pdf.extract(built)).metadata.keywords).toBe("alpha, beta");

    const editor = Pdf.Editor.read(built);
    editor.setMetadata({ keywords: "gamma" });
    expect((await Pdf.extract(await editor.toBuffer())).metadata.keywords).toBe("gamma");
    const removed = Pdf.Editor.read(built);
    removed.setMetadata({ keywords: null });
    expect(
      (await Pdf.extract(await removed.toBuffer({ incremental: true }))).metadata.keywords
    ).toBe("");

    const table = await Pdf.fromTable([["a"]], { keywords: "table" });
    expect((await Pdf.extract(table)).metadata.keywords).toBe("table");
  });

  it("writes pdf:Keywords into the XMP packet of a PDF/A document, and edits it there", async () => {
    const b = new Pdf.Builder();
    b.setPdfACompliance();
    b.setMetadata({ keywords: "archive" });
    b.addPage();
    const built = await b.toBuffer();
    expect(new TextDecoder("latin1").decode(built)).toContain(
      "<pdf:Keywords>archive</pdf:Keywords>"
    );
    const editor = Pdf.Editor.read(built);
    editor.setMetadata({ keywords: "edited" });
    const out = new TextDecoder("latin1").decode(await editor.toBuffer({ incremental: true }));
    expect(out).toContain("<pdf:Keywords>edited</pdf:Keywords>");
  });

  it("checks keywords like every other string field", () => {
    let error: unknown;
    try {
      new Pdf.Builder().setMetadata({ keywords: 5 as never });
    } catch (err) {
      error = err;
    }
    expect(error).toBeInstanceOf(PdfArgumentError);
    expect((error as PdfArgumentError).field).toBe("metadata.keywords");
  });
});
