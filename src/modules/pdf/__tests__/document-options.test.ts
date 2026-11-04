/**
 * One set of document-level options for every converter (r6 D1/D2): `Pdf.fromTable`,
 * `Pdf.fromExcel`, `Pdf.fromDocx` and `Pdf.fromChart` all take `PdfDocumentOptions` — metadata,
 * fixed dates for reproducible bytes, encryption, and the font settings — with the same
 * names and the same checks. `Pdf.fromDocx` falls back to the Word core properties, keywords
 * included; `Pdf.fromChart`'s `title` is the PDF's metadata on every path, never the chart's.
 */
import { Chart, Row, Workbook } from "@excel/index";
import { Pdf, PdfArgumentError } from "@pdf/index";
import type { PdfDocumentOptions } from "@pdf/index";
import type { DocxDocument } from "@word/types";
import { describe, expect, it } from "vitest";

const FIXED = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));

function docx(): DocxDocument {
  return {
    body: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "Hello" }] }] }],
    coreProperties: {
      title: "Core title",
      creator: "Core author",
      subject: "Core subject",
      keywords: "k1, k2"
    }
  } as DocxDocument;
}

function charts() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Row.addMany(ws, [
    ["A", 10],
    ["B", 20]
  ]);
  Chart.add(
    ws,
    {
      type: "bar",
      title: "Classic title",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
    },
    "D1:J10"
  );
  Chart.addEx(
    ws,
    {
      type: "treemap",
      title: "Ex title",
      categories: "Sheet1!$A$1:$A$2",
      series: [{ values: "Sheet1!$B$1:$B$2" }]
    },
    "D12:J22"
  );
  const [classic, ex] = Chart.list(ws);
  return { classic, ex };
}

const shared: PdfDocumentOptions = {
  title: "T",
  author: "A",
  subject: "S",
  keywords: "kw",
  creator: "C",
  creationDate: FIXED,
  modDate: FIXED,
  disableFontAutoDiscovery: true
};

describe("Pdf.fromDocx takes the document options", () => {
  it("an option wins over the core properties, which fill in the rest, keywords included", async () => {
    const fallback = (await Pdf.extract(await Pdf.fromDocx(docx()))).metadata;
    expect(fallback).toMatchObject({
      title: "Core title",
      author: "Core author",
      subject: "Core subject",
      keywords: "k1, k2"
    });
    const explicit = (await Pdf.extract(await Pdf.fromDocx(docx(), { ...shared }))).metadata;
    expect(explicit).toMatchObject({
      title: "T",
      author: "A",
      subject: "S",
      keywords: "kw",
      creator: "C"
    });
    expect(explicit.creationDate?.getTime()).toBe(FIXED.getTime());
  });

  it("with fixed dates the bytes are reproducible", async () => {
    const a = await Pdf.fromDocx(docx(), shared);
    const b = await Pdf.fromDocx(docx(), shared);
    expect(a).toEqual(b);
  });

  it("encrypts", async () => {
    const bytes = await Pdf.fromDocx(docx(), {
      ...shared,
      encryption: { ownerPassword: "owner", userPassword: "user" }
    });
    await expect(Pdf.extract(bytes)).rejects.toThrow(/password/i);
    expect((await Pdf.extract(bytes, { password: "user" })).text).toContain("Hello");
  });

  it("margins and defaultFontFamily are the new names of the flat margins and defaultFont", async () => {
    const viaNew = await Pdf.fromDocx(docx(), {
      ...shared,
      margins: { top: 100, left: 150 },
      defaultFontFamily: "Courier"
    });
    const viaOld = await Pdf.fromDocx(docx(), {
      ...shared,
      marginTop: 100,
      marginLeft: 150,
      defaultFont: "Courier"
    });
    expect(viaNew).toEqual(viaOld);
    expect(viaNew).not.toEqual(await Pdf.fromDocx(docx(), shared));
    // The new form wins for a side both name.
    const both = await Pdf.fromDocx(docx(), {
      ...shared,
      margins: { top: 100, left: 150 },
      marginTop: 10,
      defaultFontFamily: "Courier",
      defaultFont: "Times-Roman"
    });
    expect(both).toEqual(viaNew);
  });

  it("checks the new options as every converter does", async () => {
    for (const [options, field] of [
      [{ keywords: 42 }, "options.keywords"],
      [{ creationDate: new Date(Number.NaN) }, "options.creationDate"],
      [{ encryption: {} }, "options.encryption.ownerPassword"],
      [{ margins: { top: -1 } }, "options.margins.top"],
      [{ defaultFontFamily: 1 }, "options.defaultFontFamily"],
      [{ title: "a\u0000b" }, "options.title"]
    ] as const) {
      const err = await Pdf.fromDocx(docx(), options as never).catch(e => e);
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect(err.field).toBe(field);
      expect(err.message).toMatch(/^Pdf\.fromDocx: /);
    }
  });
});

describe("Pdf.fromChart takes the document options", () => {
  it("writes every metadata field, fixed dates and encryption", async () => {
    const { classic } = charts();
    const bytes = await Pdf.fromChart(classic, shared);
    const { metadata } = await Pdf.extract(bytes);
    expect(metadata).toMatchObject({ title: "T", author: "A", subject: "S", keywords: "kw" });
    expect(metadata.creator).toBe("C");
    expect(bytes).toEqual(await Pdf.fromChart(classic, shared));
    const locked = await Pdf.fromChart(classic, {
      ...shared,
      encryption: { ownerPassword: "o", userPassword: "u" }
    });
    await expect(Pdf.extract(locked)).rejects.toThrow(/password/i);
  });

  it("checks them", async () => {
    const { classic } = charts();
    const err = await Pdf.fromChart(classic, { keywords: 1 as never }).catch(e => e);
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect(err.field).toBe("options.keywords");
  });

  it.each(["vector", "raster"] as const)(
    "title is metadata only, on the %s path of both chart families (D2)",
    async path => {
      const { classic, ex } = charts();
      for (const [chart, own] of [
        [classic, "Classic title"],
        [ex, "Ex title"]
      ] as const) {
        const bytes = await Pdf.fromChart(chart, {
          ...shared,
          title: "Metadata only",
          forceRaster: path === "raster"
        });
        const read = await Pdf.extract(bytes);
        expect(read.metadata.title).toBe("Metadata only");
        expect(read.text).not.toContain("Metadata only");
        if (path === "vector") {
          expect(read.text).toContain(own);
        }
      }
    }
  );
});
