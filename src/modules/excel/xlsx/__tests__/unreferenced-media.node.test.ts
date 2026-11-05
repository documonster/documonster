/**
 * Media nothing references are left out of the package at write time; media anything references
 * — including a part preserved verbatim from a loaded file — are always kept, and image ids stay
 * stable in memory.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import {
  Cell,
  Chart,
  HeaderFooterImage,
  Image,
  Row,
  Watermark,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

const here = path.dirname(fileURLToPath(import.meta.url));
const excelDir = path.resolve(here, "../..");

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

async function mediaParts(bytes: Uint8Array): Promise<string[]> {
  const entries = await extractAll(bytes);
  return [...entries.keys()]
    .filter(name => name.startsWith("xl/media/") && !name.endsWith("/"))
    .sort();
}

/**
 * Every `xl/media/*` file name some relationship part of the package targets — optionally only
 * those whose relationship part passes `from`.
 */
async function referencedMedia(
  bytes: Uint8Array,
  from: (relsPath: string) => boolean = () => true
): Promise<Set<string>> {
  const entries = await extractAll(bytes);
  const names = new Set<string>();
  for (const [name, entry] of entries) {
    if (!name.endsWith(".rels") || !from(name)) {
      continue;
    }
    for (const match of new TextDecoder().decode(entry.data).matchAll(/Target="([^"]+)"/g)) {
      const segments = match[1].split("/");
      if (segments.at(-2) === "media") {
        names.add(`xl/media/${segments.at(-1)}`);
      }
    }
  }
  return names;
}

describe("unreferenced media are not written", () => {
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`${format}: removing the last placement drops the part`, async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      const kept = Image.add(wb, { base64: PNG, extension: "png" });
      const dropped = Image.add(wb, { base64: PNG, extension: "png" });
      Image.place(ws, kept, "A1:B2");
      Image.place(ws, dropped, "D1:E2");
      expect(Image.remove(ws, 1)).toBe(true);

      const bytes = await Workbook.toBuffer(wb, { format });
      const parts = await mediaParts(bytes);
      expect(parts).toHaveLength(1);
      expect([...(await referencedMedia(bytes))]).toEqual(parts);
      if (format === "xlsx") {
        await expectValidXlsx(bytes, { label: "unreferenced media dropped" });
      }
    });
  }

  it("keeps ids stable: re-placing a removed image before writing works", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const first = Image.add(wb, { base64: PNG, extension: "png" });
    const second = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, first, "A1:B2");
    Image.place(ws, second, "D1:E2");
    Image.remove(ws, 0);
    Image.remove(ws, 0);

    // Nothing references either image now — and both ids still resolve.
    expect(await mediaParts(await Workbook.toBuffer(wb))).toEqual([]);
    Image.place(ws, second, "G1:H2");
    const bytes = await Workbook.toBuffer(wb);
    await expectValidXlsx(bytes, { label: "re-placed after remove" });
    expect(await mediaParts(bytes)).toHaveLength(1);

    const back = Workbook.create();
    await Workbook.read(back, bytes);
    expect(Image.list(Workbook.getWorksheet(back, "S")!)).toHaveLength(1);
  });

  it("keeps media no API call released, referenced or not — as the writer always did", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    Image.add(wb, { base64: PNG, extension: "png" }); // added, never placed
    expect(await mediaParts(await Workbook.toBuffer(wb))).toHaveLength(1);
    expect(await mediaParts(await Workbook.toBuffer(wb, { format: "xlsb" }))).toHaveLength(1);
  });

  it("keeps media referenced other than by a placement", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const background = Image.add(wb, { base64: PNG, extension: "png" });
    const watermark = Image.add(wb, { base64: PNG, extension: "png" });
    const header = Image.add(wb, { base64: PNG, extension: "png" });
    Image.setBackground(ws, background);
    Watermark.set(ws, { imageId: watermark });
    const other = Workbook.addWorksheet(wb, "T");
    Watermark.set(other, { imageId: header, mode: "header" });
    const bytes = await Workbook.toBuffer(wb);
    await expectValidXlsx(bytes, { label: "background + watermarks" });
    expect(await mediaParts(bytes)).toHaveLength(3);
  });
});

describe("a read → write round trip keeps every referenced medium", () => {
  const fixtures = [
    "__tests__/data/content-type-element.xlsx",
    "__tests__/data/duplicate-row-images.xlsx",
    "__tests__/data/hyperlink-without-text.xlsx",
    "__tests__/data/malformed-comment.xlsx",
    "stream/__tests__/data/hyperlinks-cache.xlsx",
    "stream/__tests__/data/images.xlsx",
    "examples/data/themed.xlsb"
  ];
  for (const fixture of fixtures) {
    it(fixture, async () => {
      const source = new Uint8Array(readFileSync(path.join(excelDir, fixture)));
      const format = fixture.endsWith(".xlsb") ? "xlsb" : "xlsx";
      const wb = Workbook.create();
      await Workbook.read(wb, source, { format } as never);
      const written = await Workbook.toBuffer(wb, { format });

      // Every media part of the input survives — header/footer pictures included, which Excel
      // stores in a `vmlDrawingN.vml` that only the sheet's `legacyDrawingHF` identifies — and
      // every one the input referenced is still referenced. A part the input itself left
      // unreferenced (`content-type-element`'s `image1.emf`) is written as it always was.
      // Names may be reassigned on write, so counts are compared.
      const input = await mediaParts(source);
      const after = await mediaParts(written);
      const bytesOf = async (pkg: Uint8Array) => {
        const entries = await extractAll(pkg);
        return [...entries]
          .filter(([name]) => name.startsWith("xl/media/") && !name.endsWith("/"))
          .map(([, entry]) => Buffer.from(entry.data).toString("base64"))
          .sort();
      };
      // Byte for byte: the same multiset of media contents.
      expect(await bytesOf(written)).toEqual(await bytesOf(source));
      const referencedBefore = await referencedMedia(source);
      const referencedAfter = await referencedMedia(written);
      expect(input.length).toBeGreaterThan(0);
      expect(after).toHaveLength(input.length);
      expect(referencedAfter.size).toBe(referencedBefore.size);
      for (const name of referencedAfter) {
        expect(after).toContain(name);
      }
      if (format === "xlsx") {
        await expectValidXlsx(written, { label: fixture });
      }
    });
  }
});

describe("removing every placement of a read file", () => {
  const fixtures = [
    "__tests__/data/duplicate-row-images.xlsx",
    "stream/__tests__/data/images.xlsx",
    "examples/data/themed.xlsb"
  ];
  for (const fixture of fixtures) {
    it(`${fixture}: drops exactly the media nothing written references`, async () => {
      const source = new Uint8Array(readFileSync(path.join(excelDir, fixture)));
      const format = fixture.endsWith(".xlsb") ? "xlsb" : "xlsx";
      const wb = Workbook.create();
      await Workbook.read(wb, source, { format } as never);
      for (const ws of Workbook.listWorksheets(wb)) {
        while (Image.remove(ws, 0)) {
          // remove them all
        }
      }
      const written = await Workbook.toBuffer(wb, { format });
      const after = await mediaParts(written);
      const referenced = await referencedMedia(written);
      // Nothing dangles, and nothing released survives unless something still points at it.
      for (const name of referenced) {
        expect(after).toContain(name);
      }
      expect(after.length).toBeLessThanOrEqual((await mediaParts(source)).length);
      if (format === "xlsx") {
        await expectValidXlsx(written, { label: `${fixture} without placements` });
      }
    });
  }

  it("keeps a released medium that was placed again", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, id, "A1:B2");
    Image.remove(ws, 0);
    Image.place(ws, id, "D1:E2"); // re-placed: in use again
    const bytes = await Workbook.toBuffer(wb);
    expect(await mediaParts(bytes)).toHaveLength(1);
  });

  it("releases both media of a picture with an SVG companion, and drops the emptied drawing", async () => {
    // `duplicate-row-images.xlsx`'s one sheet picture is a PNG blip with an `asvg:svgBlip` naming an SVG;
    // `image3.png` belongs to the workbook's cell images and stays.
    const source = new Uint8Array(
      readFileSync(path.join(excelDir, "__tests__/data/duplicate-row-images.xlsx"))
    );
    const wb = Workbook.create();
    await Workbook.read(wb, source);
    const ws = Workbook.listWorksheets(wb).find(sheet => Image.list(sheet).length > 0)!;
    expect(Image.remove(ws, 0)).toBe(true);
    const written = await Workbook.toBuffer(wb);
    await expectValidXlsx(written);
    expect(await mediaParts(written)).toEqual(["xl/media/image3.png"]);
    const entries = await extractAll(written);
    const parts = [...entries.keys()];
    expect(parts.filter(name => name.startsWith("xl/drawings/"))).toEqual([]);
    for (const [name, entry] of entries) {
      if (/^xl\/worksheets\/(?:_rels\/)?sheet\d+\.xml(?:\.rels)?$/.test(name)) {
        expect(new TextDecoder().decode(entry.data)).not.toMatch(/<drawing\b|\/drawing"/);
      }
    }
  });

  it("keeps a drawing that still holds something after its pictures go", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, id, "A1:B2");
    const read = Workbook.create();
    await Workbook.read(read, await Workbook.toBuffer(wb));
    const sheet = Workbook.getWorksheet(read, "S")!;
    Image.remove(sheet, 0);
    Image.place(sheet, Image.add(read, { base64: PNG, extension: "png" }), "C3:D4");
    const parts = [...(await extractAll(await Workbook.toBuffer(read))).keys()];
    expect(parts.filter(name => /^xl\/drawings\/drawing\d+\.xml$/.test(name))).toHaveLength(1);
  });
});

describe("every API path that stops using a medium releases it", () => {
  /** A workbook whose sheet `S` uses `id` through `use`, and — with `shared` — sheet `T` places it too. */
  function setup(use: (ws: Worksheet.Handle, id: number) => void, shared = false) {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    use(ws, id);
    if (shared) {
      Image.place(Workbook.addWorksheet(wb, "T"), id, "A1:B2");
    }
    return { wb, ws, id };
  }

  const pictureFillChart = (ws: Worksheet.Handle, id: number) =>
    Chart.place(ws, "C1:J10", {
      type: "bar",
      series: [{ values: "S!$A$1:$A$1", pictureFill: { image: { workbookImageId: id } } }]
    });

  const paths: Record<
    string,
    {
      use: (ws: Worksheet.Handle, id: number) => void;
      drop: (wb: Workbook.Handle, ws: Worksheet.Handle, id: number) => void;
    }
  > = {
    "Workbook.removeWorksheet": {
      use: (ws, id) => Image.place(ws, id, "A1:B2"),
      drop: wb => Workbook.removeWorksheet(wb, "S")
    },
    "Workbook.removeWorksheet (background, header image, chart fill)": {
      use: (ws, id) => {
        Image.setBackground(ws, id);
        HeaderFooterImage.set(ws, { imageId: id, position: "CH" });
        pictureFillChart(ws, id);
      },
      drop: wb => Workbook.removeWorksheet(wb, "S")
    },
    "Image.setBackground replacing the background": {
      use: (ws, id) => Image.setBackground(ws, id),
      drop: (wb, ws) => Image.setBackground(ws, Image.add(wb, { base64: PNG, extension: "png" }))
    },
    "Image.removeBackground": {
      use: (ws, id) => Image.setBackground(ws, id),
      drop: (_wb, ws) => Image.removeBackground(ws)
    },
    "Chart.remove (picture fill)": {
      use: pictureFillChart,
      drop: (_wb, ws) => Chart.remove(ws, 0)
    },
    "Image.setModel naming another picture": {
      use: (ws, id) => Image.place(ws, id, "A1:B2"),
      drop: (wb, ws) => {
        const other = Image.add(wb, { base64: PNG, extension: "png" });
        const [image] = Image.list(ws);
        Image.setModel(image, { ...Image.getModel(image), imageId: String(other) } as never);
      }
    },
    "Row.duplicate overwriting a placed picture": {
      use: (ws, id) => Image.place(ws, id, "A2:B3"),
      drop: (_wb, ws) => Row.duplicate(ws, 1, 1)
    },
    "Worksheet.setModel without the picture": {
      use: (ws, id) => Image.place(ws, id, "A1:B2"),
      drop: (_wb, ws) => Worksheet.setModel(ws, { ...Worksheet.getModel(ws), media: [] })
    }
  };

  for (const [name, { use, drop }] of Object.entries(paths)) {
    it(`${name}: the part goes when nothing else uses it`, async () => {
      const { wb, ws, id } = setup(use);
      drop(wb, ws, id);
      const bytes = await Workbook.toBuffer(wb);
      expect(await mediaParts(bytes)).not.toContain(`xl/media/image${id + 1}.png`);
      expect([...(await referencedMedia(bytes))].sort()).toEqual(await mediaParts(bytes));
    });

    it(`${name}: the part stays while another sheet uses it`, async () => {
      const { wb, ws, id } = setup(use, true);
      drop(wb, ws, id);
      expect(await mediaParts(await Workbook.toBuffer(wb))).toContain(
        `xl/media/image${id + 1}.png`
      );
    });
  }

  it("Workbook.removeChartsheet and replacing its chart release a picture fill", async () => {
    for (const drop of [
      (wb: Workbook.Handle) => Workbook.removeChartsheet(wb, "C"),
      (wb: Workbook.Handle) =>
        Workbook.replaceChartsheetChart(wb, "C", {
          type: "bar",
          series: [{ values: "S!$A$1:$A$1" }]
        })
    ]) {
      const wb = Workbook.create();
      Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
      const id = Image.add(wb, { base64: PNG, extension: "png" });
      Workbook.addChartsheet(wb, "C", {
        chart: {
          type: "bar",
          series: [{ values: "S!$A$1:$A$1", pictureFill: { image: { workbookImageId: id } } }]
        }
      });
      drop(wb);
      expect(await mediaParts(await Workbook.toBuffer(wb))).toEqual([]);
    }
  });
});

it("removing a worksheet removes its charts' parts", async () => {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Workbook.addWorksheet(wb, "K");
  Cell.setValue(ws, "A1", 1);
  Chart.place(ws, "C1:J10", { type: "bar", series: [{ values: "S!$A$1:$A$1" }] });
  Chart.place(ws, "C12:J20", { type: "funnel", series: [{ values: "S!$A$1:$A$1" }] });
  Workbook.removeWorksheet(wb, "S");
  const parts = [...(await extractAll(await Workbook.toBuffer(wb))).keys()];
  expect(parts.filter(part => part.startsWith("xl/charts/"))).toEqual([]);
});

it("removing a read chart releases the medium its relationship names, whatever its id", async () => {
  const OTHER =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Image.add(wb, { base64: OTHER, extension: "png" }); // never placed: image1.png
  const fill = Image.add(wb, { base64: PNG, extension: "png" }); // image2.png
  Chart.place(ws, "C1:J10", {
    type: "bar",
    series: [{ values: "S!$A$1:$A$1", pictureFill: { image: { workbookImageId: fill } } }]
  });
  // Name the chart's picture the way another producer might: image5.png, read as media id 1.
  const entries = await extractAll(await Workbook.toBuffer(wb));
  const files: { name: string; data: Uint8Array }[] = [];
  const decoder = new TextDecoder();
  for (const [name, entry] of entries) {
    const renamed = name === "xl/media/image2.png" ? "xl/media/image5.png" : name;
    files.push({
      name: renamed,
      data: name.endsWith(".rels")
        ? new TextEncoder().encode(
            decoder.decode(entry.data).replace(/media\/image2\.png/g, "media/image5.png")
          )
        : entry.data
    });
  }
  const source = await createZip(files);

  const back = Workbook.create();
  await Workbook.read(back, source);
  const sheet = Workbook.getWorksheet(back, "S")!;
  await Workbook.toBuffer(back); // a write in between must not change what the chart names
  expect(Chart.remove(sheet, 0)).toBe(true);
  const written = await extractAll(await Workbook.toBuffer(back));
  const media = [...written]
    .filter(([name]) => name.startsWith("xl/media/"))
    .map(([, entry]) => Buffer.from(entry.data).toString("base64"));
  // The chart's picture is gone; the medium the input never referenced is kept, as promised.
  expect(media).toEqual([OTHER.slice(OTHER.indexOf(",") + 1)]);

  // And unedited, the chart still points at its own picture after the media are renumbered.
  const again = Workbook.create();
  await Workbook.read(again, source);
  const roundTrip = await Workbook.toBuffer(again);
  const parts = await extractAll(roundTrip);
  const target = /Target="\.\.\/media\/([^"]+)"/.exec(
    decoder.decode(parts.get("xl/charts/_rels/chart1.xml.rels")!.data)
  )![1];
  expect(Buffer.from(parts.get(`xl/media/${target}`)!.data).toString("base64")).toBe(
    PNG.slice(PNG.indexOf(",") + 1)
  );
  await expectValidXlsx(roundTrip, { label: "chart picture renamed on read" });
});
