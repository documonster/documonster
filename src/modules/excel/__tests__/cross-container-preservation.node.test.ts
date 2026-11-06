/**
 * A preserved part survives a change of container, and a container-specific one does not.
 *
 * **Both halves are needed, and having only the first made things worse.** Each writer decided whether a preserved
 * part's inbound edge came "from the workbook" by comparing against *its own* container's path — `xl/workbook.xml` in
 * the XLSX writer, `xl/workbook.bin` in the XLSB one. So a part read from one container and written to the other had no
 * recognised inbound edge, was judged an orphan, and was dropped.
 *
 * Measured on `cal-any_sheets.xlsb`: its theme states `accent1 = 5B9BD5`, and the XLSX conversion carried the built-in
 * `4472C4` — the preserved theme dropped, the default written in its place. Every cell whose colour is a `{ theme: n }`
 * index therefore rendered in a different colour, with nothing reported.
 *
 * Recognising the edge then kept alive the one part that must *not* travel — `xl/chartsheets/sheet1.bin`, content type
 * `application/vnd.ms-excel.chartsheet`, a BIFF12 record stream inside a SpreadsheetML package — and surfaced a second
 * defect behind it: `ContentTypesXform` declares `xl/theme/theme1.xml` unconditionally, so a *preserved* theme was
 * declared twice and the writer's own validator refused the package with `content-types-duplicate-override`. That could
 * not happen before, because a preserved theme never reached an XLSX write.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { isForeignContainerPart } from "@excel/core/opaque-part";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { validateXlsbBuffer } from "@excel/utils/xlsb-validator";
import { describe, expect, it } from "vitest";

const FIXTURE = "tmp/xlsb-corpus/cal-any_sheets.xlsb";

/** The fixture's bytes, or `undefined` when the corpus has not been fetched. */
async function fixture(): Promise<Uint8Array | undefined> {
  try {
    return Uint8Array.from(await readFile(FIXTURE));
  } catch {
    return undefined;
  }
}

/** Decoded text of one part. */
function text(parts: Awaited<ReturnType<typeof extractAll>>, path: string): string {
  return new TextDecoder().decode(parts.get(path)!.data);
}

describe("preserved parts crossing containers", () => {
  it("refuses the conversion by default, because the chartsheet is a real loss", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    // **`unsupported` defaults to `"error"` and used to be inert for XLSX.** The one thing an XLSX write cannot carry is
    // a preserved part from the other container — this fixture's `xl/chartsheets/sheet1.bin` — so a caller
    // converting it now learns rather than discovering an empty tab in Excel. Every case below passes `"ignore"`, which
    // is the same decision made explicitly.
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    await expect(Workbook.toBuffer(workbook, { format: "xlsx" })).rejects.toThrow(
      /preserved part from the other container/
    );
  });

  it("keeps the workbook's own theme instead of substituting the built-in one", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    const original = text(await extractAll(bytes), "xl/theme/theme1.xml");
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    const converted = await extractAll(
      await Workbook.toBuffer(workbook, { format: "xlsx", unsupported: "ignore" })
    );
    // Byte-for-byte, not merely "a theme is present": the failure this replaces produced a perfectly valid theme that
    // was the wrong one.
    expect(text(converted, "xl/theme/theme1.xml")).toBe(original);
    // And the accent that made it visible, asserted by value so a future default cannot quietly match.
    expect(original).toContain('<a:srgbClr val="5B9BD5"/>');
  });

  it("declares the theme exactly once", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    const converted = await extractAll(
      await Workbook.toBuffer(workbook, { format: "xlsx", unsupported: "ignore" })
    );
    const contentTypes = text(converted, "[Content_Types].xml");
    expect(contentTypes.match(/theme1\.xml/g) ?? []).toHaveLength(1);
  });

  it("does not carry a BIFF12 sheet part into a SpreadsheetML package", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    const converted = await extractAll(
      await Workbook.toBuffer(workbook, { format: "xlsx", unsupported: "ignore" })
    );
    expect([...converted.keys()].filter(name => name.endsWith(".bin"))).toEqual([]);
    // Nor may its content type be declared for a part that is not there.
    expect(text(converted, "[Content_Types].xml")).not.toContain("vnd.ms-excel.chartsheet");
  });

  it("still carries the container-neutral parts", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    const converted = await extractAll(
      await Workbook.toBuffer(workbook, { format: "xlsx", unsupported: "ignore" })
    );
    // The guard against over-correcting: a chart, its style and colour parts and its drawing are the same bytes in
    // either container, so excluding sheet parts must not take them with it.
    for (const path of [
      "xl/charts/chart1.xml",
      "xl/charts/style1.xml",
      "xl/charts/colors1.xml",
      "xl/drawings/drawing1.xml"
    ]) {
      expect([...converted.keys()]).toContain(path);
    }
  });

  it("leaves the same-container round trip byte-exact", async () => {
    const bytes = await fixture();
    if (bytes === undefined) {
      return;
    }
    // The property none of the above may cost: an unmodified XLSB comes back as the bytes it arrived as, which is the
    // strongest statement this library makes about a file whose features it does not all model.
    const workbook = Workbook.create();
    await Workbook.read(workbook, bytes);
    const again = await Workbook.toBuffer(workbook, { format: "xlsb", unsupported: "ignore" });
    expect([...again]).toEqual([...bytes]);
  });
});

describe("modelled parts crossing containers", () => {
  const TABLE_XLSB = "src/modules/excel/xlsb/__tests__/data/table-this-row.excel.xlsb";

  it("does not carry a BIFF12 table part beside the table it regenerates", async () => {
    // The table is read into the model and written again as `table1.xml`; the original `table1.bin` was kept as well,
    // with its content type and a second table relationship on the sheet.
    const workbook = Workbook.create();
    await Workbook.read(workbook, Uint8Array.from(await readFile(TABLE_XLSB)));
    const buffer = await Workbook.toBuffer(workbook, { format: "xlsx" });
    const converted = await extractAll(buffer);
    expect([...converted.keys()].filter(name => name.endsWith(".bin"))).toEqual([]);
    expect(text(converted, "[Content_Types].xml")).not.toContain("vnd.ms-excel.table");
    expect(text(converted, "xl/worksheets/_rels/sheet1.xml.rels")).not.toContain(".bin");
    expect((await validateXlsxBuffer(buffer)).problems).toEqual([]);
  });

  it("writes the table once when the same XLSB is modified and written back", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, Uint8Array.from(await readFile(TABLE_XLSB)));
    Cell.setValue(Workbook.getWorksheets(workbook)[0], "Z1", 1);
    const buffer = await Workbook.toBuffer(workbook, { format: "xlsb" });
    const parts = await extractAll(buffer);
    expect([...parts.keys()].filter(name => name.startsWith("xl/tables/"))).toHaveLength(1);
    expect((await validateXlsbBuffer(buffer)).problems).toEqual([]);
  });

  it("drops a sheet's relationship to a part that could not cross, rather than leaving it dangling", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, Uint8Array.from(await readFile(TABLE_XLSB)));
    const model = Workbook.getModel(workbook);
    const relationship = {
      id: "rId9",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
      target: "../pivotTables/pivotTable1.bin"
    };
    model.opaqueParts = [
      ...(model.opaqueParts ?? []),
      {
        path: "xl/pivotTables/pivotTable1.bin",
        data: new Uint8Array([0]),
        contentType: "application/vnd.ms-excel.pivotTable",
        sourceRelationships: [{ ...relationship, source: "xl/worksheets/sheet1.bin" }]
      }
    ];
    model.worksheets[0].opaqueRels = [relationship];
    const copy = Workbook.create();
    Workbook.setModel(copy, model);
    await expect(Workbook.toBuffer(copy, { format: "xlsx" })).rejects.toThrow(
      /pivotTable1\.bin: preserved part from the other container/
    );
    const buffer = await Workbook.toBuffer(copy, { format: "xlsx", unsupported: "ignore" });
    const converted = await extractAll(buffer);
    expect(text(converted, "xl/worksheets/_rels/sheet1.xml.rels")).not.toContain("pivotTable");
    expect((await validateXlsxBuffer(buffer)).problems).toEqual([]);
  });

  it("treats every part of the other container's own family as foreign, and nothing shared", () => {
    const part = (contentType: string) => ({ path: "x", data: new Uint8Array(), contentType });
    for (const contentType of [
      "application/vnd.ms-excel.table",
      "application/vnd.ms-excel.comments",
      "application/vnd.ms-excel.pivotTable",
      "application/vnd.ms-excel.pivotCacheDefinition",
      "application/vnd.ms-excel.pivotCacheRecords",
      "application/vnd.ms-excel.querytable",
      "application/vnd.ms-excel.connections",
      "application/vnd.ms-excel.externalLink",
      "application/vnd.ms-excel.sheetMetadata",
      "application/vnd.ms-excel.calcChain",
      "application/vnd.ms-excel.sharedStrings",
      "application/vnd.ms-excel.styles"
    ]) {
      expect(isForeignContainerPart(part(contentType), "xlsx"), contentType).toBe(true);
      expect(isForeignContainerPart(part(contentType), "xlsb"), contentType).toBe(false);
    }
    for (const contentType of [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
      "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    ]) {
      expect(isForeignContainerPart(part(contentType), "xlsb"), contentType).toBe(true);
      expect(isForeignContainerPart(part(contentType), "xlsx"), contentType).toBe(false);
    }
    for (const contentType of [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
      "application/vnd.ms-office.vbaProject",
      "application/vnd.openxmlformats-officedocument.theme+xml",
      "application/vnd.openxmlformats-officedocument.drawing+xml",
      "application/vnd.ms-excel.threadedcomments+xml",
      "application/vnd.ms-excel.rdrichvalue+xml",
      "image/png"
    ]) {
      expect(isForeignContainerPart(part(contentType), "xlsx"), contentType).toBe(false);
      expect(isForeignContainerPart(part(contentType), "xlsb"), contentType).toBe(false);
    }
  });
});
