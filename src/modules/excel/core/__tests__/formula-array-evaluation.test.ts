/**
 * A formula this library evaluates over arrays, where Excel would intersect a range if the file stored it as an
 * ordinary formula, is stored as Excel 365 stores it when typed: a dynamic array over its own cell. A formula a file
 * stored as an ordinary one keeps its meaning — it is computed with Excel's implicit intersection and written back as
 * it came.
 *
 * Reported from `scientific-analysis`: `MEDIAN(IF(Samples[Group]="Control-A",Samples[Response]))` was written as a
 * plain `<f>`, Excel showed it as `MEDIAN(IF(@Samples[Group]=…))` and computed 130.01 where the library had 100.135.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Cell, Workbook } from "@excel";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { describe, expect, it } from "vitest";

const FILTERED = 'SUM(IF(B1:B3="y",A1:A3))';

function build() {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  [10, 20, 30].forEach((value, index) => {
    Cell.setValue(sheet, `A${index + 1}`, value);
    Cell.setValue(sheet, `B${index + 1}`, index === 0 ? "x" : "y");
  });
  Cell.setValue(sheet, "D2", { formula: `=${FILTERED}` });
  Cell.setValue(sheet, "E1", { formula: "SUM(A1:A3)" });
  return workbook;
}

async function sheetXml(bytes: Uint8Array): Promise<string> {
  return new TextDecoder().decode((await extractAll(bytes)).get("xl/worksheets/sheet1.xml")!.data);
}

describe("formulas evaluated over arrays", () => {
  it("are computed over the whole range", () => {
    const workbook = build();
    calculateFormulas(workbook);
    expect(Cell.getResult(Workbook.getWorksheet(workbook, "S")!, "D2")).toBe(50);
  });

  it("are written as a dynamic array over their own cell (XLSX)", async () => {
    const bytes = await Workbook.toBuffer(build());
    const xml = await sheetXml(bytes);
    expect(xml).toMatch(
      /<c r="D2" cm="1"><f t="array" ref="D2">SUM\(IF\(B1:B3=&quot;y&quot;,A1:A3\)\)<\/f>/
    );
    expect(xml).toContain('<c r="E1"><f>SUM(A1:A3)</f>');
    expect((await extractAll(bytes)).has("xl/metadata.xml")).toBe(true);
  });

  it("leave the caller's cell as it was", async () => {
    const workbook = build();
    await Workbook.toBuffer(workbook);
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    expect(Cell.getValue(sheet, "D2")).toEqual({ formula: `=${FILTERED}` });
  });

  it.each(["xlsx", "xlsb"] as const)("read back as dynamic arrays (%s)", async format => {
    const reread = Workbook.create();
    await Workbook.read(reread, await Workbook.toBuffer(build(), { format }));
    const sheet = Workbook.getWorksheet(reread, "S")!;
    expect(Cell.getValue(sheet, "D2")).toMatchObject({ formula: FILTERED, isDynamicArray: true });
    expect(Cell.getValue(sheet, "E1")).not.toHaveProperty("isDynamicArray");
    calculateFormulas(reread);
    expect(Cell.getResult(sheet, "D2")).toBe(50);
  });

  it("are written cell by cell when filled as a shared formula", async () => {
    const workbook = build();
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    Cell.setValue(sheet, "D3", { sharedFormula: "D2" });
    const xml = await sheetXml(await Workbook.toBuffer(workbook));
    expect(xml).toMatch(
      /<c r="D3" cm="1"><f t="array" ref="D3">SUM\(IF\(B2:B4=&quot;y&quot;,A2:A4\)\)<\/f>/
    );
    expect(xml).not.toContain('t="shared"');
    const reread = Workbook.create();
    await Workbook.read(reread, await Workbook.toBuffer(workbook, { format: "xlsb" }));
    expect(Cell.getValue(Workbook.getWorksheet(reread, "S")!, "D3")).toMatchObject({
      formula: 'SUM(IF(B2:B4="y",A2:A4))',
      isDynamicArray: true
    });
  });
});

describe("a formula a file stored as an ordinary one", () => {
  /** The workbook as another producer would store it: D2 as a plain `<f>`, in row 2 of the ranges it names. */
  async function legacy(): Promise<Uint8Array> {
    const entries = await extractAll(await Workbook.toBuffer(build()));
    return createZip(
      [...entries.entries()].map(([name, file]) => ({
        name,
        data:
          name === "xl/worksheets/sheet1.xml"
            ? new TextEncoder().encode(
                new TextDecoder()
                  .decode(file.data)
                  .replace('<c r="D2" cm="1"><f t="array" ref="D2">', '<c r="D2"><f>')
              )
            : file.data
      }))
    );
  }

  it("is computed with Excel's implicit intersection", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, await legacy());
    calculateFormulas(workbook);
    // Excel compares B2 alone ("y"), so SUM(IF(TRUE,A1:A3)) — every value.
    expect(Cell.getResult(Workbook.getWorksheet(workbook, "S")!, "D2")).toBe(60);
  });

  it.each(["xlsx", "xlsb"] as const)("is written back as it came (%s)", async format => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, await legacy());
    const reread = Workbook.create();
    await Workbook.read(reread, await Workbook.toBuffer(workbook, { format }));
    const value = Cell.getValue(Workbook.getWorksheet(reread, "S")!, "D2");
    expect(value).toMatchObject({ formula: FILTERED });
    expect(value).not.toHaveProperty("isDynamicArray");
    expect(value).not.toHaveProperty("shareType");
  });
});
