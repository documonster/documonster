import { createReadStream, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";

import { ExcelCsv } from "@excel/bridge/csv.node";
import { ExcelMarkdown, readMarkdownAll } from "@excel/bridge/markdown.node";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

type Book = Workbook.Handle;

const CSV = "name,qty\nPen,3\nInk,7";
const MD_ONE = "| Name | Qty |\n| --- | --- |\n| Pen | 3 |";
const MD_MANY = `${MD_ONE}\n\ntext between\n\n| City | Pop |\n| --- | --- |\n| Oslo | 700 |`;

/** Every worksheet's name and cell values, so two workbooks can be compared whole. */
function snapshot(wb: Book): { name: string; rows: unknown[][] }[] {
  return Workbook.listWorksheets(wb).map(ws => {
    const rows: unknown[][] = [];
    for (let r = 1; r <= 3; r++) {
      rows.push([Cell.getValue(ws, r, 1), Cell.getValue(ws, r, 2)]);
    }
    return { name: Worksheet.getName(ws), rows };
  });
}

let dir: string;
let csvPath: string;
let mdOnePath: string;
let mdManyPath: string;

beforeAll(() => {
  dir = mkdtempSync(join(tmpdir(), "documonster-read-wb-"));
  csvPath = join(dir, "in.csv");
  mdOnePath = join(dir, "one.md");
  mdManyPath = join(dir, "many.md");
  writeFileSync(csvPath, CSV, "utf8");
  writeFileSync(mdOnePath, MD_ONE, "utf8");
  writeFileSync(mdManyPath, MD_MANY, "utf8");
});

afterAll(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("ExcelCsv.read returns a new workbook", () => {
  it("holds one worksheet with the data", async () => {
    const wb = await ExcelCsv.read(CSV);
    const sheets = Workbook.listWorksheets(wb);
    expect(sheets).toHaveLength(1);
    expect(Cell.getValue(sheets[0], "A1")).toBe("name");
    expect(Cell.getValue(sheets[0], "A2")).toBe("Pen");
    expect(Cell.getValue(sheets[0], "B3")).toBe(7);
  });

  it("names the worksheet with sheetName", async () => {
    const wb = await ExcelCsv.read(CSV, { sheetName: "Stock" });
    expect(Workbook.listWorksheets(wb).map(ws => Worksheet.getName(ws))).toEqual(["Stock"]);
  });

  it("matches what the deprecated workbook-first form put into a fresh workbook", async () => {
    const legacy = Workbook.create();
    const ws = await ExcelCsv.read(legacy, CSV, { sheetName: "Stock" });
    expect(Worksheet.getName(ws)).toBe("Stock");
    expect(snapshot(await ExcelCsv.read(CSV, { sheetName: "Stock" }))).toEqual(snapshot(legacy));
  });

  it("reads bytes", async () => {
    const wb = await ExcelCsv.read(new TextEncoder().encode(CSV));
    expect(snapshot(wb)).toEqual(snapshot(await ExcelCsv.read(CSV)));
  });

  it("reads a Node readable stream", async () => {
    const wb = await ExcelCsv.read(Readable.from([CSV.slice(0, 7), CSV.slice(7)]));
    expect(snapshot(wb)).toEqual(snapshot(await ExcelCsv.read(CSV)));
  });
});

describe("ExcelCsv.readFile", () => {
  it("returns a new workbook", async () => {
    const wb = await ExcelCsv.readFile(csvPath, { sheetName: "Stock" });
    expect(snapshot(wb)).toEqual(snapshot(await ExcelCsv.read(CSV, { sheetName: "Stock" })));
  });

  it("keeps the deprecated workbook-first form, which adds a worksheet", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "Existing");
    const ws = await ExcelCsv.readFile(wb, csvPath, { sheetName: "Stock" });
    expect(Worksheet.getName(ws)).toBe("Stock");
    expect(Workbook.listWorksheets(wb).map(s => Worksheet.getName(s))).toEqual([
      "Existing",
      "Stock"
    ]);
    expect(Cell.getValue(ws, "A2")).toBe("Pen");
  });
});

describe("ExcelCsv.addWorksheet", () => {
  it("adds a worksheet to an existing workbook and returns it", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "Existing");
    const ws = await ExcelCsv.addWorksheet(wb, CSV, { sheetName: "Stock" });
    expect(Worksheet.getName(ws)).toBe("Stock");
    expect(Workbook.getWorksheet(wb, "Stock")).toBe(ws);
    expect(Workbook.listWorksheets(wb)).toHaveLength(2);
    expect(Cell.getValue(ws, "B2")).toBe(3);
  });

  it("adds a file stream, as the deprecated readFile(wb, path) did", async () => {
    const viaStream = Workbook.create();
    await ExcelCsv.addWorksheet(viaStream, createReadStream(csvPath, "utf8"), {
      sheetName: "Stock"
    });
    const viaLegacy = Workbook.create();
    await ExcelCsv.readFile(viaLegacy, csvPath, { sheetName: "Stock" });
    expect(snapshot(viaStream)).toEqual(snapshot(viaLegacy));
  });
});

describe("ExcelMarkdown reads return a new workbook", () => {
  it("read takes the first table", () => {
    const wb = ExcelMarkdown.read(MD_MANY, { sheetName: "T" });
    const sheets = Workbook.listWorksheets(wb);
    expect(sheets.map(ws => Worksheet.getName(ws))).toEqual(["T"]);
    expect(Cell.getValue(sheets[0], "A2")).toBe("Pen");
  });

  it('read({ tables: "all" }) makes one worksheet per table', () => {
    const wb = ExcelMarkdown.read(MD_MANY, { sheetName: "T", tables: "all" });
    const sheets = Workbook.listWorksheets(wb);
    expect(sheets.map(ws => Worksheet.getName(ws))).toEqual(["T", "T_2"]);
    expect(Cell.getValue(sheets[1], "A2")).toBe("Oslo");
  });

  it("readFile matches read, one table or all", async () => {
    expect(snapshot(await ExcelMarkdown.readFile(mdOnePath, { sheetName: "T" }))).toEqual(
      snapshot(ExcelMarkdown.read(MD_ONE, { sheetName: "T" }))
    );
    expect(
      snapshot(await ExcelMarkdown.readFile(mdManyPath, { sheetName: "T", tables: "all" }))
    ).toEqual(snapshot(ExcelMarkdown.read(MD_MANY, { sheetName: "T", tables: "all" })));
  });

  it("match what the deprecated workbook-first forms put into a fresh workbook, which read text only", async () => {
    // The deprecated forms keep reading every cell as text; the new ones read values as ExcelCsv does, which a
    // `map` that keeps the text and `convertBr: false` turn back into the old reading.
    const asText = { sheetName: "T", map: (value: string) => value, convertBr: false };
    const one = Workbook.create();
    const ws = ExcelMarkdown.read(one, MD_ONE, { sheetName: "T" });
    expect(Worksheet.getName(ws)).toBe("T");
    expect(snapshot(ExcelMarkdown.read(MD_ONE, asText))).toEqual(snapshot(one));

    const many = Workbook.create();
    const sheets = readMarkdownAll(many, MD_MANY, { sheetName: "T" });
    expect(sheets.map(s => Worksheet.getName(s))).toEqual(["T", "T_2"]);
    expect(snapshot(ExcelMarkdown.read(MD_MANY, { ...asText, tables: "all" }))).toEqual(
      snapshot(many)
    );

    const oneFile = Workbook.create();
    const fileSheet = await ExcelMarkdown.readFile(oneFile, mdOnePath, { sheetName: "T" });
    expect(Worksheet.getName(fileSheet)).toBe("T");
    expect(snapshot(oneFile)).toEqual(snapshot(one));

    expect(
      snapshot(await ExcelMarkdown.readFile(mdManyPath, { ...asText, tables: "all" }))
    ).toEqual(snapshot(many));
  });
});

describe("ExcelMarkdown.addWorksheet / addWorksheets", () => {
  it("add to an existing workbook and return what they added", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "Existing");

    const ws = ExcelMarkdown.addWorksheet(wb, MD_ONE, { sheetName: "One" });
    expect(Worksheet.getName(ws)).toBe("One");
    expect(Workbook.getWorksheet(wb, "One")).toBe(ws);

    const added = ExcelMarkdown.addWorksheets(wb, MD_MANY, { sheetName: "Many" });
    expect(added.map(s => Worksheet.getName(s))).toEqual(["Many", "Many_2"]);
    expect(Workbook.listWorksheets(wb).map(s => Worksheet.getName(s))).toEqual([
      "Existing",
      "One",
      "Many",
      "Many_2"
    ]);
    // Read as ExcelCsv reads a cell: the number it says.
    expect(Cell.getValue(added[1], "B2")).toBe(700);
  });
});
