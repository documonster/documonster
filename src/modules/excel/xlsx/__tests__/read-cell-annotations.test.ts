import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * Hyperlinks and comments are attached to cells by address after the cell pass, rather than looked up by every
 * cell (see `WorksheetXform.reconcile`). These pin that every shape still lands where it did.
 */

async function rewriteEntry(
  buffer: Uint8Array,
  path: string,
  patch: (xml: string) => string
): Promise<Uint8Array> {
  const archive = zip();
  for (const [name, file] of await extractAll(buffer)) {
    if (file.type === "directory") {
      continue;
    }
    archive.add(name, name === path ? patch(new TextDecoder().decode(file.data)) : file.data);
  }
  return archive.bytes();
}

async function roundTrip(buffer: Uint8Array) {
  const wb = Workbook.create();
  await Workbook.read(wb, buffer);
  return Workbook.getWorksheet(wb, "Sheet1")!;
}

describe("reading hyperlinks and comments", () => {
  it("attaches each to its own cell, and nothing to its neighbours", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    for (let row = 1; row <= 50; row++) {
      for (const col of ["A", "B", "C"]) {
        Cell.setValue(ws, `${col}${row}`, `${col}${row}`);
      }
    }
    Cell.setValue(ws, "B7", { text: "external", hyperlink: "https://example.com/" });
    Cell.setValue(ws, "C40", { text: "internal", hyperlink: "#Sheet1!A1" });
    Cell.setValue(ws, "A9", { formula: "1+1", result: 2 });
    Cell.setNote(ws, "A2", "on a value");
    Cell.setNote(ws, "C50", "last cell");

    const read = await roundTrip(await Workbook.toBuffer(wb));

    expect(Cell.getValue(read, "B7")).toEqual({
      text: "external",
      hyperlink: "https://example.com/"
    });
    expect(Cell.getValue(read, "C40")).toEqual({ text: "internal", hyperlink: "#Sheet1!A1" });
    expect(Cell.getValue(read, "A9")).toMatchObject({ formula: "1+1", result: 2 });
    expect(Cell.getValue(read, "B8")).toBe("B8");
    expect(Cell.getNote(read, "A2")).toBeDefined();
    expect(Cell.getNote(read, "C50")).toBeDefined();
    expect(Cell.getNote(read, "A3")).toBeUndefined();
    expect(Cell.getValue(read, "A2")).toBe("A2");
  });

  it("keeps a comment whose cell has no <c>, including on a row that has none", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "kept");
    Cell.setValue(ws, "A3", "kept");
    Cell.setNote(ws, "B1", "beside a value");
    Cell.setNote(ws, "D2", "on an empty row");

    const patched = await rewriteEntry(
      await Workbook.toBuffer(wb),
      "xl/worksheets/sheet1.xml",
      xml =>
        xml
          .replace(/<c r="B1"[^>]*\/>|<c r="B1"[^>]*>.*?<\/c>/, "")
          .replace(/<row r="2"[^>]*\/>|<row r="2"[^>]*>.*?<\/row>/, "")
    );
    const sheetXml = new TextDecoder().decode(
      (await extractAll(patched)).get("xl/worksheets/sheet1.xml")!.data
    );
    expect(sheetXml).not.toContain('r="B1"');
    expect(sheetXml).not.toContain('r="D2"');

    const read = await roundTrip(patched);
    expect(Cell.getNote(read, "B1")).toBeDefined();
    expect(Cell.getNote(read, "D2")).toBeDefined();
    expect(Cell.getValue(read, "A1")).toBe("kept");
    expect(Cell.getValue(read, "A3")).toBe("kept");
  });

  it("attaches every annotation on a wide, densely linked sheet", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    const cols = 60;
    for (let row = 1; row <= 20; row++) {
      for (let col = 1; col <= cols; col++) {
        Cell.setValue(ws, row, col, {
          text: `t${row}-${col}`,
          hyperlink: `https://example.com/${row}/${col}`
        });
      }
    }
    const read = await roundTrip(await Workbook.toBuffer(wb));
    for (let row = 1; row <= 20; row++) {
      for (let col = 1; col <= cols; col++) {
        expect(Cell.getValue(read, row, col)).toEqual({
          text: `t${row}-${col}`,
          hyperlink: `https://example.com/${row}/${col}`
        });
      }
    }
  });

  it("keeps several comment-only cells on one row, in a row with no <c> at all", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "kept");
    for (const col of ["B", "C", "D"]) {
      Cell.setNote(ws, `${col}4`, col);
    }
    const patched = await rewriteEntry(
      await Workbook.toBuffer(wb),
      "xl/worksheets/sheet1.xml",
      xml => xml.replace(/<row r="4"[^>]*\/>|<row r="4"[^>]*>.*?<\/row>/, "")
    );
    const read = await roundTrip(patched);
    for (const col of ["B", "C", "D"]) {
      expect(Cell.getNote(read, `${col}4`)).toBeDefined();
    }
    expect(Cell.getValue(read, "A1")).toBe("kept");
  });
});
