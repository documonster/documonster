/**
 * The async whole-workbook readers take a `Blob` — a `File` from an `<input>`, a `fetch` response's `blob()` — as
 * `ExcelCsv.read` and `Archive` already do: `Workbook.read` and `Workbook.readWithDiagnostics`, XLSX and XLSB.
 */
import { Cell, ExcelArgumentError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

async function blobOf(format: "xlsx" | "xlsb"): Promise<Blob> {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 42);
  return new Blob([(await Workbook.toBuffer(wb, { format })) as BlobPart]);
}

describe("reading a Blob", () => {
  it.each(["xlsx", "xlsb"] as const)("Workbook.read: %s", async format => {
    const read = await Workbook.read(await blobOf(format));
    expect(Cell.getValue(Workbook.getWorksheet(read, "S")!, "A1")).toBe(42);
  });

  it.each(["xlsx", "xlsb"] as const)("Workbook.readWithDiagnostics: %s", async format => {
    const { workbook } = await Workbook.readWithDiagnostics(await blobOf(format));
    expect(Cell.getValue(Workbook.getWorksheet(workbook, "S")!, "A1")).toBe(42);
  });

  it("a File is a Blob", async () => {
    const file = new File([await (await blobOf("xlsx")).arrayBuffer()], "book.xlsx");
    expect(Cell.getValue(Workbook.listWorksheets(await Workbook.read(file))[0]!, "A1")).toBe(42);
  });

  it("other objects are still refused, naming Blob among what is taken", async () => {
    await expect(Workbook.read({} as never)).rejects.toThrow(ExcelArgumentError);
    await expect(Workbook.read({} as never)).rejects.toThrow(/Blob/);
  });
});
