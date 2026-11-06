/**
 * `Workbook.*` and `Column.*` take a streaming writer as the other namespaces do: what they support reaches the file,
 * the writer's state checks apply (a committed workbook; columns written with the first committed row), and what only
 * a workbook in memory can do is refused by name.
 */
import {
  Cell,
  Column,
  ExcelArgumentError,
  ExcelStreamStateError,
  ExcelUnsupportedFeatureError,
  Stream,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

async function readBack(wb: Stream.WorkbookWriter): Promise<Workbook.Handle> {
  return Workbook.read(await wb.toBuffer());
}

describe("Workbook functions on a streaming workbook writer", () => {
  it("addWorksheet / getWorksheet / listWorksheets / eachWorksheet", async () => {
    const wb = new Stream.WorkbookWriter({});
    const a = Workbook.addWorksheet(wb, "Alpha", {
      properties: { tabColor: { argb: "FFFF0000" } }
    });
    const b = Workbook.addWorksheet(wb, "Beta");
    a.addRow([1]);
    b.addRow([2]);
    expect(Workbook.getWorksheet(wb, "alpha")).toBe(a);
    expect(Workbook.getWorksheet(wb, b.id)).toBe(b);
    expect(Workbook.getWorksheet(wb)).toBe(a);
    expect(Workbook.listWorksheets(wb)).toEqual([a, b]);
    const seen: string[] = [];
    Workbook.eachWorksheet(wb, (sheet, id) => seen.push(`${id}:${sheet.name}`));
    expect(seen).toEqual([`${a.id}:Alpha`, `${b.id}:Beta`]);

    const back = await readBack(wb);
    expect(Workbook.listWorksheets(back).map(Worksheet.getName)).toEqual(["Alpha", "Beta"]);
  });

  it("addWorksheet checks its options as for a workbook, and refuses a committed writer", async () => {
    const wb = new Stream.WorkbookWriter({});
    expect(() => Workbook.addWorksheet(wb, "S", { state: "gone" as never })).toThrow(
      ExcelArgumentError
    );
    Workbook.addWorksheet(wb, "S");
    await wb.toBuffer();
    expect(() => Workbook.addWorksheet(wb, "T")).toThrow(ExcelStreamStateError);
  });

  it("the deprecated addWorksheet / getWorksheet write what the namespace writes", async () => {
    const names = async (add: (wb: Stream.WorkbookWriter) => void) => {
      const wb = new Stream.WorkbookWriter({});
      add(wb);
      return Workbook.listWorksheets(await readBack(wb)).map(Worksheet.getName);
    };
    expect(
      await names(wb => {
        wb.addWorksheet("One");
        expect(wb.getWorksheet("one")).toBe(Workbook.getWorksheet(wb, "one"));
      })
    ).toEqual(await names(wb => void Workbook.addWorksheet(wb, "One")));
  });

  it("protect / unprotect / verifyPassword, and the deprecated protect writes the same", async () => {
    for (const protect of [
      (wb: Stream.WorkbookWriter) => Workbook.protect(wb, { password: "pw", lockStructure: true }),
      (wb: Stream.WorkbookWriter) => wb.protect({ password: "pw", lockStructure: true })
    ]) {
      const wb = new Stream.WorkbookWriter({});
      Workbook.addWorksheet(wb, "S");
      await protect(wb);
      expect(await Workbook.verifyPassword(wb, "pw")).toBe(true);
      const back = await readBack(wb);
      expect(await Workbook.verifyPassword(back, "pw")).toBe(true);
      expect(back.protection).toMatchObject({ lockStructure: true });
    }

    const wb = new Stream.WorkbookWriter({});
    Workbook.addWorksheet(wb, "S");
    await Workbook.protect(wb, { password: "pw" });
    Workbook.unprotect(wb);
    expect((await readBack(wb)).protection).toBeUndefined();
    await expect(Workbook.protect(wb)).rejects.toThrow(ExcelStreamStateError);
  });

  it("addCellStyle / getCellStyle / listCellStyles / useBuiltInCellStyle", async () => {
    const wb = new Stream.WorkbookWriter({ useStyles: true });
    const entry = Workbook.addCellStyle(wb, "Totals", { font: { bold: true } });
    expect(entry).toEqual({ font: { bold: true }, name: "Totals" });
    entry.font = { italic: true };
    expect(Workbook.getCellStyle(wb, "totals")).toEqual({ font: { bold: true }, name: "Totals" });
    expect(Workbook.useBuiltInCellStyle(wb, "Heading1")).toBe("Heading 1");
    expect(Workbook.listCellStyles(wb).map(style => style.name)).toEqual(["Totals", "Heading 1"]);
    // The deprecated method registers the same entry.
    expect(wb.addCellStyle("Other", { font: { bold: true } })).toEqual(
      Workbook.addCellStyle(new Stream.WorkbookWriter({}), "Other", { font: { bold: true } })
    );

    const sheet = Workbook.addWorksheet(wb, "S");
    Stream.applyCellStyle(Stream.getRowCell(sheet.addRow([1]), 1), "Totals");
    const back = await readBack(wb);
    expect(Workbook.getCellStyle(back, "Totals")?.font).toMatchObject({ bold: true });
    expect(() => Workbook.addCellStyle(wb, "Late", {})).toThrow(ExcelStreamStateError);
  });

  it("getDefinedNames answers with the writer's own names", () => {
    const wb = new Stream.WorkbookWriter({});
    expect(Workbook.getDefinedNames(wb)).toBe(wb.definedNames);
  });

  it("what only a workbook in memory can do is refused by name", async () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = Workbook.addWorksheet(wb, "S");
    const refusals: [string, () => unknown][] = [
      ["Workbook.removeWorksheet", () => Workbook.removeWorksheet(wb as never, "S")],
      ["Workbook.listChartsheets", () => Workbook.listChartsheets(wb as never)],
      ["Workbook.importWorksheet", () => Workbook.importWorksheet(wb as never, ws as never)],
      ["Workbook.getModel", () => Workbook.getModel(wb as never)],
      ["Workbook.setModel", () => Workbook.setModel(wb as never, {} as never)],
      ["Workbook.registerPerson", () => Workbook.registerPerson(wb as never, "Ada")],
      [
        "Workbook.addExternalLink",
        () => Workbook.addExternalLink(wb as never, { target: "x.xlsx" })
      ],
      ["Workbook.removeCellStyle", () => Workbook.removeCellStyle(wb as never, "Totals")],
      ["Workbook.registerFunction", () => Workbook.registerFunction(wb as never, "F", {} as never)],
      ["Workbook.toBuffer", () => Workbook.toBuffer(wb as never)],
      ["Workbook.toStream", () => Workbook.toStream(wb as never)]
    ];
    for (const [operation, call] of refusals) {
      let error: unknown;
      try {
        await call();
      } catch (thrown) {
        error = thrown;
      }
      expect(error, operation).toBeInstanceOf(ExcelUnsupportedFeatureError);
      expect((error as ExcelUnsupportedFeatureError).operation).toBe(operation);
    }
    // The writer is untouched, and finishes as it always does.
    expect(Workbook.listWorksheets(await readBack(wb)).map(Worksheet.getName)).toEqual(["S"]);
  });
});

describe("Column functions on a streaming worksheet writer", () => {
  it("edits the columns until they are written, and the file has them", async () => {
    const wb = new Stream.WorkbookWriter({ useStyles: true });
    const ws = Workbook.addWorksheet(wb, "S");
    Column.setDefinitions(ws, [{ header: "Name", key: "name", width: 12 }, { header: "Total" }]);
    Column.setKey(ws, 2, "total");
    Column.setWidth(ws, "total", 20);
    Column.setHidden(ws, 3, true);
    Column.setOutlineLevel(ws, "D", 2);
    Column.setFont(ws, "name", { bold: true });
    Column.addPageBreak(ws, "B");
    expect(Column.getWidth(ws, "total")).toBe(20);
    expect(Column.getKey(ws, 2)).toBe("total");
    expect(Column.getNumber(ws, "total")).toBe(2);
    Stream.commitRow(ws.addRow(["Ada", 3]));

    // Written with that row: an edit now is refused, a read still answers.
    expect(() => Column.setWidth(ws, "A", 5)).toThrow(ExcelStreamStateError);
    expect(() => Column.setDefinitions(ws, [])).toThrow(ExcelStreamStateError);
    expect(() => Column.getStyle(ws, "A")).toThrow(ExcelStreamStateError);
    expect(Column.getWidth(ws, "A")).toBe(12);
    expect(Column.getFont(ws, "name")).toEqual({ bold: true });
    // A key is not written into the file, so it can still be set: later rows are read by it.
    Column.setKey(ws, "C", "note");
    Stream.commitRow(ws.addRow({ name: "Bo", total: 4, note: "late" }));

    const back = Workbook.getWorksheet(await readBack(wb), "S")!;
    expect(Cell.getValue(back, "C3")).toBe("late");
    expect(Column.getWidth(back, "A")).toBe(12);
    expect(Column.getWidth(back, "B")).toBe(20);
    expect(Column.isHidden(back, "C")).toBe(true);
    expect(Column.getOutlineLevel(back, "D")).toBe(2);
    expect(back.colBreaks.map(item => item.id)).toEqual([2]);
  });

  it("the column edits refuse a committed sheet", () => {
    const ws = Workbook.addWorksheet(new Stream.WorkbookWriter({}), "S");
    ws.commit();
    expect(() => Column.setWidth(ws, "A", 5)).toThrow(ExcelStreamStateError);
    expect(() => Column.addPageBreak(ws, "A")).toThrow(ExcelStreamStateError);
  });

  it("what reads or writes the cells of a column is refused by name", () => {
    const ws = Workbook.addWorksheet(new Stream.WorkbookWriter({}), "S");
    ws.addRow([1, 2]);
    const refusals: [string, () => unknown][] = [
      ["Column.getValues", () => Column.getValues(ws as never, "A")],
      ["Column.setValues", () => Column.setValues(ws as never, "A", [1])],
      ["Column.splice", () => Column.splice(ws as never, 1, 1)],
      ["Column.each", () => Column.each(ws as never, () => undefined)],
      ["Column.eachAddress", () => Column.eachAddress(ws as never, "A", () => undefined)],
      ["Column.has", () => Column.has(ws as never, "A")],
      ["Column.getCount", () => Column.getCount(ws as never)],
      ["Column.getLast", () => Column.getLast(ws as never)],
      ["Column.getActualCount", () => Column.getActualCount(ws as never)],
      ["Column.autoFit", () => Column.autoFit(ws as never, "A")]
    ];
    for (const [operation, call] of refusals) {
      expect(call, operation).toThrow(ExcelUnsupportedFeatureError);
    }
  });
});

describe("a named style registered on a writer", () => {
  it("keeps its registered name when the caller edits the returned entry", async () => {
    const wb = new Stream.WorkbookWriter({ useStyles: true });
    const entry = wb.addCellStyle("Totals", { font: { bold: true } });
    entry.name = "Renamed";
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.applyStyle(ws, "A1", "Totals");
    expect(Workbook.getCellStyle(wb, "Totals")?.name).toBe("Totals");
    const back = await readBack(wb);
    expect(Cell.getStyle(Workbook.getWorksheet(back, "S")!, "A1").styleName).toBe("Totals");
  });
});
