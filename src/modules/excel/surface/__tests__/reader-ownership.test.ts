/**
 * Excel's readers hand out the caller's own copy: editing what `Cell.getValue`, the values readers, `Cell.getNote`,
 * `Cell.getResult` or `Cell.getModel` return changes nothing, and a value given to a setter is not shared with the
 * cell afterwards. `getStyle` is the one live reader, by design.
 */
import { Cell, Column, Note, Range, Row, Stream, Workbook, Worksheet } from "@excel";
import type { CellRichTextValue, NoteConfig } from "@excel";
import { describe, expect, it } from "vitest";

function sheet() {
  return Workbook.addWorksheet(Workbook.create(), "S");
}

const RICH = (): CellRichTextValue => ({ richText: [{ text: "a", font: { bold: true } }] });

describe("values", () => {
  it("Cell.getValue copies rich text, dates and errors, as it already did hyperlinks and formulas", () => {
    const ws = sheet();
    Cell.setValue(ws, "A1", RICH());
    Cell.setValue(ws, "A2", new Date(Date.UTC(2024, 0, 2)));
    Cell.setValue(ws, "A3", { error: "#N/A" });

    (Cell.getValue(ws, "A1") as CellRichTextValue).richText[0]!.text = "edited";
    (Cell.getValue(ws, "A2") as Date).setUTCFullYear(1999);
    (Cell.getValue(ws, "A3") as { error: string }).error = "#REF!";

    expect(Cell.getValue(ws, "A1")).toEqual(RICH());
    expect(Cell.getValue(ws, "A2")).toEqual(new Date(Date.UTC(2024, 0, 2)));
    expect(Cell.getValue(ws, "A3")).toEqual({ error: "#N/A" });
  });

  it("a value given to Cell.setValue is not the cell's afterwards", () => {
    const ws = sheet();
    const value = RICH();
    Cell.setValue(ws, "A1", value);
    Cell.setValue(ws, "B1", value);
    value.richText[0]!.text = "edited";
    expect(Cell.getValue(ws, "A1")).toEqual(RICH());
    expect(Cell.getValue(ws, "B1")).toEqual(RICH());
  });

  it("a Date, a formula's cached result and a JSON value given to Cell.setValue are copied in", () => {
    const ws = sheet();
    const date = new Date(Date.UTC(2024, 0, 2));
    const result = new Date(Date.UTC(2024, 0, 3));
    const json = { a: 1 };
    Cell.setValue(ws, "A1", date);
    Cell.setValue(ws, "A2", { formula: "TODAY()", result });
    Cell.setValue(ws, "A3", json as never);
    date.setUTCFullYear(1999);
    result.setUTCFullYear(1999);
    json.a = 2;
    expect(Cell.getValue(ws, "A1")).toEqual(new Date(Date.UTC(2024, 0, 2)));
    expect(Cell.getResult(ws, "A2")).toEqual(new Date(Date.UTC(2024, 0, 3)));
    expect(Cell.getValue(ws, "A3")).toEqual({ a: 1 });
  });

  it("the values readers copy too", () => {
    const ws = sheet();
    Cell.setValue(ws, "A1", RICH());
    const edit = (v: unknown) => {
      (v as CellRichTextValue).richText[0]!.text = "edited";
    };
    edit(Row.getValues(ws, 1)[0]);
    edit(Column.getValues(ws, "A")[0]);
    edit(Range.getValues(ws, "A1:A1")[0]![0]);
    edit(Worksheet.getValues(ws, { base: 0 })[0]![0]);
    edit(Worksheet.getValues(ws, { base: 1 })[1]![1]);
    edit(Worksheet.toAoa(ws)[0]![0]);
    expect(Cell.getValue(ws, "A1")).toEqual(RICH());
  });

  it("Cell.getResult copies a cached date result", () => {
    const ws = sheet();
    const date = new Date(Date.UTC(2024, 0, 2));
    Cell.setValue(ws, "A1", { formula: "TODAY()", result: date });
    (Cell.getResult(ws, "A1") as Date).setUTCFullYear(1999);
    expect(Cell.getResult(ws, "A1")).toEqual(date);
  });

  it("a streamed cell's value is a copy", () => {
    const writer = new Stream.WorkbookWriter({});
    const ws = Workbook.addWorksheet(writer, "S");
    const row = ws.addRow([RICH()]);
    const cell = Stream.getRowCell(row, 1);
    (Stream.getCellValue(cell) as CellRichTextValue).richText[0]!.text = "edited";
    expect(Stream.getCellValue(cell)).toEqual(RICH());
  });
});

describe("notes", () => {
  const CONFIG = (): NoteConfig => ({ texts: [{ text: "check" }], width: 120 });

  it("Cell.getNote copies, and the configuration given to setNote is not shared", () => {
    const ws = sheet();
    const config = CONFIG();
    Cell.setNote(ws, "A1", config);
    config.width = 999;
    (Cell.getNote(ws, "A1") as NoteConfig).texts![0]!.text = "edited";
    expect(Cell.getNote(ws, "A1")).toEqual(CONFIG());
  });

  it("carries the author, and removeNote takes the note off", () => {
    const ws = sheet();
    Cell.setNote(ws, "A1", "hi", { author: "Alice" });
    expect(Cell.getNote(ws, "A1")).toBe("hi");
    expect(Cell.getNoteAuthor(ws, "A1")).toBe("Alice");
    Cell.setNote(ws, "A1", "again");
    expect(Cell.getNoteAuthor(ws, "A1")).toBeUndefined();
    expect(Cell.removeNote(ws, "A1")).toBe(true);
    expect(Cell.removeNote(ws, "A1")).toBe(false);
    expect(Cell.getNote(ws, "A1")).toBeUndefined();
  });

  it("reading creates nothing", () => {
    const ws = sheet();
    expect(Cell.getNote(ws, "C9")).toBeUndefined();
    expect(Cell.getNoteAuthor(ws, "C9")).toBeUndefined();
    expect(Cell.removeNote(ws, "C9")).toBe(false);
    expect(Cell.has(ws, "C9")).toBe(false);
  });

  it("checks its arguments", () => {
    const ws = sheet();
    expect(() => Cell.setNote(ws, "A1", "x", { author: 5 as never })).toThrow(
      /^Cell\.setNote: options\.author/
    );
    expect(() => Cell.setNote(ws, "A1", "x", 5 as never)).toThrow(/^Cell\.setNote: options/);
  });

  it("round-trips the author", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setNote(ws, "A1", "hi", { author: "Alice" });
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Cell.getNoteAuthor(back, "A1")).toBe("Alice");
  });

  it("a streamed cell's note takes an author and can be removed", () => {
    const writer = new Stream.WorkbookWriter({});
    const ws = Workbook.addWorksheet(writer, "S");
    const cell = Stream.getRowCell(ws.addRow([1]), 1);
    Stream.setCellNote(cell, "hi", { author: "Bob" });
    expect(Stream.removeCellNote(cell)).toBe(true);
    expect(Stream.removeCellNote(cell)).toBe(false);
  });
});

describe("Cell.getModel", () => {
  it("is a copy, and creates nothing", () => {
    const ws = sheet();
    Cell.setValue(ws, "A1", RICH());
    Cell.setFont(ws, "A1", { bold: true });
    const model = Cell.getModel(ws, "A1");
    (model.value as CellRichTextValue).richText[0]!.text = "edited";
    model.style!.font!.bold = false;
    expect(Cell.getValue(ws, "A1")).toEqual(RICH());
    expect(Cell.getFont(ws, "A1")).toEqual({ bold: true });
    Cell.getModel(ws, "Z99");
    expect(Cell.has(ws, "Z99")).toBe(false);
  });

  it("a note removed after a read or a save stays removed", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setNote(ws, "A1", "x");
    Cell.getModel(ws, "A1");
    await Workbook.toBuffer(wb);
    Cell.removeNote(ws, "A1");
    expect(Cell.getModel(ws, "A1").comment).toBeUndefined();
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Cell.getNote(back, "A1")).toBeUndefined();
  });

  it("round-trips through setModel", () => {
    const ws = sheet();
    Cell.setValue(ws, "A1", RICH());
    Cell.setNote(ws, "A1", "n", { author: "Al" });
    const other = sheet();
    Cell.setModel(other, "A1", Cell.getModel(ws, "A1"));
    expect(Cell.getValue(other, "A1")).toEqual(RICH());
    expect(Cell.getNote(other, "A1")).toEqual(Cell.getNote(ws, "A1"));
  });
});

describe("the deprecated comment API behaves as before", () => {
  it("setComment stores the record it is given, and getComment hands that record back", () => {
    const ws = sheet();
    const record = Note.create("hi", { author: "Al" });
    Cell.setComment(ws, "A1", record);
    expect(Cell.getComment(ws, "A1")).toBe(record);
    expect(Cell.getNote(ws, "A1")).toBe("hi");
    expect(Cell.getNoteAuthor(ws, "A1")).toBe("Al");
  });

  it("setComment takes a bare configuration, and undefined clears", () => {
    const ws = sheet();
    Cell.setComment(ws, "A1", { texts: [{ text: "t" }] });
    expect(Cell.getComment(ws, "A1")).toEqual({
      note: { texts: [{ text: "t" }] },
      author: undefined
    });
    Cell.setComment(ws, "A1", undefined);
    expect(Cell.getComment(ws, "A1")).toBeUndefined();
  });

  it("Stream.setCellComment still sets and clears", () => {
    const writer = new Stream.WorkbookWriter({});
    const ws = Workbook.addWorksheet(writer, "S");
    const cell = Stream.getRowCell(ws.addRow([1]), 1);
    Stream.setCellComment(cell, Note.create("x", { author: "A" }));
    expect(Stream.removeCellNote(cell)).toBe(true);
    Stream.setCellComment(cell, Note.create("y"));
    Stream.setCellComment(cell, undefined);
    expect(Stream.removeCellNote(cell)).toBe(false);
  });
});
