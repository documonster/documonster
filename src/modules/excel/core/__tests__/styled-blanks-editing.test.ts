/**
 * Editing a sheet read with `blankCells: "collapse"` must not tell it apart from one read with `"keep"`.
 *
 * The collapse keeps a formatted region's empty cells as rectangles instead of cells, and that is all it may change: the
 * cells are still there, with their format, for every reader of a format and for every edit. Each case here reads the
 * same file both ways, makes the same edit to both, and requires the same answers from the cells and the same file back.
 * What is allowed to differ is what the option exists to change — how many cell objects the sheet holds — so `Cell.has`
 * and the iterators are not compared.
 */
import { Cell, Column, Range, Row, Table, Workbook, Worksheet } from "@excel";
import { ExcelFormula } from "@excel/bridge/formula";
import { describe, expect, it } from "vitest";

const YELLOW = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } } as const;
const RED = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } } as const;

/** Every cell the comparisons read: past the formatted region in both directions, so a moved block is seen. */
const GRID = "A1:L24";

/**
 * Values in A1:C3, a fill over A1:H12, a second format over B6:D8 and a value further down — so the collapse leaves
 * several rectangles, some sharing rows with values, and an edit can land inside, beside or across them.
 */
async function source(format: "xlsx" | "xlsb"): Promise<Uint8Array> {
  const handle = Workbook.create();
  const sheet = Workbook.addWorksheet(handle, "S");
  Range.setValues(sheet, "A1:C3", [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ]);
  Range.setStyle(sheet, "A1:H12", { fill: YELLOW });
  Range.setStyle(sheet, "B6:D8", { numFmt: "0.00", font: { bold: true }, fill: RED });
  Cell.setValue(sheet, "A15", "after");
  return Uint8Array.from(
    await Workbook.toBuffer(
      handle,
      format === "xlsx" ? { format, validate: false } : { format, unsupported: "ignore" }
    )
  );
}

async function readAs(
  bytes: Uint8Array,
  blankCells: "keep" | "collapse"
): Promise<{ handle: Workbook.Handle; sheet: Worksheet.Handle }> {
  const handle = Workbook.create();
  await Workbook.read(handle, bytes, { blankCells });
  return { handle, sheet: Workbook.listWorksheets(handle)[0] };
}

/** What a caller can learn about each cell of {@link GRID} without creating anything. */
function cellsOf(sheet: Worksheet.Handle): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const [from, to] = GRID.split(":");
  const first = { col: from.charCodeAt(0) - 64, row: Number(from.slice(1)) };
  const last = { col: to.charCodeAt(0) - 64, row: Number(to.slice(1)) };
  for (let row = first.row; row <= last.row; row++) {
    for (let col = first.col; col <= last.col; col++) {
      const at = { row, col };
      out[`${String.fromCharCode(64 + col)}${row}`] = {
        value: Cell.getValue(sheet, at),
        style: Cell.findStyle(sheet, at),
        numFmt: Cell.getNumFmt(sheet, at),
        fill: Cell.getFill(sheet, at),
        font: Cell.getFont(sheet, at),
        note: Cell.getNote(sheet, at)
      };
    }
  }
  return out;
}

/** The file written back, read with `"keep"` so both sides are compared as cells. */
async function writtenCells(
  handle: Workbook.Handle,
  format: "xlsx" | "xlsb"
): Promise<{ cells: Record<string, unknown>; merges: unknown }> {
  const bytes = await Workbook.toBuffer(
    handle,
    format === "xlsx" ? { format, validate: false } : { format, unsupported: "ignore" }
  );
  const { sheet } = await readAs(Uint8Array.from(bytes), "keep");
  return { cells: cellsOf(sheet), merges: Worksheet.getModel(sheet).mergeCells };
}

const EDITS: Record<string, (sheet: Worksheet.Handle) => void> = {
  "Cell.setValue inside a block": sheet => Cell.setValue(sheet, "E5", 42),
  "Cell.setValue inside the second format": sheet => Cell.setValue(sheet, "C7", 1.5),
  "Cell.setFont": sheet => Cell.setFont(sheet, "C7", { italic: true }),
  "Cell.setNote": sheet => Cell.setNote(sheet, "F10", "a note"),
  "Cell.getStyle, edited in place": sheet => {
    Cell.getStyle(sheet, "G4").numFmt = "0%";
  },
  "Range.setValues across the edge": sheet =>
    Range.setValues(sheet, "G11:J13", [
      [1, 2, 3, 4],
      [5, 6, 7, 8],
      [9, 10, 11, 12]
    ]),
  "Range.setStyle": sheet => Range.setStyle(sheet, "C2:E9", { border: { top: { style: "thin" } } }),
  "Row.setStyle": sheet => Row.setStyle(sheet, 7, { font: { underline: true } }),
  "Row.setFill": sheet => Row.setFill(sheet, 4, RED),
  "Row.setValues": sheet => Row.setValues(sheet, 9, [1, 2]),
  "Row.spliceCells": sheet => Row.spliceCells(sheet, 6, 2, 1, "x", "y"),
  "Row.insert inside": sheet => Row.insert(sheet, 5, ["new"]),
  "Row.insert copying the style above, empty cells too": sheet =>
    Row.insert(sheet, 8, ["n"], { copyStyle: "above", includeEmptyCells: true }),
  "Row.insertMany copying the style below": sheet =>
    Row.insertMany(sheet, 6, [["a"], ["b"]], { copyStyle: "below", includeEmptyCells: true }),
  "Row.splice deleting across": sheet => Row.splice(sheet, 3, 4),
  "Row.splice replacing": sheet => Row.splice(sheet, 5, 2, ["r1"], ["r2"], ["r3"]),
  "Row.duplicate inserting": sheet => Row.duplicate(sheet, 7, 2, { insert: true }),
  "Row.duplicate overwriting": sheet => Row.duplicate(sheet, 7, 2),
  "Column.splice inserting": sheet => Column.splice(sheet, 3, 0, ["c1", "c2"]),
  "Column.splice deleting across": sheet => Column.splice(sheet, 2, 2),
  "Column.setStyle": sheet => Column.setStyle(sheet, "E", { numFmt: "0%" }),
  "Column.setValues": sheet => Column.setValues(sheet, "D", [1, 2, 3, 4, 5, 6]),
  "Worksheet.merge": sheet => Worksheet.merge(sheet, "E6:F8"),
  "Worksheet.merge, then unmerge": sheet => {
    Worksheet.merge(sheet, "E6:F8");
    Worksheet.unmerge(sheet, "E6:F8");
  },
  "Worksheet.setModel with its own model": sheet =>
    Worksheet.setModel(sheet, Worksheet.getModel(sheet)),
  "Table.add over the region": sheet => {
    Table.add(sheet, {
      name: "T",
      ref: "E9",
      headerRow: true,
      columns: [{ name: "One" }, { name: "Two" }],
      rows: [
        [1, 2],
        [3, 4]
      ]
    });
  }
};

describe.each(["xlsx", "xlsb"] as const)("a collapsed %s sheet, edited", format => {
  it("reads as the kept one before any edit", async () => {
    const bytes = await source(format);
    const [kept, collapsed] = await Promise.all([readAs(bytes, "keep"), readAs(bytes, "collapse")]);
    // The fixture has to be one the option changes, or every comparison below passes vacuously.
    expect(Worksheet.getModel(collapsed.sheet).styledBlankRanges?.length).toBeGreaterThan(0);
    expect(cellsOf(collapsed.sheet)).toEqual(cellsOf(kept.sheet));
  });

  it.each(Object.entries(EDITS))("%s", async (_name, edit) => {
    const bytes = await source(format);
    const [kept, collapsed] = await Promise.all([readAs(bytes, "keep"), readAs(bytes, "collapse")]);
    edit(kept.sheet);
    edit(collapsed.sheet);
    expect(cellsOf(collapsed.sheet)).toEqual(cellsOf(kept.sheet));
    expect(await writtenCells(collapsed.handle, format)).toEqual(
      await writtenCells(kept.handle, format)
    );
  });

  it("all of them, one after another", async () => {
    // Edits that interact: a row settled by one is moved by the next, a rectangle split by one is cut by another.
    const bytes = await source(format);
    const [kept, collapsed] = await Promise.all([readAs(bytes, "keep"), readAs(bytes, "collapse")]);
    for (const edit of Object.values(EDITS)) {
      // Some refuse what an earlier one left (a second merge over the first); they must refuse it alike.
      const outcome = (sheet: Worksheet.Handle): string | undefined => {
        try {
          edit(sheet);
          return undefined;
        } catch (error) {
          return String(error);
        }
      };
      expect(outcome(collapsed.sheet)).toBe(outcome(kept.sheet));
    }
    expect(cellsOf(collapsed.sheet)).toEqual(cellsOf(kept.sheet));
    expect(await writtenCells(collapsed.handle, format)).toEqual(
      await writtenCells(kept.handle, format)
    );
  });

  it("spills a dynamic array into the region", async () => {
    // The formula engine writes through its own path, and journals the cells it writes so it can put them back.
    const bytes = await source(format);
    const [kept, collapsed] = await Promise.all([readAs(bytes, "keep"), readAs(bytes, "collapse")]);
    for (const { handle, sheet } of [kept, collapsed]) {
      Cell.setValue(sheet, "E5", { formula: "SEQUENCE(3,2)" });
      ExcelFormula.calculate(handle);
    }
    expect(Cell.getValue(collapsed.sheet, "F7")).toBe(6);
    expect(cellsOf(collapsed.sheet)).toEqual(cellsOf(kept.sheet));
    expect(await writtenCells(collapsed.handle, format)).toEqual(
      await writtenCells(kept.handle, format)
    );
  });

  it("leaves the rest of a large region collapsed after an edit", async () => {
    // The saving the option exists for, kept through an edit: one row given back, not the region.
    const bytes = await source(format);
    const collapsed = await readAs(bytes, "collapse");
    Cell.setValue(collapsed.sheet, "E5", 42);
    const ranges = Worksheet.getModel(collapsed.sheet).styledBlankRanges ?? [];
    const covered = ranges.reduce(
      (cells, range) =>
        cells + (range.lastRow - range.firstRow + 1) * (range.lastColumn - range.firstColumn + 1),
      0
    );
    expect(covered).toBeGreaterThan(50);
    expect(Cell.has(collapsed.sheet, "E4")).toBe(false);
    expect(Cell.has(collapsed.sheet, "A5")).toBe(true);
  });
});
