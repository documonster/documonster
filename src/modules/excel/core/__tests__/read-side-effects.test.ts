/**
 * Reads must not mutate the workbook.
 *
 * `getCell` / `getRow` / `rowGetCell` / `getColumn` all *create* the thing they
 * are asked for, and `cellCreate` seeds a new cell with
 * `mergeCellStyle(row.style, column.style, {})` — so a materialised "empty" cell in
 * a styled row or column carries a style, gets a `styleId`, and is written out
 * (`xlsx/xform/sheet/cell-xform.ts` only skips a cell when it is both `Null`
 * and unstyled). Even unstyled, `rowGetModel` counts every existing cell record
 * towards the row's `min`/`max`, so `spans` and `<dimension>` widen too.
 *
 * The net effect is that calling a getter changed the bytes the workbook saved
 * to. Every case below is a read, so every case must be a no-op.
 */
import { captureFormulaSnapshot } from "@excel/core/formula-capture";
import {
  Anchor,
  Cell,
  Column,
  DataValidation,
  DefinedNames,
  Row,
  ValueType,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, it, expect } from "vitest";

/**
 * Everything about a workbook that a read must leave alone: the saved bytes,
 * plus the counters and container lengths that materialisation inflates but
 * serialisation happens to hide.
 */
async function fingerprint(wb: ReturnType<typeof Workbook.create>) {
  const sheets = Workbook.getWorksheets(wb).map(ws => ({
    name: Worksheet.getName(ws),
    rowCount: Worksheet.rowCount(ws),
    columnCount: Worksheet.columnCount(ws),
    columns: Worksheet.columns(ws).length,
    dimensions: { ...Worksheet.dimensions(ws) }
  }));
  const bytes = new Uint8Array(await Workbook.toBuffer(wb));
  return { sheets, bytes };
}

/**
 * A sheet shaped to make materialisation observable: column B is styled but
 * empty, and row 1 is wider than row 2, so a full-rectangle read has holes to
 * fill in a styled column.
 */
function styledSheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Column.setStyle(ws, 2, { numFmt: "0.00%" });
  Row.setStyle(ws, 3, { font: { bold: true } });
  Cell.setValue(ws, "A1", "h1");
  Cell.setValue(ws, "C1", "h3");
  Cell.setValue(ws, "A2", 1);
  Cell.setValue(ws, "C4", 4);
  return { wb, ws };
}

async function expectNoMutation(
  wb: ReturnType<typeof Workbook.create>,
  read: () => void | Promise<void>
) {
  const before = await fingerprint(wb);
  await read();
  const after = await fingerprint(wb);
  expect(after.sheets).toEqual(before.sheets);
  expect(after.bytes).toEqual(before.bytes);
}

describe("reads do not mutate the workbook", () => {
  describe("Worksheet.toJson", () => {
    it("header: 1", async () => {
      const { wb, ws } = styledSheet();
      let values: unknown;
      await expectNoMutation(wb, () => {
        values = Worksheet.toJson(ws, { header: 1 });
      });
      expect(values).toEqual([
        ["h1", null, "h3"],
        [1, null, null],
        [null, null, null],
        [null, null, 4]
      ]);
    });

    it("header: 1 with raw: false", async () => {
      const { wb, ws } = styledSheet();
      let values: unknown;
      await expectNoMutation(wb, () => {
        values = Worksheet.toJson(ws, { header: 1, raw: false });
      });
      expect(values).toEqual([
        ["h1", null, "h3"],
        ["1", null, null],
        [null, null, null],
        [null, null, "4"]
      ]);
    });

    it('header: "A"', async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => void Worksheet.toJson(ws, { header: "A" }));
    });

    it("header: string[]", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => void Worksheet.toJson(ws, { header: ["a", "b", "c"] }));
    });

    it("default first-row header", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => void Worksheet.toJson(ws));
    });

    it("with defaultValue and blankRows", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(
        wb,
        () => void Worksheet.toJson(ws, { header: 1, defaultValue: 0, blankRows: false })
      );
    });
  });

  describe("Worksheet.toAoa", () => {
    it("leaves the sheet alone", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => void Worksheet.toAoa(ws));
    });

    it("leaves the sheet alone when rows have been spliced out", async () => {
      const { wb, ws } = styledSheet();
      Worksheet.spliceRows(ws, 2, 1);
      await expectNoMutation(wb, () => void Worksheet.toAoa(ws));
    });

    it("preserves explicit empty rows rather than sparse array holes", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A3", 3);

      let values: ReturnType<typeof Worksheet.toAoa> | undefined;
      await expectNoMutation(wb, () => {
        values = Worksheet.toAoa(ws);
      });

      expect(values).toEqual([[1], [], [3]]);
      expect(Object.keys(values!)).toEqual(["0", "1", "2"]);
      expect(1 in values!).toBe(true);
    });

    it("keeps rows positioned by slot after a splice", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Cell.setValue(ws, "A3", 3);
      Worksheet.spliceRows(ws, 2, 1);

      let values: ReturnType<typeof Worksheet.toAoa> | undefined;
      await expectNoMutation(wb, () => {
        values = Worksheet.toAoa(ws);
      });

      // Row 2 now holds the old row 3, and the trailing slot is empty.
      expect(values).toEqual([[1], [3], []]);
    });

    it("does not change what a later Worksheet.getValues returns", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Cell.setValue(ws, "A3", 3);
      Worksheet.spliceRows(ws, 2, 1);

      // The materialising version of `toAoa` created the trailing row that the
      // splice had emptied, so reading the sheet one way changed what reading it
      // another way reported — `getValues` grew from 3 entries to 4.
      const before = Worksheet.getValues(ws);
      Worksheet.toAoa(ws);
      expect(Worksheet.getValues(ws)).toEqual(before);
    });
  });

  describe("Worksheet.getValues / Range.getValues", () => {
    it("leaves the sheet alone", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => void Worksheet.getValues(ws));
    });
  });

  describe("column reads", () => {
    it("Column.values does not materialise cells down the column", async () => {
      const { wb, ws } = styledSheet();
      let values: ReturnType<typeof Column.values> | undefined;
      await expectNoMutation(wb, () => {
        values = Column.values(ws, 3);
      });
      expect(values).toEqual([, "h3", , , 4]);
    });

    it("Column.getValues does not materialise cells down the column", async () => {
      const { wb, ws } = styledSheet();
      let values: ReturnType<typeof Column.getValues> | undefined;
      await expectNoMutation(wb, () => {
        values = Column.getValues(ws, 3);
      });
      expect(values).toEqual(["h3", , , 4]);
    });

    it("reading a column that was never declared does not declare it", async () => {
      const { wb, ws } = styledSheet();
      // Every other `Column.*` member resolves its reference through
      // `getColumn`, which would pad `ws._columns` out to column 26 here.
      await expectNoMutation(wb, () => {
        expect(Column.values(ws, "Z")).toEqual([]);
        expect(Column.getValues(ws, 100)).toEqual([]);
      });
    });
  });

  describe("formula snapshot capture", () => {
    it("captures without materialising", async () => {
      const { wb } = styledSheet();
      await expectNoMutation(wb, () => void captureFormulaSnapshot(wb));
    });

    it("captures without materialising after a splice leaves row holes", async () => {
      const { wb, ws } = styledSheet();
      Worksheet.spliceRows(ws, 2, 1);
      await expectNoMutation(wb, () => void captureFormulaSnapshot(wb));
    });
  });

  describe("Cell value readers", () => {
    // B is a styled column and row 3 a styled row, so a cell created at B3, B10 or
    // E3 would carry a style and be written out — the case that changed the bytes.
    const absent = ["B3", "B10", "E3", "Z100"] as const;

    it("answer for a missing cell as for an empty one, without creating it", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => {
        for (const a of absent) {
          expect(Cell.getValue(ws, a)).toBeNull();
          expect(Cell.getText(ws, a)).toBe("");
          expect(Cell.getDisplayText(ws, a)).toBe("");
          expect(Cell.getType(ws, a)).toBe(ValueType.Null);
          expect(Cell.getEffectiveType(ws, a)).toBe(ValueType.Null);
          expect(Cell.getFormula(ws, a)).toBeUndefined();
          expect(Cell.getResult(ws, a)).toBeUndefined();
          expect(Cell.getDateParts(ws, a)).toBeUndefined();
          expect(Cell.getDateKind(ws, a)).toBeUndefined();
          expect(Cell.getTemporal(ws, a)).toBeUndefined();
          expect(Cell.isMerged(ws, a)).toBe(false);
          expect(Cell.getHyperlink(ws, a)).toBeUndefined();
          expect(Cell.getNote(ws, a)).toBeUndefined();
          expect(Cell.getComment(ws, a)).toBeUndefined();
          expect(Cell.getNames(ws, a)).toEqual([]);
          expect(Cell.getValidation(ws, a)).toBeUndefined();
          expect(Cell.find(ws, a)).toBeUndefined();
        }
      });
    });

    it("the (row, col) form is just as inert", async () => {
      const { wb, ws } = styledSheet();
      await expectNoMutation(wb, () => {
        expect(Cell.getValue(ws, 10, 2)).toBeNull();
        expect(Cell.getText(ws, 3, 5)).toBe("");
        expect(Cell.getNumFmt(ws, 10, 2)).toBe("0.00%");
        expect(Cell.getTemporal(ws, 10, 2)).toBeUndefined();
        expect(Cell.getFullAddress(ws, 10, 2)).toEqual({
          sheetName: "S",
          address: "B10",
          row: 10,
          col: 2
        });
      });
    });

    it("report the number format a missing cell would inherit", async () => {
      const { wb, ws } = styledSheet();
      Row.setStyle(ws, 7, { numFmt: "0.0" });
      await expectNoMutation(wb, () => {
        expect(Cell.getNumFmt(ws, "B10")).toBe("0.00%"); // column
        expect(Cell.getNumFmt(ws, "B7")).toBe("0.0"); // row wins, as in mergeCellStyle
        expect(Cell.getNumFmt(ws, "E10")).toBeUndefined();
      });
      // …and it is the format the cell really gets once it exists.
      Cell.setValue(ws, "B10", 0.5);
      Cell.setValue(ws, "B7", 0.5);
      expect(Cell.getNumFmt(ws, "B10")).toBe("0.00%");
      expect(Cell.getNumFmt(ws, "B7")).toBe("0.0");
    });

    it("read address-keyed metadata for a missing cell", async () => {
      const { wb, ws } = styledSheet();
      DefinedNames.add(Workbook.getDefinedNames(wb), "S!$E$1:$E$9", "inputs");
      DataValidation.add(ws.dataValidations, "E1:E9", {
        type: "whole",
        operator: "between",
        formulae: [1, 9]
      });
      await expectNoMutation(wb, () => {
        expect(Cell.getNames(ws, "E5")).toEqual(["inputs"]);
        expect(Cell.getValidation(ws, "E5")).toMatchObject({ type: "whole" });
        expect(Cell.getFullAddress(ws, "E5")).toEqual({
          sheetName: "S",
          address: "E5",
          row: 5,
          col: 5
        });
      });
    });

    it("reject a malformed reference exactly as a writer does", () => {
      const { ws } = styledSheet();
      for (const bad of ["a1", "1", "A", ""]) {
        expect(() => Cell.setValue(ws, bad, 1)).toThrow();
        expect(() => Cell.getValue(ws, bad)).toThrow();
        expect(() => Cell.find(ws, bad)).toThrow();
      }
      expect(() => Cell.getValue(ws, "XFE1")).toThrow();
    });

    it("still read existing cells", () => {
      const { ws } = styledSheet();
      expect(Cell.getValue(ws, "C4")).toBe(4);
      expect(Cell.getText(ws, 1, 1)).toBe("h1");
      expect(Cell.getType(ws, "A2")).toBe(ValueType.Number);
    });

    it("style-facet readers still create the cell, because their result is edited in place", () => {
      const { ws } = styledSheet();
      expect(Cell.find(ws, "D9")).toBeUndefined();
      Cell.getStyle(ws, "D9").font = { bold: true };
      expect(Cell.find(ws, "D9")).toBeDefined();
      expect(Cell.getFont(ws, "D9")).toEqual({ bold: true });
    });
  });

  describe("Anchor geometry", () => {
    it("reading anchor geometry does not create rows or columns", async () => {
      const { wb, ws } = styledSheet();
      const anchor = Anchor.create(ws, { col: 40, row: 60 });
      await expectNoMutation(wb, () => {
        Anchor.colWidth(anchor);
        Anchor.rowHeight(anchor);
        Anchor.col(anchor);
        Anchor.row(anchor);
      });
    });
  });
});
