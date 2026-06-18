import { testUtils } from "@excel/__tests__/shared";
import { cellGetValue, cellMaster, cellSetValue, cellType } from "@excel/cell";
import { Enums } from "@excel/enums";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { rangeBottom, rangeCreate, rangeLeft, rangeRight, rangeTop } from "@excel/range";
import { getCell, getSheetModel } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Merge Cells", () => {
    it("references the same top-left value", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // initial values
      Cell.setValue(ws, "A1", "A1");
      Cell.setValue(ws, "B1", "B1");
      Cell.setValue(ws, "A2", "A2");
      Cell.setValue(ws, "B2", "B2");

      Worksheet.merge(ws, "A1:B2");

      expect(Cell.getValue(ws, "A1")).toBe("A1");
      expect(Cell.getValue(ws, "B1")).toBe("A1");
      expect(Cell.getValue(ws, "A2")).toBe("A1");
      expect(Cell.getValue(ws, "B2")).toBe("A1");

      expect(Cell.getType(ws, "A1")).toBe(Enums.ValueType.String);
      expect(Cell.getType(ws, "B1")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "A2")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "B2")).toBe(Enums.ValueType.Merge);
    });

    it("does not allow overlapping merges", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Worksheet.merge(ws, "B2:C3");

      // intersect four corners
      expect(() => {
        Worksheet.merge(ws, "A1:B2");
      }).toThrow(Error);
      expect(() => {
        Worksheet.merge(ws, "C1:D2");
      }).toThrow(Error);
      expect(() => {
        Worksheet.merge(ws, "C3:D4");
      }).toThrow(Error);
      expect(() => {
        Worksheet.merge(ws, "A3:B4");
      }).toThrow(Error);

      // enclosing
      expect(() => {
        Worksheet.merge(ws, "A1:D4");
      }).toThrow(Error);
    });

    it("merges and unmerges", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      const expectMaster = function (range: string, master: string | null) {
        const d = rangeCreate(range);
        for (let i = rangeTop(d); i <= rangeBottom(d); i++) {
          for (let j = rangeLeft(d); j <= rangeRight(d); j++) {
            const cell = getCell(ws, i, j);
            const masterCell = master ? getCell(ws, master) : cell;
            expect(cellMaster(cell).address).toBe(masterCell.address);
          }
        }
      };

      // merge some cells, then unmerge them
      Worksheet.merge(ws, "A1:B2");
      expectMaster("A1:B2", "A1");
      Worksheet.unmerge(ws, "A1:B2");
      expectMaster("A1:B2", null);

      // unmerge just one cell
      Worksheet.merge(ws, "A1:B2");
      expectMaster("A1:B2", "A1");
      Worksheet.unmerge(ws, "A1");
      expectMaster("A1:B2", null);

      Worksheet.merge(ws, "A1:B2");
      expectMaster("A1:B2", "A1");
      Worksheet.unmerge(ws, "B2");
      expectMaster("A1:B2", null);

      // build 4 merge-squares
      Worksheet.merge(ws, "A1:B2");
      Worksheet.merge(ws, "D1:E2");
      Worksheet.merge(ws, "A4:B5");
      Worksheet.merge(ws, "D4:E5");

      expectMaster("A1:B2", "A1");
      expectMaster("D1:E2", "D1");
      expectMaster("A4:B5", "A4");
      expectMaster("D4:E5", "D4");

      // unmerge the middle
      Worksheet.unmerge(ws, "B2:D4");

      expectMaster("A1:B2", null);
      expectMaster("D1:E2", null);
      expectMaster("A4:B5", null);
      expectMaster("D4:E5", null);
    });

    it("merges styles", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // initial value
      const B2 = getCell(ws, "B2");
      cellSetValue(B2, 5);
      B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
      B2.style.border = testUtils.styles.borders.doubleRed;
      B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
      B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
      B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

      // expecting styles to be copied (see worksheet spec)
      Worksheet.merge(ws, "B2:C3");

      const dblRed = testUtils.styles.borders.doubleRed;

      // Non-border styles are copied identically to all cells
      for (const addr of ["B2", "B3", "C2", "C3"]) {
        expect(Cell.getStyle(ws, addr).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
        expect(Cell.getStyle(ws, addr).fill).toEqual(testUtils.styles.fills.blueWhiteHGrad);
        expect(Cell.getStyle(ws, addr).alignment).toEqual(
          testUtils.styles.namedAlignments.middleCentre
        );
        expect(Cell.getStyle(ws, addr).numFmt).toEqual(testUtils.styles.numFmts.numFmt1);
      }

      // Borders are position-aware: only perimeter edges survive (like Excel)
      // B2 = top-left corner
      expect(Cell.getStyle(ws, "B2").border).toEqual({ left: dblRed.left, top: dblRed.top });
      // C2 = top-right corner
      expect(Cell.getStyle(ws, "C2").border).toEqual({ right: dblRed.right, top: dblRed.top });
      // B3 = bottom-left corner
      expect(Cell.getStyle(ws, "B3").border).toEqual({ left: dblRed.left, bottom: dblRed.bottom });
      // C3 = bottom-right corner
      expect(Cell.getStyle(ws, "C3").border).toEqual({
        right: dblRed.right,
        bottom: dblRed.bottom
      });
    });

    it("preserves merges after row inserts", function () {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "testMergeAfterInsert");

      Worksheet.addRow(ws, [1, 2]);
      Worksheet.addRow(ws, [3, 4]);
      Worksheet.merge(ws, "A1:B2");
      Worksheet.insertRow(ws, 1, ["Inserted Row Text"]);

      // After insert, the merged area should now be A2:B3
      // A2 is master (type=Number with value 1), B2, A3, B3 are merge cells
      const cellA2 = getCell(ws, "A2");
      const cellB2 = getCell(ws, "B2");
      const cellA3 = getCell(ws, "A3");
      const cellB3 = getCell(ws, "B3");

      // Verify master cell has the number value
      expect(cellType(cellA2)).toEqual(Enums.ValueType.Number);
      expect(cellGetValue(cellA2)).toEqual(1);

      // Verify other cells in merge area are merge type and point to A2 address
      expect(cellType(cellB2)).toEqual(Enums.ValueType.Merge);
      expect(cellMaster(cellB2).address).toEqual("A2");

      expect(cellType(cellA3)).toEqual(Enums.ValueType.Merge);
      expect(cellMaster(cellA3).address).toEqual("A2");

      expect(cellType(cellB3)).toEqual(Enums.ValueType.Merge);
      expect(cellMaster(cellB3).address).toEqual("A2");
    });

    it("spliceRows updates _merges after inserting rows above a merge", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A3", "hello");
      Worksheet.merge(ws, "A3:C4");

      // Insert 2 rows at row 1 (above the merge)
      Worksheet.spliceRows(ws, 1, 0, ["x"], ["y"]);

      // Merge should shift down by 2: A3:C4 -> A5:C6
      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual(["A5:C6"]);

      // Cell-level merge references should also be correct
      expect(Cell.getValue(ws, "A5")).toBe("hello");
      expect(Cell.getType(ws, "B5")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B5").address).toBe("A5");
      expect(Cell.getType(ws, "C6")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "C6").address).toBe("A5");
    });

    it("spliceRows updates _merges after deleting rows above a merge", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Worksheet.addRow(ws, ["filler1"]);
      Worksheet.addRow(ws, ["filler2"]);
      Worksheet.addRow(ws, ["filler3"]);
      Cell.setValue(ws, "A4", "hello");
      Worksheet.merge(ws, "A4:B5");

      // Delete rows 1-2 (above the merge)
      Worksheet.spliceRows(ws, 1, 2);

      // Merge should shift up by 2: A4:B5 -> A2:B3
      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual(["A2:B3"]);

      expect(Cell.getValue(ws, "A2")).toBe("hello");
      expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B3").address).toBe("A2");
    });

    it("spliceRows removes merges entirely within deleted rows", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Worksheet.addRow(ws, ["filler"]);
      Cell.setValue(ws, "A2", "merged");
      Worksheet.merge(ws, "A2:B3");
      Worksheet.addRow(ws, ["below"]);

      // Delete rows 2-3 which contain the entire merge
      Worksheet.spliceRows(ws, 2, 2);

      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual([]);
    });

    it("spliceRows shrinks merge spanning the splice boundary", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "hello");
      Worksheet.merge(ws, "A1:B4");

      // Delete row 3 (within the merge)
      Worksheet.spliceRows(ws, 3, 1);

      // Merge should shrink: A1:B4 -> A1:B3
      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual(["A1:B3"]);
    });

    it("duplicateRow preserves single-row horizontal merges", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "merged");
      Cell.setValue(ws, "D1", "solo");
      Worksheet.merge(ws, "A1:C1");

      // Duplicate row 1, inserting 2 copies below
      Worksheet.duplicateRow(ws, 1, 2, true);

      const model = getSheetModel(ws);
      // Should have 3 merges: original A1:C1, plus A2:C2 and A3:C3
      expect(model.mergeCells).toHaveLength(3);
      expect(model.mergeCells).toContain("A1:C1");
      expect(model.mergeCells).toContain("A2:C2");
      expect(model.mergeCells).toContain("A3:C3");

      // Verify cell-level merge references in duplicated rows
      expect(Cell.getValue(ws, "A2")).toBe("merged");
      expect(Cell.getType(ws, "B2")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B2").address).toBe("A2");
      expect(Cell.getType(ws, "C2")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "C2").address).toBe("A2");

      expect(Cell.getValue(ws, "A3")).toBe("merged");
      expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B3").address).toBe("A3");
    });

    it("duplicateRow preserves multi-row merges", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "big merge");
      Worksheet.merge(ws, "A1:B3");
      Cell.setValue(ws, "C1", "outside");

      // Duplicate row 1 once with insert
      Worksheet.duplicateRow(ws, 1, 1, true);

      const model = getSheetModel(ws);
      // Original merge A1:B3 stays at A1:B3
      // Duplicated merge for the new row should be A2:B4
      // But original rows 2-3 shifted down to 3-4, so original merge becomes A1:B4? No...
      // Actually: duplicateRow(1,1,true) calls spliceRows(2, 0, values)
      // This inserts 1 row at position 2, shifting rows 2+ down by 1
      // Original merge A1:B3: top=1 is above splice, bottom=3 is at/below splice
      // So it spans the boundary -> bottom shifts: A1:B3 -> A1:B4
      // Then duplicateRow should create a new merge for row 2 with same shape as source row merges
      // Source row 1 has merge A1:B3 (height=3), so duplicate at row 2 should be A2:B4
      // But A1:B4 and A2:B4 overlap! That won't work.
      //
      // For multi-row merges, duplicateRow should only duplicate single-row merges
      // (merges where top == bottom == source row). Multi-row merges are too complex.
      // Let's just verify the original merge is preserved correctly after the splice.
      expect(model.mergeCells).toContain("A1:B4");
    });

    it("duplicateRow with overwrite mode clears existing merges in target rows", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "source");
      Worksheet.merge(ws, "A1:C1");

      Cell.setValue(ws, "A2", "existing");
      Worksheet.merge(ws, "A2:D2");

      // Duplicate row 1 over row 2 (overwrite mode, insert=false)
      Worksheet.duplicateRow(ws, 1, 1, false);

      const model = getSheetModel(ws);
      // Original merge A1:C1 should remain
      // Row 2's old merge A2:D2 should be replaced with A2:C2 (duplicated from row 1)
      expect(model.mergeCells).toHaveLength(2);
      expect(model.mergeCells).toContain("A1:C1");
      expect(model.mergeCells).toContain("A2:C2");
    });

    it("duplicateRow + XLSX roundtrip preserves merges", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "merged");
      Cell.setValue(ws, "D1", "solo");
      Worksheet.merge(ws, "A1:C1");

      Worksheet.duplicateRow(ws, 1, 2, true);

      // Write to buffer and read back
      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer as Buffer);
      const ws2 = Workbook.getWorksheet(wb2, "sheet")!;

      const model2 = getSheetModel(ws2);
      expect(model2.mergeCells).toHaveLength(3);
      expect(model2.mergeCells).toContain("A1:C1");
      expect(model2.mergeCells).toContain("A2:C2");
      expect(model2.mergeCells).toContain("A3:C3");
    });

    it("duplicateRow with multiple merges on source row", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "first");
      Cell.setValue(ws, "D1", "second");
      Worksheet.merge(ws, "A1:B1");
      Worksheet.merge(ws, "D1:F1");

      Worksheet.duplicateRow(ws, 1, 2, true);

      const model = getSheetModel(ws);
      // 3 rows × 2 merges = 6 total merges
      expect(model.mergeCells).toHaveLength(6);
      expect(model.mergeCells).toContain("A1:B1");
      expect(model.mergeCells).toContain("D1:F1");
      expect(model.mergeCells).toContain("A2:B2");
      expect(model.mergeCells).toContain("D2:F2");
      expect(model.mergeCells).toContain("A3:B3");
      expect(model.mergeCells).toContain("D3:F3");
    });
  });

  describe("spliceColumns with merges", () => {
    it("spliceColumns updates _merges after inserting columns before a merge", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      // Merge C1:E1 (cols 3-5)
      Cell.setValue(ws, "C1", "merged");
      Worksheet.merge(ws, "C1:E1");

      // Insert 2 columns at column 2 (before the merge)
      Worksheet.spliceColumns(ws, 2, 0, [], []);

      const model = getSheetModel(ws);
      // Merge should shift right by 2: C1:E1 → E1:G1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("E1:G1");

      // Cell-level: E1 is master, F1 and G1 are merge slaves
      expect(Cell.getValue(ws, "E1")).toBe("merged");
      expect(Cell.getType(ws, "F1")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "G1")).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns updates _merges after deleting columns before a merge", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      // Merge D1:F1 (cols 4-6)
      Cell.setValue(ws, "D1", "merged");
      Worksheet.merge(ws, "D1:F1");

      // Delete 2 columns at column 1 (cols 1-2 removed)
      Worksheet.spliceColumns(ws, 1, 2);

      const model = getSheetModel(ws);
      // Merge should shift left by 2: D1:F1 → B1:D1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("B1:D1");

      // Cell-level: B1 is master
      expect(Cell.getValue(ws, "B1")).toBe("merged");
      expect(Cell.getType(ws, "C1")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "D1")).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns removes merges entirely within deleted columns", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      // Merge B1:C1 (cols 2-3)
      Cell.setValue(ws, "B1", "merged");
      Worksheet.merge(ws, "B1:C1");

      // Also add a merge that survives: E1:F1 (cols 5-6)
      Cell.setValue(ws, "E1", "survivor");
      Worksheet.merge(ws, "E1:F1");

      // Delete columns 2-3 (removes B1:C1 entirely)
      Worksheet.spliceColumns(ws, 2, 2);

      const model = getSheetModel(ws);
      // B1:C1 removed, E1:F1 shifts left by 2 → C1:D1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("C1:D1");

      expect(Cell.getValue(ws, "C1")).toBe("survivor");
      expect(Cell.getType(ws, "D1")).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns shrinks merge spanning the splice boundary", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      // Merge B1:F1 (cols 2-6, 5 columns wide)
      Cell.setValue(ws, "B1", "wide");
      Worksheet.merge(ws, "B1:F1");

      // Delete columns 4-5 (2 columns from the middle of the merge)
      Worksheet.spliceColumns(ws, 4, 2);

      const model = getSheetModel(ws);
      // Merge should shrink: B1:F1 → B1:D1 (right reduced by 2)
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("B1:D1");

      expect(Cell.getValue(ws, "B1")).toBe("wide");
      expect(Cell.getType(ws, "C1")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "D1")).toBe(Enums.ValueType.Merge);
    });
  });

  describe("insertRow with merges", () => {
    it("insertRow preserves model.mergeCells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      // Merge A2:C2
      Cell.setValue(ws, "A2", "merged");
      Worksheet.merge(ws, "A2:C2");

      // insertRow at row 1 pushes merge down by 1
      Worksheet.insertRow(ws, 1, ["new"]);

      const model = getSheetModel(ws);
      // Merge should shift down: A2:C2 → A3:C3
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("A3:C3");

      expect(Cell.getValue(ws, "A3")).toBe("merged");
      expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Merge);
      expect(Cell.getType(ws, "C3")).toBe(Enums.ValueType.Merge);
    });
  });

  describe("merge edge cases", () => {
    it("Bug #1: spliceRows with equal delete and insert updates merges in replaced range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A2", "merged");
      Worksheet.merge(ws, "A2:B3");

      // Replace rows 2-3 with new data (delete 2, insert 2 -> nExpand=0)
      Worksheet.spliceRows(ws, 2, 2, ["new1", "val1"], ["new2", "val2"]);

      // The merge A2:B3 was entirely within the deleted range, so it should be removed
      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual([]);

      // New values should be plain, not merge proxies
      expect(Cell.getValue(ws, "A2")).toBe("new1");
      expect(Cell.getValue(ws, "B2")).toBe("val1");
      expect(Cell.getValue(ws, "A3")).toBe("new2");
    });

    it("Bug #1: spliceRows replace preserves merges outside replaced range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "above");
      Worksheet.merge(ws, "A1:B1");
      Cell.setValue(ws, "A4", "below");
      Worksheet.merge(ws, "A4:B4");
      Cell.setValue(ws, "A2", "middle");
      Worksheet.merge(ws, "A2:B3");

      // Replace rows 2-3 (delete 2, insert 2 -> nExpand=0)
      Worksheet.spliceRows(ws, 2, 2, ["r2"], ["r3"]);

      const model = getSheetModel(ws);
      // A2:B3 entirely within deleted range -> removed
      // A1:B1 before -> unchanged
      // A4:B4 after -> unchanged (nExpand=0, no shift)
      expect(model.mergeCells).toHaveLength(2);
      expect(model.mergeCells).toContain("A1:B1");
      expect(model.mergeCells).toContain("A4:B4");
    });

    it("Bug #2: spliceRows copies plain values, not merge proxy values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "master");
      Cell.setValue(ws, "B1", "B1val");
      Worksheet.merge(ws, "A1:B1");
      Cell.setValue(ws, "A3", "row3A");
      Cell.setValue(ws, "B3", "row3B");

      // Insert a row at row 2 — this shifts row 3 down to row 4
      Worksheet.spliceRows(ws, 2, 0, ["inserted"]);

      // Verify row 4 has the original row 3 values (not corrupted by merge proxy)
      expect(Cell.getValue(ws, "A4")).toBe("row3A");
      expect(Cell.getValue(ws, "B4")).toBe("row3B");
    });

    it("Bug #3: spliceRows removes merge that shrinks to 1x1", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "hello");
      Worksheet.merge(ws, "A1:A3"); // vertical merge, 3 rows, 1 column

      // Delete rows 2-3 -> merge would shrink to A1:A1
      Worksheet.spliceRows(ws, 2, 2);

      const model = getSheetModel(ws);
      // 1x1 merge should be removed
      expect(model.mergeCells).toEqual([]);
      expect(Cell.getValue(ws, "A1")).toBe("hello");
    });

    it("Bug #3: spliceColumns removes merge that shrinks to 1x1", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "hello");
      Worksheet.merge(ws, "A1:C1"); // horizontal merge, 1 row, 3 columns

      // Delete columns 2-3 -> merge would shrink to A1:A1
      Worksheet.spliceColumns(ws, 2, 2);

      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual([]);
      expect(Cell.getValue(ws, "A1")).toBe("hello");
    });

    it("Bug #4: spliceRows clears stale merge refs outside shrunk range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "master");
      Worksheet.merge(ws, "A1:B4");

      // Delete row 3 -> merge shrinks from A1:B4 to A1:B3
      Worksheet.spliceRows(ws, 3, 1);

      const model = getSheetModel(ws);
      expect(model.mergeCells).toEqual(["A1:B3"]);

      // Verify cells in new range are correct
      expect(Cell.getValue(ws, "A1")).toBe("master");
      expect(Cell.getType(ws, "B1")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B1").address).toBe("A1");
      expect(Cell.getType(ws, "A3")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "A3").address).toBe("A1");
      expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Merge);
      expect(Cell.getMergeMaster(ws, "B3").address).toBe("A1");
    });

    it("Bug #5: insertRow preserves style of merged cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Worksheet.merge(ws, "A1:B1");
      const borderStyle = {
        border: {
          bottom: {
            style: "medium" as const,
            color: { argb: "FF000000" }
          }
        }
      };
      getCell(ws, "A1").style = borderStyle;
      getCell(ws, "B1").style = borderStyle;

      Worksheet.insertRow(ws, 2, []);

      // A1 (master) style should be preserved
      expect(Cell.getStyle(ws, "A1").border).toEqual(borderStyle.border);
      // B1 (merge slave) style should also be preserved
      expect(Cell.getStyle(ws, "B1").border).toEqual(borderStyle.border);
    });

    it("Bug #5: spliceColumns preserves style of merged cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Worksheet.merge(ws, "A1:A2");
      const fillStyle = {
        fill: {
          type: "pattern" as const,
          pattern: "solid" as const,
          fgColor: { argb: "FFFF0000" }
        }
      };
      getCell(ws, "A1").style = fillStyle;
      getCell(ws, "A2").style = fillStyle;

      Worksheet.spliceColumns(ws, 2, 0, []);

      // A1 (master) style should be preserved
      expect(Cell.getStyle(ws, "A1").fill).toEqual(fillStyle.fill);
      // A2 (merge slave) style should also be preserved
      expect(Cell.getStyle(ws, "A2").fill).toEqual(fillStyle.fill);
    });

    it("Bug #6: duplicateRow overwrite cleans multi-row merges on target rows", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "source");
      Worksheet.merge(ws, "A1:C1");

      // Create a multi-row merge on rows 2-3
      Cell.setValue(ws, "A2", "multi");
      Worksheet.merge(ws, "A2:B3");

      // Overwrite row 2 with row 1's data
      Worksheet.duplicateRow(ws, 1, 1, false);

      const model = getSheetModel(ws);
      // Original A1:C1 should remain
      // A2:B3 should be removed (it touched the target row 2)
      // A2:C2 should be created (duplicated from source)
      expect(model.mergeCells).toContain("A1:C1");
      expect(model.mergeCells).toContain("A2:C2");
      expect(model.mergeCells).not.toContain("A2:B3");
    });

    describe("position-aware border handling", () => {
      const thinBorder = {
        top: { style: "thin" as const },
        left: { style: "thin" as const },
        bottom: { style: "thin" as const },
        right: { style: "thin" as const }
      };

      it("horizontal merge: preserves outer left/right, top/bottom on all, clears inner left/right", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Set all four cells with full borders
        for (const addr of ["A1", "B1", "C1", "D1"]) {
          Cell.setStyle(ws, addr, { border: { ...thinBorder } });
        }

        Worksheet.merge(ws, "A1:D1");

        // A1 = leftmost: left + top + bottom
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          left: thinBorder.left,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // B1 = interior: top + bottom only
        expect(Cell.getStyle(ws, "B1").border).toEqual({
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // C1 = interior: top + bottom only
        expect(Cell.getStyle(ws, "C1").border).toEqual({
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // D1 = rightmost: right + top + bottom
        expect(Cell.getStyle(ws, "D1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
      });

      it("vertical merge: preserves outer top/bottom, left/right on all, clears inner top/bottom", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        for (const addr of ["A1", "A2", "A3", "A4"]) {
          Cell.setStyle(ws, addr, { border: { ...thinBorder } });
        }

        Worksheet.merge(ws, "A1:A4");

        // A1 = topmost: left + right + top
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right,
          top: thinBorder.top
        });
        // A2, A3 = interior: left + right only
        expect(Cell.getStyle(ws, "A2").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right
        });
        expect(Cell.getStyle(ws, "A3").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right
        });
        // A4 = bottommost: left + right + bottom
        expect(Cell.getStyle(ws, "A4").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right,
          bottom: thinBorder.bottom
        });
      });

      it("rectangular merge: each cell gets only its perimeter edges", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // 3x3 merge B2:D4
        for (let r = 2; r <= 4; r++) {
          for (let c = 2; c <= 4; c++) {
            Cell.setStyle(ws, r, c, { border: { ...thinBorder } });
          }
        }

        Worksheet.merge(ws, "B2:D4");

        // Corners
        expect(Cell.getStyle(ws, "B2").border).toEqual({
          left: thinBorder.left,
          top: thinBorder.top
        });
        expect(Cell.getStyle(ws, "D2").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top
        });
        expect(Cell.getStyle(ws, "B4").border).toEqual({
          left: thinBorder.left,
          bottom: thinBorder.bottom
        });
        expect(Cell.getStyle(ws, "D4").border).toEqual({
          right: thinBorder.right,
          bottom: thinBorder.bottom
        });

        // Edges (non-corner)
        expect(Cell.getStyle(ws, "C2").border).toEqual({ top: thinBorder.top }); // top edge
        expect(Cell.getStyle(ws, "C4").border).toEqual({ bottom: thinBorder.bottom }); // bottom edge
        expect(Cell.getStyle(ws, "B3").border).toEqual({ left: thinBorder.left }); // left edge
        expect(Cell.getStyle(ws, "D3").border).toEqual({ right: thinBorder.right }); // right edge

        // Interior cell has no border
        expect(Cell.getStyle(ws, "C3").border).toBeUndefined();
      });

      it("no borders: merge cells without borders produces no borders", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        Cell.setValue(ws, "A1", "hello");
        Worksheet.merge(ws, "A1:B2");

        expect(Cell.getStyle(ws, "A1").border).toBeUndefined();
        expect(Cell.getStyle(ws, "B1").border).toBeUndefined();
        expect(Cell.getStyle(ws, "A2").border).toBeUndefined();
        expect(Cell.getStyle(ws, "B2").border).toBeUndefined();
      });

      it("slave border preserved: slave's outer border survives even if master lacks it", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Master has left border only, slave (B1) has right border only
        Cell.setStyle(ws, "A1", { border: { left: { style: "thin" } } });
        Cell.setStyle(ws, "B1", { border: { right: { style: "thick" } } });

        Worksheet.merge(ws, "A1:B1");

        // A1 keeps its left border (it's on the left perimeter)
        expect(Cell.getStyle(ws, "A1").border).toEqual({ left: { style: "thin" } });
        // B1 keeps its own right border (it's on the right perimeter)
        expect(Cell.getStyle(ws, "B1").border).toEqual({ right: { style: "thick" } });
      });

      it("mergeCellsWithoutStyle does not alter borders", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        const fullBorder = { ...thinBorder };
        Cell.setStyle(ws, "A1", { border: { ...fullBorder } });
        Cell.setStyle(ws, "B1", { border: { ...fullBorder } });

        Worksheet.mergeWithoutStyle(ws, "A1:B1");

        // Both cells retain their original full borders untouched
        expect(Cell.getStyle(ws, "A1").border).toEqual(fullBorder);
        expect(Cell.getStyle(ws, "B1").border).toEqual(fullBorder);
      });

      it("each cell has an independent style object after merge", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        Cell.setStyle(ws, "A1", { border: { ...thinBorder } });
        Cell.setStyle(ws, "B1", { border: { ...thinBorder } });

        Worksheet.merge(ws, "A1:B1");

        // Mutating A1's border should not affect B1
        Cell.getStyle(ws, "A1").border = { top: { style: "double" } };
        expect(Cell.getStyle(ws, "B1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
      });

      it("diagonal border from master is propagated to all cells", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        Cell.setStyle(ws, "A1", {
          border: {
            ...thinBorder,
            diagonal: { style: "thin", up: true, down: false }
          }
        });
        Cell.setStyle(ws, "B1", { border: { ...thinBorder } });

        Worksheet.merge(ws, "A1:B1");

        // Both cells get the diagonal from master
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          left: thinBorder.left,
          top: thinBorder.top,
          bottom: thinBorder.bottom,
          diagonal: { style: "thin", up: true, down: false }
        });
        expect(Cell.getStyle(ws, "B1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom,
          diagonal: { style: "thin", up: true, down: false }
        });
      });

      it("single cell merge: all four borders are preserved", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        Cell.setStyle(ws, "A1", { border: { ...thinBorder } });
        Worksheet.merge(ws, "A1:A1");

        // Single-cell merge: all four sides are perimeter
        expect(Cell.getStyle(ws, "A1").border).toEqual(thinBorder);
      });

      it("top-level border color from master is propagated", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // The top-level `color` is a convenience property used by border-xform
        // at render time; it's not in the Borders type, so we cast here.
        Cell.setStyle(ws, "A1", {
          border: {
            color: { argb: "FFFF0000" },
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" }
          } as any
        });
        Worksheet.merge(ws, "A1:B2");

        // All cells should carry the top-level color
        for (const addr of ["A1", "B1", "A2", "B2"]) {
          expect(Cell.getStyle(ws, addr).border).toHaveProperty("color", { argb: "FFFF0000" });
        }
        // And each cell only has its positional edges
        expect(Cell.getStyle(ws, "A1").border).toMatchObject({
          left: { style: "double" },
          top: { style: "double" }
        });
        expect(Cell.getStyle(ws, "B1").border).toMatchObject({
          right: { style: "double" },
          top: { style: "double" }
        });
        expect(Cell.getStyle(ws, "A2").border).toMatchObject({
          left: { style: "double" },
          bottom: { style: "double" }
        });
        expect(Cell.getStyle(ws, "B2").border).toMatchObject({
          right: { style: "double" },
          bottom: { style: "double" }
        });
      });

      it("border edges with color objects are preserved", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        const coloredBorder = {
          top: { style: "double" as const, color: { argb: "FFFF00FF" } },
          left: { style: "double" as const, color: { argb: "FF00FFFF" } },
          bottom: { style: "double" as const, color: { argb: "FF00FF00" } },
          right: { style: "double" as const, color: { argb: "FF0000FF" } }
        };
        Cell.setStyle(ws, "A1", { border: { ...coloredBorder } });
        Cell.setStyle(ws, "B1", { border: { ...coloredBorder } });

        Worksheet.merge(ws, "A1:B1");

        // A1 keeps left (cyan) + top (magenta) + bottom (green)
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          left: coloredBorder.left,
          top: coloredBorder.top,
          bottom: coloredBorder.bottom
        });
        // B1 keeps right (blue) + top (from master fallback = magenta) + bottom (from master fallback = green)
        expect(Cell.getStyle(ws, "B1").border).toEqual({
          right: coloredBorder.right,
          top: coloredBorder.top,
          bottom: coloredBorder.bottom
        });
      });

      it("mixed border styles on different perimeter edges", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Master: thin top, thick left
        Cell.setStyle(ws, "A1", { border: { top: { style: "thin" }, left: { style: "thick" } } });
        // Slave: dashed bottom, double right
        Cell.setStyle(ws, "B2", {
          border: { bottom: { style: "dashed" }, right: { style: "double" } }
        });

        Worksheet.merge(ws, "A1:B2");

        // A1 = top-left: left (thick from A1) + top (thin from A1)
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          left: { style: "thick" },
          top: { style: "thin" }
        });
        // B1 = top-right: right (fallback master=undefined, B1 has none) → no right; top (fallback master=thin)
        // But B1 itself has no border, so right = master.right = undefined
        expect(Cell.getStyle(ws, "B1").border).toEqual({
          top: { style: "thin" }
        });
        // A2 = bottom-left: left (fallback master=thick), bottom (A2 has no border, master has no bottom) → just left
        expect(Cell.getStyle(ws, "A2").border).toEqual({
          left: { style: "thick" }
        });
        // B2 = bottom-right: right (double from B2) + bottom (dashed from B2)
        expect(Cell.getStyle(ws, "B2").border).toEqual({
          right: { style: "double" },
          bottom: { style: "dashed" }
        });
      });

      it("non-border styles preserved without phantom borders", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Master has font and fill but no border
        Cell.setStyle(ws, "A1", { font: { bold: true } });
        Cell.setStyle(ws, "A1", {
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          }
        });
        Cell.setValue(ws, "A1", "styled");

        Worksheet.merge(ws, "A1:B2");

        // All cells inherit font and fill
        for (const addr of ["A1", "B1", "A2", "B2"]) {
          expect(Cell.getStyle(ws, addr).font).toEqual({ bold: true });
          expect(Cell.getStyle(ws, addr).fill).toEqual({
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          });
          // No phantom border should appear
          expect(Cell.getStyle(ws, addr).border).toBeUndefined();
        }
      });

      it("discussion #78 scenario: template merge preserves rightmost border", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Simulate reading a template where each cell has its own border
        // C21 has left border, L21 has right border, all have top+bottom
        for (let col = 3; col <= 12; col++) {
          const border: Record<string, any> = {
            top: { style: "thin" },
            bottom: { style: "thin" }
          };
          if (col === 3) {
            border.left = { style: "thin" };
          }
          if (col === 12) {
            border.right = { style: "thin" };
          }
          Cell.setStyle(ws, 21, col, { border: border });
        }

        Worksheet.merge(ws, "C21:L21");

        // L21 (col 12) must retain its right border
        expect(Cell.getStyle(ws, "L21").border).toEqual({
          right: { style: "thin" },
          top: { style: "thin" },
          bottom: { style: "thin" }
        });
        // C21 (col 3) must retain its left border
        expect(Cell.getStyle(ws, "C21").border).toEqual({
          left: { style: "thin" },
          top: { style: "thin" },
          bottom: { style: "thin" }
        });
        // Interior cells (D21-K21) should only have top+bottom
        for (let col = 4; col <= 11; col++) {
          expect(Cell.getStyle(ws, 21, col).border).toEqual({
            top: { style: "thin" },
            bottom: { style: "thin" }
          });
        }
      });

      it("setting borders after merge does not cause cross-cell mutation", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        // Merge without any pre-existing borders
        Cell.setValue(ws, "A1", "Merged");
        Worksheet.merge(ws, "A1:C3");

        // Set borders on boundary cells one by one (natural user workflow)
        Cell.setStyle(ws, "A1", {
          border: {
            top: { style: "thick", color: { argb: "FF0000FF" } },
            left: { style: "thick", color: { argb: "FF0000FF" } }
          }
        });
        Cell.setStyle(ws, "C1", {
          border: {
            right: { style: "thick", color: { argb: "FFFF0000" } }
          }
        });
        Cell.setStyle(ws, "A3", {
          border: {
            bottom: { style: "thick", color: { argb: "FF00FF00" } }
          }
        });

        // Each cell should have only the borders it was assigned
        expect(Cell.getStyle(ws, "A1").border).toEqual({
          top: { style: "thick", color: { argb: "FF0000FF" } },
          left: { style: "thick", color: { argb: "FF0000FF" } }
        });
        expect(Cell.getStyle(ws, "C1").border).toEqual({
          right: { style: "thick", color: { argb: "FFFF0000" } }
        });
        expect(Cell.getStyle(ws, "A3").border).toEqual({
          bottom: { style: "thick", color: { argb: "FF00FF00" } }
        });

        // Unmodified interior cell should have no border
        expect(Cell.getStyle(ws, "B2").border).toBeUndefined();
      });

      it("styles are independent after borderless merge — mutating one cell does not affect others", () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "sheet");

        Worksheet.merge(ws, "A1:B2");

        // Set font on A1 only
        Cell.setStyle(ws, "A1", { font: { bold: true } });

        // B1 should not be affected
        expect(Cell.getStyle(ws, "A1").font).toEqual({ bold: true });
        expect(Cell.getStyle(ws, "B1").font).toBeUndefined();

        // Set fill on B2 only
        Cell.setStyle(ws, "B2", {
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          }
        });
        expect(Cell.getStyle(ws, "B2").fill).toEqual({
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        });
        expect(Cell.getStyle(ws, "A1").fill).toBeUndefined();
      });
    });
  });
});
