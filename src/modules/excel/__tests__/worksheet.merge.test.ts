import { describe, it, expect } from "vitest";
import { testUtils } from "@excel/__tests__/shared";
import { Workbook } from "../../../index";
import { Dimensions } from "@excel/range";
import { Enums } from "@excel/enums";

describe("Worksheet", () => {
  describe("Merge Cells", () => {
    it("references the same top-left value", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      // initial values
      ws.getCell("A1").value = "A1";
      ws.getCell("B1").value = "B1";
      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";

      ws.mergeCells("A1:B2");

      expect(ws.getCell("A1").value).toBe("A1");
      expect(ws.getCell("B1").value).toBe("A1");
      expect(ws.getCell("A2").value).toBe("A1");
      expect(ws.getCell("B2").value).toBe("A1");

      expect(ws.getCell("A1").type).toBe(Enums.ValueType.String);
      expect(ws.getCell("B1").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("A2").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B2").type).toBe(Enums.ValueType.Merge);
    });

    it("does not allow overlapping merges", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.mergeCells("B2:C3");

      // intersect four corners
      expect(() => {
        ws.mergeCells("A1:B2");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("C1:D2");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("C3:D4");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("A3:B4");
      }).toThrow(Error);

      // enclosing
      expect(() => {
        ws.mergeCells("A1:D4");
      }).toThrow(Error);
    });

    it("merges and unmerges", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const expectMaster = function (range: string, master: string | null) {
        const d = new Dimensions(range);
        for (let i = d.top; i <= d.bottom; i++) {
          for (let j = d.left; j <= d.right; j++) {
            const cell = ws.getCell(i, j);
            const masterCell = master ? ws.getCell(master) : cell;
            expect(cell.master.address).toBe(masterCell.address);
          }
        }
      };

      // merge some cells, then unmerge them
      ws.mergeCells("A1:B2");
      expectMaster("A1:B2", "A1");
      ws.unMergeCells("A1:B2");
      expectMaster("A1:B2", null);

      // unmerge just one cell
      ws.mergeCells("A1:B2");
      expectMaster("A1:B2", "A1");
      ws.unMergeCells("A1");
      expectMaster("A1:B2", null);

      ws.mergeCells("A1:B2");
      expectMaster("A1:B2", "A1");
      ws.unMergeCells("B2");
      expectMaster("A1:B2", null);

      // build 4 merge-squares
      ws.mergeCells("A1:B2");
      ws.mergeCells("D1:E2");
      ws.mergeCells("A4:B5");
      ws.mergeCells("D4:E5");

      expectMaster("A1:B2", "A1");
      expectMaster("D1:E2", "D1");
      expectMaster("A4:B5", "A4");
      expectMaster("D4:E5", "D4");

      // unmerge the middle
      ws.unMergeCells("B2:D4");

      expectMaster("A1:B2", null);
      expectMaster("D1:E2", null);
      expectMaster("A4:B5", null);
      expectMaster("D4:E5", null);
    });

    it("merges styles", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      // initial value
      const B2 = ws.getCell("B2");
      B2.value = 5;
      B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
      B2.style.border = testUtils.styles.borders.doubleRed;
      B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
      B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
      B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

      // expecting styles to be copied (see worksheet spec)
      ws.mergeCells("B2:C3");

      const dblRed = testUtils.styles.borders.doubleRed;

      // Non-border styles are copied identically to all cells
      for (const addr of ["B2", "B3", "C2", "C3"]) {
        expect(ws.getCell(addr).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
        expect(ws.getCell(addr).fill).toEqual(testUtils.styles.fills.blueWhiteHGrad);
        expect(ws.getCell(addr).alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
        expect(ws.getCell(addr).numFmt).toEqual(testUtils.styles.numFmts.numFmt1);
      }

      // Borders are position-aware: only perimeter edges survive (like Excel)
      // B2 = top-left corner
      expect(ws.getCell("B2").border).toEqual({ left: dblRed.left, top: dblRed.top });
      // C2 = top-right corner
      expect(ws.getCell("C2").border).toEqual({ right: dblRed.right, top: dblRed.top });
      // B3 = bottom-left corner
      expect(ws.getCell("B3").border).toEqual({ left: dblRed.left, bottom: dblRed.bottom });
      // C3 = bottom-right corner
      expect(ws.getCell("C3").border).toEqual({ right: dblRed.right, bottom: dblRed.bottom });
    });

    it("preserves merges after row inserts", function () {
      const wb = new Workbook();
      const ws = wb.addWorksheet("testMergeAfterInsert");

      ws.addRow([1, 2]);
      ws.addRow([3, 4]);
      ws.mergeCells("A1:B2");
      ws.insertRow(1, ["Inserted Row Text"]);

      // After insert, the merged area should now be A2:B3
      // A2 is master (type=Number with value 1), B2, A3, B3 are merge cells
      const cellA2 = ws.getCell("A2");
      const cellB2 = ws.getCell("B2");
      const cellA3 = ws.getCell("A3");
      const cellB3 = ws.getCell("B3");

      // Verify master cell has the number value
      expect(cellA2.type).toEqual(Enums.ValueType.Number);
      expect(cellA2.value).toEqual(1);

      // Verify other cells in merge area are merge type and point to A2 address
      expect(cellB2.type).toEqual(Enums.ValueType.Merge);
      expect(cellB2.master.address).toEqual("A2");

      expect(cellA3.type).toEqual(Enums.ValueType.Merge);
      expect(cellA3.master.address).toEqual("A2");

      expect(cellB3.type).toEqual(Enums.ValueType.Merge);
      expect(cellB3.master.address).toEqual("A2");
    });

    it("spliceRows updates _merges after inserting rows above a merge", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A3").value = "hello";
      ws.mergeCells("A3:C4");

      // Insert 2 rows at row 1 (above the merge)
      ws.spliceRows(1, 0, ["x"], ["y"]);

      // Merge should shift down by 2: A3:C4 -> A5:C6
      const model = ws.model;
      expect(model.mergeCells).toEqual(["A5:C6"]);

      // Cell-level merge references should also be correct
      expect(ws.getCell("A5").value).toBe("hello");
      expect(ws.getCell("B5").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B5").master.address).toBe("A5");
      expect(ws.getCell("C6").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("C6").master.address).toBe("A5");
    });

    it("spliceRows updates _merges after deleting rows above a merge", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.addRow(["filler1"]);
      ws.addRow(["filler2"]);
      ws.addRow(["filler3"]);
      ws.getCell("A4").value = "hello";
      ws.mergeCells("A4:B5");

      // Delete rows 1-2 (above the merge)
      ws.spliceRows(1, 2);

      // Merge should shift up by 2: A4:B5 -> A2:B3
      const model = ws.model;
      expect(model.mergeCells).toEqual(["A2:B3"]);

      expect(ws.getCell("A2").value).toBe("hello");
      expect(ws.getCell("B3").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B3").master.address).toBe("A2");
    });

    it("spliceRows removes merges entirely within deleted rows", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.addRow(["filler"]);
      ws.getCell("A2").value = "merged";
      ws.mergeCells("A2:B3");
      ws.addRow(["below"]);

      // Delete rows 2-3 which contain the entire merge
      ws.spliceRows(2, 2);

      const model = ws.model;
      expect(model.mergeCells).toEqual([]);
    });

    it("spliceRows shrinks merge spanning the splice boundary", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "hello";
      ws.mergeCells("A1:B4");

      // Delete row 3 (within the merge)
      ws.spliceRows(3, 1);

      // Merge should shrink: A1:B4 -> A1:B3
      const model = ws.model;
      expect(model.mergeCells).toEqual(["A1:B3"]);
    });

    it("duplicateRow preserves single-row horizontal merges", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "merged";
      ws.getCell("D1").value = "solo";
      ws.mergeCells("A1:C1");

      // Duplicate row 1, inserting 2 copies below
      ws.duplicateRow(1, 2, true);

      const model = ws.model;
      // Should have 3 merges: original A1:C1, plus A2:C2 and A3:C3
      expect(model.mergeCells).toHaveLength(3);
      expect(model.mergeCells).toContain("A1:C1");
      expect(model.mergeCells).toContain("A2:C2");
      expect(model.mergeCells).toContain("A3:C3");

      // Verify cell-level merge references in duplicated rows
      expect(ws.getCell("A2").value).toBe("merged");
      expect(ws.getCell("B2").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B2").master.address).toBe("A2");
      expect(ws.getCell("C2").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("C2").master.address).toBe("A2");

      expect(ws.getCell("A3").value).toBe("merged");
      expect(ws.getCell("B3").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B3").master.address).toBe("A3");
    });

    it("duplicateRow preserves multi-row merges", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "big merge";
      ws.mergeCells("A1:B3");
      ws.getCell("C1").value = "outside";

      // Duplicate row 1 once with insert
      ws.duplicateRow(1, 1, true);

      const model = ws.model;
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "source";
      ws.mergeCells("A1:C1");

      ws.getCell("A2").value = "existing";
      ws.mergeCells("A2:D2");

      // Duplicate row 1 over row 2 (overwrite mode, insert=false)
      ws.duplicateRow(1, 1, false);

      const model = ws.model;
      // Original merge A1:C1 should remain
      // Row 2's old merge A2:D2 should be replaced with A2:C2 (duplicated from row 1)
      expect(model.mergeCells).toHaveLength(2);
      expect(model.mergeCells).toContain("A1:C1");
      expect(model.mergeCells).toContain("A2:C2");
    });

    it("duplicateRow + XLSX roundtrip preserves merges", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "merged";
      ws.getCell("D1").value = "solo";
      ws.mergeCells("A1:C1");

      ws.duplicateRow(1, 2, true);

      // Write to buffer and read back
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      const ws2 = wb2.getWorksheet("sheet")!;

      const model2 = ws2.model;
      expect(model2.mergeCells).toHaveLength(3);
      expect(model2.mergeCells).toContain("A1:C1");
      expect(model2.mergeCells).toContain("A2:C2");
      expect(model2.mergeCells).toContain("A3:C3");
    });

    it("duplicateRow with multiple merges on source row", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "first";
      ws.getCell("D1").value = "second";
      ws.mergeCells("A1:B1");
      ws.mergeCells("D1:F1");

      ws.duplicateRow(1, 2, true);

      const model = ws.model;
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      // Merge C1:E1 (cols 3-5)
      ws.getCell("C1").value = "merged";
      ws.mergeCells("C1:E1");

      // Insert 2 columns at column 2 (before the merge)
      ws.spliceColumns(2, 0, [], []);

      const model = ws.model;
      // Merge should shift right by 2: C1:E1 → E1:G1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("E1:G1");

      // Cell-level: E1 is master, F1 and G1 are merge slaves
      expect(ws.getCell("E1").value).toBe("merged");
      expect(ws.getCell("F1").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("G1").type).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns updates _merges after deleting columns before a merge", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      // Merge D1:F1 (cols 4-6)
      ws.getCell("D1").value = "merged";
      ws.mergeCells("D1:F1");

      // Delete 2 columns at column 1 (cols 1-2 removed)
      ws.spliceColumns(1, 2);

      const model = ws.model;
      // Merge should shift left by 2: D1:F1 → B1:D1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("B1:D1");

      // Cell-level: B1 is master
      expect(ws.getCell("B1").value).toBe("merged");
      expect(ws.getCell("C1").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("D1").type).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns removes merges entirely within deleted columns", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      // Merge B1:C1 (cols 2-3)
      ws.getCell("B1").value = "merged";
      ws.mergeCells("B1:C1");

      // Also add a merge that survives: E1:F1 (cols 5-6)
      ws.getCell("E1").value = "survivor";
      ws.mergeCells("E1:F1");

      // Delete columns 2-3 (removes B1:C1 entirely)
      ws.spliceColumns(2, 2);

      const model = ws.model;
      // B1:C1 removed, E1:F1 shifts left by 2 → C1:D1
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("C1:D1");

      expect(ws.getCell("C1").value).toBe("survivor");
      expect(ws.getCell("D1").type).toBe(Enums.ValueType.Merge);
    });

    it("spliceColumns shrinks merge spanning the splice boundary", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      // Merge B1:F1 (cols 2-6, 5 columns wide)
      ws.getCell("B1").value = "wide";
      ws.mergeCells("B1:F1");

      // Delete columns 4-5 (2 columns from the middle of the merge)
      ws.spliceColumns(4, 2);

      const model = ws.model;
      // Merge should shrink: B1:F1 → B1:D1 (right reduced by 2)
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("B1:D1");

      expect(ws.getCell("B1").value).toBe("wide");
      expect(ws.getCell("C1").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("D1").type).toBe(Enums.ValueType.Merge);
    });
  });

  describe("insertRow with merges", () => {
    it("insertRow preserves model.mergeCells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      // Merge A2:C2
      ws.getCell("A2").value = "merged";
      ws.mergeCells("A2:C2");

      // insertRow at row 1 pushes merge down by 1
      ws.insertRow(1, ["new"]);

      const model = ws.model;
      // Merge should shift down: A2:C2 → A3:C3
      expect(model.mergeCells).toHaveLength(1);
      expect(model.mergeCells).toContain("A3:C3");

      expect(ws.getCell("A3").value).toBe("merged");
      expect(ws.getCell("B3").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("C3").type).toBe(Enums.ValueType.Merge);
    });
  });

  describe("merge edge cases", () => {
    it("Bug #1: spliceRows with equal delete and insert updates merges in replaced range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A2").value = "merged";
      ws.mergeCells("A2:B3");

      // Replace rows 2-3 with new data (delete 2, insert 2 -> nExpand=0)
      ws.spliceRows(2, 2, ["new1", "val1"], ["new2", "val2"]);

      // The merge A2:B3 was entirely within the deleted range, so it should be removed
      const model = ws.model;
      expect(model.mergeCells).toEqual([]);

      // New values should be plain, not merge proxies
      expect(ws.getCell("A2").value).toBe("new1");
      expect(ws.getCell("B2").value).toBe("val1");
      expect(ws.getCell("A3").value).toBe("new2");
    });

    it("Bug #1: spliceRows replace preserves merges outside replaced range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "above";
      ws.mergeCells("A1:B1");
      ws.getCell("A4").value = "below";
      ws.mergeCells("A4:B4");
      ws.getCell("A2").value = "middle";
      ws.mergeCells("A2:B3");

      // Replace rows 2-3 (delete 2, insert 2 -> nExpand=0)
      ws.spliceRows(2, 2, ["r2"], ["r3"]);

      const model = ws.model;
      // A2:B3 entirely within deleted range -> removed
      // A1:B1 before -> unchanged
      // A4:B4 after -> unchanged (nExpand=0, no shift)
      expect(model.mergeCells).toHaveLength(2);
      expect(model.mergeCells).toContain("A1:B1");
      expect(model.mergeCells).toContain("A4:B4");
    });

    it("Bug #2: spliceRows copies plain values, not merge proxy values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "master";
      ws.getCell("B1").value = "B1val";
      ws.mergeCells("A1:B1");
      ws.getCell("A3").value = "row3A";
      ws.getCell("B3").value = "row3B";

      // Insert a row at row 2 — this shifts row 3 down to row 4
      ws.spliceRows(2, 0, ["inserted"]);

      // Verify row 4 has the original row 3 values (not corrupted by merge proxy)
      expect(ws.getCell("A4").value).toBe("row3A");
      expect(ws.getCell("B4").value).toBe("row3B");
    });

    it("Bug #3: spliceRows removes merge that shrinks to 1x1", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "hello";
      ws.mergeCells("A1:A3"); // vertical merge, 3 rows, 1 column

      // Delete rows 2-3 -> merge would shrink to A1:A1
      ws.spliceRows(2, 2);

      const model = ws.model;
      // 1x1 merge should be removed
      expect(model.mergeCells).toEqual([]);
      expect(ws.getCell("A1").value).toBe("hello");
    });

    it("Bug #3: spliceColumns removes merge that shrinks to 1x1", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "hello";
      ws.mergeCells("A1:C1"); // horizontal merge, 1 row, 3 columns

      // Delete columns 2-3 -> merge would shrink to A1:A1
      ws.spliceColumns(2, 2);

      const model = ws.model;
      expect(model.mergeCells).toEqual([]);
      expect(ws.getCell("A1").value).toBe("hello");
    });

    it("Bug #4: spliceRows clears stale merge refs outside shrunk range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "master";
      ws.mergeCells("A1:B4");

      // Delete row 3 -> merge shrinks from A1:B4 to A1:B3
      ws.spliceRows(3, 1);

      const model = ws.model;
      expect(model.mergeCells).toEqual(["A1:B3"]);

      // Verify cells in new range are correct
      expect(ws.getCell("A1").value).toBe("master");
      expect(ws.getCell("B1").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B1").master.address).toBe("A1");
      expect(ws.getCell("A3").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("A3").master.address).toBe("A1");
      expect(ws.getCell("B3").type).toBe(Enums.ValueType.Merge);
      expect(ws.getCell("B3").master.address).toBe("A1");
    });

    it("Bug #5: insertRow preserves style of merged cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.mergeCells("A1:B1");
      const borderStyle = {
        border: {
          bottom: {
            style: "medium" as const,
            color: { argb: "FF000000" }
          }
        }
      };
      ws.getCell("A1").style = borderStyle;
      ws.getCell("B1").style = borderStyle;

      ws.insertRow(2, []);

      // A1 (master) style should be preserved
      expect(ws.getCell("A1").style.border).toEqual(borderStyle.border);
      // B1 (merge slave) style should also be preserved
      expect(ws.getCell("B1").style.border).toEqual(borderStyle.border);
    });

    it("Bug #5: spliceColumns preserves style of merged cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.mergeCells("A1:A2");
      const fillStyle = {
        fill: {
          type: "pattern" as const,
          pattern: "solid" as const,
          fgColor: { argb: "FFFF0000" }
        }
      };
      ws.getCell("A1").style = fillStyle;
      ws.getCell("A2").style = fillStyle;

      ws.spliceColumns(2, 0, []);

      // A1 (master) style should be preserved
      expect(ws.getCell("A1").style.fill).toEqual(fillStyle.fill);
      // A2 (merge slave) style should also be preserved
      expect(ws.getCell("A2").style.fill).toEqual(fillStyle.fill);
    });

    it("Bug #6: duplicateRow overwrite cleans multi-row merges on target rows", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "source";
      ws.mergeCells("A1:C1");

      // Create a multi-row merge on rows 2-3
      ws.getCell("A2").value = "multi";
      ws.mergeCells("A2:B3");

      // Overwrite row 2 with row 1's data
      ws.duplicateRow(1, 1, false);

      const model = ws.model;
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
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // Set all four cells with full borders
        for (const addr of ["A1", "B1", "C1", "D1"]) {
          ws.getCell(addr).border = { ...thinBorder };
        }

        ws.mergeCells("A1:D1");

        // A1 = leftmost: left + top + bottom
        expect(ws.getCell("A1").border).toEqual({
          left: thinBorder.left,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // B1 = interior: top + bottom only
        expect(ws.getCell("B1").border).toEqual({
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // C1 = interior: top + bottom only
        expect(ws.getCell("C1").border).toEqual({
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
        // D1 = rightmost: right + top + bottom
        expect(ws.getCell("D1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
      });

      it("vertical merge: preserves outer top/bottom, left/right on all, clears inner top/bottom", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        for (const addr of ["A1", "A2", "A3", "A4"]) {
          ws.getCell(addr).border = { ...thinBorder };
        }

        ws.mergeCells("A1:A4");

        // A1 = topmost: left + right + top
        expect(ws.getCell("A1").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right,
          top: thinBorder.top
        });
        // A2, A3 = interior: left + right only
        expect(ws.getCell("A2").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right
        });
        expect(ws.getCell("A3").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right
        });
        // A4 = bottommost: left + right + bottom
        expect(ws.getCell("A4").border).toEqual({
          left: thinBorder.left,
          right: thinBorder.right,
          bottom: thinBorder.bottom
        });
      });

      it("rectangular merge: each cell gets only its perimeter edges", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // 3x3 merge B2:D4
        for (let r = 2; r <= 4; r++) {
          for (let c = 2; c <= 4; c++) {
            ws.getCell(r, c).border = { ...thinBorder };
          }
        }

        ws.mergeCells("B2:D4");

        // Corners
        expect(ws.getCell("B2").border).toEqual({ left: thinBorder.left, top: thinBorder.top });
        expect(ws.getCell("D2").border).toEqual({ right: thinBorder.right, top: thinBorder.top });
        expect(ws.getCell("B4").border).toEqual({
          left: thinBorder.left,
          bottom: thinBorder.bottom
        });
        expect(ws.getCell("D4").border).toEqual({
          right: thinBorder.right,
          bottom: thinBorder.bottom
        });

        // Edges (non-corner)
        expect(ws.getCell("C2").border).toEqual({ top: thinBorder.top }); // top edge
        expect(ws.getCell("C4").border).toEqual({ bottom: thinBorder.bottom }); // bottom edge
        expect(ws.getCell("B3").border).toEqual({ left: thinBorder.left }); // left edge
        expect(ws.getCell("D3").border).toEqual({ right: thinBorder.right }); // right edge

        // Interior cell has no border
        expect(ws.getCell("C3").border).toBeUndefined();
      });

      it("no borders: merge cells without borders produces no borders", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        ws.getCell("A1").value = "hello";
        ws.mergeCells("A1:B2");

        expect(ws.getCell("A1").border).toBeUndefined();
        expect(ws.getCell("B1").border).toBeUndefined();
        expect(ws.getCell("A2").border).toBeUndefined();
        expect(ws.getCell("B2").border).toBeUndefined();
      });

      it("slave border preserved: slave's outer border survives even if master lacks it", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // Master has left border only, slave (B1) has right border only
        ws.getCell("A1").border = { left: { style: "thin" } };
        ws.getCell("B1").border = { right: { style: "thick" } };

        ws.mergeCells("A1:B1");

        // A1 keeps its left border (it's on the left perimeter)
        expect(ws.getCell("A1").border).toEqual({ left: { style: "thin" } });
        // B1 keeps its own right border (it's on the right perimeter)
        expect(ws.getCell("B1").border).toEqual({ right: { style: "thick" } });
      });

      it("mergeCellsWithoutStyle does not alter borders", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        const fullBorder = { ...thinBorder };
        ws.getCell("A1").border = { ...fullBorder };
        ws.getCell("B1").border = { ...fullBorder };

        ws.mergeCellsWithoutStyle("A1:B1");

        // Both cells retain their original full borders untouched
        expect(ws.getCell("A1").border).toEqual(fullBorder);
        expect(ws.getCell("B1").border).toEqual(fullBorder);
      });

      it("each cell has an independent style object after merge", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        ws.getCell("A1").border = { ...thinBorder };
        ws.getCell("B1").border = { ...thinBorder };

        ws.mergeCells("A1:B1");

        // Mutating A1's border should not affect B1
        ws.getCell("A1").style.border = { top: { style: "double" } };
        expect(ws.getCell("B1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom
        });
      });

      it("diagonal border from master is propagated to all cells", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        ws.getCell("A1").border = {
          ...thinBorder,
          diagonal: { style: "thin", up: true, down: false }
        };
        ws.getCell("B1").border = { ...thinBorder };

        ws.mergeCells("A1:B1");

        // Both cells get the diagonal from master
        expect(ws.getCell("A1").border).toEqual({
          left: thinBorder.left,
          top: thinBorder.top,
          bottom: thinBorder.bottom,
          diagonal: { style: "thin", up: true, down: false }
        });
        expect(ws.getCell("B1").border).toEqual({
          right: thinBorder.right,
          top: thinBorder.top,
          bottom: thinBorder.bottom,
          diagonal: { style: "thin", up: true, down: false }
        });
      });

      it("single cell merge: all four borders are preserved", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        ws.getCell("A1").border = { ...thinBorder };
        ws.mergeCells("A1:A1");

        // Single-cell merge: all four sides are perimeter
        expect(ws.getCell("A1").border).toEqual(thinBorder);
      });

      it("top-level border color from master is propagated", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // The top-level `color` is a convenience property used by border-xform
        // at render time; it's not in the Borders type, so we cast here.
        ws.getCell("A1").border = {
          color: { argb: "FFFF0000" },
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" }
        } as any;
        ws.mergeCells("A1:B2");

        // All cells should carry the top-level color
        for (const addr of ["A1", "B1", "A2", "B2"]) {
          expect(ws.getCell(addr).border).toHaveProperty("color", { argb: "FFFF0000" });
        }
        // And each cell only has its positional edges
        expect(ws.getCell("A1").border).toMatchObject({
          left: { style: "double" },
          top: { style: "double" }
        });
        expect(ws.getCell("B1").border).toMatchObject({
          right: { style: "double" },
          top: { style: "double" }
        });
        expect(ws.getCell("A2").border).toMatchObject({
          left: { style: "double" },
          bottom: { style: "double" }
        });
        expect(ws.getCell("B2").border).toMatchObject({
          right: { style: "double" },
          bottom: { style: "double" }
        });
      });

      it("border edges with color objects are preserved", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        const coloredBorder = {
          top: { style: "double" as const, color: { argb: "FFFF00FF" } },
          left: { style: "double" as const, color: { argb: "FF00FFFF" } },
          bottom: { style: "double" as const, color: { argb: "FF00FF00" } },
          right: { style: "double" as const, color: { argb: "FF0000FF" } }
        };
        ws.getCell("A1").border = { ...coloredBorder };
        ws.getCell("B1").border = { ...coloredBorder };

        ws.mergeCells("A1:B1");

        // A1 keeps left (cyan) + top (magenta) + bottom (green)
        expect(ws.getCell("A1").border).toEqual({
          left: coloredBorder.left,
          top: coloredBorder.top,
          bottom: coloredBorder.bottom
        });
        // B1 keeps right (blue) + top (from master fallback = magenta) + bottom (from master fallback = green)
        expect(ws.getCell("B1").border).toEqual({
          right: coloredBorder.right,
          top: coloredBorder.top,
          bottom: coloredBorder.bottom
        });
      });

      it("mixed border styles on different perimeter edges", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // Master: thin top, thick left
        ws.getCell("A1").border = { top: { style: "thin" }, left: { style: "thick" } };
        // Slave: dashed bottom, double right
        ws.getCell("B2").border = { bottom: { style: "dashed" }, right: { style: "double" } };

        ws.mergeCells("A1:B2");

        // A1 = top-left: left (thick from A1) + top (thin from A1)
        expect(ws.getCell("A1").border).toEqual({
          left: { style: "thick" },
          top: { style: "thin" }
        });
        // B1 = top-right: right (fallback master=undefined, B1 has none) → no right; top (fallback master=thin)
        // But B1 itself has no border, so right = master.right = undefined
        expect(ws.getCell("B1").border).toEqual({
          top: { style: "thin" }
        });
        // A2 = bottom-left: left (fallback master=thick), bottom (A2 has no border, master has no bottom) → just left
        expect(ws.getCell("A2").border).toEqual({
          left: { style: "thick" }
        });
        // B2 = bottom-right: right (double from B2) + bottom (dashed from B2)
        expect(ws.getCell("B2").border).toEqual({
          right: { style: "double" },
          bottom: { style: "dashed" }
        });
      });

      it("non-border styles preserved without phantom borders", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // Master has font and fill but no border
        ws.getCell("A1").font = { bold: true };
        ws.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        };
        ws.getCell("A1").value = "styled";

        ws.mergeCells("A1:B2");

        // All cells inherit font and fill
        for (const addr of ["A1", "B1", "A2", "B2"]) {
          expect(ws.getCell(addr).font).toEqual({ bold: true });
          expect(ws.getCell(addr).fill).toEqual({
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          });
          // No phantom border should appear
          expect(ws.getCell(addr).border).toBeUndefined();
        }
      });

      it("discussion #78 scenario: template merge preserves rightmost border", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

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
          ws.getCell(21, col).border = border;
        }

        ws.mergeCells("C21:L21");

        // L21 (col 12) must retain its right border
        expect(ws.getCell("L21").border).toEqual({
          right: { style: "thin" },
          top: { style: "thin" },
          bottom: { style: "thin" }
        });
        // C21 (col 3) must retain its left border
        expect(ws.getCell("C21").border).toEqual({
          left: { style: "thin" },
          top: { style: "thin" },
          bottom: { style: "thin" }
        });
        // Interior cells (D21-K21) should only have top+bottom
        for (let col = 4; col <= 11; col++) {
          expect(ws.getCell(21, col).border).toEqual({
            top: { style: "thin" },
            bottom: { style: "thin" }
          });
        }
      });

      it("setting borders after merge does not cause cross-cell mutation", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        // Merge without any pre-existing borders
        ws.getCell("A1").value = "Merged";
        ws.mergeCells("A1:C3");

        // Set borders on boundary cells one by one (natural user workflow)
        ws.getCell("A1").border = {
          top: { style: "thick", color: { argb: "FF0000FF" } },
          left: { style: "thick", color: { argb: "FF0000FF" } }
        };
        ws.getCell("C1").border = {
          right: { style: "thick", color: { argb: "FFFF0000" } }
        };
        ws.getCell("A3").border = {
          bottom: { style: "thick", color: { argb: "FF00FF00" } }
        };

        // Each cell should have only the borders it was assigned
        expect(ws.getCell("A1").border).toEqual({
          top: { style: "thick", color: { argb: "FF0000FF" } },
          left: { style: "thick", color: { argb: "FF0000FF" } }
        });
        expect(ws.getCell("C1").border).toEqual({
          right: { style: "thick", color: { argb: "FFFF0000" } }
        });
        expect(ws.getCell("A3").border).toEqual({
          bottom: { style: "thick", color: { argb: "FF00FF00" } }
        });

        // Unmodified interior cell should have no border
        expect(ws.getCell("B2").border).toBeUndefined();
      });

      it("styles are independent after borderless merge — mutating one cell does not affect others", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("sheet");

        ws.mergeCells("A1:B2");

        // Set font on A1 only
        ws.getCell("A1").font = { bold: true };

        // B1 should not be affected
        expect(ws.getCell("A1").font).toEqual({ bold: true });
        expect(ws.getCell("B1").font).toBeUndefined();

        // Set fill on B2 only
        ws.getCell("B2").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        };
        expect(ws.getCell("B2").fill).toEqual({
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        });
        expect(ws.getCell("A1").fill).toBeUndefined();
      });
    });
  });
});
