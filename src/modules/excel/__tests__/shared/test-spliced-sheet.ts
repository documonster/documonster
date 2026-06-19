import { addSheetTo } from "@excel/__tests__/shared/add-sheet-to";
import {
  cellGetValue,
  cellSetBorder,
  cellSetFill,
  cellSetNumFmt,
  cellSetValue,
  cellType,
  cellName,
  cellSetName
} from "@excel/core/cell";
import { ValueType } from "@excel/core/enums";
import { rowSetAlignment, rowSetBorder, rowValues } from "@excel/core/row";
import { getCell, setColumns } from "@excel/core/worksheet";
import { Workbook, Worksheet } from "@excel/index";
import { expect } from "vitest";

export const splice = {
  rows: {
    removeOnly: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-remove-only");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

        Worksheet.spliceRows(ws, 2, 1);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-remove-only")!;
        expect(ws).toBeDefined();

        expect(cellGetValue(getCell(ws, "A1"))).toBe("1,1");
        expect(cellType(getCell(ws, "A1"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "B1"))).toBe("1,2");
        expect(cellType(getCell(ws, "B1"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "C1"))).toBe("1,3");
        expect(cellType(getCell(ws, "C1"))).toBe(ValueType.String);

        expect(cellType(getCell(ws, "A2"))).toBe(ValueType.Null);
        expect(cellType(getCell(ws, "B2"))).toBe(ValueType.Null);
        expect(cellType(getCell(ws, "C2"))).toBe(ValueType.Null);

        expect(cellGetValue(getCell(ws, "A3"))).toBe(4.1);
        expect(cellType(getCell(ws, "A3"))).toBe(ValueType.Number);
        expect(cellType(getCell(ws, "B3"))).toBe(ValueType.Null);
        expect(cellGetValue(getCell(ws, "C3"))).toBe(4.3);
        expect(cellType(getCell(ws, "C3"))).toBe(ValueType.Number);

        expect(cellGetValue(getCell(ws, "A4"))).toBe("5,1");
        expect(cellType(getCell(ws, "A4"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "B4"))).toBe("5,2");
        expect(cellType(getCell(ws, "B4"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "C4"))).toBe("5,3");
        expect(cellType(getCell(ws, "C4"))).toBe(ValueType.String);

        Worksheet.addRow(ws, ["5,1b", "5,2b", "5,3b"]);
        expect(cellGetValue(getCell(ws, "A5"))).toBe("5,1b");
        expect(cellType(getCell(ws, "A5"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "B5"))).toBe("5,2b");
        expect(cellType(getCell(ws, "B5"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "C5"))).toBe("5,3b");
        expect(cellType(getCell(ws, "C5"))).toBe(ValueType.String);
      }
    },
    insertFewer: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-insert-fewer");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

        Worksheet.spliceRows(ws, 2, 2, ["one", "two", "three"]);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-insert-fewer")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, 4.1, , 4.3]);
        expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, "5,1", "5,2", "5,3"]);
      }
    },
    insertSame: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-insert-same");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

        Worksheet.spliceRows(ws, 2, 2, ["one", "two", "three"], ["une", "deux", "trois"]);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-insert-same")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "une", "deux", "trois"]);
        expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, 4.1, , 4.3]);
        expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([, "5,1", "5,2", "5,3"]);
      }
    },
    insertMore: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-insert-more");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

        Worksheet.spliceRows(
          ws,
          2,
          2,
          ["one", "two", "three"],
          ["une", "deux", "trois"],
          ["uno", "due", "tre"]
        );
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-insert-more")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "une", "deux", "trois"]);
        expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, "uno", "due", "tre"]);
        expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([, 4.1, , 4.3]);
        expect(rowValues(Worksheet.getRow(ws, 6))).toEqual([, "5,1", "5,2", "5,3"]);
      }
    },
    removeStyle: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-remove-style");
        Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
        Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);
        Worksheet.addRow(ws, ["4,1", "4,2", "4,3", "4,4"]);

        cellSetNumFmt(getCell(ws, "A1"), "# ?/?");
        cellSetFill(getCell(ws, "B2"), {
          type: "pattern",
          pattern: "darkVertical",
          fgColor: { argb: "FFFF0000" }
        });
        rowSetBorder(Worksheet.getRow(ws, 3), {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        });
        rowSetAlignment(Worksheet.getRow(ws, 4), {
          horizontal: "left",
          vertical: "middle"
        });

        // remove rows 2 & 3
        Worksheet.spliceRows(ws, 2, 2);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-remove-style")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).to.deep.equal([, "1,1", "1,2", "1,3", "1,4"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).to.deep.equal([, "4,1", "4,2", "4,3", "4,4"]);

        expect(getCell(ws, "A1").style).to.deep.equal({
          numFmt: "# ?/?"
        });
        expect(Worksheet.getRow(ws, 2).style).to.deep.equal({
          alignment: {
            horizontal: "left",
            vertical: "middle"
          }
        });
      }
    },
    insertStyle: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-insert-style");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
        cellSetFill(getCell(ws, "A2"), {
          type: "pattern",
          pattern: "darkVertical",
          fgColor: { argb: "FFFF0000" }
        });
        rowSetAlignment(Worksheet.getRow(ws, 2), {
          horizontal: "left",
          vertical: "middle"
        });

        Worksheet.spliceRows(ws, 2, 0, ["one", "two", "three"]);
        cellSetBorder(getCell(ws, "A2"), {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        });
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-insert-style")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "2,1", "2,2", "2,3"]);

        expect(Worksheet.getRow(ws, 3).style.alignment).to.deep.equal({
          horizontal: "left",
          vertical: "middle"
        });
        expect(getCell(ws, "A2").style.border).to.deep.equal({
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        });
        expect(getCell(ws, "A3").style.alignment).to.deep.equal({
          horizontal: "left",
          vertical: "middle"
        });
        expect(getCell(ws, "A3").style.fill).to.deep.equal({
          type: "pattern",
          pattern: "darkVertical",
          fgColor: { argb: "FFFF0000" }
        });
      }
    },
    replaceStyle: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-row-replace-style");
        Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
        Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);

        cellSetNumFmt(getCell(ws, "B1"), "top");
        cellSetNumFmt(getCell(ws, "B2"), "middle");
        cellSetNumFmt(getCell(ws, "B3"), "bottom");

        rowSetAlignment(Worksheet.getRow(ws, 1), {
          horizontal: "left",
          vertical: "top"
        });
        rowSetAlignment(Worksheet.getRow(ws, 2), {
          horizontal: "center",
          vertical: "middle"
        });
        rowSetAlignment(Worksheet.getRow(ws, 3), {
          horizontal: "right",
          vertical: "bottom"
        });

        // remove rows 2 & 3
        Worksheet.spliceRows(ws, 2, 1, ["two-one", "two-two", "two-three", "two-four"]);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-row-replace-style")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).to.deep.equal([, "1,1", "1,2", "1,3", "1,4"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).to.deep.equal([
          ,
          "two-one",
          "two-two",
          "two-three",
          "two-four"
        ]);
        expect(rowValues(Worksheet.getRow(ws, 3))).to.deep.equal([, "3,1", "3,2", "3,3", "3,4"]);

        expect(getCell(ws, "B1").style).to.deep.equal({
          numFmt: "top",
          alignment: {
            horizontal: "left",
            vertical: "top"
          }
        });
        expect(getCell(ws, "B2").style).toEqual({});
        expect(getCell(ws, "B3").style).to.deep.equal({
          numFmt: "bottom",
          alignment: {
            horizontal: "right",
            vertical: "bottom"
          }
        });
        expect(Worksheet.getRow(ws, 1).style).to.deep.equal({
          alignment: {
            horizontal: "left",
            vertical: "top"
          }
        });
        expect(Worksheet.getRow(ws, 2).style).toEqual({});
        expect(Worksheet.getRow(ws, 3).style).to.deep.equal({
          alignment: {
            horizontal: "right",
            vertical: "bottom"
          }
        });
      }
    },
    removeDefinedNames: {
      addSheet(wb: any) {
        const wsSquare = addSheetTo(wb, "splice-row-remove-name-square");
        Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
        Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
        Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
        Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3, 4].forEach(row => {
            cellSetName(getCell(wsSquare, col + row), "square");
          });
        });

        Worksheet.spliceRows(wsSquare, 2, 2);

        const wsSingles = addSheetTo(wb, "splice-row-remove-name-singles");
        cellSetValue(getCell(wsSingles, "A1"), "1,1");
        cellSetValue(getCell(wsSingles, "A4"), "4,1");
        cellSetValue(getCell(wsSingles, "D1"), "1,4");
        cellSetValue(getCell(wsSingles, "D4"), "4,4");

        ["A", "D"].forEach(col => {
          [1, 4].forEach(row => {
            cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
          });
        });

        Worksheet.spliceRows(wsSingles, 2, 2);
      },

      checkSheet(wb: any) {
        const wsSquare = Workbook.getWorksheet(wb, "splice-row-remove-name-square")!;
        expect(wsSquare).toBeDefined();

        expect(rowValues(Worksheet.getRow(wsSquare, 1))).to.deep.equal([
          ,
          "1,1",
          "1,2",
          "1,3",
          "1,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 2))).to.deep.equal([
          ,
          "4,1",
          "4,2",
          "4,3",
          "4,4"
        ]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3].forEach(row => {
            if (row === 3) {
              expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
            } else {
              expect(cellName(getCell(wsSquare, col + row))).toBe("square");
            }
          });
        });

        const wsSingles = Workbook.getWorksheet(wb, "splice-row-remove-name-singles")!;
        expect(wsSingles).toBeDefined();

        expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
        expect(rowValues(Worksheet.getRow(wsSingles, 2))).toEqual([, "4,1", , , "4,4"]);

        expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
        expect(cellName(getCell(wsSingles, "A2"))).toBe("single-A4");
        expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
        expect(cellName(getCell(wsSingles, "D2"))).toBe("single-D4");
      }
    },
    insertDefinedNames: {
      addSheet(wb: any) {
        const wsSquare = addSheetTo(wb, "splice-row-insert-name-square");
        Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
        Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
        Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
        Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3, 4].forEach(row => {
            cellSetName(getCell(wsSquare, col + row), "square");
          });
        });

        Worksheet.spliceRows(wsSquare, 3, 0, ["foo", "bar", "baz", "qux"]);

        const wsSingles = addSheetTo(wb, "splice-row-insert-name-singles");
        cellSetValue(getCell(wsSingles, "A1"), "1,1");
        cellSetValue(getCell(wsSingles, "A4"), "4,1");
        cellSetValue(getCell(wsSingles, "D1"), "1,4");
        cellSetValue(getCell(wsSingles, "D4"), "4,4");

        ["A", "D"].forEach(col => {
          [1, 4].forEach(row => {
            cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
          });
        });

        Worksheet.spliceRows(wsSingles, 3, 0, ["foo", "bar", "baz", "qux"]);
      },

      checkSheet(wb: any) {
        const wsSquare = Workbook.getWorksheet(wb, "splice-row-insert-name-square")!;
        expect(wsSquare).toBeDefined();

        expect(rowValues(Worksheet.getRow(wsSquare, 1))).to.deep.equal([
          ,
          "1,1",
          "1,2",
          "1,3",
          "1,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 2))).to.deep.equal([
          ,
          "2,1",
          "2,2",
          "2,3",
          "2,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 3))).to.deep.equal([
          ,
          "foo",
          "bar",
          "baz",
          "qux"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 4))).to.deep.equal([
          ,
          "3,1",
          "3,2",
          "3,3",
          "3,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 5))).to.deep.equal([
          ,
          "4,1",
          "4,2",
          "4,3",
          "4,4"
        ]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3, 4, 5].forEach(row => {
            if (row === 3) {
              expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
            } else {
              expect(cellName(getCell(wsSquare, col + row))).toBe("square");
            }
          });
        });

        const wsSingles = Workbook.getWorksheet(wb, "splice-row-insert-name-singles")!;
        expect(wsSingles).toBeDefined();
        expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
        expect(rowValues(Worksheet.getRow(wsSingles, 3))).to.deep.equal([
          ,
          "foo",
          "bar",
          "baz",
          "qux"
        ]);
        expect(rowValues(Worksheet.getRow(wsSingles, 5))).toEqual([, "4,1", , , "4,4"]);

        expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
        expect(cellName(getCell(wsSingles, "A5"))).toBe("single-A4");
        expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
        expect(cellName(getCell(wsSingles, "D5"))).toBe("single-D4");
      }
    },
    replaceDefinedNames: {
      addSheet(wb: any) {
        const wsSquare = addSheetTo(wb, "splice-row-replace-name-square");
        Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
        Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
        Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
        Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3, 4].forEach(row => {
            cellSetName(getCell(wsSquare, col + row), "square");
          });
        });

        Worksheet.spliceRows(wsSquare, 2, 1, ["foo", "bar", "baz", "qux"]);

        const wsSingles = addSheetTo(wb, "splice-row-replace-name-singles");
        cellSetValue(getCell(wsSingles, "A1"), "1,1");
        cellSetValue(getCell(wsSingles, "A4"), "4,1");
        cellSetValue(getCell(wsSingles, "D1"), "1,4");
        cellSetValue(getCell(wsSingles, "D4"), "4,4");

        ["A", "D"].forEach(col => {
          [1, 4].forEach(row => {
            cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
          });
        });

        Worksheet.spliceRows(wsSingles, 2, 1, ["foo", "bar", "baz", "qux"]);
      },

      checkSheet(wb: any) {
        const wsSquare = Workbook.getWorksheet(wb, "splice-row-replace-name-square")!;
        expect(wsSquare).toBeDefined();

        expect(rowValues(Worksheet.getRow(wsSquare, 1))).to.deep.equal([
          ,
          "1,1",
          "1,2",
          "1,3",
          "1,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 2))).to.deep.equal([
          ,
          "foo",
          "bar",
          "baz",
          "qux"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 3))).to.deep.equal([
          ,
          "3,1",
          "3,2",
          "3,3",
          "3,4"
        ]);
        expect(rowValues(Worksheet.getRow(wsSquare, 4))).to.deep.equal([
          ,
          "4,1",
          "4,2",
          "4,3",
          "4,4"
        ]);

        ["A", "B", "C", "D"].forEach(col => {
          [1, 2, 3, 4].forEach(row => {
            if (row === 2) {
              expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
            } else {
              expect(cellName(getCell(wsSquare, col + row))).toBe("square");
            }
          });
        });

        const wsSingles = Workbook.getWorksheet(wb, "splice-row-replace-name-singles")!;
        expect(wsSingles).toBeDefined();

        expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
        expect(rowValues(Worksheet.getRow(wsSingles, 2))).to.deep.equal([
          ,
          "foo",
          "bar",
          "baz",
          "qux"
        ]);
        expect(rowValues(Worksheet.getRow(wsSingles, 4))).toEqual([, "4,1", , , "4,4"]);

        expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
        expect(cellName(getCell(wsSingles, "A4"))).toBe("single-A4");
        expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
        expect(cellName(getCell(wsSingles, "D4"))).toBe("single-D4");
      }
    }
  },
  columns: {
    removeOnly: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-column-remove-only");

        setColumns(ws, [
          { key: "id", width: 10 },
          { key: "name", width: 32 },
          { key: "dob", width: 10 }
        ]);

        Worksheet.addRow(ws, { id: "id1", name: "name1", dob: "dob1" });
        Worksheet.addRow(ws, { id: 2, dob: "dob2" });
        Worksheet.addRow(ws, { name: "name3", dob: 3 });

        Worksheet.spliceColumns(ws, 2, 1);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-column-remove-only")!;
        expect(ws).toBeDefined();

        expect(cellGetValue(getCell(ws, "A1"))).toBe("id1");
        expect(cellType(getCell(ws, "A1"))).toBe(ValueType.String);
        expect(cellGetValue(getCell(ws, "B1"))).toBe("dob1");
        expect(cellType(getCell(ws, "B1"))).toBe(ValueType.String);
        expect(cellType(getCell(ws, "C1"))).toBe(ValueType.Null);

        expect(cellGetValue(getCell(ws, "A2"))).toBe(2);
        expect(cellType(getCell(ws, "A2"))).toBe(ValueType.Number);
        expect(cellGetValue(getCell(ws, "B2"))).toBe("dob2");
        expect(cellType(getCell(ws, "B2"))).toBe(ValueType.String);
        expect(cellType(getCell(ws, "C2"))).toBe(ValueType.Null);

        expect(cellType(getCell(ws, "A3"))).toBe(ValueType.Null);
        expect(cellGetValue(getCell(ws, "B3"))).toBe(3);
        expect(cellType(getCell(ws, "B3"))).toBe(ValueType.Number);
        expect(cellType(getCell(ws, "C3"))).toBe(ValueType.Null);
      }
    },
    insertFewer: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-column-insert-fewer");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        cellSetValue(getCell(ws, "E4"), 4.5);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

        Worksheet.spliceColumns(ws, 2, 2, ["one", "two", "three", "four", "five"]);
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-column-insert-fewer")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).to.deep.equal([, "1,1", "one", "1,4", "1,5"]);
        expect(rowValues(Worksheet.getRow(ws, 2))).to.deep.equal([, "2,1", "two", "2,4", "2,5"]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, , "three"]);
        expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, 4.1, "four", , 4.5]);
        expect(rowValues(Worksheet.getRow(ws, 5))).to.deep.equal([, "5,1", "five", "5,4", "5,5"]);
      }
    },
    insertSame: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-column-insert-same");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        cellSetValue(getCell(ws, "E4"), 4.5);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

        Worksheet.spliceColumns(
          ws,
          2,
          2,
          ["one", "two", "three", "four", "five"],
          ["une", "deux", "trois", "quatre", "cinq"]
        );
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-column-insert-same")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).to.deep.equal([
          ,
          "1,1",
          "one",
          "une",
          "1,4",
          "1,5"
        ]);
        expect(rowValues(Worksheet.getRow(ws, 2))).to.deep.equal([
          ,
          "2,1",
          "two",
          "deux",
          "2,4",
          "2,5"
        ]);
        expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, , "three", "trois"]);
        expect(rowValues(Worksheet.getRow(ws, 4))).to.deep.equal([, 4.1, "four", "quatre", , 4.5]);
        expect(rowValues(Worksheet.getRow(ws, 5))).to.deep.equal([
          ,
          "5,1",
          "five",
          "cinq",
          "5,4",
          "5,5"
        ]);
      }
    },
    insertMore: {
      addSheet(wb: any) {
        const ws = addSheetTo(wb, "splice-column-insert-more");

        Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
        Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
        cellSetValue(getCell(ws, "A4"), 4.1);
        cellSetValue(getCell(ws, "C4"), 4.3);
        cellSetValue(getCell(ws, "E4"), 4.5);
        Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

        Worksheet.spliceColumns(
          ws,
          2,
          2,
          ["one", "two", "three", "four", "five"],
          ["une", "deux", "trois", "quatre", "cinq"],
          ["uno", "due", "tre", "quatro", "cinque"]
        );
      },

      checkSheet(wb: any) {
        const ws = Workbook.getWorksheet(wb, "splice-column-insert-more")!;
        expect(ws).toBeDefined();

        expect(rowValues(Worksheet.getRow(ws, 1))).to.deep.equal([
          ,
          "1,1",
          "one",
          "une",
          "uno",
          "1,4",
          "1,5"
        ]);
        expect(rowValues(Worksheet.getRow(ws, 2))).to.deep.equal([
          ,
          "2,1",
          "two",
          "deux",
          "due",
          "2,4",
          "2,5"
        ]);
        expect(rowValues(Worksheet.getRow(ws, 3))).to.deep.equal([, , "three", "trois", "tre"]);
        expect(rowValues(Worksheet.getRow(ws, 4))).to.deep.equal([
          ,
          4.1,
          "four",
          "quatre",
          "quatro",
          ,
          4.5
        ]);
        expect(rowValues(Worksheet.getRow(ws, 5))).to.deep.equal([
          ,
          "5,1",
          "five",
          "cinq",
          "cinque",
          "5,4",
          "5,5"
        ]);
      }
    }
  }
};
