import { addSheetTo, cellAt, mergeAt } from "@excel/__tests__/shared/add-sheet-to";
import headerFooterJson from "@excel/__tests__/shared/data/header-footer.json" with { type: "json" };
import pageSetupJson from "@excel/__tests__/shared/data/page-setup.json" with { type: "json" };
import propertiesJson from "@excel/__tests__/shared/data/sheet-properties.json" with { type: "json" };
import testValuesJson from "@excel/__tests__/shared/data/sheet-values.json" with { type: "json" };
import stylesJson from "@excel/__tests__/shared/data/styles.json" with { type: "json" };
import { fix } from "@excel/__tests__/shared/tools";
import {
  cellFill,
  cellSetFill,
  cellSetValue,
  cellAlignment,
  cellBorder,
  cellFont,
  cellFormula,
  cellGetValue,
  cellMaster,
  cellNumFmt,
  cellSetAlignment,
  cellSetBorder,
  cellSetFont,
  cellSetNumFmt,
  cellType
} from "@excel/cell";
import { columnCollapsed } from "@excel/column";
import { ValueType } from "@excel/enums";
import { Workbook, Worksheet } from "@excel/index";
import { rowCollapsed, rowOutlineLevel, rowSetOutlineLevel } from "@excel/row";
import { rowCommit, rowGetCell, getColumn, fillFormula } from "@excel/worksheet";

const testValues = fix(testValuesJson);
const styles = fix(stylesJson);
const properties = fix(propertiesJson);
const pageSetup = fix(pageSetupJson);
const headerFooter = fix(headerFooterJson);

/** Value accessor returning `any` so callers can read concrete value shapes (.getTime/.formula/...) without union narrowing. */
function gv(ws: any, addr: string | number): any {
  return cellGetValue(cellAt(ws, addr));
}

export const values = {
  addSheet(wb: any, options: any) {
    // call it sheet1 so this sheet can be used for csv testing
    const ws = addSheetTo(wb, "sheet1", {
      properties: properties,
      pageSetup: pageSetup,
      headerFooter: headerFooter
    });

    cellSetValue(cellAt(ws, "J10"), 1);
    getColumn(ws, 10).outlineLevel = 1;
    rowSetOutlineLevel(Worksheet.getRow(ws, 10), 1);

    cellSetValue(cellAt(ws, "A1"), 7);
    cellSetValue(cellAt(ws, "B1"), (testValues as any).str);
    cellSetValue(cellAt(ws, "C1"), (testValues as any).date);
    cellSetValue(cellAt(ws, "D1"), (testValues as any).formulas[0]);
    cellSetValue(cellAt(ws, "E1"), (testValues as any).formulas[1]);
    cellSetValue(cellAt(ws, "F1"), (testValues as any).hyperlink);
    cellSetValue(cellAt(ws, "G1"), (testValues as any).str2);
    cellSetValue(cellAt(ws, "H1"), (testValues as any).json.raw);
    cellSetValue(cellAt(ws, "I1"), true);
    cellSetValue(cellAt(ws, "J1"), false);
    cellSetValue(cellAt(ws, "K1"), (testValues as any).Errors.NotApplicable);
    cellSetValue(cellAt(ws, "L1"), (testValues as any).Errors.Value);

    rowCommit(Worksheet.getRow(ws, 1));

    // merge cell square with numerical value
    cellSetValue(cellAt(ws, "A2"), 5);
    mergeAt(ws, "A2:B3");

    // merge cell square with null value
    mergeAt(ws, "C2:D3");
    rowCommit(Worksheet.getRow(ws, 3));

    cellSetValue(cellAt(ws, "A4"), 1.5);
    cellSetNumFmt(cellAt(ws, "A4"), (testValues as any).numFmt1);
    cellSetBorder(cellAt(ws, "A4"), (styles as any).borders.thin);
    cellSetValue(cellAt(ws, "C4"), 1.5);
    cellSetNumFmt(cellAt(ws, "C4"), (testValues as any).numFmt2);
    cellSetBorder(cellAt(ws, "C4"), (styles as any).borders.doubleRed);
    cellSetValue(cellAt(ws, "E4"), 1.5);
    cellSetBorder(cellAt(ws, "E4"), (styles as any).borders.thickRainbow);
    rowCommit(Worksheet.getRow(ws, 4));

    // test fonts and formats
    cellSetValue(cellAt(ws, "A5"), (testValues as any).str);
    cellSetFont(cellAt(ws, "A5"), (styles as any).fonts.arialBlackUI14);
    cellSetValue(cellAt(ws, "B5"), (testValues as any).str);
    cellSetFont(cellAt(ws, "B5"), (styles as any).fonts.broadwayRedOutline20);
    cellSetValue(cellAt(ws, "C5"), (testValues as any).str);
    cellSetFont(cellAt(ws, "C5"), (styles as any).fonts.comicSansUdB16);

    cellSetValue(cellAt(ws, "D5"), 1.6);
    cellSetNumFmt(cellAt(ws, "D5"), (testValues as any).numFmt1);
    cellSetFont(cellAt(ws, "D5"), (styles as any).fonts.arialBlackUI14);

    cellSetValue(cellAt(ws, "E5"), 1.6);
    cellSetNumFmt(cellAt(ws, "E5"), (testValues as any).numFmt2);
    cellSetFont(cellAt(ws, "E5"), (styles as any).fonts.broadwayRedOutline20);

    cellSetValue(cellAt(ws, "F5"), (testValues as any).date);
    cellSetNumFmt(cellAt(ws, "F5"), (testValues as any).numFmtDate);
    cellSetFont(cellAt(ws, "F5"), (styles as any).fonts.comicSansUdB16);
    rowCommit(Worksheet.getRow(ws, 5));

    Worksheet.getRow(ws, 6).height = 42;
    (styles as any).alignments.forEach((alignment: any, index: number) => {
      const rowNumber = 6;
      const colNumber = index + 1;
      const cell = cellAt(ws, rowNumber, colNumber);
      cellSetValue(cell, alignment.text);
      cellSetAlignment(cell, alignment.alignment);
    });
    rowCommit(Worksheet.getRow(ws, 6));

    if (options.checkBadAlignments) {
      (styles as any).badAlignments.forEach((alignment: any, index: number) => {
        const rowNumber = 7;
        const colNumber = index + 1;
        const cell = cellAt(ws, rowNumber, colNumber);
        cellSetValue(cell, alignment.text);
        cellSetAlignment(cell, alignment.alignment);
      });
    }
    rowCommit(Worksheet.getRow(ws, 7));

    const row8 = Worksheet.getRow(ws, 8);
    row8.height = 40;
    cellSetValue(rowGetCell(row8, 1), "Blue White Horizontal Gradient");
    cellSetFill(rowGetCell(row8, 1), (styles as any).fills.blueWhiteHGrad);
    cellSetValue(rowGetCell(row8, 2), "Red Dark Vertical");
    cellSetFill(rowGetCell(row8, 2), (styles as any).fills.redDarkVertical);
    cellSetValue(rowGetCell(row8, 3), "Red Green Dark Trellis");
    cellSetFill(rowGetCell(row8, 3), (styles as any).fills.redGreenDarkTrellis);
    cellSetValue(rowGetCell(row8, 4), "RGB Path Gradient");
    cellSetFill(rowGetCell(row8, 4), (styles as any).fills.rgbPathGrad);
    rowCommit(row8);

    // Old Shared Formula
    cellSetValue(cellAt(ws, "A9"), 1);
    cellSetValue(cellAt(ws, "B9"), { formula: "A9+1", result: 2 });
    cellSetValue(cellAt(ws, "C9"), { sharedFormula: "B9", result: 3 });
    cellSetValue(cellAt(ws, "D9"), { sharedFormula: "B9", result: 4 });
    cellSetValue(cellAt(ws, "E9"), { sharedFormula: "B9", result: 5 });

    // Streaming writers don't support fillFormula; record worksheets do (it's
    // a flat helper rather than a method now).
    if (typeof (ws as any).commit !== "function") {
      // Fill Formula Shared
      fillFormula(ws, "A10:E10", "A9", [1, 2, 3, 4, 5]);

      // Array Formula
      fillFormula(ws, "A11:E11", "A9", [1, 1, 1, 1, 1], "array");
    }
  },

  checkSheet(wb: any, options: any) {
    const ws = Workbook.getWorksheet(wb, "sheet1")!;
    expect(ws).toBeDefined();

    if (options.checkSheetProperties) {
      expect(getColumn(ws, 10).outlineLevel).toBe(1);
      expect(columnCollapsed(getColumn(ws, 10))).toBe(true);
      expect(rowOutlineLevel(Worksheet.getRow(ws, 10))).toBe(1);
      expect(rowCollapsed(Worksheet.getRow(ws, 10))).toBe(true);
      expect(ws.properties.outlineLevelCol).toBe(1);
      expect(ws.properties.outlineLevelRow).toBe(1);
      expect(ws.properties.tabColor).toEqual({ argb: "FF00FF00" });
      expect(ws.properties).toEqual(properties);
      expect(ws.pageSetup).toEqual(pageSetup);
      expect(ws.headerFooter).toEqual(headerFooter);
    }

    expect(gv(ws, "A1")).toBe(7);
    expect(cellType(cellAt(ws, "A1"))).toBe(ValueType.Number);
    expect(gv(ws, "B1")).toBe((testValues as any).str);
    expect(cellType(cellAt(ws, "B1"))).toBe(ValueType.String);
    expect(Math.abs(gv(ws, "C1").getTime() - (testValues as any).date.getTime())).to.be.below(
      options.dateAccuracy
    );
    expect(cellType(cellAt(ws, "C1"))).toBe(ValueType.Date);

    if (options.checkFormulas) {
      expect(gv(ws, "D1")).toEqual((testValues as any).formulas[0]);
      expect(cellType(cellAt(ws, "D1"))).toBe(ValueType.Formula);
      expect(gv(ws, "E1").formula).to.equal((testValues as any).formulas[1].formula);
      expect(gv(ws, "E1").value).toBeUndefined();
      expect(cellType(cellAt(ws, "E1"))).toBe(ValueType.Formula);
      expect(gv(ws, "F1")).toEqual((testValues as any).hyperlink);
      expect(cellType(cellAt(ws, "F1"))).toBe(ValueType.Hyperlink);
      expect(gv(ws, "G1")).toBe((testValues as any).str2);
    } else {
      expect(gv(ws, "D1")).to.equal((testValues as any).formulas[0].result);
      expect(cellType(cellAt(ws, "D1"))).toBe(ValueType.Number);
      expect(gv(ws, "E1")).toBeNull();
      expect(cellType(cellAt(ws, "E1"))).toBe(ValueType.Null);
      expect(gv(ws, "F1")).to.deep.equal((testValues as any).hyperlink.hyperlink);
      expect(cellType(cellAt(ws, "F1"))).toBe(ValueType.String);
      expect(gv(ws, "G1")).toBe((testValues as any).str2);
    }

    expect(gv(ws, "H1")).toBe((testValues as any).json.string);
    expect(cellType(cellAt(ws, "H1"))).toBe(ValueType.String);

    expect(gv(ws, "I1")).toBe(true);
    expect(cellType(cellAt(ws, "I1"))).toBe(ValueType.Boolean);
    expect(gv(ws, "J1")).toBe(false);
    expect(cellType(cellAt(ws, "J1"))).toBe(ValueType.Boolean);

    expect(gv(ws, "K1")).to.deep.equal((testValues as any).Errors.NotApplicable);
    expect(cellType(cellAt(ws, "K1"))).toBe(ValueType.Error);
    expect(gv(ws, "L1")).toEqual((testValues as any).Errors.Value);
    expect(cellType(cellAt(ws, "L1"))).toBe(ValueType.Error);

    // A2:B3
    expect(gv(ws, "A2")).toBe(5);
    expect(cellType(cellAt(ws, "A2"))).toBe(ValueType.Number);
    expect(cellMaster(cellAt(ws, "A2"))).toBe(cellAt(ws, "A2"));

    if (options.checkMerges) {
      expect(gv(ws, "A3")).toBe(5);
      expect(cellType(cellAt(ws, "A3"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "A3"))).toBe(cellAt(ws, "A2"));

      expect(gv(ws, "B2")).toBe(5);
      expect(cellType(cellAt(ws, "B2"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "B2"))).toBe(cellAt(ws, "A2"));

      expect(gv(ws, "B3")).toBe(5);
      expect(cellType(cellAt(ws, "B3"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "B3"))).toBe(cellAt(ws, "A2"));

      // C2:D3
      expect(gv(ws, "C2")).toBeNull();
      expect(cellType(cellAt(ws, "C2"))).toBe(ValueType.Null);
      expect(cellMaster(cellAt(ws, "C2"))).toBe(cellAt(ws, "C2"));

      expect(gv(ws, "D2")).toBeNull();
      expect(cellType(cellAt(ws, "D2"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "D2"))).toBe(cellAt(ws, "C2"));

      expect(gv(ws, "C3")).toBeNull();
      expect(cellType(cellAt(ws, "C3"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "C3"))).toBe(cellAt(ws, "C2"));

      expect(gv(ws, "D3")).toBeNull();
      expect(cellType(cellAt(ws, "D3"))).toBe(ValueType.Merge);
      expect(cellMaster(cellAt(ws, "D3"))).toBe(cellAt(ws, "C2"));
    }

    if (options.checkStyles) {
      expect(cellNumFmt(cellAt(ws, "A4"))).toBe((testValues as any).numFmt1);
      expect(cellType(cellAt(ws, "A4"))).toBe(ValueType.Number);
      expect(cellBorder(cellAt(ws, "A4"))).toEqual((styles as any).borders.thin);
      expect(cellNumFmt(cellAt(ws, "C4"))).toBe((testValues as any).numFmt2);
      expect(cellType(cellAt(ws, "C4"))).toBe(ValueType.Number);
      expect(cellBorder(cellAt(ws, "C4"))).to.deep.equal((styles as any).borders.doubleRed);
      expect(cellBorder(cellAt(ws, "E4"))).to.deep.equal((styles as any).borders.thickRainbow);

      // test fonts and formats
      expect(gv(ws, "A5")).toBe((testValues as any).str);
      expect(cellType(cellAt(ws, "A5"))).toBe(ValueType.String);
      expect(gv(ws, "B5")).toBe((testValues as any).str);
      expect(cellType(cellAt(ws, "B5"))).toBe(ValueType.String);
      expect(cellFont(cellAt(ws, "B5"))).to.deep.equal((styles as any).fonts.broadwayRedOutline20);
      expect(gv(ws, "C5")).toBe((testValues as any).str);
      expect(cellType(cellAt(ws, "C5"))).toBe(ValueType.String);
      expect(cellFont(cellAt(ws, "C5"))).to.deep.equal((styles as any).fonts.comicSansUdB16);

      expect(Math.abs(gv(ws, "D5") - 1.6)).to.be.below(0.00000001);
      expect(cellType(cellAt(ws, "D5"))).toBe(ValueType.Number);
      expect(cellNumFmt(cellAt(ws, "D5"))).toBe((testValues as any).numFmt1);
      expect(cellFont(cellAt(ws, "D5"))).to.deep.equal((styles as any).fonts.arialBlackUI14);

      expect(Math.abs(gv(ws, "E5") - 1.6)).to.be.below(0.00000001);
      expect(cellType(cellAt(ws, "E5"))).toBe(ValueType.Number);
      expect(cellNumFmt(cellAt(ws, "E5"))).toBe((testValues as any).numFmt2);
      expect(cellFont(cellAt(ws, "E5"))).to.deep.equal((styles as any).fonts.broadwayRedOutline20);

      expect(Math.abs(gv(ws, "F5").getTime() - (testValues as any).date.getTime())).to.be.below(
        options.dateAccuracy
      );
      expect(cellType(cellAt(ws, "F5"))).toBe(ValueType.Date);
      expect(cellNumFmt(cellAt(ws, "F5"))).toBe((testValues as any).numFmtDate);
      expect(cellFont(cellAt(ws, "F5"))).to.deep.equal((styles as any).fonts.comicSansUdB16);

      expect(Worksheet.getRow(ws, 5).height).toBeUndefined();
      expect(Worksheet.getRow(ws, 6).height).toBe(42);
      (styles as any).alignments.forEach((alignment: any, index: number) => {
        const rowNumber = 6;
        const colNumber = index + 1;
        const cell = cellAt(ws, rowNumber, colNumber);
        expect(cellGetValue(cell)).toBe(alignment.text);
        expect(cellAlignment(cell)).toEqual(alignment.alignment);
      });

      if (options.checkBadAlignments) {
        (styles as any).badAlignments.forEach((alignment: any, index: number) => {
          const rowNumber = 7;
          const colNumber = index + 1;
          const cell = cellAt(ws, rowNumber, colNumber);
          expect(cellGetValue(cell)).toBe(alignment.text);
          expect(cellAlignment(cell)).toBeUndefined();
        });
      }

      const row8 = Worksheet.getRow(ws, 8);
      expect(row8.height).toBe(40);
      expect(cellFill(rowGetCell(row8, 1))).to.deep.equal((styles as any).fills.blueWhiteHGrad);
      expect(cellFill(rowGetCell(row8, 2))).to.deep.equal((styles as any).fills.redDarkVertical);
      expect(cellFill(rowGetCell(row8, 3))).to.deep.equal(
        (styles as any).fills.redGreenDarkTrellis
      );
      expect(cellFill(rowGetCell(row8, 4))).toEqual((styles as any).fills.rgbPathGrad);

      if (options.checkFormulas) {
        // Shared Formula
        expect(gv(ws, "A9")).toBe(1);
        expect(cellType(cellAt(ws, "A9"))).toBe(ValueType.Number);

        expect(gv(ws, "B9")).to.deep.equal({
          shareType: "shared",
          ref: "B9:E9",
          formula: "A9+1",
          result: 2
        });
        expect(cellType(cellAt(ws, "B9"))).toBe(ValueType.Formula);

        ["C9", "D9", "E9"].forEach((address, index) => {
          expect(gv(ws, address)).to.deep.equal({
            sharedFormula: "B9",
            result: index + 3
          });
          expect(cellType(cellAt(ws, address))).toBe(ValueType.Formula);
        });

        if (gv(ws, "A10")) {
          // Fill Formula Shared
          expect(gv(ws, "A10")).to.deep.equal({
            shareType: "shared",
            ref: "A10:E10",
            formula: "A9",
            result: 1
          });
          ["B10", "C10", "D10", "E10"].forEach((address, index) => {
            expect(gv(ws, address)).to.deep.equal({
              sharedFormula: "A10",
              result: index + 2
            });
            expect(cellFormula(cellAt(ws, address))).toBe(`${address[0]}9`);
          });

          // Array Formula
          expect(gv(ws, "A11")).to.deep.equal({
            shareType: "array",
            ref: "A11:E11",
            formula: "A9",
            result: 1
          });
          ["B11", "C11", "D11", "E11"].forEach(address => {
            expect(gv(ws, address)).toBe(1);
          });
        }
      } else {
        ["A9", "B9", "C9", "D9", "E9"].forEach((address, index) => {
          expect(gv(ws, address)).toBe(index + 1);
          expect(cellType(cellAt(ws, address))).toBe(ValueType.Number);
        });
      }
    }
  }
};
