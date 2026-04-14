import headerFooterJson from "@excel/__tests__/shared/data/header-footer.json" with { type: "json" };
import pageSetupJson from "@excel/__tests__/shared/data/page-setup.json" with { type: "json" };
import propertiesJson from "@excel/__tests__/shared/data/sheet-properties.json" with { type: "json" };
import testValuesJson from "@excel/__tests__/shared/data/sheet-values.json" with { type: "json" };
import stylesJson from "@excel/__tests__/shared/data/styles.json" with { type: "json" };
import { fix } from "@excel/__tests__/shared/tools";
import { ValueType } from "@excel/enums";

const testValues = fix(testValuesJson);
const styles = fix(stylesJson);
const properties = fix(propertiesJson);
const pageSetup = fix(pageSetupJson);
const headerFooter = fix(headerFooterJson);

export const values = {
  addSheet(wb: any, options: any) {
    // call it sheet1 so this sheet can be used for csv testing
    const ws = wb.addWorksheet("sheet1", {
      properties: properties,
      pageSetup: pageSetup,
      headerFooter: headerFooter
    });

    ws.getCell("J10").value = 1;
    ws.getColumn(10).outlineLevel = 1;
    ws.getRow(10).outlineLevel = 1;

    ws.getCell("A1").value = 7;
    ws.getCell("B1").value = (testValues as any).str;
    ws.getCell("C1").value = (testValues as any).date;
    ws.getCell("D1").value = (testValues as any).formulas[0];
    ws.getCell("E1").value = (testValues as any).formulas[1];
    ws.getCell("F1").value = (testValues as any).hyperlink;
    ws.getCell("G1").value = (testValues as any).str2;
    ws.getCell("H1").value = (testValues as any).json.raw;
    ws.getCell("I1").value = true;
    ws.getCell("J1").value = false;
    ws.getCell("K1").value = (testValues as any).Errors.NotApplicable;
    ws.getCell("L1").value = (testValues as any).Errors.Value;

    ws.getRow(1).commit();

    // merge cell square with numerical value
    ws.getCell("A2").value = 5;
    ws.mergeCells("A2:B3");

    // merge cell square with null value
    ws.mergeCells("C2:D3");
    ws.getRow(3).commit();

    ws.getCell("A4").value = 1.5;
    ws.getCell("A4").numFmt = (testValues as any).numFmt1;
    ws.getCell("A4").border = (styles as any).borders.thin;
    ws.getCell("C4").value = 1.5;
    ws.getCell("C4").numFmt = (testValues as any).numFmt2;
    ws.getCell("C4").border = (styles as any).borders.doubleRed;
    ws.getCell("E4").value = 1.5;
    ws.getCell("E4").border = (styles as any).borders.thickRainbow;
    ws.getRow(4).commit();

    // test fonts and formats
    ws.getCell("A5").value = (testValues as any).str;
    ws.getCell("A5").font = (styles as any).fonts.arialBlackUI14;
    ws.getCell("B5").value = (testValues as any).str;
    ws.getCell("B5").font = (styles as any).fonts.broadwayRedOutline20;
    ws.getCell("C5").value = (testValues as any).str;
    ws.getCell("C5").font = (styles as any).fonts.comicSansUdB16;

    ws.getCell("D5").value = 1.6;
    ws.getCell("D5").numFmt = (testValues as any).numFmt1;
    ws.getCell("D5").font = (styles as any).fonts.arialBlackUI14;

    ws.getCell("E5").value = 1.6;
    ws.getCell("E5").numFmt = (testValues as any).numFmt2;
    ws.getCell("E5").font = (styles as any).fonts.broadwayRedOutline20;

    ws.getCell("F5").value = (testValues as any).date;
    ws.getCell("F5").numFmt = (testValues as any).numFmtDate;
    ws.getCell("F5").font = (styles as any).fonts.comicSansUdB16;
    ws.getRow(5).commit();

    ws.getRow(6).height = 42;
    (styles as any).alignments.forEach((alignment: any, index: number) => {
      const rowNumber = 6;
      const colNumber = index + 1;
      const cell = ws.getCell(rowNumber, colNumber);
      cell.value = alignment.text;
      cell.alignment = alignment.alignment;
    });
    ws.getRow(6).commit();

    if (options.checkBadAlignments) {
      (styles as any).badAlignments.forEach((alignment: any, index: number) => {
        const rowNumber = 7;
        const colNumber = index + 1;
        const cell = ws.getCell(rowNumber, colNumber);
        cell.value = alignment.text;
        cell.alignment = alignment.alignment;
      });
    }
    ws.getRow(7).commit();

    const row8 = ws.getRow(8);
    row8.height = 40;
    row8.getCell(1).value = "Blue White Horizontal Gradient";
    row8.getCell(1).fill = (styles as any).fills.blueWhiteHGrad;
    row8.getCell(2).value = "Red Dark Vertical";
    row8.getCell(2).fill = (styles as any).fills.redDarkVertical;
    row8.getCell(3).value = "Red Green Dark Trellis";
    row8.getCell(3).fill = (styles as any).fills.redGreenDarkTrellis;
    row8.getCell(4).value = "RGB Path Gradient";
    row8.getCell(4).fill = (styles as any).fills.rgbPathGrad;
    row8.commit();

    // Old Shared Formula
    ws.getCell("A9").value = 1;
    ws.getCell("B9").value = { formula: "A9+1", result: 2 };
    ws.getCell("C9").value = { sharedFormula: "B9", result: 3 };
    ws.getCell("D9").value = { sharedFormula: "B9", result: 4 };
    ws.getCell("E9").value = { sharedFormula: "B9", result: 5 };

    if ((ws as any).fillFormula) {
      // Fill Formula Shared
      (ws as any).fillFormula("A10:E10", "A9", [1, 2, 3, 4, 5]);

      // Array Formula
      (ws as any).fillFormula("A11:E11", "A9", [1, 1, 1, 1, 1], "array");
    }
  },

  checkSheet(wb: any, options: any) {
    const ws = wb.getWorksheet("sheet1");
    expect(ws).toBeDefined();

    if (options.checkSheetProperties) {
      expect(ws.getColumn(10).outlineLevel).toBe(1);
      expect(ws.getColumn(10).collapsed).toBe(true);
      expect(ws.getRow(10).outlineLevel).toBe(1);
      expect(ws.getRow(10).collapsed).toBe(true);
      expect(ws.properties.outlineLevelCol).toBe(1);
      expect(ws.properties.outlineLevelRow).toBe(1);
      expect(ws.properties.tabColor).toEqual({ argb: "FF00FF00" });
      expect(ws.properties).toEqual(properties);
      expect(ws.pageSetup).toEqual(pageSetup);
      expect(ws.headerFooter).toEqual(headerFooter);
    }

    expect(ws.getCell("A1").value).toBe(7);
    expect(ws.getCell("A1").type).toBe(ValueType.Number);
    expect(ws.getCell("B1").value).toBe((testValues as any).str);
    expect(ws.getCell("B1").type).toBe(ValueType.String);
    expect(
      Math.abs(ws.getCell("C1").value.getTime() - (testValues as any).date.getTime())
    ).to.be.below(options.dateAccuracy);
    expect(ws.getCell("C1").type).toBe(ValueType.Date);

    if (options.checkFormulas) {
      expect(ws.getCell("D1").value).toEqual((testValues as any).formulas[0]);
      expect(ws.getCell("D1").type).toBe(ValueType.Formula);
      expect(ws.getCell("E1").value.formula).to.equal((testValues as any).formulas[1].formula);
      expect(ws.getCell("E1").value.value).toBeUndefined();
      expect(ws.getCell("E1").type).toBe(ValueType.Formula);
      expect(ws.getCell("F1").value).toEqual((testValues as any).hyperlink);
      expect(ws.getCell("F1").type).toBe(ValueType.Hyperlink);
      expect(ws.getCell("G1").value).toBe((testValues as any).str2);
    } else {
      expect(ws.getCell("D1").value).to.equal((testValues as any).formulas[0].result);
      expect(ws.getCell("D1").type).toBe(ValueType.Number);
      expect(ws.getCell("E1").value).toBeNull();
      expect(ws.getCell("E1").type).toBe(ValueType.Null);
      expect(ws.getCell("F1").value).to.deep.equal((testValues as any).hyperlink.hyperlink);
      expect(ws.getCell("F1").type).toBe(ValueType.String);
      expect(ws.getCell("G1").value).toBe((testValues as any).str2);
    }

    expect(ws.getCell("H1").value).toBe((testValues as any).json.string);
    expect(ws.getCell("H1").type).toBe(ValueType.String);

    expect(ws.getCell("I1").value).toBe(true);
    expect(ws.getCell("I1").type).toBe(ValueType.Boolean);
    expect(ws.getCell("J1").value).toBe(false);
    expect(ws.getCell("J1").type).toBe(ValueType.Boolean);

    expect(ws.getCell("K1").value).to.deep.equal((testValues as any).Errors.NotApplicable);
    expect(ws.getCell("K1").type).toBe(ValueType.Error);
    expect(ws.getCell("L1").value).toEqual((testValues as any).Errors.Value);
    expect(ws.getCell("L1").type).toBe(ValueType.Error);

    // A2:B3
    expect(ws.getCell("A2").value).toBe(5);
    expect(ws.getCell("A2").type).toBe(ValueType.Number);
    expect(ws.getCell("A2").master).toBe(ws.getCell("A2"));

    if (options.checkMerges) {
      expect(ws.getCell("A3").value).toBe(5);
      expect(ws.getCell("A3").type).toBe(ValueType.Merge);
      expect(ws.getCell("A3").master).toBe(ws.getCell("A2"));

      expect(ws.getCell("B2").value).toBe(5);
      expect(ws.getCell("B2").type).toBe(ValueType.Merge);
      expect(ws.getCell("B2").master).toBe(ws.getCell("A2"));

      expect(ws.getCell("B3").value).toBe(5);
      expect(ws.getCell("B3").type).toBe(ValueType.Merge);
      expect(ws.getCell("B3").master).toBe(ws.getCell("A2"));

      // C2:D3
      expect(ws.getCell("C2").value).toBeNull();
      expect(ws.getCell("C2").type).toBe(ValueType.Null);
      expect(ws.getCell("C2").master).toBe(ws.getCell("C2"));

      expect(ws.getCell("D2").value).toBeNull();
      expect(ws.getCell("D2").type).toBe(ValueType.Merge);
      expect(ws.getCell("D2").master).toBe(ws.getCell("C2"));

      expect(ws.getCell("C3").value).toBeNull();
      expect(ws.getCell("C3").type).toBe(ValueType.Merge);
      expect(ws.getCell("C3").master).toBe(ws.getCell("C2"));

      expect(ws.getCell("D3").value).toBeNull();
      expect(ws.getCell("D3").type).toBe(ValueType.Merge);
      expect(ws.getCell("D3").master).toBe(ws.getCell("C2"));
    }

    if (options.checkStyles) {
      expect(ws.getCell("A4").numFmt).toBe((testValues as any).numFmt1);
      expect(ws.getCell("A4").type).toBe(ValueType.Number);
      expect(ws.getCell("A4").border).toEqual((styles as any).borders.thin);
      expect(ws.getCell("C4").numFmt).toBe((testValues as any).numFmt2);
      expect(ws.getCell("C4").type).toBe(ValueType.Number);
      expect(ws.getCell("C4").border).to.deep.equal((styles as any).borders.doubleRed);
      expect(ws.getCell("E4").border).to.deep.equal((styles as any).borders.thickRainbow);

      // test fonts and formats
      expect(ws.getCell("A5").value).toBe((testValues as any).str);
      expect(ws.getCell("A5").type).toBe(ValueType.String);
      expect(ws.getCell("B5").value).toBe((testValues as any).str);
      expect(ws.getCell("B5").type).toBe(ValueType.String);
      expect(ws.getCell("B5").font).to.deep.equal((styles as any).fonts.broadwayRedOutline20);
      expect(ws.getCell("C5").value).toBe((testValues as any).str);
      expect(ws.getCell("C5").type).toBe(ValueType.String);
      expect(ws.getCell("C5").font).to.deep.equal((styles as any).fonts.comicSansUdB16);

      expect(Math.abs(ws.getCell("D5").value - 1.6)).to.be.below(0.00000001);
      expect(ws.getCell("D5").type).toBe(ValueType.Number);
      expect(ws.getCell("D5").numFmt).toBe((testValues as any).numFmt1);
      expect(ws.getCell("D5").font).to.deep.equal((styles as any).fonts.arialBlackUI14);

      expect(Math.abs(ws.getCell("E5").value - 1.6)).to.be.below(0.00000001);
      expect(ws.getCell("E5").type).toBe(ValueType.Number);
      expect(ws.getCell("E5").numFmt).toBe((testValues as any).numFmt2);
      expect(ws.getCell("E5").font).to.deep.equal((styles as any).fonts.broadwayRedOutline20);

      expect(
        Math.abs(ws.getCell("F5").value.getTime() - (testValues as any).date.getTime())
      ).to.be.below(options.dateAccuracy);
      expect(ws.getCell("F5").type).toBe(ValueType.Date);
      expect(ws.getCell("F5").numFmt).toBe((testValues as any).numFmtDate);
      expect(ws.getCell("F5").font).to.deep.equal((styles as any).fonts.comicSansUdB16);

      expect(ws.getRow(5).height).toBeUndefined();
      expect(ws.getRow(6).height).toBe(42);
      (styles as any).alignments.forEach((alignment: any, index: number) => {
        const rowNumber = 6;
        const colNumber = index + 1;
        const cell = ws.getCell(rowNumber, colNumber);
        expect(cell.value).toBe(alignment.text);
        expect(cell.alignment).toEqual(alignment.alignment);
      });

      if (options.checkBadAlignments) {
        (styles as any).badAlignments.forEach((alignment: any, index: number) => {
          const rowNumber = 7;
          const colNumber = index + 1;
          const cell = ws.getCell(rowNumber, colNumber);
          expect(cell.value).toBe(alignment.text);
          expect(cell.alignment).toBeUndefined();
        });
      }

      const row8 = ws.getRow(8);
      expect(row8.height).toBe(40);
      expect(row8.getCell(1).fill).to.deep.equal((styles as any).fills.blueWhiteHGrad);
      expect(row8.getCell(2).fill).to.deep.equal((styles as any).fills.redDarkVertical);
      expect(row8.getCell(3).fill).to.deep.equal((styles as any).fills.redGreenDarkTrellis);
      expect(row8.getCell(4).fill).toEqual((styles as any).fills.rgbPathGrad);

      if (options.checkFormulas) {
        // Shared Formula
        expect(ws.getCell("A9").value).toBe(1);
        expect(ws.getCell("A9").type).toBe(ValueType.Number);

        expect(ws.getCell("B9").value).to.deep.equal({
          shareType: "shared",
          ref: "B9:E9",
          formula: "A9+1",
          result: 2
        });
        expect(ws.getCell("B9").type).toBe(ValueType.Formula);

        ["C9", "D9", "E9"].forEach((address, index) => {
          expect(ws.getCell(address).value).to.deep.equal({
            sharedFormula: "B9",
            result: index + 3
          });
          expect(ws.getCell(address).type).toBe(ValueType.Formula);
        });

        if (ws.getCell("A10").value) {
          // Fill Formula Shared
          expect(ws.getCell("A10").value).to.deep.equal({
            shareType: "shared",
            ref: "A10:E10",
            formula: "A9",
            result: 1
          });
          ["B10", "C10", "D10", "E10"].forEach((address, index) => {
            expect(ws.getCell(address).value).to.deep.equal({
              sharedFormula: "A10",
              result: index + 2
            });
            expect(ws.getCell(address).formula).toBe(`${address[0]}9`);
          });

          // Array Formula
          expect(ws.getCell("A11").value).to.deep.equal({
            shareType: "array",
            ref: "A11:E11",
            formula: "A9",
            result: 1
          });
          ["B11", "C11", "D11", "E11"].forEach(address => {
            expect(ws.getCell(address).value).toBe(1);
          });
        }
      } else {
        ["A9", "B9", "C9", "D9", "E9"].forEach((address, index) => {
          expect(ws.getCell(address).value).toBe(index + 1);
          expect(ws.getCell(address).type).toBe(ValueType.Number);
        });
      }
    }
  }
};
