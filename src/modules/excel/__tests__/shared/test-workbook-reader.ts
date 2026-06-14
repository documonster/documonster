import testValuesJson from "@excel/__tests__/shared/data/sheet-values.json" with { type: "json" };
import { fix } from "@excel/__tests__/shared/tools";
import { cellGetValue, cellType } from "@excel/cell";
import { rowGetCell } from "@excel/worksheet";
const testValues = fix(testValuesJson);
import { ValueType } from "@excel/enums";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { dateToExcel } from "@utils/utils";

function fillFormula(f: any) {
  return Object.assign({ formula: undefined }, f);
}

const streamedValues = {
  B1: { sharedString: 0 },
  C1: dateToExcel((testValues as any).date),
  D1: fillFormula((testValues as any).formulas[0]),
  E1: fillFormula((testValues as any).formulas[1]),
  F1: { sharedString: 1 },
  G1: { sharedString: 2 }
};
import pageSetupJson from "@excel/__tests__/shared/data/page-setup.json" with { type: "json" };
import propertiesJson from "@excel/__tests__/shared/data/sheet-properties.json" with { type: "json" };
import stylesJson from "@excel/__tests__/shared/data/styles.json" with { type: "json" };

const testWorkbookReader = {
  testValues: fix(testValuesJson),
  styles: fix(stylesJson),
  properties: fix(propertiesJson),
  pageSetup: fix(pageSetupJson),

  checkBook(filename: string): Promise<void> {
    const wb = new WorkbookReader(filename, {
      entries: "emit",
      worksheets: "emit",
      sharedStrings: "ignore"
    });

    // expectations
    const dateAccuracy = 0.00001;

    return new Promise<void>((resolve, reject) => {
      let rowCount = 0;

      wb.on("worksheet", ws => {
        // Sheet name stored in workbook. Not guaranteed here
        // expect(ws.name).toBe('blort');
        ws.on("row", row => {
          rowCount++;
          try {
            switch (row.number) {
              case 1:
                expect(cellGetValue(rowGetCell(row, "A"))).toBe(7);
                expect(cellType(rowGetCell(row, "A"))).to.equal(ValueType.Number);
                expect(cellGetValue(rowGetCell(row, "B"))).toEqual(streamedValues.B1);
                expect(cellType(rowGetCell(row, "B"))).to.equal(ValueType.String);
                expect(
                  Math.abs((cellGetValue(rowGetCell(row, "C")) as number) - streamedValues.C1)
                ).to.be.below(dateAccuracy);
                expect(cellType(rowGetCell(row, "C"))).to.equal(ValueType.Number);

                expect(cellGetValue(rowGetCell(row, "D"))).toEqual(streamedValues.D1);
                expect(cellType(rowGetCell(row, "D"))).to.equal(ValueType.Formula);
                expect(cellGetValue(rowGetCell(row, "E"))).toEqual(streamedValues.E1);
                expect(cellType(rowGetCell(row, "E"))).to.equal(ValueType.Formula);
                expect(cellGetValue(rowGetCell(row, "F"))).toEqual(streamedValues.F1);
                expect(cellType(rowGetCell(row, "F"))).to.equal(ValueType.SharedString);
                expect(cellGetValue(rowGetCell(row, "G"))).toEqual(streamedValues.G1);
                break;

              case 2:
                // A2:B3
                expect(cellGetValue(rowGetCell(row, "A"))).toBe(5);
                expect(cellType(rowGetCell(row, "A"))).to.equal(ValueType.Number);

                expect(cellType(rowGetCell(row, "B"))).toBe(ValueType.Null);

                // C2:D3
                expect(cellGetValue(rowGetCell(row, "C"))).toBeNull();
                expect(cellType(rowGetCell(row, "C"))).toBe(ValueType.Null);

                expect(cellGetValue(rowGetCell(row, "D"))).toBeNull();
                expect(cellType(rowGetCell(row, "D"))).toBe(ValueType.Null);

                break;

              case 3:
                expect(cellGetValue(rowGetCell(row, "A"))).toBe(null);
                expect(cellType(rowGetCell(row, "A"))).toBe(ValueType.Null);

                expect(cellGetValue(rowGetCell(row, "B"))).toBe(null);
                expect(cellType(rowGetCell(row, "B"))).toBe(ValueType.Null);

                expect(cellGetValue(rowGetCell(row, "C"))).toBeNull();
                expect(cellType(rowGetCell(row, "C"))).toBe(ValueType.Null);

                expect(cellGetValue(rowGetCell(row, "D"))).toBeNull();
                expect(cellType(rowGetCell(row, "D"))).toBe(ValueType.Null);
                break;

              case 4:
                expect(cellType(rowGetCell(row, "A"))).to.equal(ValueType.Number);
                expect(cellType(rowGetCell(row, "C"))).to.equal(ValueType.Number);
                break;

              case 5:
                // test fonts and formats
                expect(cellGetValue(rowGetCell(row, "A"))).toEqual(streamedValues.B1);
                expect(cellType(rowGetCell(row, "A"))).to.equal(ValueType.String);
                expect(cellGetValue(rowGetCell(row, "B"))).toEqual(streamedValues.B1);
                expect(cellType(rowGetCell(row, "B"))).to.equal(ValueType.String);
                expect(cellGetValue(rowGetCell(row, "C"))).toEqual(streamedValues.B1);
                expect(cellType(rowGetCell(row, "C"))).to.equal(ValueType.String);

                expect(Math.abs((cellGetValue(rowGetCell(row, "D")) as number) - 1.6)).to.be.below(
                  0.00000001
                );
                expect(cellType(rowGetCell(row, "D"))).to.equal(ValueType.Number);

                expect(Math.abs((cellGetValue(rowGetCell(row, "E")) as number) - 1.6)).to.be.below(
                  0.00000001
                );
                expect(cellType(rowGetCell(row, "E"))).to.equal(ValueType.Number);

                expect(
                  Math.abs((cellGetValue(rowGetCell(row, "F")) as number) - streamedValues.C1)
                ).to.be.below(dateAccuracy);
                expect(cellType(rowGetCell(row, "F"))).to.equal(ValueType.Number);
                break;

              case 6:
                expect(row.height).toBe(42);
                break;

              case 7:
                break;

              case 8:
                expect(row.height).toBe(40);
                break;

              default:
                break;
            }
          } catch (error) {
            reject(error);
          }
        });
      });
      wb.on("end", () => {
        try {
          expect(rowCount).toBe(11);
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      wb.read();
    });
  }
};

export { testWorkbookReader };
