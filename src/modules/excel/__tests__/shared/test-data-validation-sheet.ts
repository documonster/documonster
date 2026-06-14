import { addSheetTo } from "@excel/__tests__/shared/add-sheet-to";
import dataValidationsJson from "@excel/__tests__/shared/data/data-validations.json" with { type: "json" };
import { fix, concatenateFormula } from "@excel/__tests__/shared/tools";
import { cellDataValidation, cellSetDataValidation, cellSetValue, cellSetName } from "@excel/cell";
import { Workbook } from "@excel/index";
import { getCell } from "@excel/worksheet";

const self = {
  dataValidations: fix(dataValidationsJson),
  createDataValidations(type: string, operator: string) {
    const dataValidation: any = {
      type,
      operator,
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: [(self.dataValidations as any).values[type].v1]
    };
    switch (operator) {
      case "between":
      case "notBetween":
        dataValidation.formulae.push((self.dataValidations as any).values[type].v2);
        break;
      default:
        break;
    }
    return dataValidation;
  },

  addSheet(wb: any) {
    const ws = addSheetTo(wb, "data-validations");

    // named list
    cellSetValue(getCell(ws, "D1"), "Hewie");
    cellSetName(getCell(ws, "D1"), "Nephews");
    cellSetValue(getCell(ws, "E1"), "Dewie");
    cellSetName(getCell(ws, "E1"), "Nephews");
    cellSetValue(getCell(ws, "F1"), "Louie");
    cellSetName(getCell(ws, "F1"), "Nephews");
    cellSetValue(getCell(ws, "A1"), concatenateFormula("Named List"));
    cellSetDataValidation(getCell(ws, "B1"), (self.dataValidations as any).B1);

    cellSetValue(getCell(ws, "A3"), concatenateFormula("Literal List"));
    cellSetDataValidation(getCell(ws, "B3"), (self.dataValidations as any).B3);

    cellSetValue(getCell(ws, "D5"), "Tom");
    cellSetValue(getCell(ws, "E5"), "Dick");
    cellSetValue(getCell(ws, "F5"), "Harry");
    cellSetValue(getCell(ws, "A5"), concatenateFormula("Range List"));
    cellSetDataValidation(getCell(ws, "B5"), (self.dataValidations as any).B5);

    (self.dataValidations as any).operators.forEach((operator: string, cIndex: number) => {
      const col = 3 + cIndex;
      cellSetValue(getCell(ws, 7, col), concatenateFormula(operator));
    });
    (self.dataValidations as any).types.forEach((type: string, rIndex: number) => {
      const row = 8 + rIndex;
      cellSetValue(getCell(ws, row, 1), concatenateFormula(type));
      (self.dataValidations as any).operators.forEach((operator: string, cIndex: number) => {
        const col = 3 + cIndex;
        cellSetDataValidation(getCell(ws, row, col), self.createDataValidations(type, operator));
      });
    });

    cellSetValue(getCell(ws, "A13"), concatenateFormula("Prompt"));
    cellSetDataValidation(getCell(ws, "B13"), (self.dataValidations as any).B13);

    cellSetValue(getCell(ws, "D13"), concatenateFormula("Error"));
    cellSetDataValidation(getCell(ws, "E13"), (self.dataValidations as any).E13);

    cellSetValue(getCell(ws, "A15"), concatenateFormula("Terse"));
    cellSetDataValidation(getCell(ws, "B15"), (self.dataValidations as any).B15);

    cellSetValue(getCell(ws, "A17"), concatenateFormula("Decimal"));
    cellSetDataValidation(getCell(ws, "B17"), (self.dataValidations as any).B17);

    cellSetValue(getCell(ws, "A19"), concatenateFormula("Any"));
    cellSetDataValidation(getCell(ws, "B19"), (self.dataValidations as any).B19);

    cellSetValue(getCell(ws, "A20"), new Date());
    cellSetDataValidation(getCell(ws, "A20"), {
      type: "date",
      operator: "greaterThan",
      showErrorMessage: true,
      allowBlank: true,
      formulae: [new Date(2016, 0, 1)]
    });

    // two rows of the same validation to test dataValidation optimisation
    ["A22", "A23"].forEach(address => {
      cellSetValue(getCell(ws, address), concatenateFormula("Five Numbers"));
    });
    ["B22", "C22", "D22", "E22", "F22", "B23", "C23", "D23", "E23", "F23"].forEach(address => {
      cellSetDataValidation(
        getCell(ws, address),
        JSON.parse(JSON.stringify((self.dataValidations as any).shared))
      );
    });
  },

  checkSheet(wb: any) {
    const ws = Workbook.getWorksheet(wb, "data-validations")!;
    expect(ws).toBeDefined();

    expect(cellDataValidation(getCell(ws, "B1"))).to.deep.equal((self.dataValidations as any).B1);
    expect(cellDataValidation(getCell(ws, "B3"))).to.deep.equal((self.dataValidations as any).B3);
    expect(cellDataValidation(getCell(ws, "B5"))).to.deep.equal((self.dataValidations as any).B5);

    (self.dataValidations as any).types.forEach((type: string, rIndex: number) => {
      const row = 8 + rIndex;
      cellSetValue(getCell(ws, row, 1), concatenateFormula(type));
      (self.dataValidations as any).operators.forEach((operator: string, cIndex: number) => {
        const col = 3 + cIndex;
        expect(cellDataValidation(getCell(ws, row, col))).to.deep.equal(
          self.createDataValidations(type, operator)
        );
      });
    });

    expect(cellDataValidation(getCell(ws, "B13"))).to.deep.equal((self.dataValidations as any).B13);
    expect(cellDataValidation(getCell(ws, "E13"))).to.deep.equal((self.dataValidations as any).E13);
    expect(cellDataValidation(getCell(ws, "B15"))).to.deep.equal((self.dataValidations as any).B15);
    expect(cellDataValidation(getCell(ws, "B17"))).to.deep.equal((self.dataValidations as any).B17);
    expect(cellDataValidation(getCell(ws, "B19"))).to.deep.equal((self.dataValidations as any).B19);

    // two rows of the same validation to test dataValidation optimisation
    ["B22", "C22", "D22", "E22", "F22", "B23", "C23", "D23", "E23", "F23"].forEach(address => {
      expect(cellDataValidation(getCell(ws, address))).to.deep.equal(
        (self.dataValidations as any).shared
      );
    });
  }
};

const dataValidations = self;
export { dataValidations };
