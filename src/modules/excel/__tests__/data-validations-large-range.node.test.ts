import fs from "fs";
import path from "path";

import { cellDataValidation, cellSetDataValidation } from "@excel/cell";
import { dataValidationFind, dataValidationRemove } from "@excel/data-validations";
import { Cell, Workbook } from "@excel/index";
import type { DataValidationWithFormulae } from "@excel/types";
import { getCell } from "@excel/worksheet";
import { testOutDir } from "@test/utils";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

describe("DataValidation Large Range Performance", () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(testOutDir(), "documonster-test-"));
  });

  afterEach(() => {
    if (tempDir) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  describe("parsing large range validations", () => {
    it("should parse entire column validation (B2:B1048576) efficiently", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:B2:B1048576"] = {
        type: "list",
        formulae: ["Option1,Option2,Option3"],
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true
      };

      const filePath = path.join(tempDir, "large-range.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      const start = performance.now();
      await Workbook.readFile(wb2, filePath);
      const elapsed = performance.now() - start;

      expect(elapsed).toBeLessThan(500);

      const ws2 = Workbook.getWorksheet(wb2, "Test")!;
      expect(cellDataValidation(getCell(ws2, "B2"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "B100"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "B1000000"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "B1"))).toBeUndefined();
      expect(cellDataValidation(getCell(ws2, "A2"))).toBeUndefined();
    });

    it("should parse multiple entire column validations efficiently", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:B2:B1048576"] = {
        type: "list",
        formulae: ["A,B,C"],
        allowBlank: true
      };
      ws.dataValidations.model["range:C2:C1048576"] = {
        type: "list",
        formulae: ["X,Y,Z"],
        allowBlank: true
      };
      ws.dataValidations.model["range:D2:D1048576"] = {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
        allowBlank: true
      };
      ws.dataValidations.model["range:E2:E1048576"] = {
        type: "decimal",
        operator: "greaterThan",
        formulae: [0],
        allowBlank: true
      };
      ws.dataValidations.model["range:F2:F1048576"] = {
        type: "date",
        operator: "greaterThan",
        formulae: [new Date("2020-01-01")],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "multi-column.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      const start = performance.now();
      await Workbook.readFile(wb2, filePath);
      const elapsed = performance.now() - start;

      expect(elapsed).toBeLessThan(1000);

      const ws2 = Workbook.getWorksheet(wb2, "Test")!;
      expect(cellDataValidation(getCell(ws2, "B100"))?.type).toBe("list");
      expect(cellDataValidation(getCell(ws2, "C100"))?.type).toBe("list");
      expect(cellDataValidation(getCell(ws2, "D100"))?.type).toBe("whole");
      expect(cellDataValidation(getCell(ws2, "E100"))?.type).toBe("decimal");
      expect(cellDataValidation(getCell(ws2, "F100"))?.type).toBe("date");
    });
  });

  describe("small range validations (range-based lookup)", () => {
    it("should allow cell lookup for small ranges", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      for (let row = 1; row <= 10; row++) {
        for (let col = 1; col <= 10; col++) {
          cellSetDataValidation(getCell(ws, row, col), {
            type: "list",
            formulae: ["Yes,No"],
            allowBlank: true
          });
        }
      }

      const filePath = path.join(tempDir, "small-range.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "J10"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "K1"))).toBeUndefined();
    });

    it("should handle moderately sized ranges", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["A1:Y40"] = {
        type: "list",
        formulae: ["Test"],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "threshold.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "Y40"))).toBeDefined();
    });

    it("should NOT expand range just over threshold (1001+ cells)", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:A1:Y41"] = {
        type: "list",
        formulae: ["Test"],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "over-threshold.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "Y41"))).toBeDefined();
    });
  });

  describe("edge cases", () => {
    it("should handle single cell validation", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      cellSetDataValidation(getCell(ws, "A1"), {
        type: "whole",
        operator: "between",
        formulae: [1, 10]
      });

      const filePath = path.join(tempDir, "single-cell.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))?.type).toBe("whole");
      expect(cellDataValidation(getCell(ws2, "A2"))).toBeUndefined();
    });

    it("should handle multiple disjoint ranges in same sqref (E4:L9 N4:U9)", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      const validation = {
        type: "list" as const,
        formulae: ["A,B,C"],
        allowBlank: true
      };

      for (let row = 4; row <= 9; row++) {
        for (let col = 5; col <= 12; col++) {
          cellSetDataValidation(getCell(ws, row, col), validation);
        }
      }
      for (let row = 4; row <= 9; row++) {
        for (let col = 14; col <= 21; col++) {
          cellSetDataValidation(getCell(ws, row, col), validation);
        }
      }

      const filePath = path.join(tempDir, "disjoint.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "E4"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "L9"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "N4"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "U9"))).toBeDefined();
      expect(cellDataValidation(getCell(ws2, "M4"))).toBeUndefined();
    });

    it("should handle validation without type (any type)", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      cellSetDataValidation(getCell(ws, "A1"), {
        type: "any",
        promptTitle: "Input",
        prompt: "Enter any value"
      });

      const filePath = path.join(tempDir, "any-type.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))?.type).toBe("any");
      expect(cellDataValidation(getCell(ws2, "A1"))?.prompt).toBe("Enter any value");
    });

    it("should handle large range with formula reference to another sheet", async () => {
      const wb = Workbook.create();
      const ws1 = Workbook.addWorksheet(wb, "Data");
      const ws2 = Workbook.addWorksheet(wb, "Input");

      Cell.setValue(ws1, "A1", "Option1");
      Cell.setValue(ws1, "A2", "Option2");
      Cell.setValue(ws1, "A3", "Option3");

      ws2.dataValidations.model["range:A1:A1048576"] = {
        type: "list",
        formulae: ["Data!$A$1:$A$3"],
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true
      };

      const filePath = path.join(tempDir, "cross-sheet-ref.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb3 = Workbook.create();
      await Workbook.readFile(wb3, filePath);
      const ws2Loaded = Workbook.getWorksheet(wb3, "Input")!;

      const dv = cellDataValidation(getCell(ws2Loaded, "A100")) as
        | DataValidationWithFormulae
        | undefined;
      expect(dv?.type).toBe("list");
      expect(dv?.formulae?.[0]).toBe("Data!$A$1:$A$3");
    });

    it("should handle overlapping ranges correctly (last one wins)", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:A1:Z100"] = {
        type: "list",
        formulae: ["Large"],
        allowBlank: true
      };

      cellSetDataValidation(getCell(ws, "B2"), {
        type: "list",
        formulae: ["Specific"],
        allowBlank: true
      });

      const filePath = path.join(tempDir, "overlap.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      const dvB2 = cellDataValidation(getCell(ws2, "B2")) as DataValidationWithFormulae | undefined;
      const dvA1 = cellDataValidation(getCell(ws2, "A1")) as DataValidationWithFormulae | undefined;
      const dvZ100 = cellDataValidation(getCell(ws2, "Z100")) as
        | DataValidationWithFormulae
        | undefined;
      expect(dvB2?.formulae?.[0]).toBe("Specific");
      expect(dvA1?.formulae?.[0]).toBe("Large");
      expect(dvZ100?.formulae?.[0]).toBe("Large");
    });

    it("should preserve all validation properties through round-trip", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      const fullValidation = {
        type: "list" as const,
        formulae: ["A,B,C"],
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        promptTitle: "Select Value",
        prompt: "Please select a value from the list",
        errorStyle: "warning" as const,
        errorTitle: "Invalid Input",
        error: "The value you entered is not valid"
      };

      ws.dataValidations.model["range:A1:A1048576"] = fullValidation;

      const filePath = path.join(tempDir, "full-props.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      const dv = cellDataValidation(getCell(ws2, "A500")) as DataValidationWithFormulae | undefined;
      expect(dv?.type).toBe("list");
      expect(dv?.formulae).toEqual(["A,B,C"]);
      expect(dv?.allowBlank).toBe(true);
      expect(dv?.showInputMessage).toBe(true);
      expect(dv?.showErrorMessage).toBe(true);
      expect(dv?.promptTitle).toBe("Select Value");
      expect(dv?.prompt).toBe("Please select a value from the list");
      expect(dv?.errorStyle).toBe("warning");
      expect(dv?.errorTitle).toBe("Invalid Input");
      expect(dv?.error).toBe("The value you entered is not valid");
    });

    it("should handle removing validation from a cell", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      cellSetDataValidation(getCell(ws, "A1"), {
        type: "list",
        formulae: ["Test"]
      });
      cellSetDataValidation(getCell(ws, "A2"), {
        type: "list",
        formulae: ["Test"]
      });

      dataValidationRemove(ws.dataValidations, "A1");

      expect(cellDataValidation(getCell(ws, "A1"))).toBeUndefined();
      expect(cellDataValidation(getCell(ws, "A2"))).toBeDefined();

      const filePath = path.join(tempDir, "remove.xlsx");
      await Workbook.writeFile(wb, filePath);

      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, filePath);
      const ws2 = Workbook.getWorksheet(wb2, "Test")!;

      expect(cellDataValidation(getCell(ws2, "A1"))).toBeUndefined();
      expect(cellDataValidation(getCell(ws2, "A2"))).toBeDefined();
    });
  });

  describe("DataValidations.find edge cases", () => {
    it("should return undefined for empty model", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      expect(dataValidationFind(ws.dataValidations, "A1")).toBeUndefined();
      expect(dataValidationFind(ws.dataValidations, "Z999")).toBeUndefined();
    });

    it("should handle find with only range: keys", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:B2:D10"] = {
        type: "list",
        formulae: ["Test"]
      };

      expect(dataValidationFind(ws.dataValidations, "B2")).toBeDefined();
      expect(dataValidationFind(ws.dataValidations, "C5")).toBeDefined();
      expect(dataValidationFind(ws.dataValidations, "D10")).toBeDefined();

      expect(dataValidationFind(ws.dataValidations, "A1")).toBeUndefined();
      expect(dataValidationFind(ws.dataValidations, "B1")).toBeUndefined();
      expect(dataValidationFind(ws.dataValidations, "E2")).toBeUndefined();
      expect(dataValidationFind(ws.dataValidations, "B11")).toBeUndefined();
    });

    it("should prioritize direct match over range match", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      ws.dataValidations.model["range:A1:Z100"] = {
        type: "list",
        formulae: ["Range"]
      };
      ws.dataValidations.model["B5"] = {
        type: "list",
        formulae: ["Direct"]
      };

      const b5 = dataValidationFind(ws.dataValidations, "B5");
      const b6 = dataValidationFind(ws.dataValidations, "B6");
      expect(b5?.type !== "any" ? b5?.formulae?.[0] : undefined).toBe("Direct");
      expect(b6?.type !== "any" ? b6?.formulae?.[0] : undefined).toBe("Range");
    });
  });
});
