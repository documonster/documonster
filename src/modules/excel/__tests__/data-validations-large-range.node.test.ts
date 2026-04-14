import fs from "fs";
import path from "path";

import type { DataValidationWithFormulae } from "@excel/types";
import { testOutDir } from "@test/utils";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

import { Workbook } from "../../../index";

describe("DataValidation Large Range Performance", () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(testOutDir(), "excelts-test-"));
  });

  afterEach(() => {
    if (tempDir) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  describe("parsing large range validations", () => {
    it("should parse entire column validation (B2:B1048576) efficiently", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["range:B2:B1048576"] = {
        type: "list",
        formulae: ["Option1,Option2,Option3"],
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true
      };

      const filePath = path.join(tempDir, "large-range.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      const start = performance.now();
      await wb2.xlsx.readFile(filePath);
      const elapsed = performance.now() - start;

      expect(elapsed).toBeLessThan(500);

      const ws2 = wb2.getWorksheet("Test");
      expect(ws2?.getCell("B2").dataValidation).toBeDefined();
      expect(ws2?.getCell("B100").dataValidation).toBeDefined();
      expect(ws2?.getCell("B1000000").dataValidation).toBeDefined();
      expect(ws2?.getCell("B1").dataValidation).toBeUndefined();
      expect(ws2?.getCell("A2").dataValidation).toBeUndefined();
    });

    it("should parse multiple entire column validations efficiently", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

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
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      const start = performance.now();
      await wb2.xlsx.readFile(filePath);
      const elapsed = performance.now() - start;

      expect(elapsed).toBeLessThan(1000);

      const ws2 = wb2.getWorksheet("Test");
      expect(ws2?.getCell("B100").dataValidation?.type).toBe("list");
      expect(ws2?.getCell("C100").dataValidation?.type).toBe("list");
      expect(ws2?.getCell("D100").dataValidation?.type).toBe("whole");
      expect(ws2?.getCell("E100").dataValidation?.type).toBe("decimal");
      expect(ws2?.getCell("F100").dataValidation?.type).toBe("date");
    });
  });

  describe("small range validations (range-based lookup)", () => {
    it("should allow cell lookup for small ranges", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      for (let row = 1; row <= 10; row++) {
        for (let col = 1; col <= 10; col++) {
          ws.getCell(row, col).dataValidation = {
            type: "list",
            formulae: ["Yes,No"],
            allowBlank: true
          };
        }
      }

      const filePath = path.join(tempDir, "small-range.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation).toBeDefined();
      expect(ws2?.getCell("J10").dataValidation).toBeDefined();
      expect(ws2?.getCell("K1").dataValidation).toBeUndefined();
    });

    it("should handle moderately sized ranges", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["A1:Y40"] = {
        type: "list",
        formulae: ["Test"],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "threshold.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation).toBeDefined();
      expect(ws2?.getCell("Y40").dataValidation).toBeDefined();
    });

    it("should NOT expand range just over threshold (1001+ cells)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["range:A1:Y41"] = {
        type: "list",
        formulae: ["Test"],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "over-threshold.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation).toBeDefined();
      expect(ws2?.getCell("Y41").dataValidation).toBeDefined();
    });
  });

  describe("edge cases", () => {
    it("should handle single cell validation", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.getCell("A1").dataValidation = {
        type: "whole",
        operator: "between",
        formulae: [1, 10]
      };

      const filePath = path.join(tempDir, "single-cell.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation?.type).toBe("whole");
      expect(ws2?.getCell("A2").dataValidation).toBeUndefined();
    });

    it("should handle multiple disjoint ranges in same sqref (E4:L9 N4:U9)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      const validation = {
        type: "list" as const,
        formulae: ["A,B,C"],
        allowBlank: true
      };

      for (let row = 4; row <= 9; row++) {
        for (let col = 5; col <= 12; col++) {
          ws.getCell(row, col).dataValidation = validation;
        }
      }
      for (let row = 4; row <= 9; row++) {
        for (let col = 14; col <= 21; col++) {
          ws.getCell(row, col).dataValidation = validation;
        }
      }

      const filePath = path.join(tempDir, "disjoint.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("E4").dataValidation).toBeDefined();
      expect(ws2?.getCell("L9").dataValidation).toBeDefined();
      expect(ws2?.getCell("N4").dataValidation).toBeDefined();
      expect(ws2?.getCell("U9").dataValidation).toBeDefined();
      expect(ws2?.getCell("M4").dataValidation).toBeUndefined();
    });

    it("should handle validation without type (any type)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.getCell("A1").dataValidation = {
        type: "any",
        promptTitle: "Input",
        prompt: "Enter any value"
      };

      const filePath = path.join(tempDir, "any-type.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation?.type).toBe("any");
      expect(ws2?.getCell("A1").dataValidation?.prompt).toBe("Enter any value");
    });

    it("should handle large range with formula reference to another sheet", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Data");
      const ws2 = wb.addWorksheet("Input");

      ws1.getCell("A1").value = "Option1";
      ws1.getCell("A2").value = "Option2";
      ws1.getCell("A3").value = "Option3";

      ws2.dataValidations.model["range:A1:A1048576"] = {
        type: "list",
        formulae: ["Data!$A$1:$A$3"],
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true
      };

      const filePath = path.join(tempDir, "cross-sheet-ref.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb3 = new Workbook();
      await wb3.xlsx.readFile(filePath);
      const ws2Loaded = wb3.getWorksheet("Input");

      const dv = ws2Loaded?.getCell("A100").dataValidation as
        | DataValidationWithFormulae
        | undefined;
      expect(dv?.type).toBe("list");
      expect(dv?.formulae?.[0]).toBe("Data!$A$1:$A$3");
    });

    it("should handle overlapping ranges correctly (last one wins)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["range:A1:Z100"] = {
        type: "list",
        formulae: ["Large"],
        allowBlank: true
      };

      ws.getCell("B2").dataValidation = {
        type: "list",
        formulae: ["Specific"],
        allowBlank: true
      };

      const filePath = path.join(tempDir, "overlap.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      const dvB2 = ws2?.getCell("B2").dataValidation as DataValidationWithFormulae | undefined;
      const dvA1 = ws2?.getCell("A1").dataValidation as DataValidationWithFormulae | undefined;
      const dvZ100 = ws2?.getCell("Z100").dataValidation as DataValidationWithFormulae | undefined;
      expect(dvB2?.formulae?.[0]).toBe("Specific");
      expect(dvA1?.formulae?.[0]).toBe("Large");
      expect(dvZ100?.formulae?.[0]).toBe("Large");
    });

    it("should preserve all validation properties through round-trip", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

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
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      const dv = ws2?.getCell("A500").dataValidation as DataValidationWithFormulae | undefined;
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.getCell("A1").dataValidation = {
        type: "list",
        formulae: ["Test"]
      };
      ws.getCell("A2").dataValidation = {
        type: "list",
        formulae: ["Test"]
      };

      ws.dataValidations.remove("A1");

      expect(ws.getCell("A1").dataValidation).toBeUndefined();
      expect(ws.getCell("A2").dataValidation).toBeDefined();

      const filePath = path.join(tempDir, "remove.xlsx");
      await wb.xlsx.writeFile(filePath);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filePath);
      const ws2 = wb2.getWorksheet("Test");

      expect(ws2?.getCell("A1").dataValidation).toBeUndefined();
      expect(ws2?.getCell("A2").dataValidation).toBeDefined();
    });
  });

  describe("DataValidations.find edge cases", () => {
    it("should return undefined for empty model", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      expect(ws.dataValidations.find("A1")).toBeUndefined();
      expect(ws.dataValidations.find("Z999")).toBeUndefined();
    });

    it("should handle find with only range: keys", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["range:B2:D10"] = {
        type: "list",
        formulae: ["Test"]
      };

      expect(ws.dataValidations.find("B2")).toBeDefined();
      expect(ws.dataValidations.find("C5")).toBeDefined();
      expect(ws.dataValidations.find("D10")).toBeDefined();

      expect(ws.dataValidations.find("A1")).toBeUndefined();
      expect(ws.dataValidations.find("B1")).toBeUndefined();
      expect(ws.dataValidations.find("E2")).toBeUndefined();
      expect(ws.dataValidations.find("B11")).toBeUndefined();
    });

    it("should prioritize direct match over range match", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Test");

      ws.dataValidations.model["range:A1:Z100"] = {
        type: "list",
        formulae: ["Range"]
      };
      ws.dataValidations.model["B5"] = {
        type: "list",
        formulae: ["Direct"]
      };

      const b5 = ws.dataValidations.find("B5");
      const b6 = ws.dataValidations.find("B6");
      expect(b5?.type !== "any" ? b5?.formulae?.[0] : undefined).toBe("Direct");
      expect(b6?.type !== "any" ? b6?.formulae?.[0] : undefined).toBe("Range");
    });
  });
});
