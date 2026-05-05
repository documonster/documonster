import { readFileSync } from "fs";

import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";
import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");

describe("DataValidations", () => {
  it("reads a workbook with dataValidation missing type", async () => {
    const wb = new Workbook();
    await wb.xlsx.readFile(excelTestDataPath("data-validation-missing-type.xlsx"));

    // Should load without error and have at least one worksheet
    expect(wb.worksheets.length).toBeGreaterThan(0);
  });

  it("writes a full-column validation without throwing", async () => {
    const TEST_XLSX_FILE_NAME = testFilePath("data-validation-full-column.test");

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    const range = "A2:A1048576";
    ws.dataValidations.model[range] = {
      allowBlank: true,
      error: "Please use the drop down to select a valid value",
      errorTitle: "Invalid Selection",
      formulae: ['"Apples,Bananas,Oranges"'],
      showErrorMessage: true,
      type: "list"
    };

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
    await expectValidXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));

    // Read back and verify validation survived
    const wb2 = new Workbook();
    await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    const dvKeys = Object.keys(ws2.dataValidations.model);
    expect(dvKeys.length).toBe(1);
    expect(ws2.dataValidations.model[dvKeys[0]]?.type).toBe("list");
  });

  it("reads and writes data validations", async () => {
    const TEST_XLSX_FILE_NAME = testFilePath("pr-1204.data-validations.test");

    const wb = new Workbook();
    await wb.xlsx.readFile(excelTestDataPath("data-validation-text-length.xlsx"));

    const expected = {
      E1: {
        type: "textLength",
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: "greaterThan"
      },
      E4: {
        type: "textLength",
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: "greaterThan"
      }
    };

    const ws = wb.getWorksheet(1);
    expect(ws!.dataValidations.model).toEqual(expected);

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
    await expectValidXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));
  });

  describe("ignoreNodes", () => {
    it("readFile ignores dataValidations without blowing up memory", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("data-validations-large.xlsx"), {
        ignoreNodes: ["dataValidations"]
      });
      // Should load successfully and have worksheets, but no data validations
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });

    it("load(buffer) ignores dataValidations without blowing up memory", async () => {
      const buffer = readFileSync(excelTestDataPath("data-validations-large.xlsx"));
      const wb = new Workbook();
      await wb.xlsx.load(buffer, {
        ignoreNodes: ["dataValidations"]
      });
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });
  });
});
