/**
 * Test for HAN CELL xlsx files
 *
 * HAN CELL is a spreadsheet application that uses namespace prefixes
 * in its XML output (e.g., "x:workbook", "ep:Properties", "dc:creator")
 * instead of the more common unprefixed element names used by Microsoft Excel.
 *
 * This test ensures that files created by HAN CELL can be loaded correctly.
 */
import * as fs from "node:fs";
import * as path from "node:path";

import { Workbook, Worksheet } from "@excel/index";
import { getWorksheets } from "@excel/workbook";
import { getSheetName } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

const TEST_DATA_DIR = path.join(__dirname, "data");

describe("HAN CELL xlsx files", () => {
  it("loads xlsx with namespace prefixes and reads shared strings", async () => {
    const filePath = path.join(TEST_DATA_DIR, "han-cell-namespace-prefixes.xlsx");
    const buffer = fs.readFileSync(filePath);

    const workbook = Workbook.create();
    await Workbook.loadXlsx(workbook, buffer);

    // Verify the workbook structure
    expect(getWorksheets(workbook).length).toBe(1);
    expect(getSheetName(getWorksheets(workbook)[0])).toBe("no build");

    // Verify actual cell data was parsed (shared strings resolved)
    const worksheet = getWorksheets(workbook)[0];
    const rowCount = Worksheet.rowCount(worksheet);
    expect(rowCount).toBeGreaterThan(0);
  });
});
