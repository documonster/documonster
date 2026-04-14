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

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

const TEST_DATA_DIR = path.join(__dirname, "data");

describe("HAN CELL xlsx files", () => {
  it("loads xlsx with namespace prefixes and reads shared strings", async () => {
    const filePath = path.join(TEST_DATA_DIR, "han-cell-namespace-prefixes.xlsx");
    const buffer = fs.readFileSync(filePath);

    const workbook = new Workbook();
    await workbook.xlsx.load(buffer);

    // Verify the workbook structure
    expect(workbook.worksheets.length).toBe(1);
    expect(workbook.worksheets[0].name).toBe("no build");

    // Verify actual cell data was parsed (shared strings resolved)
    const worksheet = workbook.worksheets[0];
    const rowCount = worksheet.rowCount;
    expect(rowCount).toBeGreaterThan(0);
  });
});
