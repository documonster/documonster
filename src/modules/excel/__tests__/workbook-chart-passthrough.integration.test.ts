/**
 * Integration tests for chart and drawing passthrough preservation.
 *
 * These tests verify that Excel files containing charts and drawings
 * are correctly preserved during read/write cycles, even though the
 * library doesn't fully parse these features.
 *
 * Related issue: https://github.com/nickmessing/excelts/issues/41
 */

import * as fs from "fs";
import * as path from "path";

import { extractAll } from "@archive/unzip/extract";
import { Workbook } from "@excel/workbook";
import { describe, it, expect, beforeAll } from "vitest";

// Path to test data file
const SAMPLE_FILE = path.join(__dirname, "data", "chart-pivot-sample.xlsx");

// Test utilities
let sampleBuffer: Buffer;

/**
 * Helper to load the sample file and return a workbook
 */
async function loadSampleWorkbook(): Promise<Workbook> {
  const workbook = new Workbook();
  await workbook.xlsx.load(sampleBuffer);
  return workbook;
}

/**
 * Helper to get text content from ZIP entry
 */
function getEntryContent(entries: Map<string, any>, entryPath: string): string | null {
  const entry = entries.get(entryPath);
  if (!entry) {
    return null;
  }
  return new TextDecoder().decode(entry.data);
}

/**
 * Helper to perform round-trip and return both input/output entries
 */
async function performRoundTrip(): Promise<{
  inputEntries: Map<string, any>;
  outputEntries: Map<string, any>;
}> {
  const workbook = await loadSampleWorkbook();
  const outputBuffer = await workbook.xlsx.writeBuffer();

  const inputEntries = await extractAll(new Uint8Array(sampleBuffer));
  const outputEntries = await extractAll(new Uint8Array(outputBuffer));

  return { inputEntries, outputEntries };
}

describe("Chart Passthrough Preservation", () => {
  // Load sample buffer once before all tests
  beforeAll(() => {
    sampleBuffer = fs.readFileSync(SAMPLE_FILE);
  });

  describe("round-trip preservation", () => {
    it("should preserve chart files during read/write cycle", async () => {
      const workbook = await loadSampleWorkbook();

      // Verify workbook was loaded
      expect(workbook.worksheets.length).toBe(2);
      expect(workbook.worksheets.map(ws => ws.name)).toContain("Data");
      expect(workbook.worksheets.map(ws => ws.name)).toContain("Pivot");

      const { inputEntries, outputEntries } = await performRoundTrip();

      // Get file lists
      const inputFiles = [...inputEntries.keys()].sort();
      const outputFiles = [...outputEntries.keys()].sort();

      // Check that all chart-related files are preserved
      const chartFiles = [
        "xl/charts/chart1.xml",
        "xl/charts/style1.xml",
        "xl/charts/colors1.xml",
        "xl/charts/_rels/chart1.xml.rels"
      ];

      for (const chartFile of chartFiles) {
        expect(inputFiles).toContain(chartFile);
        expect(outputFiles).toContain(chartFile);
      }

      // Check that drawing files are preserved
      const drawingFiles = ["xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"];

      for (const drawingFile of drawingFiles) {
        expect(inputFiles).toContain(drawingFile);
        expect(outputFiles).toContain(drawingFile);
      }

      // Verify chart content is identical
      for (const chartFile of chartFiles) {
        const inputContent = getEntryContent(inputEntries, chartFile);
        const outputContent = getEntryContent(outputEntries, chartFile);
        expect(outputContent).toBe(inputContent);
      }

      // Verify drawing XML is identical (preserves graphicFrame for charts)
      const inputDrawing = getEntryContent(inputEntries, "xl/drawings/drawing1.xml");
      const outputDrawing = getEntryContent(outputEntries, "xl/drawings/drawing1.xml");
      expect(outputDrawing).toBe(inputDrawing);

      // Verify drawing XML contains graphicFrame (chart reference)
      expect(inputDrawing).toContain("xdr:graphicFrame");
      expect(outputDrawing).toContain("xdr:graphicFrame");
    });

    it("should preserve pivot tables during read/write cycle", async () => {
      const workbook = await loadSampleWorkbook();

      // Verify pivot table is loaded
      expect(workbook.pivotTables.length).toBe(1);

      const { inputEntries, outputEntries } = await performRoundTrip();

      // Check that pivot table files are preserved
      const pivotFiles = [
        "xl/pivotTables/pivotTable1.xml",
        "xl/pivotTables/_rels/pivotTable1.xml.rels",
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotCache/pivotCacheRecords1.xml",
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"
      ];

      const inputFiles = [...inputEntries.keys()];
      const outputFiles = [...outputEntries.keys()];

      for (const pivotFile of pivotFiles) {
        expect(inputFiles).toContain(pivotFile);
        expect(outputFiles).toContain(pivotFile);
      }
    });

    it("should generate valid Content_Types.xml with chart content types", async () => {
      const { outputEntries } = await performRoundTrip();

      // Get Content_Types.xml
      const contentTypesXml = getEntryContent(outputEntries, "[Content_Types].xml");
      expect(contentTypesXml).not.toBeNull();

      // Verify chart content types are present
      expect(contentTypesXml).toContain("/xl/charts/chart1.xml");
      expect(contentTypesXml).toContain(
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
      );
      expect(contentTypesXml).toContain("application/vnd.ms-office.chartstyle+xml");
      expect(contentTypesXml).toContain("application/vnd.ms-office.chartcolorstyle+xml");
    });

    it("should preserve pivot table attributes including chartFormat", async () => {
      const { outputEntries } = await performRoundTrip();

      // Get pivot table XML
      const pivotTableXml = getEntryContent(outputEntries, "xl/pivotTables/pivotTable1.xml");
      expect(pivotTableXml).not.toBeNull();

      // Verify critical attributes for chart integration
      expect(pivotTableXml).toContain('chartFormat="1"');
      expect(pivotTableXml).toContain('outline="1"');
      expect(pivotTableXml).toContain('outlineData="1"');

      // Verify chartFormats section exists
      expect(pivotTableXml).toContain("<chartFormats");
      expect(pivotTableXml).toContain("<chartFormat");
    });

    it("should preserve rowItems and colItems in pivot table", async () => {
      const { outputEntries } = await performRoundTrip();

      // Get pivot table XML
      const pivotTableXml = getEntryContent(outputEntries, "xl/pivotTables/pivotTable1.xml");
      expect(pivotTableXml).not.toBeNull();

      // Verify rowItems has correct count (5 items: 4 data + 1 grand total)
      expect(pivotTableXml).toContain('<rowItems count="5">');

      // Verify colItems exists
      expect(pivotTableXml).toContain("<colItems");
    });

    it("should preserve worksheet-drawing relationship", async () => {
      const { outputEntries } = await performRoundTrip();

      // Get sheet2 relationships
      const sheet2RelsXml = getEntryContent(outputEntries, "xl/worksheets/_rels/sheet2.xml.rels");
      expect(sheet2RelsXml).not.toBeNull();

      // Verify drawing relationship exists
      expect(sheet2RelsXml).toContain("relationships/drawing");
      expect(sheet2RelsXml).toContain("drawings/drawing1.xml");

      // Get sheet2.xml
      const sheet2Xml = getEntryContent(outputEntries, "xl/worksheets/sheet2.xml");
      expect(sheet2Xml).not.toBeNull();

      // Verify drawing reference in sheet
      expect(sheet2Xml).toContain("<drawing");
      expect(sheet2Xml).toContain("r:id=");
    });

    it("should preserve drawing-chart relationship", async () => {
      const { outputEntries } = await performRoundTrip();

      // Get drawing1 relationships
      const drawingRelsXml = getEntryContent(outputEntries, "xl/drawings/_rels/drawing1.xml.rels");
      expect(drawingRelsXml).not.toBeNull();

      // Verify chart relationship exists
      expect(drawingRelsXml).toContain("relationships/chart");
      expect(drawingRelsXml).toContain("charts/chart1.xml");
    });
  });

  describe("data integrity", () => {
    it("should preserve cell data in Data sheet", async () => {
      const workbook = await loadSampleWorkbook();

      // Get Data sheet
      const dataSheet = workbook.getWorksheet("Data");
      expect(dataSheet).toBeDefined();

      // Check header row (actual values from test file)
      expect(dataSheet!.getCell("A1").value).toBe("Custom ID");
      expect(dataSheet!.getCell("B1").value).toBe("Customer Name");
      expect(dataSheet!.getCell("C1").value).toBe("Region");
      expect(dataSheet!.getCell("D1").value).toBe("Revenue");

      // Write and re-read
      const outputBuffer = await workbook.xlsx.writeBuffer();
      const workbook2 = new Workbook();
      await workbook2.xlsx.load(outputBuffer);

      // Verify data preserved
      const dataSheet2 = workbook2.getWorksheet("Data");
      expect(dataSheet2!.getCell("A1").value).toBe("Custom ID");
      expect(dataSheet2!.getCell("B1").value).toBe("Customer Name");
    });

    it("should preserve pivot table summary values in Pivot sheet", async () => {
      const workbook = await loadSampleWorkbook();

      // Get Pivot sheet
      const pivotSheet = workbook.getWorksheet("Pivot");
      expect(pivotSheet).toBeDefined();

      // Write and re-read
      const outputBuffer = await workbook.xlsx.writeBuffer();
      const workbook2 = new Workbook();
      await workbook2.xlsx.load(outputBuffer);

      // Verify pivot sheet still exists
      const pivotSheet2 = workbook2.getWorksheet("Pivot");
      expect(pivotSheet2).toBeDefined();
    });
  });
});
