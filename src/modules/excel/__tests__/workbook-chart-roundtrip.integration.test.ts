/**
 * Integration tests for chart round-trip preservation.
 *
 * These tests verify that Excel files containing charts and drawings
 * are correctly preserved during read/write cycles using native
 * chart parsing and rendering (no raw XML passthrough).
 *
 * Related issue: https://github.com/nickmessing/excelts/issues/41
 */

import * as fs from "fs";
import * as path from "path";

import { extractAll } from "@archive/unzip/extract";
import {
  applyChartPreset,
  EXCEL_CHART_PRESETS,
  type ChartExType,
  type ExcelChartPreset
} from "@excel/chart";
import { Workbook } from "@excel/workbook";
import { describe, it, expect, beforeAll } from "vitest";

import { smokeRoundTripWithLibreOffice } from "./helpers/libreoffice-smoke";
import { auditOoxmlPackage } from "./helpers/ooxml-package-audit";

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

describe("Chart Round-Trip Preservation", () => {
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

      // Verify chart content is preserved
      // Chart XML is now parsed and re-rendered (not byte-for-byte passthrough),
      // so we verify structural correctness instead of exact match.
      const outputChartXml = getEntryContent(outputEntries, "xl/charts/chart1.xml");
      expect(outputChartXml).toContain("c:chartSpace");
      expect(outputChartXml).toContain("c:barChart");
      expect(outputChartXml).toContain('c:barDir val="col"');
      expect(outputChartXml).toContain('c:grouping val="clustered"');
      expect(outputChartXml).toContain("Pivot!$B$1");
      expect(outputChartXml).toContain("Pivot!$A$2:$A$6");
      expect(outputChartXml).toContain("c:catAx");
      expect(outputChartXml).toContain("c:valAx");
      expect(outputChartXml).toContain("c:legend");
      expect(outputChartXml).toContain("c:printSettings");

      // Style and colors are raw byte passthrough — should match exactly
      for (const file of ["xl/charts/style1.xml", "xl/charts/colors1.xml"]) {
        const inputContent = getEntryContent(inputEntries, file);
        const outputContent = getEntryContent(outputEntries, file);
        expect(outputContent).toBe(inputContent);
      }

      // Chart rels should be structurally correct
      const outputChartRels = getEntryContent(outputEntries, "xl/charts/_rels/chart1.xml.rels");
      expect(outputChartRels).toContain("Relationship");
      expect(outputChartRels).toContain("chartStyle");
      expect(outputChartRels).toContain("chartColorStyle");

      // Verify drawing XML contains graphicFrame (chart reference)
      // Drawing XML is regenerated, so verify structure rather than exact match
      const outputDrawing = getEntryContent(outputEntries, "xl/drawings/drawing1.xml");
      expect(outputDrawing).toContain("xdr:graphicFrame");
      expect(outputDrawing).toContain("xdr:wsDr");
      const inputDrawing = getEntryContent(inputEntries, "xl/drawings/drawing1.xml");
      expect(inputDrawing).toContain("xdr:graphicFrame");
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

    it("validates generated chart relationships and content types for classic, ChartEx, and chartsheets", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Data");
      ws.addRows([
        ["A", 10],
        ["B", 20]
      ]);
      ws.addChart(
        {
          type: "bar",
          series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }],
          chartStyle: { id: 10 },
          chartColors: { method: "cycle", colors: [{ srgb: "4472C4" }] }
        },
        "D1:J10"
      );
      ws.addChartEx(
        { type: "waterfall", categories: "Data!$A$1:$A$2", series: [{ values: "Data!$B$1:$B$2" }] },
        "D12:J22"
      );
      wb.addChartsheet("Chart Sheet", {
        chart: {
          type: "funnel",
          categories: "Data!$A$1:$A$2",
          series: [{ values: "Data!$B$1:$B$2" }]
        }
      });

      const output = await wb.xlsx.writeBuffer();
      const entries = await extractAll(new Uint8Array(output));
      const contentTypes = getEntryContent(entries, "[Content_Types].xml")!;
      expect(contentTypes).toContain("/xl/charts/chart1.xml");
      expect(contentTypes).toContain("/xl/charts/chartEx1.xml");
      expect(contentTypes).toContain("/xl/chartsheets/sheet1.xml");

      const drawingRels = getEntryContent(entries, "xl/drawings/_rels/drawing1.xml.rels")!;
      expect(drawingRels).toContain("../charts/chart1.xml");
      expect(drawingRels).toContain("../charts/chartEx1.xml");
      const chartRels = getEntryContent(entries, "xl/charts/_rels/chart1.xml.rels")!;
      expect(chartRels).toContain("chartStyle");
      expect(chartRels).toContain("chartColorStyle");
      const chartsheetRels = getEntryContent(entries, "xl/chartsheets/_rels/sheet1.xml.rels")!;
      expect(chartsheetRels).toContain("../drawings/");
    });

    it("audits ChartEx data, axis, and external-data relationship references", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Data");
      ws.addRows([
        ["A", 10],
        ["B", 20]
      ]);
      ws.addChartEx(
        { type: "waterfall", categories: "Data!$A$1:$A$2", series: [{ values: "Data!$B$1:$B$2" }] },
        "D1:J10"
      );
      const entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
      const chartExEntry = entries.get("xl/charts/chartEx1.xml")!;
      chartExEntry.data = new TextEncoder().encode(
        getEntryContent(entries, "xl/charts/chartEx1.xml")!
          .replace('<cx:dataId val="1"/>', '<cx:dataId val="99"/>')
          .replace('<cx:axisId val="1"/>', '<cx:axisId val="42"/>')
          .replace(
            "<cx:chartData>",
            '<cx:chartData><cx:externalData r:id="rMissing" autoUpdate="1"/>'
          )
      );

      const audit = auditOoxmlPackage(entries);
      expect(audit.errors.join("\n")).toContain("missing cx:data id 99");
      expect(audit.errors.join("\n")).toContain("missing cx:axis id 42");
      expect(audit.errors.join("\n")).toContain("missing relationship rMissing");
    });

    it("audits package structure for the full classic preset and ChartEx matrix", async () => {
      const presets: ExcelChartPreset[] = [...EXCEL_CHART_PRESETS];
      const chartExTypes: ChartExType[] = [
        "sunburst",
        "treemap",
        "waterfall",
        "funnel",
        "histogram",
        "pareto",
        "boxWhisker",
        "regionMap"
      ];
      const wb = new Workbook();
      const ws = wb.addWorksheet("Data");
      ws.addRows([
        ["Cat", "Region", "Open", "High", "Low", "Close", "Value", "Value2", "Size"],
        ["A", "North", 10, 15, 8, 12, 30, 3, 5],
        ["B", "North", 20, 25, 18, 21, 20, 2, 7],
        ["C", "South", 12, 18, 9, 16, 10, 1, 9]
      ]);
      let row = 1;
      for (const preset of presets) {
        const presetType = applyChartPreset(preset, { series: [] }).type;
        const noAxis =
          presetType === "ofPie" ||
          presetType === "pie" ||
          presetType === "pie3D" ||
          presetType === "doughnut";
        const noTrendline = noAxis || presetType === "surface" || presetType === "surface3D";
        const base = {
          name: String(preset),
          categories: "Data!$A$2:$A$4",
          values: "Data!$G$2:$G$4",
          fill: "4472C4",
          line: "ED7D31",
          dataLabels: { showVal: true, position: "outEnd" as const },
          trendline: noTrendline ? undefined : { type: "linear" as const },
          errorBars: noAxis ? undefined : { type: "fixedVal" as const, value: 1 },
          dataPoints: [{ index: 0, fill: "FFC000" }]
        };
        const series = preset.startsWith("scatter")
          ? [
              {
                name: String(preset),
                xValues: "Data!$H$2:$H$4",
                values: "Data!$G$2:$G$4",
                marker: { symbol: "circle" as const, size: 6 }
              }
            ]
          : preset.startsWith("bubble")
            ? [{ ...base, xValues: "Data!$H$2:$H$4", bubbleSize: "Data!$I$2:$I$4" }]
            : preset.startsWith("stock")
              ? [
                  { name: "Open", categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" },
                  { name: "High", categories: "Data!$A$2:$A$4", values: "Data!$D$2:$D$4" },
                  { name: "Low", categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" },
                  { name: "Close", categories: "Data!$A$2:$A$4", values: "Data!$F$2:$F$4" }
                ]
              : [base];
        const options = applyChartPreset(preset, {
          series,
          title: `Preset ${preset}`,
          valueAxis: noAxis
            ? undefined
            : { numFmt: "#,##0", majorGridlines: true, lineColor: "70AD47" },
          categoryAxis: noAxis ? undefined : { textRotation: -30 },
          plotAreaOptions: { layout: { manualLayout: { x: 0.1, y: 0.1, w: 0.8, h: 0.8 } } }
        });
        ws.addChart(options, `K${row}:Q${row + 8}`);
        row += 10;
      }
      for (const type of chartExTypes) {
        ws.addChartEx(
          {
            type,
            categories: "Data!$A$2:$A$4",
            series: [
              {
                name: type,
                values: "Data!$G$2:$G$4",
                hierarchy:
                  type === "sunburst" || type === "treemap" ? ["Data!$B$2:$B$4"] : undefined,
                subtotals: type === "waterfall" ? [2] : undefined
              }
            ]
          },
          `S${row}:Y${row + 8}`
        );
        row += 10;
      }

      const entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
      const audit = auditOoxmlPackage(entries);
      expect(audit.errors, audit.errors.join("\n")).toEqual([]);
      const contentTypes = getEntryContent(entries, "[Content_Types].xml")!;
      for (let i = 1; i <= presets.length; i++) {
        expect(entries.get(`xl/charts/chart${i}.xml`)).toBeDefined();
        expect(contentTypes).toContain(`/xl/charts/chart${i}.xml`);
      }
      for (let i = 1; i <= chartExTypes.length; i++) {
        expect(entries.get(`xl/charts/chartEx${i}.xml`)).toBeDefined();
        expect(contentTypes).toContain(`/xl/charts/chartEx${i}.xml`);
      }
      const drawingRels = getEntryContent(entries, "xl/drawings/_rels/drawing1.xml.rels")!;
      expect((drawingRels.match(/relationships\/chart"/g) ?? []).length).toBe(presets.length);
      expect((drawingRels.match(/relationships\/chartEx/g) ?? []).length).toBe(chartExTypes.length);
      const firstChart = getEntryContent(entries, "xl/charts/chart1.xml")!;
      expect(firstChart).toContain("<c:dLbls>");
      expect(firstChart).toContain("<c:spPr>");
      expect(firstChart).toContain("<c:numFmt");
    });

    it("audits a synthetic chart fixture matrix with combo, pivot chart, chartsheet, and ChartEx", async () => {
      const wb = new Workbook();
      const data = wb.addWorksheet("Data");
      data.addRows([
        ["Region", "Product", "Revenue", "Units", "Growth"],
        ["West", "A", 10, 1, 0.1],
        ["West", "B", 20, 2, 0.2],
        ["East", "A", 30, 3, 0.15]
      ]);
      data.addComboChart(
        {
          groups: [
            {
              type: "bar",
              series: [
                {
                  name: "Revenue",
                  categories: "Data!$A$2:$A$4",
                  values: "Data!$C$2:$C$4",
                  dataLabels: { showVal: true },
                  dataPoints: [{ index: 1, fill: "FFC000" }]
                }
              ]
            },
            {
              type: "line",
              useSecondaryAxis: true,
              series: [
                {
                  name: "Growth",
                  categories: "Data!$A$2:$A$4",
                  values: "Data!$E$2:$E$4",
                  marker: { symbol: "diamond", size: 8 },
                  trendline: { type: "linear", displayEq: true },
                  errorBars: { type: "fixedVal", value: 0.05 }
                }
              ]
            }
          ],
          title: "Revenue vs Growth"
        },
        "G1:N12"
      );

      const pivotSheet = wb.addWorksheet("Pivot");
      const pivotTable = pivotSheet.addPivotTable({
        sourceSheet: data,
        rows: ["Region"],
        values: ["Revenue"],
        metric: "sum"
      });
      pivotSheet.addPivotChart(
        pivotTable,
        {
          type: "bar",
          series: [
            {
              name: "Pivot Revenue",
              categories: "'Pivot'!$A$4:$A$5",
              values: "'Pivot'!$B$4:$B$5"
            }
          ],
          pivotChartOptions: {
            refreshOnOpen: true,
            dropZonesVisible: true,
            dropZoneCategories: true
          }
        },
        "D1:J10"
      );
      wb.addChartsheet("Classic ChartSheet", {
        chart: {
          type: "pie",
          series: [{ categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" }],
          title: "Pie Sheet"
        }
      });
      wb.addChartsheet("Modern ChartSheet", {
        chart: {
          type: "pareto",
          categories: "Data!$A$2:$A$4",
          series: [{ values: "Data!$C$2:$C$4" }]
        }
      });

      const entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
      const audit = auditOoxmlPackage(entries);
      expect(audit.errors, audit.errors.join("\n")).toEqual([]);
      const contentTypes = getEntryContent(entries, "[Content_Types].xml")!;
      expect(entries.get("xl/charts/chart1.xml")).toBeDefined();
      expect(entries.get("xl/charts/chart2.xml")).toBeDefined();
      expect(entries.get("xl/charts/chart3.xml")).toBeDefined();
      expect(entries.get("xl/charts/chartEx1.xml")).toBeDefined();
      expect(entries.get("xl/chartsheets/sheet1.xml")).toBeDefined();
      expect(entries.get("xl/chartsheets/sheet2.xml")).toBeDefined();
      expect(contentTypes).toContain("/xl/chartsheets/sheet1.xml");
      expect(contentTypes).toContain("/xl/charts/chartEx1.xml");
      const comboXml = getEntryContent(entries, "xl/charts/chart1.xml")!;
      expect(comboXml).toContain("<c:barChart>");
      expect(comboXml).toContain("<c:lineChart>");
      expect(comboXml).toContain("<c:trendline>");
      expect(comboXml).toContain("<c:errBars>");
      const pivotXml = getEntryContent(entries, "xl/charts/chart2.xml")!;
      expect(pivotXml).toContain("<c:pivotSource>");
      // Pivot chart options now emit the MS standard c14:pivotOptions
      // extension (ECMA-376 MS-XLSX §2.3.11) instead of a private excelts
      // namespace. Earlier versions wrote `excelts:pivotChartOptions` that
      // Excel silently ignored.
      expect(pivotXml).toContain("{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}");
      expect(pivotXml).toContain("<c14:pivotOptions>");
      expect(pivotXml).toContain('<c14:dropZoneCategories val="1"/>');
      expect(pivotXml).toContain('<c14:dropZonesVisible val="1"/>');
      const chartExXml = getEntryContent(entries, "xl/charts/chartEx1.xml")!;
      expect(chartExXml).toContain("paretoLine");
    });

    it("optionally smoke-opens chart workbooks with LibreOffice headless", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Data");
      ws.addRows([
        ["A", 10],
        ["B", 20]
      ]);
      ws.addChart(
        { type: "bar", series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }] },
        "D1:J10"
      );
      ws.addChartEx(
        { type: "funnel", categories: "Data!$A$1:$A$2", series: [{ values: "Data!$B$1:$B$2" }] },
        "D12:J22"
      );
      const result = await smokeRoundTripWithLibreOffice(
        new Uint8Array(await wb.xlsx.writeBuffer())
      );
      if (!result.available) {
        expect(result.skipped).toBeTruthy();
        return;
      }
      const entries = await extractAll(result.output!);
      expect(entries.get("xl/workbook.xml")).toBeDefined();
      expect(
        entries.get("xl/charts/chart1.xml") ?? entries.get("xl/charts/chartEx1.xml")
      ).toBeDefined();
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
    it("should access charts via high-level API after load", async () => {
      const workbook = await loadSampleWorkbook();

      // Find the worksheet that contains the chart (Pivot sheet)
      const pivotSheet = workbook.getWorksheet("Pivot");
      expect(pivotSheet).toBeDefined();

      const charts = pivotSheet!.getCharts();
      expect(charts.length).toBe(1);

      const chart = charts[0];
      expect(chart.chartNumber).toBe(1);

      // This chart has an auto-generated title (no explicit c:tx), so title is undefined
      expect(chart.title).toBeUndefined();

      // Verify chart model is accessible
      expect(chart.chartModel).toBeDefined();
      expect(chart.chartTypes.length).toBeGreaterThan(0);
      expect(chart.chartTypes[0].type).toBe("bar");
      expect(chart.axes.length).toBeGreaterThan(0);
    });

    it("should read chart title after setting it programmatically", async () => {
      const workbook = await loadSampleWorkbook();
      const pivotSheet = workbook.getWorksheet("Pivot");
      const chart = pivotSheet!.getCharts()[0];

      // Set a title
      chart.title = "Revenue by Region";
      expect(chart.title).toBe("Revenue by Region");

      // Round-trip and verify
      const outputBuffer = await workbook.xlsx.writeBuffer();
      const workbook2 = new Workbook();
      await workbook2.xlsx.load(outputBuffer);

      const chart2 = workbook2.getWorksheet("Pivot")!.getCharts()[0];
      expect(chart2.title).toBe("Revenue by Region");
    });

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

  // =========================================================================
  // Self-generated per-type fixture matrix
  // =========================================================================
  //
  // The chart-oracle test harness can load real .xlsx files produced by
  // Excel / WPS / LibreOffice from the optional `EXCELTS_ENTERPRISE_CORPUS_DIR`
  // directory; those fixtures cannot be generated from code by definition.
  // This `self-generated fixture matrix` block is the self-consistent
  // baseline that runs on every CI pipeline: one workbook per chart type,
  // round-tripped through writeBuffer + xlsx.load, with the resulting
  // model fields asserted against the authored values. It catches
  // "library round-trip regressed this type" before the real-file oracle
  // catches "Excel would refuse to open it".
  describe("self-generated per-type fixture matrix", () => {
    type ClassicTypeCase = { preset: ExcelChartPreset; expectedBuilderType: string };
    const classicCases: ClassicTypeCase[] = EXCEL_CHART_PRESETS.map(preset => ({
      preset,
      expectedBuilderType: applyChartPreset(preset, { series: [] }).type as string
    }));
    const chartExTypes: ChartExType[] = [
      "sunburst",
      "treemap",
      "waterfall",
      "funnel",
      "histogram",
      "pareto",
      "boxWhisker",
      "regionMap"
    ];

    it("round-trips each classic chart preset through writeBuffer + load", async () => {
      for (const { preset, expectedBuilderType } of classicCases) {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.addRows([
          ["A", 1, 1, 1, 1, 1, 1],
          ["B", 2, 2, 2, 2, 2, 2],
          ["C", 3, 3, 3, 3, 3, 3]
        ]);
        const noAxis =
          expectedBuilderType === "pie" ||
          expectedBuilderType === "pie3D" ||
          expectedBuilderType === "doughnut" ||
          expectedBuilderType === "ofPie";

        // Stock charts want four sibling series (Open/High/Low/Close) not one.
        const series =
          expectedBuilderType === "stock"
            ? [
                { name: "Open", categories: "S!$A$1:$A$3", values: "S!$C$1:$C$3" },
                { name: "High", categories: "S!$A$1:$A$3", values: "S!$D$1:$D$3" },
                { name: "Low", categories: "S!$A$1:$A$3", values: "S!$E$1:$E$3" },
                { name: "Close", categories: "S!$A$1:$A$3", values: "S!$F$1:$F$3" }
              ]
            : expectedBuilderType === "bubble"
              ? [
                  {
                    name: "S",
                    xValues: "S!$B$1:$B$3",
                    values: "S!$C$1:$C$3",
                    bubbleSize: "S!$D$1:$D$3"
                  }
                ]
              : expectedBuilderType === "scatter"
                ? [{ name: "S", xValues: "S!$B$1:$B$3", values: "S!$C$1:$C$3" }]
                : [{ name: "S", categories: "S!$A$1:$A$3", values: "S!$B$1:$B$3" }];

        const opts = applyChartPreset(preset, { series, title: `preset:${preset}` });
        ws.addChart(opts as never, "D1:J10");
        const buf = await wb.xlsx.writeBuffer();
        // Audit passes.
        const zip = await extractAll(new Uint8Array(buf));
        const audit = auditOoxmlPackage(zip);
        expect(audit.errors, `preset=${preset}: ${audit.errors.join("\n")}`).toEqual([]);

        // Load back and assert the primary chart type survives.
        const wb2 = new Workbook();
        await wb2.xlsx.load(buf);
        const loaded = wb2.getWorksheet("S")!.getCharts()[0];
        const loadedType = loaded.chartModel?.chart.plotArea.chartTypes[0].type;
        expect(loadedType, `preset=${preset} expected chart-type ${expectedBuilderType}`).toBe(
          expectedBuilderType
        );
        if (!noAxis) {
          expect(loaded.chartModel?.chart.plotArea.axes.length ?? 0).toBeGreaterThan(0);
        }
      }
    });

    it("round-trips each ChartEx type through writeBuffer + load", async () => {
      for (const type of chartExTypes) {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.addRows([
          ["A", 10],
          ["B", 20],
          ["C", 15]
        ]);
        // Every ChartEx type accepts the same minimal options surface.
        ws.addChartEx(
          {
            type,
            categories: "S!$A$1:$A$3",
            series: [{ values: "S!$B$1:$B$3" }],
            ...(type === "histogram" || type === "pareto" ? { binning: { binCount: 3 } } : {})
          },
          "D1:J10"
        );
        const buf = await wb.xlsx.writeBuffer();
        const zip = await extractAll(new Uint8Array(buf));
        const audit = auditOoxmlPackage(zip);
        expect(audit.errors, `type=${type}: ${audit.errors.join("\n")}`).toEqual([]);

        const wb2 = new Workbook();
        await wb2.xlsx.load(buf);
        const loaded = wb2.getWorksheet("S")!.getCharts()[0];
        const plotArea = loaded.chartExModel?.chartSpace.chart.plotArea;
        const firstSeries = plotArea?.plotAreaRegion?.series?.[0] ?? plotArea?.series?.[0];
        expect(
          firstSeries,
          `type=${type} should have at least one series after round-trip`
        ).toBeDefined();
        // `histogram` and `pareto` internally use clusteredColumn layoutId; other
        // types use their direct name. Accept either so this matrix
        // asserts the right shape without coupling to Excel's internal
        // normalisation.
        const expectedLayoutId =
          type === "histogram" || type === "pareto" ? "clusteredColumn" : type;
        expect(firstSeries!.layoutId).toBe(expectedLayoutId);
      }
    });
  });
});
