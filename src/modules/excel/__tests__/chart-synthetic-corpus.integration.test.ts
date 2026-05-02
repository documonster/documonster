/**
 * Synthetic chart corpus integration tests.
 *
 * Single home for the per-preset / per-ChartEx-layoutId package audit
 * matrix and the auto-mode LibreOffice open-validation gate. Built on
 * top of the `synthetic-fixtures` builders so per-type coverage is
 * data-driven and reusable from other integration files.
 *
 * Real-fixture round-trip lives in
 * `workbook-chart-roundtrip.integration.test.ts`. Mutation fidelity,
 * chartsheet, combo/axis matrix, and pivot chart fidelity each have
 * their own dedicated integration file.
 */

import { extractAll } from "@archive/unzip/extract";
import { applyChartPreset, EXCEL_CHART_PRESETS } from "@excel/chart";
import { Workbook } from "@excel/workbook";
import { beforeAll, describe, expect, it } from "vitest";

import { runLibreOfficeOpenValidationAuto } from "./helpers/external-oracle";
import { auditOoxmlPackage } from "./helpers/ooxml-package-audit";
import {
  buildChartExFixtures,
  buildClassicPresetFixtures,
  buildComboAxisFixtures,
  buildChartsheetFixtures,
  buildPivotChartFixtures,
  CHART_EX_TYPES,
  type SyntheticFixture
} from "./helpers/synthetic-fixtures";
import { entryText } from "./helpers/zip-text";

// Top-level fixture cache. Built once per file via `beforeAll` so the
// per-fixture matrix tests don't pay the xlsx-serialisation cost N
// times. Tests must NOT mutate `fixture.bytes` in place — they reload
// via `Workbook.load(...)` which copies the buffer.
let classicFixtures: SyntheticFixture[];
let chartExFixtures: SyntheticFixture[];
let comboAxisFixtures: SyntheticFixture[];
let chartsheetFixtures: SyntheticFixture[];
let pivotChartFixtures: SyntheticFixture[];

beforeAll(async () => {
  [classicFixtures, chartExFixtures, comboAxisFixtures, chartsheetFixtures, pivotChartFixtures] =
    await Promise.all([
      buildClassicPresetFixtures(),
      buildChartExFixtures(),
      buildComboAxisFixtures(),
      buildChartsheetFixtures(),
      buildPivotChartFixtures()
    ]);
});

async function expectAuditClean(fixture: SyntheticFixture): Promise<void> {
  const entries = await extractAll(fixture.bytes);
  const audit = auditOoxmlPackage(entries);
  expect(
    audit.errors,
    `${fixture.id} (${fixture.description})\n${audit.errors.join("\n")}`
  ).toEqual([]);
}

async function expectOpensInLibreOffice(fixture: SyntheticFixture): Promise<void> {
  const result = await runLibreOfficeOpenValidationAuto(fixture.bytes, `${fixture.id}.xlsx`);
  if (!result.available) {
    expect(result.skipped, fixture.id).toBeTruthy();
    return;
  }
  expect(result.exitCode, `${fixture.id}: ${result.stderr ?? ""}`).toBe(0);
  expect(result.outputs.length, fixture.id).toBeGreaterThan(0);
}

describe("Synthetic chart corpus", () => {
  describe("audit baseline", () => {
    it("keeps the package audit clean for every classic preset", async () => {
      const fixtures = classicFixtures;
      expect(fixtures.length).toBe(EXCEL_CHART_PRESETS.length);
      for (const fixture of fixtures) {
        await expectAuditClean(fixture);
      }
    });

    it("keeps the package audit clean for every ChartEx layoutId", async () => {
      const fixtures = chartExFixtures;
      expect(fixtures.length).toBe(CHART_EX_TYPES.length);
      for (const fixture of fixtures) {
        await expectAuditClean(fixture);
      }
    });

    it("keeps the package audit clean across combo/axis variants", async () => {
      const fixtures = comboAxisFixtures;
      expect(fixtures.length).toBeGreaterThan(0);
      for (const fixture of fixtures) {
        await expectAuditClean(fixture);
      }
    });

    it("keeps the package audit clean for chartsheet hosting", async () => {
      const fixtures = chartsheetFixtures;
      for (const fixture of fixtures) {
        await expectAuditClean(fixture);
      }
    });

    it("keeps the package audit clean for pivot chart variants", async () => {
      const fixtures = pivotChartFixtures;
      for (const fixture of fixtures) {
        await expectAuditClean(fixture);
      }
    });
  });

  describe("synthetic marker decoration", () => {
    it("preserves the SYNTHETIC marker AND vendor extLst on every classic chart part", async () => {
      const fixtures = classicFixtures;
      for (const fixture of fixtures) {
        const entries = await extractAll(fixture.bytes);
        const chartPaths = [...entries.keys()].filter(name =>
          /^xl\/charts\/chart\d+[.]xml$/.test(name)
        );
        expect(chartPaths.length, fixture.id).toBeGreaterThan(0);
        for (const path of chartPaths) {
          const xml = entryText(entries, path)!;
          expect(xml, `${fixture.id} ${path} marker`).toContain("SYNTHETIC-FIXTURE");
          expect(xml, `${fixture.id} ${path} extLst`).toContain("{synthetic-roundtrip-fixture}");
        }
      }
    });
  });

  describe("classic preset model survives load → write", () => {
    it("retains primary chart-type for every preset after one round trip", async () => {
      const fixtures = classicFixtures;
      for (const fixture of fixtures) {
        const expectedType = applyChartPreset(
          fixture.id.replace(/^classic-/, "") as Parameters<typeof applyChartPreset>[0],
          { series: [] }
        ).type;
        const wb = new Workbook();
        await wb.xlsx.load(fixture.bytes);
        const sheet = wb.getWorksheet("Data");
        expect(sheet, fixture.id).toBeDefined();
        const charts = sheet!.getCharts();
        expect(charts.length, fixture.id).toBeGreaterThanOrEqual(1);
        const loadedType = charts[0].chartModel?.chart.plotArea.chartTypes[0].type;
        expect(loadedType, `${fixture.id} expected ${expectedType}`).toBe(expectedType);
      }
    });
  });

  describe("ChartEx layoutId survives load → write", () => {
    it("retains a series with the correct layoutId for every ChartEx fixture", async () => {
      const fixtures = chartExFixtures;
      for (const fixture of fixtures) {
        const layoutType = fixture.id.replace(/^chartEx-/, "");
        const wb = new Workbook();
        await wb.xlsx.load(fixture.bytes);
        const charts = wb.getWorksheet("Data")!.getCharts();
        expect(charts.length, fixture.id).toBeGreaterThanOrEqual(1);
        const plotArea = charts[0].chartExModel?.chartSpace.chart.plotArea;
        const firstSeries = plotArea?.plotAreaRegion?.series?.[0] ?? plotArea?.series?.[0];
        expect(firstSeries, fixture.id).toBeDefined();
        // Histogram and Pareto normalise to clusteredColumn internally.
        const expectedLayout =
          layoutType === "histogram" || layoutType === "pareto" ? "clusteredColumn" : layoutType;
        expect(firstSeries!.layoutId, fixture.id).toBe(expectedLayout);
      }
    });
  });

  describe("LibreOffice open-validation (auto)", () => {
    it("opens every classic preset fixture if LibreOffice is installed", async () => {
      const fixtures = classicFixtures;
      for (const fixture of fixtures) {
        await expectOpensInLibreOffice(fixture);
      }
    });

    it("opens every ChartEx fixture if LibreOffice is installed", async () => {
      const fixtures = chartExFixtures;
      for (const fixture of fixtures) {
        await expectOpensInLibreOffice(fixture);
      }
    });

    it("opens combo/axis variants if LibreOffice is installed", async () => {
      const fixtures = comboAxisFixtures;
      for (const fixture of fixtures) {
        await expectOpensInLibreOffice(fixture);
      }
    });

    it("opens chartsheet variants if LibreOffice is installed", async () => {
      const fixtures = chartsheetFixtures;
      for (const fixture of fixtures) {
        await expectOpensInLibreOffice(fixture);
      }
    });

    it("opens pivot chart variants if LibreOffice is installed", async () => {
      const fixtures = pivotChartFixtures;
      for (const fixture of fixtures) {
        await expectOpensInLibreOffice(fixture);
      }
    });
  });

  // -----------------------------------------------------------------------
  // Edge-case audit cases that don't fit the per-fixture matrix above
  // -----------------------------------------------------------------------

  describe("audit edge cases", () => {
    it("validates rels and content types when classic + ChartEx + chartsheet all coexist", async () => {
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

      const entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
      const contentTypes = entryText(entries, "[Content_Types].xml")!;
      expect(contentTypes).toContain("/xl/charts/chart1.xml");
      expect(contentTypes).toContain("/xl/charts/chartEx1.xml");
      expect(contentTypes).toContain("/xl/chartsheets/sheet1.xml");

      const drawingRels = entryText(entries, "xl/drawings/_rels/drawing1.xml.rels")!;
      expect(drawingRels).toContain("../charts/chart1.xml");
      expect(drawingRels).toContain("../charts/chartEx1.xml");
      const chartRels = entryText(entries, "xl/charts/_rels/chart1.xml.rels")!;
      expect(chartRels).toContain("chartStyle");
      expect(chartRels).toContain("chartColorStyle");
      const chartsheetRels = entryText(entries, "xl/chartsheets/_rels/sheet1.xml.rels")!;
      expect(chartsheetRels).toContain("../drawings/");
    });

    it("reports missing cx:data id, cx:axis id, and externalData rel references", async () => {
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
      const originalXml = new TextDecoder().decode(chartExEntry.data);
      const axisIdMatches = [...originalXml.matchAll(/<cx:axisId val="(\d+)"\/>/g)];
      expect(axisIdMatches.length).toBeGreaterThanOrEqual(2);
      const secondAxisId = axisIdMatches[1][1];
      chartExEntry.data = new TextEncoder().encode(
        originalXml
          .replace('<cx:dataId val="1"/>', '<cx:dataId val="99"/>')
          .replace(`<cx:axisId val="${secondAxisId}"/>`, '<cx:axisId val="42"/>')
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
  });
});
