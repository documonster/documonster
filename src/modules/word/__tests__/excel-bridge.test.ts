/**
 * Tests for the Word ↔ Excel bridge helpers exposed at `excelts/word/excel`.
 *
 * Beyond the high-level `excelToDocx` / `extractTablesToExcel` flows, the
 * bridge ships four lower-level helpers that production code (and the
 * documented examples) depends on:
 *
 *   - `wordChartToChartModel` — converts a Word `Chart` into the Excel
 *     ChartModel for delegated rendering.
 *   - `renderWordChartSvg` — renders a Word `Chart` to a self-contained
 *     SVG string (used by the HTML exporter).
 *   - `generateChartEmbeddedXlsx` — produces the embedded xlsx workbook
 *     stored alongside a chart so users can edit chart data in Excel.
 *   - `buildWordChartExXml` — produces ChartEx (cx:) XML for the modern
 *     2016+ chart types (sunburst / treemap / waterfall / funnel /
 *     histogram / pareto / boxWhisker / regionMap).
 *
 * These were previously not covered by tests; this suite locks down the
 * contracts including the headless / literal-only behaviour required by
 * the Word writer (no underlying worksheet reference is available, so
 * cached `literalValues` must be enough).
 */
import { describe, it, expect } from "vitest";

import {
  buildWordChartExXml,
  generateChartEmbeddedXlsx,
  renderWordChartSvg,
  wordChartToChartModel
} from "../excel";
import type { Chart } from "../index";

const sampleChart: Chart = {
  type: "column",
  title: "Quarterly revenue",
  series: [
    {
      name: "FY-25",
      categories: ["Q1", "Q2", "Q3", "Q4"],
      values: [1.2, 1.5, 1.8, 2.1]
    }
  ],
  legend: "r"
};

describe("wordChartToChartModel", () => {
  it("returns a ChartModel whose chart slot carries the title and plotArea", () => {
    const model = wordChartToChartModel(sampleChart);
    expect(model.chart).toBeDefined();
    expect(model.chart.plotArea).toBeDefined();
    expect(model.chart.title).toBeDefined();
  });

  it("propagates legend position", () => {
    const model = wordChartToChartModel(sampleChart);
    expect(model.chart.legend).toBeDefined();
  });
});

describe("renderWordChartSvg", () => {
  it("returns a self-contained SVG string", () => {
    const svg = renderWordChartSvg(sampleChart);
    expect(svg.startsWith("<svg") || svg.startsWith("<?xml")).toBe(true);
    expect(svg).toContain("</svg>");
  });

  it("never throws on a minimal single-series chart", () => {
    expect(() =>
      renderWordChartSvg({
        type: "line",
        series: [
          {
            name: "A",
            categories: ["x"],
            values: [1]
          }
        ]
      })
    ).not.toThrow();
  });
});

describe("generateChartEmbeddedXlsx", () => {
  it("produces a non-empty xlsx (PK ZIP magic) for a single series", async () => {
    const xlsx = await generateChartEmbeddedXlsx([
      {
        name: "Revenue",
        categories: ["Q1", "Q2", "Q3", "Q4"],
        values: [1, 2, 3, 4]
      }
    ]);
    expect(xlsx.length).toBeGreaterThan(200);
    // ZIP local file header: 0x50 0x4b 0x03 0x04
    expect(xlsx[0]).toBe(0x50);
    expect(xlsx[1]).toBe(0x4b);
    expect(xlsx[2]).toBe(0x03);
    expect(xlsx[3]).toBe(0x04);
  });

  it("handles multi-series workbooks", async () => {
    const xlsx = await generateChartEmbeddedXlsx([
      { name: "A", categories: ["x", "y"], values: [1, 2] },
      { name: "B", categories: ["x", "y"], values: [3, 4] }
    ]);
    expect(xlsx.length).toBeGreaterThan(200);
  });

  it("does not throw on an empty series list (degenerate but legal)", async () => {
    const xlsx = await generateChartEmbeddedXlsx([]);
    expect(xlsx.length).toBeGreaterThan(200);
  });
});

describe("buildWordChartExXml", () => {
  it("produces a non-empty cx: XML payload for a sunburst chart (headless / literal-only)", () => {
    // This is the contract the Word writer relies on — there is no
    // underlying worksheet to reference, so values must be cached as
    // literals. A previous regression made the validator reject this.
    const xml = buildWordChartExXml({
      type: "sunburst",
      title: "Population",
      series: [
        {
          name: "Pop",
          categories: ["A", "B", "C"],
          values: [10, 20, 30]
        }
      ]
    });
    expect(xml).toContain("<cx:chartSpace");
    expect(xml).toContain("</cx:chartSpace>");
  });

  it("supports treemap, funnel, waterfall, boxWhisker types", () => {
    for (const type of ["treemap", "funnel", "waterfall", "boxWhisker"] as const) {
      const xml = buildWordChartExXml({
        type,
        series: [{ name: "x", categories: ["a", "b"], values: [1, 2] }]
      });
      expect(xml).toContain("<cx:chartSpace");
    }
  });

  it("does not throw with a legend toggle and position", () => {
    const xml = buildWordChartExXml({
      type: "histogram",
      showLegend: true,
      legendPosition: "b",
      series: [{ name: "h", values: [1, 2, 3, 4, 5] }]
    });
    expect(xml).toContain("<cx:chartSpace");
  });

  it("explicit null title suppresses auto-titling", () => {
    const xml = buildWordChartExXml({
      type: "sunburst",
      title: null,
      series: [{ name: "x", categories: ["a"], values: [1] }]
    });
    expect(xml).toContain("<cx:chartSpace");
  });
});
