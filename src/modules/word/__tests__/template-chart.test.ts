/**
 * DOCX Module - Template Chart Data Binding Tests
 */

import { describe, it, expect } from "vitest";

import { bindChartData } from "../template/template-chart";
import type { ChartBinding } from "../template/template-chart";
import type { DocxDocument, ChartContent, Chart, BodyContent } from "../types";

// =============================================================================
// Helpers
// =============================================================================

function makeChart(opts: {
  type?: string;
  title?: string;
  series?: { name: string; categories: string[]; values: number[]; color?: string }[];
  name?: string;
}): ChartContent {
  const chart: Chart = {
    type: (opts.type ?? "bar") as Chart["type"],
    title: opts.title ?? "Default Title",
    series: (opts.series ?? [
      { name: "Series1", categories: ["A", "B", "C"], values: [1, 2, 3] }
    ]) as Chart["series"]
  };
  return {
    type: "chart",
    chart,
    ...(opts.name ? { name: opts.name } : {})
  };
}

function makeParagraph(text: string): BodyContent {
  return {
    type: "paragraph",
    runs: [{ content: [{ type: "text", text }] }]
  } as unknown as BodyContent;
}

function makeDoc(body: BodyContent[]): DocxDocument {
  return { body };
}

// =============================================================================
// bindChartData - by index
// =============================================================================

describe("bindChartData - by index", () => {
  it("binds data to chart at index 0", () => {
    const chart = makeChart({ type: "bar" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "Revenue", values: [10, 20, 30] }],
        categories: ["Q1", "Q2", "Q3"]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series[0].name).toBe("Revenue");
    expect(updated.chart.series[0].values).toEqual([10, 20, 30]);
    expect(updated.chart.series[0].categories).toEqual(["Q1", "Q2", "Q3"]);
  });

  it("binds data to correct chart when multiple charts exist", () => {
    const chart0 = makeChart({ type: "bar", title: "Chart A" });
    const chart1 = makeChart({ type: "line", title: "Chart B" });
    const doc = makeDoc([chart0, makeParagraph("text"), chart1]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 1,
        series: [{ name: "Visits", values: [100, 200] }],
        categories: ["Jan", "Feb"]
      }
    ];

    const result = bindChartData(doc, bindings);
    // Chart 0 unchanged
    const first = result.body[0] as ChartContent;
    expect(first.chart.title).toBe("Chart A");
    expect(first.chart.series[0].name).toBe("Series1");

    // Chart 1 updated
    const second = result.body[2] as ChartContent;
    expect(second.chart.series[0].name).toBe("Visits");
    expect(second.chart.series[0].values).toEqual([100, 200]);
  });
});

// =============================================================================
// bindChartData - by name
// =============================================================================

describe("bindChartData - by name", () => {
  it("binds data to chart matched by name", () => {
    const chart = makeChart({ type: "column", name: "SalesChart" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: "SalesChart",
        series: [{ name: "Sales", values: [5, 10, 15] }],
        categories: ["X", "Y", "Z"]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series[0].name).toBe("Sales");
    expect(updated.chart.series[0].categories).toEqual(["X", "Y", "Z"]);
  });

  it("does not bind when name does not match", () => {
    const chart = makeChart({ type: "bar", name: "ChartA" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: "ChartB",
        series: [{ name: "New", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    const unchanged = result.body[0] as ChartContent;
    expect(unchanged.chart.series[0].name).toBe("Series1");
  });
});

// =============================================================================
// bindChartData - series, categories, title replacement
// =============================================================================

describe("bindChartData - data replacement", () => {
  it("replaces series completely", () => {
    const chart = makeChart({
      type: "line",
      series: [
        { name: "Old1", categories: ["A"], values: [1] },
        { name: "Old2", categories: ["A"], values: [2] }
      ]
    });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [
          { name: "New1", values: [10, 20] },
          { name: "New2", values: [30, 40] },
          { name: "New3", values: [50, 60] }
        ],
        categories: ["X", "Y"]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series).toHaveLength(3);
    expect(updated.chart.series[0].name).toBe("New1");
    expect(updated.chart.series[1].name).toBe("New2");
    expect(updated.chart.series[2].name).toBe("New3");
  });

  it("replaces categories", () => {
    const chart = makeChart({ type: "bar" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "S", values: [1, 2] }],
        categories: ["Cat1", "Cat2"]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series[0].categories).toEqual(["Cat1", "Cat2"]);
  });

  it("uses original categories when not specified in binding", () => {
    const chart = makeChart({
      type: "bar",
      series: [{ name: "Orig", categories: ["A", "B"], values: [1, 2] }]
    });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "New", values: [10, 20] }]
        // no categories specified
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series[0].categories).toEqual(["A", "B"]);
  });

  it("replaces chart title", () => {
    const chart = makeChart({ type: "pie", title: "Old Title" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "Slice", values: [40, 60] }],
        title: "New Title"
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.title).toBe("New Title");
  });

  it("preserves original title when not specified in binding", () => {
    const chart = makeChart({ type: "doughnut", title: "Keep Me" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "S", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.title).toBe("Keep Me");
  });

  it("applies series color when provided", () => {
    const chart = makeChart({ type: "column" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "Colored", values: [1, 2], color: "FF0000" }],
        categories: ["A", "B"]
      }
    ];

    const result = bindChartData(doc, bindings);
    const updated = result.body[0] as ChartContent;
    expect(updated.chart.series[0].color).toBe("FF0000");
  });
});

// =============================================================================
// bindChartData - no charts in document
// =============================================================================

describe("bindChartData - no charts in document", () => {
  it("returns document unchanged when body has no charts", () => {
    const doc = makeDoc([makeParagraph("Hello"), makeParagraph("World")]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "S", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    expect(result.body).toHaveLength(2);
    expect(result.body).toEqual(doc.body);
  });

  it("returns document unchanged when bindings array is empty", () => {
    const chart = makeChart({ type: "bar" });
    const doc = makeDoc([chart]);

    const result = bindChartData(doc, []);
    expect(result).toBe(doc); // same reference — early return
  });
});

// =============================================================================
// bindChartData - invalid chart reference
// =============================================================================

describe("bindChartData - invalid chart reference", () => {
  it("returns chart unchanged for out-of-range index", () => {
    const chart = makeChart({ type: "bar" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 99,
        series: [{ name: "S", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    const unchanged = result.body[0] as ChartContent;
    expect(unchanged.chart.series[0].name).toBe("Series1");
  });

  it("returns chart unchanged for non-matching name", () => {
    const chart = makeChart({ type: "line", name: "MyChart" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: "NonExistent",
        series: [{ name: "S", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    const unchanged = result.body[0] as ChartContent;
    expect(unchanged.chart.series[0].name).toBe("Series1");
  });

  it("skips unsupported chart types (e.g. scatter)", () => {
    const chart = makeChart({ type: "scatter" });
    const doc = makeDoc([chart]);

    const bindings: ChartBinding[] = [
      {
        chartRef: 0,
        series: [{ name: "S", values: [1] }]
      }
    ];

    const result = bindChartData(doc, bindings);
    const unchanged = result.body[0] as ChartContent;
    expect(unchanged.chart.series[0].name).toBe("Series1");
  });
});
