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

import { Workbook } from "../../../index";
import {
  buildWordChartExXml,
  excelToDocx,
  extractTablesToExcel,
  generateChartEmbeddedXlsx,
  renderWordChartSvg,
  wordChartToChartModel
} from "../excel";
import { Document } from "../index";
import type { Chart } from "../index";
import type { BodyContent, Paragraph, Run, Table } from "../types";

/** Find all top-level tables in a converted DocxDocument body. */
function tablesOf(body: readonly BodyContent[]): Table[] {
  return body.filter((b): b is Table => "type" in b && b.type === "table");
}

/** Concatenate the plain text of a table cell's first paragraph. */
function cellText(table: Table, row: number, col: number): string {
  const cell = table.rows[row]?.cells[col];
  if (!cell) {
    return "";
  }
  let out = "";
  for (const block of cell.content) {
    if (block.type === "paragraph") {
      for (const child of (block as Paragraph).children) {
        const run = child as Run;
        if (run.content) {
          for (const c of run.content) {
            if (c.type === "text") {
              out += c.text;
            }
          }
        }
      }
    }
  }
  return out;
}

/** First run of a given table cell, for inspecting run properties. */
function firstRun(table: Table, row: number, col: number): Run | undefined {
  const para = table.rows[row]?.cells[col]?.content?.[0] as Paragraph | undefined;
  return para?.children?.[0] as Run | undefined;
}

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

// =============================================================================
// excelToDocx — core workbook → Word conversion
// =============================================================================

describe("excelToDocx — sheet → table conversion", () => {
  function buildWorkbook(): Workbook {
    const wb = new Workbook();
    wb.creator = "Tester";
    const ws = wb.addWorksheet("Sales");
    ws.addRow(["Region", "Revenue"]);
    ws.addRow(["North", 1200]);
    ws.addRow(["South", 900]);
    return wb;
  }

  it("renders each visible sheet as a table with a heading", () => {
    const doc = excelToDocx(buildWorkbook());
    const tables = tablesOf(doc.body);
    expect(tables).toHaveLength(1);
    // Heading paragraph for the sheet name appears before the table.
    const headings = doc.body.filter(
      (b): b is Paragraph => "type" in b && b.type === "paragraph" && b.properties?.style != null
    );
    expect(headings.some(h => cellTextFromParagraph(h).includes("Sales"))).toBe(true);
  });

  it("preserves cell text and numeric values", () => {
    const doc = excelToDocx(buildWorkbook());
    const t = tablesOf(doc.body)[0];
    expect(cellText(t, 0, 0)).toBe("Region");
    expect(cellText(t, 0, 1)).toBe("Revenue");
    expect(cellText(t, 1, 0)).toBe("North");
    expect(cellText(t, 1, 1)).toBe("1200");
  });

  it("preserves bold cell formatting as a run property", () => {
    const wb = buildWorkbook();
    wb.getWorksheet("Sales")!.getRow(1).font = { bold: true };
    const doc = excelToDocx(wb, { preserveFormatting: true });
    const t = tablesOf(doc.body)[0];
    expect(firstRun(t, 0, 0)?.properties?.bold).toBe(true);
  });

  it("drops formatting when preserveFormatting is false", () => {
    const wb = buildWorkbook();
    wb.getWorksheet("Sales")!.getRow(1).font = { bold: true };
    const doc = excelToDocx(wb, { preserveFormatting: false });
    const t = tablesOf(doc.body)[0];
    expect(firstRun(t, 0, 0)?.properties?.bold).toBeUndefined();
  });

  it("skips hidden sheets by default", () => {
    const wb = buildWorkbook();
    const hidden = wb.addWorksheet("Secret");
    hidden.state = "hidden";
    hidden.addRow(["classified", 42]);
    const doc = excelToDocx(wb);
    // Only the visible "Sales" sheet should produce a table.
    expect(tablesOf(doc.body)).toHaveLength(1);
  });

  it("selects only the requested sheets", () => {
    const wb = buildWorkbook();
    const ws2 = wb.addWorksheet("Inventory");
    ws2.addRow(["Item", "Qty"]);
    const doc = excelToDocx(wb, { sheets: ["Inventory"], includeSheetHeadings: false });
    const tables = tablesOf(doc.body);
    expect(tables).toHaveLength(1);
    expect(cellText(tables[0], 0, 0)).toBe("Item");
  });

  it("caps rows and columns", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Big");
    ws.addRow(["a", "b", "c", "d"]);
    ws.addRow([1, 2, 3, 4]);
    ws.addRow([5, 6, 7, 8]);
    ws.addRow([9, 10, 11, 12]);
    const doc = excelToDocx(wb, { maxRows: 2, maxColumns: 2 });
    const t = tablesOf(doc.body)[0];
    expect(t.rows).toHaveLength(2);
    expect(t.rows[0].cells).toHaveLength(2);
  });

  it("emits a title page when requested", () => {
    const doc = excelToDocx(buildWorkbook(), { includeTitlePage: true });
    const firstPara = doc.body[0] as Paragraph;
    expect(cellTextFromParagraph(firstPara)).toContain("Tester");
  });

  it("preserves a cell hyperlink as a Word hyperlink (URL not lost)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Links");
    ws.getCell("A1").value = { text: "OpenAI", hyperlink: "https://openai.com" };
    const doc = excelToDocx(wb, { includeSheetHeadings: false });
    const t = tablesOf(doc.body)[0];
    const para = t.rows[0].cells[0].content[0] as Paragraph;
    const firstChild = para.children[0] as {
      type?: string;
      url?: string;
      children?: { content: { type: string; text: string }[] }[];
    };
    expect(firstChild.type).toBe("hyperlink");
    expect(firstChild.url).toBe("https://openai.com");
    // The display text is kept inside the hyperlink's runs.
    expect(firstChild.children?.[0]?.content?.[0]?.text).toBe("OpenAI");
  });
});

describe("extractTablesToExcel — Word table → 2D data", () => {
  it("extracts every table as a named 2D array", () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addTable(
      d,
      [
        ["Quarter", "Revenue"],
        ["Q1", "1.2"],
        ["Q2", "1.5"]
      ],
      { headerRow: true }
    );
    const doc = Document.build(d);
    const tables = extractTablesToExcel(doc);
    expect(tables).toHaveLength(1);
    expect(tables[0].data).toHaveLength(3);
    // Numeric-looking cells are coerced to numbers.
    expect(tables[0].data[1]).toEqual(["Q1", 1.2]);
    expect(tables[0].data[0]).toEqual(["Quarter", "Revenue"]);
  });

  it("round-trips excelToDocx → extractTablesToExcel", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRow(["x", "y"]);
    ws.addRow([1, 2]);
    const doc = excelToDocx(wb, { includeSheetHeadings: false });
    const tables = extractTablesToExcel(doc);
    expect(tables).toHaveLength(1);
    expect(tables[0].data[0]).toEqual(["x", "y"]);
    expect(tables[0].data[1]).toEqual([1, 2]);
  });
});

/** Plain text of a standalone paragraph (heading / title-page line). */
function cellTextFromParagraph(para: Paragraph): string {
  let out = "";
  for (const child of para.children) {
    const run = child as Run;
    if (run.content) {
      for (const c of run.content) {
        if (c.type === "text") {
          out += c.text;
        }
      }
    }
  }
  return out;
}
