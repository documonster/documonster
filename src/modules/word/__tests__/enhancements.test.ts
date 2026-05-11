/**
 * DOCX Module - Enhancement Tests
 *
 * Tests covering:
 * 1. Chart data editing (modify series data, round-trip verify)
 * 2. RTL/BiDi support verification
 * 3. Large file streaming performance
 */

import { describe, it, expect } from "vitest";

import { Document, chart, packageDocx, readDocx, toBuffer } from "../index";
import type { ChartContent, DocxDocument, Paragraph, Run } from "../types";

// =============================================================================
// 1. Chart Data Editing
// =============================================================================

describe("Chart data editing round-trip", () => {
  it("should preserve modified chart series data through round-trip", async () => {
    // Step 1: Create a document with a chart
    const h = Document.create();
    Document.addParagraph(h, "Chart test");
    Document.addContent(
      h,
      chart({
        type: "bar",
        title: "Sales",
        series: [
          { name: "Revenue", categories: ["Q1", "Q2", "Q3", "Q4"], values: [100, 200, 300, 400] }
        ]
      })
    );
    const doc1 = Document.build(h);

    // Step 2: Package and read back
    const buffer1 = await toBuffer(doc1);
    const parsed1 = await readDocx(buffer1);

    // Verify initial chart data
    const chartItems1 = parsed1.body.filter(b => b.type === "chart");
    expect(chartItems1.length).toBe(1);
    const chartContent1 = chartItems1[0] as ChartContent;
    expect(chartContent1.chart.series[0]!.values).toEqual([100, 200, 300, 400]);

    // Step 3: Modify chart series data
    const modifiedDoc: DocxDocument = {
      ...parsed1,
      body: parsed1.body.map(item => {
        if (item.type === "chart") {
          const chartItem = item as ChartContent;
          return {
            ...chartItem,
            chart: {
              ...chartItem.chart,
              series: chartItem.chart.series.map(s => ({
                ...s,
                values: [500, 600, 700, 800]
              }))
            }
          };
        }
        return item;
      })
    };

    // Step 4: Re-package
    const buffer2 = await packageDocx(modifiedDoc);

    // Step 5: Read back and verify updated data
    const parsed2 = await readDocx(buffer2);
    const chartItems2 = parsed2.body.filter(b => b.type === "chart");
    expect(chartItems2.length).toBe(1);
    const chartContent2 = chartItems2[0] as ChartContent;
    expect(chartContent2.chart.series[0]!.values).toEqual([500, 600, 700, 800]);
    expect(chartContent2.chart.series[0]!.categories).toEqual(["Q1", "Q2", "Q3", "Q4"]);
    expect(chartContent2.chart.series[0]!.name).toBe("Revenue");
    expect(chartContent2.chart.title).toBe("Sales");
  });

  it("should support modifying multiple series", async () => {
    const h = Document.create();
    Document.addContent(
      h,
      chart({
        type: "line",
        title: "Multi-Series",
        series: [
          { name: "Series A", categories: ["X", "Y", "Z"], values: [1, 2, 3] },
          { name: "Series B", categories: ["X", "Y", "Z"], values: [4, 5, 6] }
        ]
      })
    );
    const doc = Document.build(h);
    const buffer1 = await toBuffer(doc);
    const parsed1 = await readDocx(buffer1);

    // Modify values
    const modifiedDoc: DocxDocument = {
      ...parsed1,
      body: parsed1.body.map(item => {
        if (item.type === "chart") {
          const chartItem = item as ChartContent;
          return {
            ...chartItem,
            chart: {
              ...chartItem.chart,
              series: chartItem.chart.series.map((s, i) => ({
                ...s,
                values: i === 0 ? [10, 20, 30] : [40, 50, 60]
              }))
            }
          };
        }
        return item;
      })
    };

    const buffer2 = await packageDocx(modifiedDoc);
    const parsed2 = await readDocx(buffer2);
    const chartItems = parsed2.body.filter(b => b.type === "chart");
    expect(chartItems.length).toBe(1);
    const cc = chartItems[0] as ChartContent;
    expect(cc.chart.series.length).toBe(2);
    expect(cc.chart.series[0]!.values).toEqual([10, 20, 30]);
    expect(cc.chart.series[1]!.values).toEqual([40, 50, 60]);
  });

  it("should support adding new data points to a series", async () => {
    const h = Document.create();
    Document.addContent(
      h,
      chart({
        type: "column",
        series: [{ name: "Data", categories: ["A", "B"], values: [10, 20] }]
      })
    );
    const doc = Document.build(h);
    const buffer1 = await toBuffer(doc);
    const parsed1 = await readDocx(buffer1);

    // Add more data points
    const modifiedDoc: DocxDocument = {
      ...parsed1,
      body: parsed1.body.map(item => {
        if (item.type === "chart") {
          const chartItem = item as ChartContent;
          return {
            ...chartItem,
            chart: {
              ...chartItem.chart,
              series: [
                {
                  ...chartItem.chart.series[0]!,
                  categories: ["A", "B", "C", "D"],
                  values: [10, 20, 30, 40]
                }
              ]
            }
          };
        }
        return item;
      })
    };

    const buffer2 = await packageDocx(modifiedDoc);
    const parsed2 = await readDocx(buffer2);
    const chartItems = parsed2.body.filter(b => b.type === "chart");
    const cc = chartItems[0] as ChartContent;
    expect(cc.chart.series[0]!.categories).toEqual(["A", "B", "C", "D"]);
    expect(cc.chart.series[0]!.values).toEqual([10, 20, 30, 40]);
  });
});

// =============================================================================
// 2. RTL/BiDi Support Verification
// =============================================================================

describe("RTL/BiDi support round-trip", () => {
  it("should preserve paragraph bidi flag", async () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { bidi: true },
          children: [{ content: [{ type: "text", text: "مرحبا" }] } as Run]
        }
      ]
    };
    const buffer = await packageDocx(doc);
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    expect(para.properties?.bidi).toBe(true);
  });

  it("should preserve run rightToLeft property", async () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [
            {
              properties: { rightToLeft: true },
              content: [{ type: "text", text: "שלום" }]
            } as Run
          ]
        }
      ]
    };
    const buffer = await packageDocx(doc);
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.properties?.rightToLeft).toBe(true);
  });

  it("should preserve section textDirection tbRl", async () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "text", text: "vertical" }] } as Run]
        }
      ],
      sectionProperties: {
        textDirection: "tbRl"
      }
    };
    const buffer = await packageDocx(doc);
    const parsed = await readDocx(buffer);
    expect(parsed.sectionProperties?.textDirection).toBe("tbRl");
  });

  it("should preserve section bidi with textDirection combined", async () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { bidi: true },
          children: [{ content: [{ type: "text", text: "RTL" }] } as Run]
        }
      ],
      sectionProperties: {
        bidi: true,
        textDirection: "tbRl"
      }
    };
    const buffer = await packageDocx(doc);
    const parsed = await readDocx(buffer);
    expect(parsed.sectionProperties?.bidi).toBe(true);
    expect(parsed.sectionProperties?.textDirection).toBe("tbRl");
    const para = parsed.body[0] as Paragraph;
    expect(para.properties?.bidi).toBe(true);
  });
});

// =============================================================================
// 3. Large File Streaming Performance
// =============================================================================

describe("Large file performance", () => {
  it("should handle 10000 paragraphs within reasonable time", { timeout: 30000 }, async () => {
    const PARA_COUNT = 10000;

    // Create a document with 10000 paragraphs
    const body: Paragraph[] = [];
    for (let i = 0; i < PARA_COUNT; i++) {
      body.push({
        type: "paragraph",
        children: [
          {
            content: [{ type: "text", text: `Paragraph number ${i + 1} with some content.` }]
          } as Run
        ]
      });
    }
    const doc: DocxDocument = { body };

    // Time the packaging
    const startPackage = performance.now();
    const buffer = await packageDocx(doc);
    const packageTime = performance.now() - startPackage;

    // Should complete within 5 seconds
    expect(packageTime).toBeLessThan(5000);

    // Verify the buffer is a valid ZIP (starts with PK signature)
    expect(buffer[0]).toBe(0x50); // 'P'
    expect(buffer[1]).toBe(0x4b); // 'K'

    // Read back and verify paragraph count
    const startRead = performance.now();
    const parsed = await readDocx(buffer);
    const readTime = performance.now() - startRead;

    // Reading should also be within reasonable time
    expect(readTime).toBeLessThan(10000);

    // Verify body length
    expect(parsed.body.length).toBe(PARA_COUNT);

    // Spot-check some paragraphs
    const firstPara = parsed.body[0] as Paragraph;
    const firstRun = firstPara.children[0] as Run;
    expect(firstRun.content[0]).toEqual({
      type: "text",
      text: "Paragraph number 1 with some content."
    });

    const lastPara = parsed.body[PARA_COUNT - 1] as Paragraph;
    const lastRun = lastPara.children[0] as Run;
    expect(lastRun.content[0]).toEqual({
      type: "text",
      text: `Paragraph number ${PARA_COUNT} with some content.`
    });
  });
});
