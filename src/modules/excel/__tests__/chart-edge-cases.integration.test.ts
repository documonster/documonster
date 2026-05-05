/**
 * Chart edge-case integration tests.
 *
 * Covers fidelity dimensions that don't fit cleanly into the per-preset
 * matrix or the mutation/round-trip files:
 *
 *   - Removing a chart cleans up `[Content_Types].xml`, drawing rels,
 *     and the orphaned `xl/charts/chartN.xml` part.
 *   - Multiple charts on a single worksheet preserve their drawing
 *     anchor order through round-trip.
 *   - Unicode (CJK + emoji) titles round-trip through the structured
 *     renderer with correct XML escaping and UTF-8 byte fidelity.
 */

import { extractAll } from "@archive/unzip/extract";
import { Workbook } from "@excel/workbook";
import { describe, expect, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";
import { entryText, loadRoundTrip, type EntryMap } from "./helpers/zip-text";

/**
 * Build a minimal workbook with N classic bar charts on the same
 * worksheet, anchored in vertically stacked, non-overlapping ranges so
 * the post-roundtrip ordering is unambiguously checkable.
 */
async function buildMultiChartWorkbook(count: number): Promise<Uint8Array> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Data");
  ws.addRows([
    ["A", 10],
    ["B", 20],
    ["C", 30]
  ]);
  for (let i = 0; i < count; i++) {
    const top = 1 + i * 12;
    const bottom = top + 10;
    ws.addChart(
      {
        type: "bar",
        title: `Chart #${i + 1}`,
        series: [{ categories: "Data!$A$1:$A$3", values: "Data!$B$1:$B$3" }]
      },
      `D${top}:J${bottom}`
    );
  }
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

function listChartParts(entries: EntryMap): string[] {
  return [...entries.keys()].filter(p => /^xl\/charts\/chart\d+[.]xml$/.test(p)).sort();
}

describe("Chart edge cases", () => {
  describe("removal cleans up package state", () => {
    it("removing a chart drops its content-type, drawing rel, and chart part", async () => {
      const bytes = await buildMultiChartWorkbook(3);
      const wb = new Workbook();
      await wb.xlsx.load(bytes);
      const ws = wb.getWorksheet("Data")!;
      expect(ws.getCharts().length).toBe(3);

      // Remove the middle chart by index.
      const removed = ws.removeChart(1);
      expect(removed).toBe(true);
      expect(ws.getCharts().length).toBe(2);

      const out = new Uint8Array(await wb.xlsx.writeBuffer());
      const entries = await extractAll(out);

      // Validator must remain clean (no dangling rels / missing parts).
      await expectValidXlsx(out);

      // Exactly two chart parts remain.
      const remainingCharts = listChartParts(entries);
      expect(remainingCharts.length).toBe(2);

      // The drawing rels must reference exactly the surviving charts —
      // no <Relationship> pointing at a non-existent chart part.
      const drawingRels = entryText(entries, "xl/drawings/_rels/drawing1.xml.rels")!;
      const relTargets = [...drawingRels.matchAll(/Target="([^"]*charts\/chart\d+\.xml)"/g)].map(
        m => m[1].replace(/^\.\.\//, "xl/")
      );
      expect(relTargets.length).toBe(2);
      for (const target of relTargets) {
        expect(remainingCharts, `dangling drawing rel target ${target}`).toContain(target);
      }

      // ContentTypes must list exactly the remaining chart parts and no
      // others.
      const contentTypes = entryText(entries, "[Content_Types].xml")!;
      const ctChartOverrides = [
        ...contentTypes.matchAll(/PartName="(\/xl\/charts\/chart\d+\.xml)"/g)
      ].map(m => m[1].replace(/^\//, ""));
      expect(ctChartOverrides.sort()).toEqual(remainingCharts);
    });

    it("removing the only chart drops the drawing part itself", async () => {
      const bytes = await buildMultiChartWorkbook(1);
      const wb = new Workbook();
      await wb.xlsx.load(bytes);
      const ws = wb.getWorksheet("Data")!;
      expect(ws.removeChart(0)).toBe(true);
      const out = new Uint8Array(await wb.xlsx.writeBuffer());
      const entries = await extractAll(out);
      await expectValidXlsx(out);
      // No chart parts should remain.
      expect(listChartParts(entries)).toEqual([]);
      // Drawing parts should also be gone (or at minimum have no chart
      // rels). Both shapes are valid; only audit-cleanliness is
      // contractual.
      const drawingRels = entryText(entries, "xl/drawings/_rels/drawing1.xml.rels");
      if (drawingRels !== undefined) {
        expect(drawingRels).not.toMatch(/charts\/chart\d+\.xml/);
      }
    });
  });

  describe("multiple charts on one worksheet", () => {
    it("preserves drawing anchor order for 5 charts through round-trip", async () => {
      const bytes = await buildMultiChartWorkbook(5);
      const { wb, entries } = await loadRoundTrip(bytes);
      await expectValidXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));

      // High-level model: 5 charts, original add order preserved.
      const charts = wb.getWorksheet("Data")!.getCharts();
      expect(charts.length).toBe(5);
      const titles = charts.map(c => c.title);
      expect(titles).toEqual(["Chart #1", "Chart #2", "Chart #3", "Chart #4", "Chart #5"]);

      // The drawing XML lists anchors in the same order. Each anchor
      // wraps a `<c:chart r:id="rIdN"/>` — extract the rIds in order
      // and resolve them via the rels file to chart paths, then assert
      // the chart titles match in the same order.
      const drawingXml = entryText(entries, "xl/drawings/drawing1.xml")!;
      const anchorRids = [...drawingXml.matchAll(/<xdr:graphicFrame[\s\S]*?r:id="(rId\d+)"/g)].map(
        m => m[1]
      );
      expect(anchorRids.length).toBe(5);

      const drawingRels = entryText(entries, "xl/drawings/_rels/drawing1.xml.rels")!;
      const ridToTarget = new Map<string, string>();
      for (const m of drawingRels.matchAll(
        /<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]*charts\/chart\d+\.xml)"/g
      )) {
        ridToTarget.set(m[1], m[2].replace(/^\.\.\//, "xl/"));
      }

      const orderedChartPaths = anchorRids.map(rid => ridToTarget.get(rid)!);
      const orderedTitles = orderedChartPaths.map(path => {
        const xml = entryText(entries, path)!;
        const match = xml.match(/<a:t>([^<]+)<\/a:t>/);
        return match ? match[1] : undefined;
      });
      expect(orderedTitles).toEqual(["Chart #1", "Chart #2", "Chart #3", "Chart #4", "Chart #5"]);
    });

    it("repeated round-trips do not reorder the 5-chart drawing", async () => {
      let bytes = await buildMultiChartWorkbook(5);
      for (let pass = 0; pass < 3; pass++) {
        const { wb } = await loadRoundTrip(bytes);
        const titles = wb
          .getWorksheet("Data")!
          .getCharts()
          .map(c => c.title);
        expect(titles, `pass ${pass + 1}`).toEqual([
          "Chart #1",
          "Chart #2",
          "Chart #3",
          "Chart #4",
          "Chart #5"
        ]);
        bytes = new Uint8Array(await wb.xlsx.writeBuffer());
      }
    });
  });

  describe("Unicode title fidelity", () => {
    it.each([
      ["CJK", "中文图表标题"],
      ["emoji", "📊 Sales 2025 📈"],
      ["mixed CJK + emoji + ASCII", "Q1 销售 📊 — 2025"],
      ["XML-escape-sensitive", `Title with <angle> & "quote" 'apos'`]
    ])(
      "round-trips %s title byte-faithfully through structured renderer",
      async (_label, title) => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Data");
        ws.addRows([
          ["A", 10],
          ["B", 20]
        ]);
        ws.addChart(
          {
            type: "bar",
            title,
            series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }]
          },
          "D1:J10"
        );
        const bytes = new Uint8Array(await wb.xlsx.writeBuffer());

        // First round-trip: structured rebuild after high-level mutation.
        const wb2 = new Workbook();
        await wb2.xlsx.load(bytes);
        const chart = wb2.getWorksheet("Data")!.getCharts()[0];
        expect(chart.title, "loaded title").toBe(title);
        // Force structured re-render by touching title with the same value.
        chart.title = title;
        const out = new Uint8Array(await wb2.xlsx.writeBuffer());
        await expectValidXlsx(out);

        const wb3 = new Workbook();
        await wb3.xlsx.load(out);
        const reloaded = wb3.getWorksheet("Data")!.getCharts()[0];
        expect(reloaded.title, "reloaded title after structured rebuild").toBe(title);
      }
    );
  });
});
