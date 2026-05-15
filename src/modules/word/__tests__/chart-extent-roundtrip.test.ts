/**
 * Chart drawing extent must survive a round-trip:
 *
 *   Chart.width/height
 *      → writer emits `<wp:extent cx cy>`
 *      → reader replaces the OpaqueDrawing with ChartContent and
 *        recovers `<wp:extent>` back into `chart.width` / `chart.height`
 *
 * Without the reader-side extent recovery the chart would re-emerge
 * with `chart.width === undefined`, forcing every consumer (the layout
 * engine, the PDF bridge, the SVG renderer) to fall back to the
 * default 6"×3.5" frame regardless of what the source document
 * actually specified.
 */

import { Document, packageDocx, readDocx, chart as chartBuilder } from "@word/index";
import type { ChartContent } from "@word/index";
import { describe, it, expect } from "vitest";

describe("chart drawing extent round-trip", () => {
  it("recovers chart.width / chart.height from <wp:extent> on read", async () => {
    // 5"×3" — distinct from the engine's 6"×3.5" default so the test
    // can prove the value came from the document, not from a fallback.
    const widthEmu = 5 * 914_400;
    const heightEmu = 3 * 914_400;

    const h = Document.create();
    const ch = chartBuilder({
      type: "bar",
      series: [
        {
          name: "Revenue",
          categories: ["Q1", "Q2"],
          values: [10, 20]
        }
      ],
      title: "Sized Chart",
      width: widthEmu,
      height: heightEmu
    });
    Document.addContent(h, ch);

    const bytes = await packageDocx(Document.build(h));
    const reread = await readDocx(bytes);

    const reChart = reread.body.find(b => b.type === "chart") as ChartContent | undefined;
    expect(reChart).toBeDefined();
    expect(reChart!.chart.width).toBe(widthEmu);
    expect(reChart!.chart.height).toBe(heightEmu);
  });

  it("leaves chart.width undefined when source document had no <wp:extent>", () => {
    // This is the negative case: a chart constructed in code without
    // a width must NOT magically gain one after a round-trip — that
    // would mean we silently injected a default at read time, which
    // would mask author intent on the next round-trip.
    const ch = chartBuilder({ type: "bar", series: [] });
    expect(ch.chart.width).toBeUndefined();
    expect(ch.chart.height).toBeUndefined();
  });
});
