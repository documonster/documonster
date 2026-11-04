import * as ChartEntry from "@excel/chart/index";
import { ChartClassic, ChartEx, DrawingMl } from "@excel/chart/index";
import { drawChartPdf } from "@excel/chart/render/chart-renderer";
import { describe, expect, it } from "vitest";

function recordingSurface(): { surface: ChartEntry.ChartPdfDrawingSurface; calls: string[] } {
  const calls: string[] = [];
  const surface = {
    drawRect(o: { x: number; y: number; width: number; height: number }) {
      calls.push(`rect ${o.x} ${o.y} ${o.width} ${o.height}`);
      return surface;
    },
    drawLine() {
      calls.push("line");
      return surface;
    },
    drawText(text: string) {
      calls.push(`text ${text}`);
      return surface;
    },
    drawCircle() {
      calls.push("circle");
      return surface;
    },
    drawPath() {
      calls.push("path");
      return surface;
    }
  } as unknown as ChartEntry.ChartPdfDrawingSurface;
  return { surface, calls };
}

const model = ChartClassic.createModel({
  type: "bar",
  series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
});

describe("documonster/chart namespaces", () => {
  it("the published drawChartPdf draws the same picture as the internal one, chart area first", () => {
    const viaEntry = recordingSurface();
    const viaInternal = recordingSurface();
    const rect = { x: 10, y: 20, width: 300, height: 180 };
    ChartEntry.drawChartPdf(viaEntry.surface, model, rect);
    drawChartPdf(viaInternal.surface, model, rect);
    expect(viaEntry.calls).toEqual(viaInternal.calls);
    // The chart area, its 0.75 pt border inside the box.
    expect(viaEntry.calls[0]).toBe("rect 10.375 20.375 299.25 179.25");
  });

  it("ChartEx.toXml returns the part XML", () => {
    const ex = ChartEx.createModel({
      type: "histogram",
      series: [{ values: "Sheet1!$A$1:$A$3" }]
    });
    expect(ChartEx.toXml(ex)).toContain("<cx:chartSpace");
    expect(ChartEx.fromXml(ChartEx.toXml(ex)).chartSpace).toBeDefined();
  });

  it("keeps the deprecated flat names as forwards", () => {
    expect(ChartEntry.buildChartModel).toBe(ChartClassic.createModel);
    expect(ChartEntry.renderChartSvg).toBe(ChartClassic.toSvg);
    expect(ChartEntry.renderChartPng).toBe(ChartClassic.toPng);
    expect(ChartEntry.buildChartScene).toBe(ChartClassic.Backend.toScene);
    expect(ChartEntry.parseChartColors).toBe(ChartClassic.fromColorsXml);
    expect(ChartEntry.buildChartStyle).toBe(ChartClassic.toStyleXml);
    expect(ChartEntry.chartOptionsFromRows).toBe(ChartClassic.createOptionsFromRows);
    expect(ChartEntry.drawChartPdf).toBe(drawChartPdf);
    expect(ChartEntry.renderChartExSvg).toBe(ChartEx.toSvg);
    expect(ChartEntry.renderChartExPng).toBe(ChartEx.toPng);
    expect(ChartEntry.renderChartEx).toBe(ChartEx.toXml);
    expect(ChartEntry.parseChartEx).toBe(ChartEx.fromXml);
    expect(typeof ChartEntry.drawChartExPdf).toBe("function");
    expect(typeof ChartEntry.canRenderChartExAsVectorPdf).toBe("function");
    expect(ChartEntry.buildEffectFilter).toBe(ChartClassic.Backend.toSvgFilter);
    expect(ChartEntry.buildChartExModel).toBe(ChartEx.createModel);
    expect(ChartEntry.parseSpPr).toBe(DrawingMl.parseSpPr);
    expect(ChartEntry.buildSpPr).toBe(DrawingMl.createSpPr);
    expect(ChartEntry.buildTxPr).toBe(DrawingMl.createTxPr);
    expect(ChartEntry.setSpPrFill).toBe(DrawingMl.setSpPrFill);
  });
});
