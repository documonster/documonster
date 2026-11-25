import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import { Document, DocxArgumentError } from "@word/index";
import type { ChartContent, ChartSeries, ChartType } from "@word/index";
import { checkDocxElementOrder } from "@word/writer/order-check";
import { describe, expect, it } from "vitest";

/**
 * Word writes its charts itself (`chart-writer.ts`), not through the workbook's chart engine, so the order its
 * chart elements come out in is its own — and it wrote a trendline's `dispEq` before its `dispRSqr`. Every chart type
 * here carries every option the model has, and each chart part is held to the chart rules the workbook validator
 * uses (`@xml/ooxml-chart-order`).
 */

const TYPES: readonly ChartType[] = [
  "bar",
  "barStacked",
  "barPercentStacked",
  "column",
  "columnStacked",
  "columnPercentStacked",
  "line",
  "lineStacked",
  "lineMarked",
  "pie",
  "pie3D",
  "doughnut",
  "area",
  "areaStacked",
  "scatter",
  "scatterSmooth",
  "radar",
  "radarFilled",
  "bubble",
  "stock",
  "surface",
  "surface3D",
  "surfaceWireframe",
  "surfaceWireframe3D"
];

const CATEGORIES = ["Q1", "Q2", "Q3", "Q4"];

/** Chart types whose series have neither trendlines nor error bars (`CT_PieSer`, `CT_RadarSer`, `CT_SurfaceSer`). */
const NO_ANALYSIS = new Set<ChartType>([
  "pie",
  "pie3D",
  "doughnut",
  "radar",
  "radarFilled",
  "surface",
  "surface3D",
  "surfaceWireframe",
  "surfaceWireframe3D"
]);

/** Chart types whose series take an error bar in each direction (`CT_ScatterSer`, `CT_BubbleSer`, `CT_AreaSer`). */
const TWO_ERROR_BARS = new Set<ChartType>([
  "scatter",
  "scatterSmooth",
  "bubble",
  "area",
  "areaStacked"
]);

function series(name: string, extra: Partial<ChartSeries>): ChartSeries {
  return { name, categories: CATEGORIES, values: [1, 3, 2, 4], color: "4472C4", ...extra };
}

function fullChart(type: ChartType): ChartContent {
  const analysis = !NO_ANALYSIS.has(type);
  return {
    type: "chart",
    chart: {
      type,
      title: `${type} chart`,
      legend: "b",
      style: 2,
      plotAreaColor: "FFFFFF",
      chartAreaColor: "F2F2F2",
      view3d: true,
      dataLabels: { showValue: true, showCategory: true, showSerName: true, showPercent: true },
      categoryAxis: { title: "Quarter", numberFormat: "General" },
      valueAxis: { title: "Value", min: 0, max: 10, majorUnit: 2, numberFormat: "0.0" },
      secondaryValueAxis: { title: "Secondary", min: 0 },
      series: [
        series("Trend", {
          showDataLabels: true,
          pointColors: ["FF0000", "00FF00"],
          ...(analysis && {
            trendline: {
              type: "polynomial",
              order: 2,
              displayEquation: true,
              displayRSquared: true
            },
            errorBars: { direction: "y", type: "percentage", value: 5 }
          })
        }),
        series("Average", {
          ...(analysis && {
            trendline: {
              type: "movingAvg",
              period: 2,
              displayEquation: true,
              displayRSquared: true
            },
            errorBars: {
              direction: TWO_ERROR_BARS.has(type) ? "both" : "y",
              type: "fixedVal",
              value: 1
            }
          })
        }),
        series("Secondary", { plotOnSecondaryAxis: true, chartType: "line" })
      ],
      secondaryType: "line",
      secondarySeries: [series("Overlay", { trendline: { type: "linear", displayRSquared: true } })]
    }
  };
}

describe("Word chart element order", () => {
  for (const type of TYPES) {
    it(`writes a ${type} chart with every option in schema order`, async () => {
      const bytes = await Document.toBuffer({ body: [fullChart(type)] });
      const parts = new Map<string, string>();
      for (const [path, entry] of await extractAll(bytes)) {
        if (/^word\/charts\/chart\d+\.xml$/.test(path)) {
          parts.set(path, new TextDecoder().decode(entry.data));
        }
      }
      expect(parts.size).toBeGreaterThan(0);
      expect(checkDocxElementOrder(parts)).toEqual([]);
    });
  }

  for (const type of NO_ANALYSIS) {
    it(`refuses a trendline or error bars on a ${type} chart's series, which have none`, async () => {
      const chart = (extra: Partial<ChartSeries>): ChartContent => ({
        type: "chart",
        chart: { type, series: [series("S", extra)] }
      });
      await expect(
        Document.toBuffer({ body: [chart({ trendline: { type: "linear" } })] })
      ).rejects.toThrow(DocxArgumentError);
      await expect(
        Document.toBuffer({ body: [chart({ errorBars: { direction: "y", type: "stdErr" } })] })
      ).rejects.toThrow(DocxArgumentError);
    });
  }
});

describe("Word chart error bars", () => {
  const chart = (type: ChartType): ChartContent => ({
    type: "chart",
    chart: {
      type,
      series: [series("S", { errorBars: { direction: "both", type: "fixedVal", value: 1 } })]
    }
  });

  for (const type of ["bar", "column", "line", "stock"] as const) {
    it(`refuses both directions on a ${type} chart's series, which has room for one error bar`, async () => {
      await expect(Document.toBuffer({ body: [chart(type)] })).rejects.toThrow(DocxArgumentError);
    });
  }

  for (const type of TWO_ERROR_BARS) {
    it(`writes both directions on a ${type} chart's series as two, and reads them back as both`, async () => {
      const bytes = await Document.toBuffer({ body: [chart(type)] });
      let xml = "";
      for (const [path, entry] of await extractAll(bytes)) {
        if (/^word\/charts\/chart\d+\.xml$/.test(path)) {
          xml = new TextDecoder().decode(entry.data);
        }
      }
      expect(xml.match(/<c:errBars>/g)).toHaveLength(2);
      const back = (await Document.read(bytes)).body[0] as ChartContent;
      expect(back.chart.series[0]!.errorBars?.direction).toBe("both");
    });
  }
});

describe("two error bars that differ", () => {
  it("are read as the first, and written back as read", async () => {
    const bytes = await Document.toBuffer({
      body: [
        {
          type: "chart",
          chart: {
            type: "scatter",
            series: [series("S", { errorBars: { direction: "both", type: "fixedVal", value: 1 } })]
          }
        }
      ]
    });
    const files = new Map<string, Uint8Array>();
    let chartPath = "";
    for (const [path, entry] of await extractAll(bytes)) {
      files.set(path, entry.data);
      if (/^word\/charts\/chart\d+\.xml$/.test(path)) {
        chartPath = path;
      }
    }
    // Make the second bar a percentage: no longer the same bar in two directions.
    const original = new TextDecoder().decode(files.get(chartPath)!);
    const second = original.lastIndexOf('<c:errValType val="fixedVal"/>');
    const edited =
      original.slice(0, second) +
      '<c:errValType val="percentage"/>' +
      original.slice(second + '<c:errValType val="fixedVal"/>'.length);
    expect(edited).not.toBe(original);
    files.set(chartPath, new TextEncoder().encode(edited));
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    const read = await Document.read(await archive.toBuffer());
    const errorBars = (read.body[0] as ChartContent).chart.series[0]!.errorBars;
    expect(errorBars).toEqual({ direction: "x", type: "fixedVal", value: 1 });

    let written = "";
    for (const [path, entry] of await extractAll(await Document.toBuffer(read))) {
      if (/^word\/charts\/chart\d+\.xml$/.test(path)) {
        written = new TextDecoder().decode(entry.data);
      }
    }
    expect(written.match(/<c:errBars>/g)).toHaveLength(2);
    expect(written).toContain('<c:errValType val="percentage"/>');
  });
});

describe("Word chart axes and data, as Office opens them", () => {
  async function chartXml(chart: ChartContent): Promise<string> {
    for (const [path, entry] of await extractAll(await Document.toBuffer({ body: [chart] }))) {
      if (/^word\/charts\/chart\d+\.xml$/.test(path)) {
        return new TextDecoder().decode(entry.data);
      }
    }
    throw new Error("no chart part");
  }
  const twoSeries = (type: ChartType, categories = CATEGORIES): ChartContent => ({
    type: "chart",
    chart: {
      type,
      series: [series("A", { categories }), series("B", { categories, values: [5, 6, 7, 8] })]
    }
  });

  for (const type of ["scatter", "scatterSmooth", "bubble"] as const) {
    it(`plots a ${type} chart on two value axes, from xVal and yVal`, async () => {
      const xml = await chartXml(twoSeries(type));
      expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
      expect(xml).not.toContain("<c:catAx>");
      expect(xml).not.toMatch(/<c:(?:cat|val)>/);
      expect(xml).toContain('<c:crossBetween val="midCat"/>');
    });
  }

  it("writes text x values as text and numeric ones as numbers", async () => {
    expect(await chartXml(twoSeries("scatter"))).toContain("<c:xVal><c:strRef>");
    const numeric = await chartXml(twoSeries("scatter", ["1", "2.5", "4", "8"]));
    expect(numeric).toContain("<c:xVal><c:numRef>");
    expect(numeric).not.toContain("<c:xVal><c:strRef>");
  });

  it("points each series' values at its own column, as its name is", async () => {
    const xml = await chartXml(twoSeries("line"));
    expect(xml).toContain("<c:f>Sheet1!$B$1</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$B$2:$B$5</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$C$1</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$C$2:$C$5</c:f>");
  });

  for (const type of ["surface", "surface3D", "surfaceWireframe", "surfaceWireframe3D"] as const) {
    it(`gives a ${type} chart the series axis it is drawn over, once, also when read back`, async () => {
      const bytes = await Document.toBuffer({ body: [twoSeries(type)] });
      const read = await Document.read(bytes);
      for (const xml of [
        await chartXml(twoSeries(type)),
        await chartXml(read.body[0] as ChartContent)
      ]) {
        expect(xml.match(/<c:serAx>/g)).toHaveLength(1);
        const group = /<c:surface(?:3D)?Chart>[\s\S]*<\/c:surface(?:3D)?Chart>/.exec(xml)![0];
        expect(group.match(/<c:axId /g)).toHaveLength(3);
      }
    });
  }
});
