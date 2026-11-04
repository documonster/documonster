/**
 * A Word chart read from a file and written back loses nothing: every element
 * of the part comes back — those `Chart` types through their fields, the rest
 * as read — with the part's relationships. Setting a field rewrites only the
 * element it governs.
 */

import { readFileSync } from "node:fs";

import { unpackZipSync } from "@archive/zip-one-shot";
import { describe, expect, it } from "vitest";

import { Build, Document, Io } from "../index";
import type { Chart, ChartContent, DocxDocument } from "../index";
import { canonicalXml, chartPartDiff, docxWithChartPart, packageText } from "./fixtures/chart-docx";
import type { ChartPackage } from "./fixtures/chart-docx";
import {
  HAND_BUILT_CHART,
  HAND_BUILT_RELATED_PARTS,
  HAND_BUILT_RELS
} from "./fixtures/chart-parts";

const CHART = "word/charts/chart1.xml";
const RELS = "word/charts/_rels/chart1.xml.rels";

const HAND_BUILT: ChartPackage = {
  chartXml: HAND_BUILT_CHART,
  relsXml: HAND_BUILT_RELS,
  parts: HAND_BUILT_RELATED_PARTS
};

/** The chart Excel saved in `pivot-format-dxf.xlsx`, with its style and colour parts. */
function excelSaved(): ChartPackage {
  const xlsx = unpackZipSync(
    readFileSync(new URL("../../excel/__tests__/data/pivot-format-dxf.xlsx", import.meta.url))
  );
  const text = (path: string) => new TextDecoder().decode(xlsx.get(path)!.data);
  return {
    chartXml: text("xl/charts/chart1.xml"),
    relsXml: text("xl/charts/_rels/chart1.xml.rels"),
    parts: new Map([
      [
        "word/charts/style1.xml",
        {
          data: text("xl/charts/style1.xml"),
          contentType: "application/vnd.ms-office.chartstyle+xml"
        }
      ],
      [
        "word/charts/colors1.xml",
        {
          data: text("xl/charts/colors1.xml"),
          contentType: "application/vnd.ms-office.chartcolorstyle+xml"
        }
      ]
    ])
  };
}

function chartOf(doc: DocxDocument): ChartContent {
  return doc.body.find(b => b.type === "chart") as ChartContent;
}

function editChart(doc: DocxDocument, edit: (chart: Chart) => Chart): void {
  const content = chartOf(doc);
  (doc.body as unknown[])[doc.body.indexOf(content)] = { ...content, chart: edit(content.chart) };
}

async function written(doc: DocxDocument): Promise<string> {
  return (await packageText(doc)).get(CHART)!;
}

describe("a read chart part is written back whole", () => {
  for (const [name, pkg] of [
    ["hand-built, every untyped element", () => HAND_BUILT],
    ["saved by Excel", excelSaved]
  ] as const) {
    it(`${name}: the same element tree, and its relationships`, async () => {
      const source = pkg();
      const doc = await Io.read(await docxWithChartPart(source));
      const out = await packageText(doc);
      expect(chartPartDiff(source.chartXml, out.get(CHART)!)).toEqual({ lost: [], added: [] });
      expect(canonicalXml(out.get(CHART)!)).toBe(canonicalXml(source.chartXml));
      expect(canonicalXml(out.get(RELS)!)).toBe(canonicalXml(source.relsXml!));
      for (const path of source.parts!.keys()) {
        expect(out.has(path), path).toBe(true);
      }
    });
  }

  it("is stable: a second round trip writes the same bytes", async () => {
    const first = await written(await Io.read(await docxWithChartPart(HAND_BUILT)));
    const again = await written(
      await Io.read(await docxWithChartPart({ ...HAND_BUILT, chartXml: first }))
    );
    expect(again).toBe(first);
  });

  it("reads what it can type through the fields", async () => {
    const chart = chartOf(await Io.read(await docxWithChartPart(HAND_BUILT))).chart;
    expect(chart.type).toBe("bar");
    expect(chart.style).toBe(102);
    expect(chart.title).toBe("Umsatz und Marge");
    expect(chart.view3d).toBe(true);
    expect(chart.legend).toBe("t");
    expect(chart.series.map(s => s.name)).toEqual(["Umsatz", "Marge"]);
    expect(chart.series[1]).toMatchObject({ chartType: "lineMarked", plotOnSecondaryAxis: true });
    expect(chart.categoryAxis).toMatchObject({ title: "Region", numberFormat: "@" });
    expect(chart.valueAxis).toMatchObject({ min: 1, max: 1000, majorUnit: 10 });
    expect(chart.secondaryValueAxis).toMatchObject({ max: 0.5, numberFormat: "0%" });
  });
});

describe("a field that is set rewrites only what it governs", () => {
  async function editAndDiff(edit: (chart: Chart) => Chart) {
    const doc = await Io.read(await docxWithChartPart(HAND_BUILT));
    editChart(doc, edit);
    return chartPartDiff(HAND_BUILT_CHART, await written(doc));
  }

  it("series values: the values are written from the field, nothing else moves", async () => {
    const { lost, added } = await editAndDiff(chart => ({
      ...chart,
      series: chart.series.map((s, i) => (i === 0 ? { ...s, values: [1, 2, 3, 4] } : s))
    }));
    expect(lost.every(path => path.includes("/bar3DChart/ser/val"))).toBe(true);
    expect(added.every(path => path.includes("/bar3DChart/ser/val"))).toBe(true);
    expect(added).toContain(
      '/chartSpace/chart/plotArea/bar3DChart/ser/val/numRef/numCache/pt/v "4"'
    );
  });

  it("an axis minimum: only c:min, inside the kept c:scaling", async () => {
    const { lost, added } = await editAndDiff(chart => ({
      ...chart,
      valueAxis: { ...chart.valueAxis, min: 5 }
    }));
    expect(lost).toEqual(["/chartSpace/chart/plotArea/valAx/scaling/min[val=1]"]);
    expect(added).toEqual(["/chartSpace/chart/plotArea/valAx/scaling/min[val=5]"]);
  });

  it("the title text: only c:tx; its layout and formatting stay", async () => {
    const doc = await Io.read(await docxWithChartPart(HAND_BUILT));
    editChart(doc, chart => ({ ...chart, title: "Neu" }));
    const { lost, added } = chartPartDiff(HAND_BUILT_CHART, await written(doc));
    expect(lost.every(path => path.startsWith("/chartSpace/chart/title/tx"))).toBe(true);
    expect(added).toContain('/chartSpace/chart/title/tx/rich/p/r/t "Neu"');
    expect(await written(doc)).toContain("<c:manualLayout>");
  });

  it("hiding a series' data labels removes them; the other series keeps its own", async () => {
    const { lost, added } = await editAndDiff(chart => ({
      ...chart,
      series: chart.series.map((s, i) => (i === 0 ? { ...s, showDataLabels: false } : s))
    }));
    expect(lost.length).toBeGreaterThan(0);
    expect(
      lost.every(path => path.startsWith("/chartSpace/chart/plotArea/bar3DChart/ser/dLbls"))
    ).toBe(true);
    expect(added).toEqual([]);
  });

  it("the legend position: only c:legendPos", async () => {
    const { lost, added } = await editAndDiff(chart => ({ ...chart, legend: "b" }));
    expect(lost).toEqual(["/chartSpace/chart/legend/legendPos[val=t]"]);
    expect(added).toEqual(["/chartSpace/chart/legend/legendPos[val=b]"]);
  });

  it("chart data labels: the flags are written, the kept formatting of the group's labels stays", async () => {
    const { lost, added } = await editAndDiff(chart => ({
      ...chart,
      dataLabels: { showValue: true }
    }));
    expect(lost).toEqual(["/chartSpace/chart/plotArea/bar3DChart/dLbls/showVal[val=0]"]);
    expect(added).toContain("/chartSpace/chart/plotArea/bar3DChart/dLbls/showVal[val=1]");
    // The chart's labels apply to every group: the line group, which had none, gains them.
    expect(added).toContain("/chartSpace/chart/plotArea/lineChart/dLbls/showVal[val=1]");
    expect(added.every(path => path.includes("/dLbls"))).toBe(true);
  });
});

describe("relationships of a read chart", () => {
  it("drops one whose target is not written, and the element that referred to it", async () => {
    const doc = await Io.read(await docxWithChartPart(HAND_BUILT));
    const out = await packageText(doc, { securityPolicy: { preserveOleObjects: false } });
    expect(out.get(RELS)).not.toContain('Id="rId3"');
    expect(out.get(RELS)).toContain('Id="rId4"');
    expect(out.get(CHART)).not.toContain("externalData");
    expect(out.get(CHART)).toContain('<c:userShapes r:id="rId4"/>');
  });

  it("a chart part the model cannot read stays as it was, with its relationships", async () => {
    const ofPie = HAND_BUILT_CHART.replace(/<c:bar3DChart>[\s\S]*<\/c:lineChart>/, "")
      .replace(/<c:catAx>[\s\S]*<\/c:valAx>(?=<c:dTable>)/, "")
      .replace(
        "<c:dTable>",
        '<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:gapWidth val="100"/><c:secondPieSize val="75"/></c:ofPieChart><c:dTable>'
      );
    const doc = await Io.read(await docxWithChartPart({ ...HAND_BUILT, chartXml: ofPie }));
    expect(doc.body.some(b => b.type === "chart")).toBe(false);
    const out = await packageText(doc);
    const part = [...out].find(
      ([path, text]) => path.endsWith(".xml") && text.includes("ofPieChart")
    );
    expect(part).toBeDefined();
    expect(canonicalXml(part![1])).toBe(canonicalXml(ofPie));
  });
});

describe("pointColors is read and written by point index", () => {
  const onePoint = (idx: number) =>
    HAND_BUILT_CHART.replace('<c:dPt><c:idx val="3"/>', `<c:dPt><c:idx val="${idx}"/>`);

  it("a lone c:dPt idx=3 reads as pointColors[3], points 0–2 holes", async () => {
    const series = chartOf(await Io.read(await docxWithChartPart(HAND_BUILT))).chart.series[0]!;
    expect(series.pointColors).toHaveLength(4);
    expect(series.pointColors![3]).toBe("FF8800");
    expect([0, 1, 2].some(i => i in series.pointColors!)).toBe(false);
  });

  it("a c:dPt past the series' points colours nothing, and is kept", async () => {
    const doc = await Io.read(
      await docxWithChartPart({ ...HAND_BUILT, chartXml: onePoint(4000000000) })
    );
    expect(chartOf(doc).chart.series[0]!.pointColors).toBeUndefined();
    expect(await written(doc)).toContain('<c:dPt><c:idx val="4000000000"/>');
  });

  it("writes one c:dPt per coloured point, at its index", async () => {
    const h = Document.create();
    const sparse: string[] = [];
    sparse[2] = "FF0000";
    Document.addContent(
      h,
      Build.chart({
        type: "pie",
        series: [{ name: "s", categories: ["a", "b", "c"], values: [1, 2, 3], pointColors: sparse }]
      })
    );
    Document.addContent(
      h,
      Build.chart({
        type: "pie",
        series: [
          {
            name: "t",
            categories: ["a", "b"],
            values: [1, 2],
            pointColors: [undefined as never, "00FF00"]
          }
        ]
      })
    );
    const out = await packageText(h);
    const points = (xml: string) =>
      [...xml.matchAll(/<c:dPt><c:idx val="(\d+)"\/>/g)].map(m => m[1]);
    expect(points(out.get("word/charts/chart1.xml")!)).toEqual(["2"]);
    expect(points(out.get("word/charts/chart2.xml")!)).toEqual(["1"]);
  });

  it("setting one point's colour rewrites only that c:dPt", async () => {
    const doc = await Io.read(await docxWithChartPart(HAND_BUILT));
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map((s, i) => {
        if (i !== 0) {
          return s;
        }
        const pointColors = [...s.pointColors!];
        pointColors[1] = "112233";
        return { ...s, pointColors };
      })
    }));
    const { lost, added } = chartPartDiff(HAND_BUILT_CHART, await written(doc));
    expect(lost).toEqual([]);
    expect(added).toContain("/chartSpace/chart/plotArea/bar3DChart/ser/dPt/idx[val=1]");
    expect(added).toContain(
      "/chartSpace/chart/plotArea/bar3DChart/ser/dPt/spPr/solidFill/srgbClr[val=112233]"
    );
  });
});
