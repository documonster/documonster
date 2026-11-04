/**
 * A Word chart part keeps, on read, the formatting its model cannot hold — a
 * series' or point's whole `c:spPr` (scheme colours, every DrawingML colour
 * transform, gradients, line styles, effects), the plot and chart area's
 * `c:spPr`, the chart's `c:txPr`, the axes' gridlines, `c:spPr` and `c:txPr`
 * and the legend's — and writes it back unchanged. The typed field an element
 * belongs to (`color`, `pointColors`, `plotAreaColor`, `chartAreaColor`) stays
 * authoritative: once it is set to anything other than what was read, the
 * element is written from the field.
 */

import { packZipSync, unpackZipSync } from "@archive/zip-one-shot";
import { describe, expect, it } from "vitest";

import { WordExcel } from "../excel";
import { Build, Document, Io } from "../index";
import type { Chart, ChartContent, DocxDocument } from "../index";

/** Every colour transform of ECMA-376 §20.1.2.3, in one scheme colour. */
const ALL_TRANSFORMS =
  '<a:schemeClr val="accent1"><a:tint val="80000"/><a:shade val="90000"/><a:comp/><a:inv/>' +
  '<a:gray/><a:alpha val="50000"/><a:alphaOff val="1000"/><a:alphaMod val="90000"/>' +
  '<a:hue val="60000"/><a:hueOff val="1000"/><a:hueMod val="90000"/><a:sat val="50000"/>' +
  '<a:satOff val="10000"/><a:satMod val="120000"/><a:lum val="50000"/><a:lumOff val="25000"/>' +
  '<a:lumMod val="75000"/><a:red val="10000"/><a:redOff val="1000"/><a:redMod val="80000"/>' +
  '<a:green val="10000"/><a:greenOff val="1000"/><a:greenMod val="80000"/><a:blue val="10000"/>' +
  '<a:blueOff val="1000"/><a:blueMod val="80000"/><a:gamma/><a:invGamma/></a:schemeClr>';

const S1_SPPR =
  `<c:spPr><a:solidFill>${ALL_TRANSFORMS}</a:solidFill>` +
  '<a:ln w="12700" cap="rnd"><a:solidFill><a:srgbClr val="00FF00"><a:satMod val="120000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:ln>' +
  '<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>';
const S1_DPT_2 =
  '<c:dPt><c:idx val="2"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"><a:lumMod val="60000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></c:spPr></c:dPt>';
const S1_DPT_3 =
  '<c:dPt><c:idx val="3"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"><a:hueOff val="1000"/></a:srgbClr></a:solidFill></c:spPr></c:dPt>';
const S2_SPPR =
  '<c:spPr><a:solidFill><a:srgbClr val="C00000"><a:lumMod val="50000"/></a:srgbClr></a:solidFill></c:spPr>';
const GRIDLINES =
  '<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
const AXIS_SPPR =
  '<c:spPr><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr>';
const AXIS_TXPR =
  '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
const PLOT_SPPR =
  '<c:spPr><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr></a:solidFill></c:spPr>';
const LEGEND_TXPR =
  '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
const CHART_SPPR =
  '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="90000"/></a:srgbClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>';
const CHART_TXPR =
  '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';

function literal(tag: "strLit" | "numLit", values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:${tag}><c:ptCount val="${values.length}"/>${pts}</c:${tag}>`;
}

const KEPT = {
  S1_SPPR,
  S1_DPT_2,
  S1_DPT_3,
  S2_SPPR,
  GRIDLINES,
  AXIS_SPPR,
  AXIS_TXPR,
  PLOT_SPPR,
  LEGEND_TXPR,
  CHART_SPPR,
  CHART_TXPR
};

const CHART_PART =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/>' +
  '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx>${S1_SPPR}<c:invertIfNegative val="0"/>${S1_DPT_2}${S1_DPT_3}` +
  `<c:cat>${literal("strLit", ["a", "b", "c", "d"])}</c:cat><c:val>${literal("numLit", ["1", "2", "3", "4"])}</c:val></c:ser>` +
  `<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>S2</c:v></c:tx>${S2_SPPR}` +
  `<c:cat>${literal("strLit", ["a", "b", "c", "d"])}</c:cat><c:val>${literal("numLit", ["5", "6", "7", "8"])}</c:val></c:ser>` +
  '<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
  `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${GRIDLINES}` +
  `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${AXIS_SPPR}${AXIS_TXPR}` +
  '<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
  '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/>' +
  '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>' +
  `${PLOT_SPPR}</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/>${LEGEND_TXPR}</c:legend>` +
  `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${CHART_SPPR}${CHART_TXPR}</c:chartSpace>`;

/** A document whose chart part is `chartXml`. */
async function docxWithChart(chartXml: string): Promise<Uint8Array> {
  const h = Document.create();
  Document.addContent(
    h,
    Build.chart({ type: "column", series: [{ name: "x", categories: ["a"], values: [1] }] })
  );
  const files = unpackZipSync(await Io.toBuffer(h));
  const entries = new Map<string, Uint8Array | string>();
  for (const [path, entry] of files) {
    entries.set(path, path === "word/charts/chart1.xml" ? chartXml : entry.data);
  }
  return packZipSync(entries);
}

async function chartPartOf(doc: DocxDocument, options?: Parameters<typeof Io.toBuffer>[1]) {
  const files = unpackZipSync(await Io.toBuffer(doc, options));
  return new TextDecoder().decode(files.get("word/charts/chart1.xml")!.data);
}

function chartContent(doc: DocxDocument): ChartContent {
  return doc.body.find(b => b.type === "chart") as ChartContent;
}

/** Replace the document's chart with `edit(chart)`. */
function editChart(doc: DocxDocument, edit: (chart: Chart) => Chart): void {
  const content = chartContent(doc);
  const index = doc.body.indexOf(content);
  (doc.body as unknown[])[index] = { ...content, chart: edit(content.chart) };
}

async function readChartDoc(): Promise<DocxDocument> {
  return Io.read(await docxWithChart(CHART_PART));
}

describe("a Word chart part keeps the formatting its model cannot hold", () => {
  it("writes every kept element back byte for byte", async () => {
    const written = await chartPartOf(await readChartDoc());
    for (const [name, element] of Object.entries(KEPT)) {
      expect(written, name).toContain(element);
    }
  });

  it("keeps every colour transform and the scheme colour, in order", async () => {
    const written = await chartPartOf(await readChartDoc());
    expect(written).toContain(ALL_TRANSFORMS);
  });

  it("puts each kept element at its schema position", async () => {
    const written = await chartPartOf(await readChartDoc());
    const order = (...parts: string[]) => {
      const at = parts.map(p => written.indexOf(p));
      expect(at.every(i => i >= 0)).toBe(true);
      expect([...at].sort((a, b) => a - b)).toEqual(at);
    };
    // c:ser: tx, spPr, dPt, cat
    order("<c:v>S1</c:v>", S1_SPPR, S1_DPT_2, S1_DPT_3, "<c:cat>");
    // c:catAx: axPos, majorGridlines, numFmt, tickLblPos, spPr, txPr, crossAx
    order('<c:axPos val="b"/>', GRIDLINES, "<c:numFmt", "<c:tickLblPos", AXIS_SPPR, AXIS_TXPR);
    // c:plotArea: last axis, spPr, close
    order("</c:valAx>", PLOT_SPPR, "</c:plotArea>");
    // c:legend: overlay, txPr
    order("<c:overlay", LEGEND_TXPR, "</c:legend>");
    // c:chartSpace: chart, spPr, txPr
    order("</c:chart>", CHART_SPPR, CHART_TXPR, "</c:chartSpace>");
  });

  it("is stable: a second round trip writes the same part", async () => {
    const first = await chartPartOf(await readChartDoc());
    const again = await chartPartOf(await Io.read(await docxWithChart(first)));
    expect(again).toBe(first);
  });

  it("a series colour that is set takes over its spPr", async () => {
    const doc = await readChartDoc();
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map((s, i) => (i === 0 ? { ...s, color: "FF0000" } : s))
    }));
    const written = await chartPartOf(doc);
    expect(written).not.toContain(S1_SPPR);
    expect(written).not.toContain(ALL_TRANSFORMS);
    expect(written).toContain(
      '<c:tx><c:v>S1</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>'
    );
    // The other series and the rest are untouched.
    expect(written).toContain(S2_SPPR);
    expect(written).toContain(S1_DPT_2);
    expect(written).toContain(PLOT_SPPR);
  });

  it("a series colour read as an sRGB value is kept while it still equals it", async () => {
    const doc = await readChartDoc();
    expect(chartContent(doc).chart.series[1]!.color).toBe("C00000");
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map(s => ({ ...s, values: s.values.map(v => v * 2) }))
    }));
    expect(await chartPartOf(doc)).toContain(S2_SPPR);
  });

  it("removing a series colour removes its spPr", async () => {
    const doc = await readChartDoc();
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map((s, i) => {
        if (i !== 1) {
          return s;
        }
        const { color: _color, ...rest } = s;
        return rest;
      })
    }));
    const written = await chartPartOf(doc);
    expect(written).not.toContain(S2_SPPR);
    expect(written).not.toContain('val="C00000"');
  });

  it("point colours that are set take over their points, point by point", async () => {
    const doc = await readChartDoc();
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map((s, i) => (i === 0 ? { ...s, pointColors: ["00FF00"] } : s))
    }));
    const written = await chartPartOf(doc);
    // Point 0 is coloured from the field; point 3's colour was removed; point 2
    // (a gradient, no colour read) still has no colour in the field, so stays.
    expect(written).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>'
    );
    expect(written).not.toContain(S1_DPT_3);
    expect(written).toContain(S1_DPT_2);
    expect(written).toContain(S1_SPPR);
  });

  it("plot and chart area colours that are set take over their spPr", async () => {
    const doc = await readChartDoc();
    editChart(doc, chart => ({ ...chart, plotAreaColor: "EEEEEE", chartAreaColor: "DDDDDD" }));
    const written = await chartPartOf(doc);
    expect(written).not.toContain(PLOT_SPPR);
    expect(written).not.toContain(CHART_SPPR);
    expect(written).toContain(
      '<c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr></c:plotArea>'
    );
    expect(written).toContain(
      '</c:chart><c:spPr><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill></c:spPr>'
    );
    // Formatting no field governs stays.
    expect(written).toContain(CHART_TXPR);
    expect(written).toContain(GRIDLINES);
  });

  it("writes nothing it kept under rawXmlPolicy 'strip'", async () => {
    const written = await chartPartOf(await readChartDoc(), {
      securityPolicy: { rawXmlPolicy: "strip" }
    });
    for (const [name, element] of Object.entries(KEPT)) {
      expect(written, name).not.toContain(element);
    }
    // The fields are still written.
    expect(written).toContain('<a:srgbClr val="C00000"/>');
  });

  it("does not keep an element that refers to a relationship of the chart part", async () => {
    const blip =
      '<c:spPr><a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>';
    const doc = await Io.read(await docxWithChart(CHART_PART.replace(S2_SPPR, blip)));
    const written = await chartPartOf(doc);
    expect(written).not.toContain("rId9");
    expect(written).toContain(S1_SPPR);
  });
});

describe("a read Word chart is drawn with the formatting it kept", () => {
  it("hands the kept spPr, points and areas to the Excel chart model", async () => {
    const chart = chartContent(await readChartDoc()).chart;
    const model = WordExcel.toChartModel(chart);
    const group = model.chart.plotArea.chartTypes[0]!;
    const [s1, s2] = group.series as unknown as {
      spPr?: { _rawXml?: string };
      dataPoints?: { index: number; spPr?: { _rawXml?: string } }[];
    }[];
    expect(s1!.spPr?._rawXml).toBe(S1_SPPR);
    expect(s1!.dataPoints?.map(p => p.index)).toEqual([2, 3]);
    expect(s1!.dataPoints?.[1]?.spPr?._rawXml).toBe(
      '<c:spPr><a:solidFill><a:srgbClr val="ABCDEF"><a:hueOff val="1000"/></a:srgbClr></a:solidFill></c:spPr>'
    );
    expect(s2!.spPr?._rawXml).toBe(S2_SPPR);
    expect(model.chart.plotArea.spPr?._rawXml).toBe(PLOT_SPPR);
    expect(model.spPr?._rawXml).toBe(CHART_SPPR);
    expect(model.txPr?._rawXml).toBe(CHART_TXPR);
    expect(model.chart.plotArea.axes[0]!.majorGridlines?._rawXml).toBe(
      GRIDLINES.replace(/^<c:majorGridlines>|<\/c:majorGridlines>$/g, "")
    );
  });

  it("draws a kept colour with its transforms, and a set colour as set", async () => {
    const doc = await readChartDoc();
    // C00000 with lumMod 50% is #600000.
    expect(WordExcel.toChartSvg(chartContent(doc).chart).toUpperCase()).toContain("#600000");
    editChart(doc, chart => ({
      ...chart,
      series: chart.series.map((s, i) => (i === 1 ? { ...s, color: "FF0000" } : s))
    }));
    const svg = WordExcel.toChartSvg(chartContent(doc).chart).toUpperCase();
    expect(svg).not.toContain("#600000");
    expect(svg).toContain("#FF0000");
  });

  it("draws a series colour set in code", () => {
    const chart = Build.chart({
      type: "column",
      series: [{ name: "x", categories: ["a"], values: [1], color: "12AB34" }]
    }).chart;
    expect(WordExcel.toChartSvg(chart).toUpperCase()).toContain("#12AB34");
  });
});
