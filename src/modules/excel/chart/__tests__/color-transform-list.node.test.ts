/**
 * A chart colour keeps every DrawingML transform it carries (ECMA-376 §20.1.2.3), in document order.
 *
 * The chart colour model held six — `tint`, `shade`, `satMod`, `lumMod`, `lumOff`, `alpha` — so a colour that went
 * through it lost the rest: `satOff`, `hueMod`, `hueOff`, `alphaMod`, `alphaOff`, `comp`, `inv`, `gray`, `gamma`,
 * `invGamma`, `red` / `green` / `blue` and their `Mod` / `Off`. An unchanged chart part was written back from its raw
 * XML and kept them, but parsing (`DrawingMl.parseSpPr`), any structured edit (`setSpPrLine`, `setSpPrFill`), the
 * colours sidecar (`ChartClassic.fromColorsXml` → `toColorsXml`) and the renderer all went through the six. Now each
 * colour carries `transforms`, the ordered list, and the six fields are views of it.
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { Archive } from "@archive";
import { ChartClassic, DrawingMl } from "@excel/chart/index";
import type { ChartColor, ChartColorTransform } from "@excel/chart/index";
import { resolveChartColor } from "@excel/chart/shared/chart-utils";
import { Cell, Chart, Workbook } from "@excel/index";
import { applyColorTransforms } from "@utils/theme-colors";
import { describe, expect, it } from "vitest";

/** Every transform, each kind once (the six included, out of the writers' order), then a repeat. */
const EVERY: ChartColorTransform[] = [
  { kind: "lumOff", val: 5000 },
  { kind: "satOff", val: 10000 },
  { kind: "hue", val: 1200000 },
  { kind: "hueMod", val: 90000 },
  { kind: "hueOff", val: 600000 },
  { kind: "sat", val: 80000 },
  { kind: "lum", val: 60000 },
  { kind: "alpha", val: 90000 },
  { kind: "alphaMod", val: 50000 },
  { kind: "alphaOff", val: -10000 },
  { kind: "comp" },
  { kind: "inv" },
  { kind: "gray" },
  { kind: "gamma" },
  { kind: "invGamma" },
  { kind: "red", val: 50000 },
  { kind: "redMod", val: 80000 },
  { kind: "redOff", val: 2000 },
  { kind: "green", val: 40000 },
  { kind: "greenMod", val: 70000 },
  { kind: "greenOff", val: 1000 },
  { kind: "blue", val: 30000 },
  { kind: "blueMod", val: 90000 },
  { kind: "blueOff", val: 3000 },
  { kind: "tint", val: 80000 },
  { kind: "shade", val: 95000 },
  { kind: "satMod", val: 120000 },
  { kind: "lumMod", val: 75000 },
  { kind: "lumMod", val: 90000 }
];

function transformsXml(list: readonly ChartColorTransform[]): string {
  return list
    .map(t => (t.val === undefined ? `<a:${t.kind}/>` : `<a:${t.kind} val="${t.val}"/>`))
    .join("");
}

/** The transform children of every colour element in `xml`, in order, as text. */
function colourChildren(xml: string): string[] {
  return [
    ...xml.matchAll(
      /<a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:\1>)/g
    )
  ].map(m => m[2] ?? "");
}

const ALL_XML = transformsXml(EVERY);
const RAW_SPPR = `<c:spPr><a:solidFill><a:schemeClr val="accent1">${ALL_XML}</a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000">${ALL_XML}</a:srgbClr></a:solidFill></a:ln></c:spPr>`;

describe("reading", () => {
  it("DrawingMl.parseSpPr keeps every transform in order, and sets the six named fields from it", () => {
    const parsed = DrawingMl.parseSpPr({ _rawXml: RAW_SPPR } as never);
    const fill = parsed.fill!.solid!;
    expect(fill.transforms).toEqual(EVERY);
    expect(fill).toMatchObject({
      theme: 4,
      lumOff: 5000,
      alpha: 90000,
      tint: 0.8,
      shade: 95000,
      satMod: 120000,
      lumMod: 75000
    });
    expect(parsed.line!.color!.transforms).toEqual(EVERY);
  });

  it("ChartClassic.fromColorsXml keeps them on palette entries and variations", () => {
    const colors = ChartClassic.fromColorsXml(
      `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1">${ALL_XML}</a:schemeClr><cs:variation>${ALL_XML}</cs:variation></cs:colorStyle>`
    );
    expect(colors.colors![0]!.transforms).toEqual(EVERY);
    expect(colors.variations![0]!.transforms).toEqual(EVERY);
  });
});

describe("writing", () => {
  it("a structured edit of a read spPr writes every transform back, in order", async () => {
    const edited = DrawingMl.setSpPrLine({ _rawXml: RAW_SPPR } as never, {
      width: 12700,
      color: { srgb: "00FF00" }
    });
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Chart.place(ws, "C1:H10", {
      type: "bar",
      series: [{ values: "S!$A$1:$A$2", spPr: edited }]
    } as never);
    const xml = new TextDecoder().decode(
      (await Archive.unpackZip(await Workbook.toBuffer(wb))).get("xl/charts/chart1.xml")!.data
    );
    const series = xml.slice(xml.indexOf("<c:ser>"), xml.indexOf("</c:ser>"));
    expect(colourChildren(series)).toContain(ALL_XML);
  });

  it("a colour built with transforms is written with them, in both chart families", async () => {
    const color: ChartColor = { srgb: "4472C4", transforms: EVERY };
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Chart.place(ws, "C1:H10", {
      type: "bar",
      series: [{ values: "S!$A$1:$A$2", spPr: { fill: { solid: color } } }]
    } as never);
    Chart.place(ws, "C12:H22", {
      type: "waterfall",
      series: [{ values: "S!$A$1:$A$2", spPr: { fill: { solid: color } } }]
    } as never);
    const files = await Archive.unpackZip(await Workbook.toBuffer(wb));
    const classic = new TextDecoder().decode(files.get("xl/charts/chart1.xml")!.data);
    const chartEx = new TextDecoder().decode(
      [...files].find(([path]) => /xl\/charts\/chartEx\d+\.xml$/.test(path))![1].data
    );
    expect(colourChildren(classic)).toContain(ALL_XML);
    expect(colourChildren(chartEx)).toContain(ALL_XML);
  });

  it("ChartClassic.toColorsXml writes them back", () => {
    const model = ChartClassic.fromColorsXml(
      `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1">${ALL_XML}</a:schemeClr><cs:variation>${ALL_XML}</cs:variation></cs:colorStyle>`
    );
    const xml = ChartClassic.toColorsXml(model);
    expect(colourChildren(xml)).toEqual([ALL_XML]);
    expect(xml).toContain(`<cs:variation>${ALL_XML}</cs:variation>`);
  });
});

describe("the six named fields are views of the list", () => {
  const base = (): ChartColor => ({
    srgb: "FF0000",
    lumMod: 75000,
    transforms: [{ kind: "lumMod", val: 75000 }, { kind: "comp" }]
  });

  async function written(color: ChartColor): Promise<string> {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Chart.place(ws, "C1:H10", {
      type: "bar",
      series: [{ values: "S!$A$1:$A$1", spPr: { fill: { solid: color } } }]
    } as never);
    const xml = new TextDecoder().decode(
      (await Archive.unpackZip(await Workbook.toBuffer(wb))).get("xl/charts/chart1.xml")!.data
    );
    return colourChildren(xml.slice(xml.indexOf("<c:ser>"), xml.indexOf("</c:ser>")))[0]!;
  }

  it.each<[string, ChartColor, string]>([
    ["unchanged, the list as it is", base(), '<a:lumMod val="75000"/><a:comp/>'],
    [
      "a changed field replaces its first entry in place",
      { ...base(), lumMod: 50000 },
      '<a:lumMod val="50000"/><a:comp/>'
    ],
    [
      "an absent field leaves the list's entries",
      { ...base(), lumMod: undefined },
      '<a:lumMod val="75000"/><a:comp/>'
    ],
    [
      "a transform removed from the list is gone",
      { srgb: "FF0000", transforms: [{ kind: "comp" }] },
      "<a:comp/>"
    ],
    [
      "a new field is appended",
      { ...base(), tint: 0.5 },
      '<a:lumMod val="75000"/><a:comp/><a:tint val="50000"/>'
    ],
    [
      "without the list, the six fields are written as before",
      { srgb: "FF0000", lumMod: 75000, lumOff: 25000 },
      '<a:lumMod val="75000"/><a:lumOff val="25000"/>'
    ]
  ])("%s", async (_name, color, expected) => {
    expect(await written(color)).toBe(expected);
  });

  it("the renderer reads the same list", () => {
    expect(resolveChartColor({ srgb: "FF0000", transforms: [{ kind: "comp" }] })).toBe("#00FFFF");
    // `lumMod` 50 % of red's 50 % luminance, then the complement: the named field wins over the list's 75 %.
    expect(resolveChartColor({ ...base(), lumMod: 50000 })).toBe(
      resolveChartColor({
        srgb: "FF0000",
        transforms: [{ kind: "lumMod", val: 50000 }, { kind: "comp" }]
      })
    );
  });
});

describe("rendering applies the whole list", () => {
  it("resolveChartColor goes through applyColorTransforms with every transform", () => {
    // A 120° hue turn takes red to green; the six fields have no way to say it.
    expect(
      resolveChartColor({ srgb: "FF0000", transforms: [{ kind: "hueOff", val: 7200000 }] })
    ).toBe("#00FF00");
    expect(resolveChartColor({ srgb: "4472C4", transforms: EVERY })).toBe(
      `#${["r", "g", "b"]
        .map(c =>
          Math.round(
            (
              applyColorTransforms(
                { r: 0x44 / 255, g: 0x72 / 255, b: 0xc4 / 255 },
                EVERY
              ) as unknown as Record<string, number>
            )[c]! * 255
          )
            .toString(16)
            .padStart(2, "0")
        )
        .join("")
        .toUpperCase()}`
    );
  });

  it("an SVG draws the transformed colour", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    const chart = Chart.place(ws, "C1:H10", {
      type: "bar",
      series: [
        {
          values: "S!$A$1:$A$1",
          spPr: {
            fill: { solid: { srgb: "FF0000", transforms: [{ kind: "hueOff", val: 7200000 }] } }
          }
        }
      ]
    } as never);
    expect(Chart.toSvg(chart, { width: 400, height: 300 }).toLowerCase()).toContain(
      'fill="#00ff00"'
    );
  });
});

describe("an Excel-saved chart and its colours sidecar", () => {
  const FIXTURE = resolve(__dirname, "../../__tests__/data/workbook-roundtrip.xlsx");

  it("every colour of chart1.xml reads and writes back to the same transforms, in order", async () => {
    const files = await Archive.unpackZip(readFileSync(FIXTURE));
    const xml = new TextDecoder().decode(files.get("xl/charts/chart1.xml")!.data);
    const blocks = [...xml.matchAll(/<c:spPr>[\s\S]*?<\/c:spPr>/g)].map(m => m[0]);
    expect(blocks.length).toBeGreaterThan(3);
    for (const raw of blocks) {
      // Parsed, then rebuilt as a fresh structured spPr — the path every edit takes.
      const rebuilt = DrawingMl.createSpPr(DrawingMl.parseSpPr({ _rawXml: raw } as never));
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Chart.place(ws, "C1:H10", {
        type: "bar",
        series: [{ values: "S!$A$1:$A$1", spPr: rebuilt }]
      } as never);
      const out = new TextDecoder().decode(
        (await Archive.unpackZip(await Workbook.toBuffer(wb))).get("xl/charts/chart1.xml")!.data
      );
      const series = out.slice(out.indexOf("<c:ser>"), out.indexOf("</c:ser>"));
      for (const children of colourChildren(raw)) {
        expect(colourChildren(series)).toContain(children);
      }
    }
  });

  it("colors1.xml reads and writes back to the same palette and variations", async () => {
    const files = await Archive.unpackZip(readFileSync(FIXTURE));
    const raw = new TextDecoder().decode(files.get("xl/charts/colors1.xml")!.data);
    const written = ChartClassic.toColorsXml(ChartClassic.fromColorsXml(raw));
    expect(colourChildren(written)).toEqual(colourChildren(raw));
    const variations = (text: string) =>
      [...text.matchAll(/<cs:variation(?:\/>|>([\s\S]*?)<\/cs:variation>)/g)].map(m => m[1] ?? "");
    expect(variations(written)).toEqual(variations(raw));
  });
});
