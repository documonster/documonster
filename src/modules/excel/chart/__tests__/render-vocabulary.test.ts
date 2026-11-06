/**
 * The chart renderers speak the drawing engine's vocabulary (r6 D3/D4):
 *
 * - `background` takes what `Draw.toSvg` / `Draw.toRgba` take — a CSS colour or `{ r, g, b, a }`
 *   — and the deprecated `backgroundColor` forwards to it;
 * - `toPng` takes `fonts` / `useSystemFonts` with Draw's meaning, so a chart PNG can be pinned
 *   to the caller's fonts;
 * - one route puts a chart on a PDF page: `toDrawList` + `Pdf.createDrawSurface` +
 *   `Draw.render`, as for every other producer. `ChartClassic.drawPdf`, `ChartEx.drawPdf` and
 *   `ChartEx.canDrawVectorPdf` are gone from the surface; the published flat forwards stay.
 */
import { encodePng } from "@archive/png";
import { Draw } from "@draw/index";
import * as ChartEntry from "@excel/chart/index";
import { ChartClassic, ChartEx, ExcelChartOptionsError } from "@excel/chart/index";
import { drawChartExPdf } from "@excel/chart/render/chart-ex-renderer";
import { drawChartPdf } from "@excel/chart/render/chart-renderer";
import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

const classic = ChartClassic.createModel({
  type: "bar",
  title: "Revenue",
  series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
});

const ex = ChartEx.createModel({
  type: "treemap",
  series: [{ values: "Sheet1!$B$1:$B$2", literalValues: [10, 20], literalCategories: ["A", "B"] }]
} as never);

function refusal(run: () => unknown): ExcelChartOptionsError {
  try {
    run();
  } catch (err) {
    return err as ExcelChartOptionsError;
  }
  throw new Error("expected a refusal");
}

function pixel(list: ReturnType<typeof ChartClassic.toDrawList>, x: number, y: number) {
  const image = Draw.toRgba(list, { useSystemFonts: false });
  const i = (y * image.width + x) * 4;
  return Array.from(image.data.subarray(i, i + 4));
}

describe("ChartRenderOptions.background", () => {
  it.each([
    ["ChartClassic", (o: object) => ChartClassic.toDrawList(classic, o)],
    ["ChartEx", (o: object) => ChartEx.toDrawList(ex, o)]
  ])("%s takes a CSS colour or { r, g, b, a }, as Draw does", (_, toDrawList) => {
    expect(pixel(toDrawList({ background: "#ff0000" }), 1, 1)).toEqual([255, 0, 0, 255]);
    expect(pixel(toDrawList({ background: { r: 0, g: 0, b: 1, a: 1 } }), 1, 1)).toEqual([
      0, 0, 255, 255
    ]);
    expect(pixel(toDrawList({ background: "transparent" }), 1, 1)[3]).toBe(0);
    // The deprecated name forwards; `background` wins when both are given.
    expect(pixel(toDrawList({ backgroundColor: "#00ff00" }), 1, 1)).toEqual([0, 255, 0, 255]);
    expect(pixel(toDrawList({ backgroundColor: "#00ff00", background: "#ff0000" }), 1, 1)).toEqual([
      255, 0, 0, 255
    ]);
  });

  it("an { r, g, b, a } background reaches the SVG, with its alpha", () => {
    for (const svg of [
      ChartClassic.toSvg(classic, { background: { r: 1, g: 0, b: 0, a: 0.5 } }),
      ChartEx.toSvg(ex, { background: { r: 1, g: 0, b: 0, a: 0.5 } })
    ]) {
      expect(svg).toMatch(/fill="(?:#ff0000|rgba\(100%, 0%, 0%, 0\.5\))"/);
    }
  });

  it("refuses a background Draw would refuse, as a ExcelChartOptionsError naming it", () => {
    for (const bad of ["not-a-colour", 42, { r: 2, g: 0, b: 0, a: 1 }]) {
      const err = refusal(() => ChartClassic.toSvg(classic, { background: bad as never }));
      expect(err).toBeInstanceOf(ExcelChartOptionsError);
      expect(err.field).toMatch(/^options\.background/);
      expect(err.message).toMatch(/^ChartClassic\.toSvg: options\.background/);
    }
    expect(refusal(() => ChartEx.toSvg(ex, { backgroundColor: "nope" })).field).toBe(
      "options.backgroundColor"
    );
  });
});

describe("toPng takes fonts and useSystemFonts", () => {
  it("passes them to the rasteriser: with no system fonts the output is the same everywhere", async () => {
    const a = await ChartClassic.toPng(classic, { useSystemFonts: false, fonts: [] });
    const b = await ChartClassic.toPng(classic, { useSystemFonts: false });
    expect(a).toEqual(b);
    // The same pixels the drawing engine gives for the same request: had the option been
    // dropped, `toPng` would have drawn with whatever the host has installed.
    const viaDraw = Draw.toRgba(ChartClassic.toDrawList(classic), { useSystemFonts: false });
    expect(a).toEqual(encodePng(viaDraw.data, viaDraw.width, viaDraw.height));
    await expect(ChartEx.toPng(ex, { useSystemFonts: false })).resolves.toBeInstanceOf(Uint8Array);
  });

  it("checks them", () => {
    expect(refusal(() => ChartClassic.toSvg(classic, { fonts: "x" as never })).field).toBe(
      "options.fonts"
    );
    expect(refusal(() => ChartClassic.toSvg(classic, { fonts: [42] as never })).field).toBe(
      "options.fonts[0]"
    );
    expect(refusal(() => ChartEx.toSvg(ex, { useSystemFonts: "no" as never })).field).toBe(
      "options.useSystemFonts"
    );
  });
});

describe("one route puts a chart on a PDF page", () => {
  it("drawPdf and canDrawVectorPdf are not on the namespaces", () => {
    expect("drawPdf" in ChartClassic).toBe(false);
    expect("drawPdf" in ChartEx).toBe(false);
    expect("canDrawVectorPdf" in ChartEx).toBe(false);
  });

  it("toDrawList + Pdf.createDrawSurface + Draw.render draws the chart, title included", async () => {
    for (const list of [
      ChartClassic.toDrawList(classic, { width: 300, height: 200 }),
      ChartEx.toDrawList(ex, { width: 300, height: 200, title: "Shares" })
    ]) {
      const builder = new Pdf.Builder();
      const page = builder.addPage({ width: 400, height: 300 });
      Draw.render(list, Pdf.createDrawSurface(page, { x: 50, y: 50, width: 300, height: 200 }));
      const text = (await Pdf.extract(await builder.toBuffer())).text;
      expect(text).toMatch(/Revenue|Shares/);
    }
  });

  it("keeps the published flat forwards, with the behaviour they had", () => {
    expect(ChartEntry.drawChartPdf).toBe(drawChartPdf);
    expect(typeof ChartEntry.drawChartExPdf).toBe("function");
    expect(ChartEntry.canRenderChartExAsVectorPdf(ex)).toBe(true);
    const calls: string[] = [];
    const surface = {
      drawRect: (o: { stroke?: unknown }) => {
        calls.push(`rect ${o.stroke ? "framed" : ""}`);
        return surface;
      },
      drawLine: () => surface,
      drawText: (t: string) => {
        calls.push(`text ${t}`);
        return surface;
      }
    };
    ChartEntry.drawChartExPdf(surface as never, ex, { x: 0, y: 0, width: 300, height: 200 });
    expect(calls[0]).toBe("rect framed");
    const internal: string[] = [];
    const recorder = {
      drawRect: (o: { stroke?: unknown }) => {
        internal.push(`rect ${o.stroke ? "framed" : ""}`);
        return recorder;
      },
      drawLine: () => recorder,
      drawText: (t: string) => {
        internal.push(`text ${t}`);
        return recorder;
      }
    };
    drawChartExPdf(recorder as never, ex, { x: 0, y: 0, width: 300, height: 200 });
    expect(calls).toEqual(internal);
    // Its arguments are checked, naming the call as written.
    expect(
      refusal(() => ChartEntry.drawChartExPdf({} as never, ex, { x: 0, y: 0, width: 1, height: 1 }))
        .message
    ).toMatch(/^drawChartExPdf: surface/);
  });
});
