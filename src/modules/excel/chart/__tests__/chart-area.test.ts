/**
 * The chart area — the chart's own fill and border — is part of the picture (r6, the
 * "grey frame" decision). It comes from the model (`c:chartSpace/c:spPr`, `cx:chartSpace/cx:spPr`)
 * and is in the display list, so SVG, PNG and PDF all show the same thing; with no `spPr`
 * it is what Office draws for a chart area left to its automatic format.
 *
 * Office's automatic chart area, as Excel writes it out explicitly
 * (`src/modules/excel/__tests__/data/*.xlsx`) and as the library's own default chart style
 * (`chart-sidecar.ts`, `cs:chartArea`) and LibreOffice's MSO-compatible importer
 * (`oox/source/drawingml/chart/objectformatter.cxx`) all give it: a `bg1` fill (`lt1` for
 * styles 33–40, `dk1` for 41–48) and a 0.75 pt `tx1` 15 % / 85 % luminance line, which is
 * `#D9D9D9` in the default theme. `c:roundedCorners` absent means rounded corners in Excel;
 * the library writes `val="0"`.
 */
import { Draw } from "@draw/index";
import type { DrawList, DrawNode } from "@draw/index";
import { ChartClassic, ChartEx } from "@excel/chart/index";
import type { ChartModel } from "@excel/chart/model/types";
import { drawChartExPdf } from "@excel/chart/render/chart-ex-renderer";
import { drawChartPdf } from "@excel/chart/render/chart-renderer";
import { describe, expect, it } from "vitest";

const D9 = 0xd9 / 255;

function classic(extra: Partial<ChartModel> = {}): ChartModel {
  return {
    ...ChartClassic.createModel({
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
    }),
    ...extra
  };
}

function ex(spPr?: object) {
  const model = ChartEx.createModel({
    type: "treemap",
    series: [{ values: "Sheet1!$B$1:$B$2", literalValues: [10, 20], literalCategories: ["A", "B"] }]
  } as never);
  if (spPr) {
    model.chartSpace.spPr = spPr as never;
  }
  return model;
}

/** The chart-area node: the first node, covering the whole list. */
function frame(list: DrawList): Extract<DrawNode, { kind: "rect" }> {
  const [first] = list.children;
  expect(first?.kind).toBe("rect");
  return first as Extract<DrawNode, { kind: "rect" }>;
}

function close(a: { r: number; g: number; b: number; a: number } | undefined, rgb: number[]) {
  expect(a).toBeDefined();
  expect(a!.r).toBeCloseTo(rgb[0], 2);
  expect(a!.g).toBeCloseTo(rgb[1], 2);
  expect(a!.b).toBeCloseTo(rgb[2], 2);
}

describe("the chart area with no spPr is Office's automatic one", () => {
  it.each([
    ["ChartClassic", () => ChartClassic.toDrawList(classic(), { width: 300, height: 200 })],
    ["ChartEx", () => ChartEx.toDrawList(ex(), { width: 300, height: 200 })]
  ])("%s: white fill, 0.75 pt #D9D9D9 border inside the box, square corners", (_, list) => {
    const node = frame(list());
    close(node.paint.fill, [1, 1, 1]);
    close(node.paint.stroke, [D9, D9, D9]);
    expect(node.paint.strokeWidth).toBe(0.75);
    // The border sits inside the chart's box, as Office draws it, not half outside.
    expect(node).toMatchObject({ x: 0.375, y: 0.375, width: 299.25, height: 199.25 });
    expect(node.rx ?? 0).toBe(0);
  });

  it("a dark chart style (41–48) has a dark automatic fill", () => {
    close(frame(ChartClassic.toDrawList(classic({ style: 42 }))).paint.fill, [0, 0, 0]);
    close(frame(ChartClassic.toDrawList(classic({ style: 34 }))).paint.fill, [1, 1, 1]);
  });

  it("c:roundedCorners absent means rounded corners, as Excel shows them; val=0 square", () => {
    expect(
      frame(ChartClassic.toDrawList(classic({ roundedCorners: undefined }))).rx
    ).toBeGreaterThan(0);
    expect(frame(ChartClassic.toDrawList(classic({ roundedCorners: true }))).rx).toBeGreaterThan(0);
    expect(frame(ChartClassic.toDrawList(classic({ roundedCorners: false }))).rx ?? 0).toBe(0);
    // The builder writes `<c:roundedCorners val="0"/>`, so a built chart is square.
    expect(classic().roundedCorners).toBe(false);
  });
});

describe("the chart area follows the model's spPr", () => {
  it("solid fill and line, with their width", () => {
    const node = frame(
      ChartClassic.toDrawList(
        classic({
          spPr: {
            fill: { solid: { srgb: "FFEECC" } },
            line: { color: { srgb: "336699" }, width: 25400 }
          }
        }),
        { width: 300, height: 200 }
      )
    );
    close(node.paint.fill, [1, 0xee / 255, 0xcc / 255]);
    close(node.paint.stroke, [0x33 / 255, 0x66 / 255, 0x99 / 255]);
    expect(node.paint.strokeWidth).toBe(2);
    expect(node).toMatchObject({ x: 1, y: 1, width: 298, height: 198 });
  });

  it("noFill and ln noFill draw neither", () => {
    const covering = (list: DrawList) =>
      list.children.filter(
        n => n.kind === "rect" && n.width >= list.width - 2 && n.height >= list.height - 2
      );
    const none = { fill: { noFill: true }, line: { noFill: true } };
    expect(covering(ChartClassic.toDrawList(classic({ spPr: none })))).toEqual([]);
    expect(covering(ChartEx.toDrawList(ex(none)))).toEqual([]);
  });

  it("a line without a colour keeps the automatic colour; a fill-only spPr keeps the automatic line", () => {
    const node = frame(
      ChartClassic.toDrawList(classic({ spPr: { fill: { solid: { srgb: "000080" } } } }))
    );
    close(node.paint.fill, [0, 0, 0x80 / 255]);
    close(node.paint.stroke, [D9, D9, D9]);
  });

  it("applies the lumMod / lumOff Excel writes for its automatic border", () => {
    const node = frame(
      ChartClassic.toDrawList(
        classic({
          spPr: {
            fill: { solid: { theme: 1 } },
            line: { width: 9525, color: { theme: 0, lumMod: 15000, lumOff: 85000 } }
          }
        })
      )
    );
    close(node.paint.fill, [1, 1, 1]);
    close(node.paint.stroke, [D9, D9, D9]);
  });

  it("an explicit background option replaces the fill and keeps the border", () => {
    const node = frame(ChartClassic.toDrawList(classic(), { background: "#ff0000" }));
    close(node.paint.fill, [1, 0, 0]);
    close(node.paint.stroke, [D9, D9, D9]);
    const clear = frame(ChartEx.toDrawList(ex(), { background: "transparent" }));
    expect(clear.paint.fill).toBeUndefined();
    close(clear.paint.stroke, [D9, D9, D9]);
  });
});

describe("one picture on every route", () => {
  function recorder() {
    const calls: unknown[] = [];
    const surface: Record<string, unknown> = {};
    for (const method of ["drawRect", "drawLine", "drawText", "drawPath", "drawCircle"]) {
      surface[method] = (...args: unknown[]) => {
        calls.push([method, ...args]);
        return surface;
      };
    }
    return { surface: surface as never, calls };
  }

  it("the PDF bridges draw the display list's chart area, not an invented grey frame", () => {
    for (const draw of [
      (s: never) => drawChartPdf(s, classic(), { x: 10, y: 20, width: 300, height: 200 }),
      (s: never) => drawChartExPdf(s, ex(), { x: 10, y: 20, width: 300, height: 200 })
    ]) {
      const { surface, calls } = recorder();
      draw(surface);
      const [method, first] = calls[0] as [string, Record<string, unknown>];
      expect(method).toBe("drawRect");
      expect(first.stroke).toMatchObject({ r: expect.closeTo(D9, 2) });
      const grey = calls.filter(
        c => (c as [string, { stroke?: { r: number } }])[1]?.stroke?.r === 0.8
      );
      expect(grey).toEqual([]);
      // Exactly one chart-area rect, not a frame plus a background.
      const full = calls.filter(
        c =>
          (c as [string, { width?: number }])[0] === "drawRect" &&
          ((c as [string, { width: number }])[1].width ?? 0) > 290
      );
      expect(full).toHaveLength(1);
    }
  });

  it("SVG and PNG show the same border", () => {
    const list = ChartClassic.toDrawList(classic(), { width: 120, height: 80 });
    expect(ChartClassic.toSvg(classic(), { width: 120, height: 80 })).toMatch(/stroke="#d9d9d9"/i);
    const image = Draw.toRgba(list, { useSystemFonts: false });
    const at = (x: number, y: number) => image.data[(y * image.width + x) * 4];
    // The border's left column is light grey; the interior is white.
    expect(at(0, 40)).toBeLessThan(250);
    expect(at(0, 40)).toBeGreaterThan(200);
    expect(at(5, 40)).toBe(255);
  });
});
