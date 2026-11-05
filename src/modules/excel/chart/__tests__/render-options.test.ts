/**
 * Chart render options are checked at every public entry (R4P-B1 / X4), chart argument errors
 * carry a `field` and name the public call (R4P-S3 / X7, X6), and a chart can be had as a
 * display list (R4P-G3).
 */
import { Draw } from "@draw/index";
import {
  ChartClassic,
  ChartEx,
  ExcelChartOptionsError,
  ExcelError,
  isExcelError
} from "@excel/chart/index";
import { Chart, ExcelArgumentError, Row, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const classic = ChartClassic.createModel({
  type: "bar",
  series: [{ name: "a", values: "Sheet1!$B$1:$B$3", categories: "Sheet1!$A$1:$A$3" }]
});
const modern = ChartEx.createModel({
  type: "funnel",
  categories: "Sheet1!$A$1:$A$3",
  series: [{ values: "Sheet1!$B$1:$B$3" }]
});

async function refusal(run: () => unknown): Promise<ExcelChartOptionsError> {
  try {
    await run();
  } catch (err) {
    expect(err).toBeInstanceOf(ExcelChartOptionsError);
    return err as ExcelChartOptionsError;
  }
  throw new Error("expected a ExcelChartOptionsError");
}

function placedChart() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.add(ws, ["a", 1]);
  Row.add(ws, ["b", 2]);
  return Chart.place(ws, "D1:H10", {
    type: "bar",
    series: [{ values: "S!$B$1:$B$2", categories: "S!$A$1:$A$2" }]
  });
}

const INVALID: Array<[Record<string, unknown>, string]> = [
  [{ width: Number.NaN }, "options.width"],
  [{ width: -5 }, "options.width"],
  [{ height: 0 }, "options.height"],
  [{ height: Number.POSITIVE_INFINITY }, "options.height"],
  [{ scale: -1 }, "options.scale"],
  [{ dpi: 0 }, "options.dpi"],
  [{ title: 42 }, "options.title"],
  [{ backgroundColor: 42 }, "options.backgroundColor"],
  [{ backgroundColor: "nonsense" }, "options.backgroundColor"],
  [{ regionMap: 42 }, "options.regionMap"],
  [{ regionMap: { topology: { objects: {} }, objectName: "" } }, "options.regionMap.objectName"],
  [{ regionMap: { topology: {}, objectName: "x" } }, "options.regionMap.topology.objects"],
  [
    {
      regionMap: {
        topology: { arcs: [], objects: { x: { type: "GeometryCollection", geometries: [] } } },
        objectName: "x",
        match: "name"
      }
    },
    "options.regionMap.match"
  ]
];

describe("chart render options are checked at the entry (R4P-B1, X4)", () => {
  const entries: Array<[string, (options: unknown) => unknown]> = [
    ["ChartClassic.toSvg", options => ChartClassic.toSvg(classic, options as never)],
    ["ChartClassic.toPng", options => ChartClassic.toPng(classic, options as never)],
    ["ChartClassic.toDrawList", options => ChartClassic.toDrawList(classic, options as never)],
    [
      "ChartClassic.Backend.toScene",
      options => ChartClassic.Backend.toScene(classic, options as never)
    ],
    ["ChartEx.toSvg", options => ChartEx.toSvg(modern, options as never)],
    ["ChartEx.toPng", options => ChartEx.toPng(modern, options as never)],
    ["ChartEx.toDrawList", options => ChartEx.toDrawList(modern, options as never)],
    ["Chart.toSvg", options => Chart.toSvg(placedChart(), options as never)],
    ["Chart.toPng", options => Chart.toPng(placedChart(), options as never)]
  ];

  it.each(entries)("%s refuses an option with no valid reading", async (where, run) => {
    for (const [options, field] of INVALID) {
      const err = await refusal(() => run(options));
      expect(err.field).toBe(field);
      expect(err.message.startsWith(`${where}: ${field} must `)).toBe(true);
    }
    const notAnObject = await refusal(() => run(42));
    expect(notAnObject.field).toBe("options");
  });

  it("writes no NaN or negative size into the SVG it used to write them into", () => {
    // `width="NaN"`, `width="-5"` and a 1-px-wide PNG were all produced before.
    const svg = ChartEx.toSvg(modern, { width: 300, height: 200 });
    expect(svg).toMatch(/^<svg [^>]*width="300" height="200"/);
    expect(ChartClassic.toSvg(classic, { backgroundColor: "transparent" })).toContain("<svg");
  });

  it("names the model when it is not one", async () => {
    expect((await refusal(() => ChartClassic.toSvg(modern as never))).field).toBe("model");
    const err = await refusal(() => ChartEx.toPng(classic as never));
    expect(err.field).toBe("model");
    expect(err.message.startsWith("ChartEx.toPng: model must be")).toBe(true);
  });
});

describe("ExcelChartOptionsError carries a field and names the public call (R4P-S3, X7, X6)", () => {
  it("is an ExcelArgumentError, so every existing catch still answers", () => {
    const err = new ExcelChartOptionsError("options.type", "bad");
    expect(err).toBeInstanceOf(ExcelArgumentError);
    expect(err).toBeInstanceOf(ExcelError);
    expect(isExcelError(err)).toBe(true);
    expect(err.field).toBe("options.type");
    expect(err.name).toBe("ExcelChartOptionsError");
    expect(err.toJSON()).toMatchObject({ field: "options.type", message: "bad" });
    // The published two-argument form still compiles and runs; it has no field to give.
    const legacy = new ExcelChartOptionsError("message only", { cause: 1 });
    expect(legacy.message).toBe("message only");
    expect(legacy.field).toBe("");
    expect(legacy.cause).toBe(1);
  });

  it("names the field and the call for the builders", async () => {
    for (const [run, where, field] of [
      [
        () => ChartClassic.createModel({ type: "pie", series: [] } as never),
        "ChartClassic.createModel",
        "options.series"
      ],
      [
        () =>
          ChartClassic.createModel({
            type: "pie",
            series: [{ values: "S!$A$1", trendline: { type: "linear" } }]
          } as never),
        "ChartClassic.createModel",
        "options.series[0].trendline"
      ],
      [
        () => ChartEx.createModel({ type: "funnel", series: [] } as never),
        "ChartEx.createModel",
        "options.series"
      ],
      [
        () => ChartClassic.applyPreset("nope" as never, {} as never),
        "ChartClassic.applyPreset",
        "preset"
      ],
      [() => ChartEx.applyPreset("nope" as never, {} as never), "ChartEx.applyPreset", "preset"]
    ] as const) {
      const err = await refusal(run);
      expect(err.field).toBe(field);
      expect(err.message.startsWith(`${where}: `)).toBe(true);
    }
  });

  it("replaces an internal name at the start of a builder's message with the public call", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    let error: unknown;
    try {
      ChartClassic.createOptionsFromRows(ws, [], null as never);
    } catch (err) {
      error = err;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect(
      (error as Error).message.startsWith("ChartClassic.createOptionsFromRows: options ")
    ).toBe(true);
  });
});

describe("a chart as a display list (R4P-G3)", () => {
  it("is the picture toSvg serialises, and Draw draws it with the caller's fonts", () => {
    const list = ChartClassic.toDrawList(classic, { width: 320, height: 200 });
    expect([list.width, list.height]).toEqual([320, 200]);
    expect(Draw.toSvg(list)).toContain("<svg");
    const image = Draw.toRgba(list, { useSystemFonts: false });
    expect([image.width, image.height]).toEqual([320, 200]);
    const modernList = ChartEx.toDrawList(modern, { width: 300, height: 180 });
    expect([modernList.width, modernList.height]).toEqual([300, 180]);
    expect(Draw.toRgba(modernList).width).toBe(300);
  });
});
