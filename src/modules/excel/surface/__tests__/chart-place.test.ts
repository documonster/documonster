/**
 * `Chart.place(ws, range, options)` is the one way to add a chart; the option builders beside it cover what the
 * deprecated `add*` family's variants did. Each deprecated name must still build the same chart as its replacement.
 */
import { ChartClassic, ChartEx } from "@excel/chart/index";
import { Cell, Chart, ExcelArgumentError, Row, Table, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Row.addMany(ws, [
    ["Month", "Revenue"],
    ["Jan", 1],
    ["Feb", 2],
    ["Mar", 3]
  ]);
  return { wb, ws };
}

const series = [{ name: "Revenue", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }];

/** The model of the only chart on a fresh sheet built by `build`. */
function modelOf(build: (ws: ReturnType<typeof sheet>["ws"]) => unknown) {
  const { ws } = sheet();
  build(ws);
  const [chart] = Chart.list(ws);
  return Chart.isEx(chart!) ? Chart.getExModel(chart!) : Chart.getModel(chart!);
}

describe("Chart.place", () => {
  it("takes the placement before the options and returns the handle", () => {
    const { ws } = sheet();
    const chart = Chart.place(ws, "D1:K16", { type: "line", series });
    expect(Chart.list(ws)).toEqual([chart]);
    expect(Chart.getAnchorModel(chart)).toBeDefined();
  });

  it("refuses a missing chart and a malformed range at the call", () => {
    const { ws } = sheet();
    expect(() => Chart.place(ws, "D1:K16", undefined as never)).toThrow(ExcelArgumentError);
    expect(() =>
      Chart.place(ws, { tl: { col: -1, row: 0 } } as never, { type: "line", series })
    ).toThrow(/range\.tl\.col/);
  });

  it.each([
    ["addColumn", { type: "bar", barDir: "col" }],
    ["addBar", { type: "bar", barDir: "bar" }],
    ["addLine", { type: "line" }],
    ["addPie", { type: "pie" }],
    ["addHistogram", { type: "histogram" }],
    ["addWaterfall", { type: "waterfall" }]
  ] as const)("builds what the deprecated Chart.%s built", (name, type) => {
    const add = Chart[name] as (ws: unknown, options: unknown, range: string) => number;
    expect(modelOf(ws => Chart.place(ws, "D1:K16", { ...type, series } as never))).toEqual(
      modelOf(ws => add(ws, { series }, "D1:K16"))
    );
  });
});

describe("the option builders", () => {
  it("applyPreset picks the classic or the Office 2016 preset by its name", () => {
    expect(Chart.applyPreset("lineMarkers", { series })).toEqual(
      ChartClassic.applyPreset("lineMarkers", { series })
    );
    expect(Chart.applyPreset("boxAndWhisker", { series })).toEqual(
      ChartEx.applyPreset("boxAndWhisker", { series })
    );
    expect(
      modelOf(ws => Chart.place(ws, "D1:K16", Chart.applyPreset("lineMarkers", { series })))
    ).toEqual(modelOf(ws => Chart.addPreset(ws, "lineMarkers", { series }, "D1:K16")));
  });

  it("createOptionsFromTable and createOptionsFromRows pick the family by the type", () => {
    const { ws } = sheet();
    const table = Table.add(ws, {
      name: "Kpi",
      ref: "D1",
      headerRow: true,
      columns: [{ name: "Month" }, { name: "Revenue" }],
      rows: [
        ["Jan", 1],
        ["Feb", 2]
      ]
    });
    const classic = {
      type: "bar" as const,
      barDir: "col" as const,
      categoryColumn: "Month",
      valueColumns: ["Revenue"]
    };
    expect(Chart.createOptionsFromTable(ws, table, classic)).toEqual(
      ChartClassic.createOptionsFromTable(ws, table, classic)
    );
    const funnel = { type: "funnel" as const, categoryColumn: "Month", valueColumns: ["Revenue"] };
    expect(Chart.createOptionsFromTable(ws, table, funnel)).toEqual(
      ChartEx.createOptionsFromTable(ws, table, funnel)
    );

    const rows = [
      { day: "Mon", visits: 3 },
      { day: "Tue", visits: 4 }
    ];
    const bar = { type: "bar", barDir: "col", x: "day", y: "visits", startCell: "H1" } as const;
    expect(Chart.createOptionsFromRows(ws, rows, bar)).toEqual(
      ChartClassic.createOptionsFromRows(ws, rows, bar)
    );
    const histogram = { type: "histogram", x: "day", y: "visits", startCell: "K1" } as const;
    expect(Chart.createOptionsFromRows(ws, rows, histogram)).toEqual(
      ChartEx.createOptionsFromRows(ws, rows, histogram)
    );
  });
});

describe("createOptionsFromRows writes nothing unless it can write everything", () => {
  const snapshot = (ws: ReturnType<typeof sheet>["ws"]) =>
    ["A1", "B1", "A2", "B2", "D1", "E1", "D2", "E2", "D3"].map(a => Cell.getValue(ws, a));

  it.each([
    [
      "a value no cell holds",
      [
        { x: "a", y: 1 },
        { x: "b", y: {} }
      ],
      {}
    ],
    [
      "a number no workbook holds",
      [
        { x: "a", y: 1 },
        { x: "b", y: NaN }
      ],
      {}
    ],
    [
      "an invalid Date",
      [
        { x: "a", y: 1 },
        { x: "b", y: new Date("x") }
      ],
      {}
    ],
    ["a start cell that is not one", [{ x: "a", y: 1 }], { startCell: "nope" }],
    [
      "rows past the end of the sheet",
      [
        { x: "a", y: 1 },
        { x: "b", y: 2 }
      ],
      { startCell: "D1048576" }
    ],
    ["columns past the end of the sheet", [{ x: "a", y: 1 }], { startCell: "XFD1" }]
  ])("leaves the sheet as it was for %s", (_name, rows, extra) => {
    const { ws } = sheet();
    Cell.setValue(ws, "D1", "KEEP");
    const before = snapshot(ws);
    expect(() =>
      Chart.createOptionsFromRows(ws, rows as never, {
        type: "line",
        x: "x",
        y: "y",
        startCell: "D1",
        ...extra
      })
    ).toThrow();
    expect(snapshot(ws)).toEqual(before);
    expect(Cell.getValue(ws, "D1048576")).toBeNull();
  });

  it("refuses what is not options or rows by name", () => {
    const { ws } = sheet();
    expect(() =>
      Chart.createOptionsFromRows(ws, null as never, { type: "line", x: "x", y: "y" })
    ).toThrow(/rows/);
    expect(() => Chart.createOptionsFromRows(ws, [], null as never)).toThrow(ExcelArgumentError);
    expect(() => Chart.createOptionsFromTable(ws, "T", 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Chart.applyPreset("lineMarkers", 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Chart.place(ws, "D1:K16", 42 as never)).toThrow(ExcelArgumentError);
  });
});

describe("the option builders take a family chosen at run time", () => {
  it("applyPreset, createOptionsFromRows and createOptionsFromTable accept either family's name", () => {
    const { ws } = sheet();
    const pick = (classic: boolean) =>
      (classic ? "lineMarkers" : "boxAndWhisker") as "lineMarkers" | "boxAndWhisker";
    for (const classic of [true, false]) {
      const chart = Chart.place(ws, "D1:K16", Chart.applyPreset(pick(classic), { series }));
      expect(Chart.isEx(chart)).toBe(!classic);
      const type = (classic ? "line" : "histogram") as "line" | "histogram";
      const rows = Chart.createOptionsFromRows(ws, [{ x: "a", y: 1 }], {
        type,
        x: "x",
        y: "y",
        startCell: "M1"
      });
      expect(Chart.isEx(Chart.place(ws, "D20:K36", rows))).toBe(!classic);
    }
  });
});

describe("the option builders take the data callers have", () => {
  interface SalesRow {
    month: string;
    revenue: number;
  }

  it("rows of an interface type, in a readonly array", () => {
    const { ws } = sheet();
    const rows: readonly SalesRow[] = [
      { month: "Jan", revenue: 10 },
      { month: "Feb", revenue: 20 }
    ];
    const options = Chart.createOptionsFromRows(ws, rows, {
      type: "line",
      x: "month",
      y: "revenue",
      startCell: "M1"
    });
    expect(options.series).toEqual([
      { name: "revenue", categories: "Data!$M$2:$M$3", values: "Data!$N$2:$N$3" }
    ]);
  });

  it("refuses a table from another sheet, which charted different cells by reference form", () => {
    const { wb, ws } = sheet();
    const other = Workbook.addWorksheet(wb, "Other");
    const table = Table.add(ws, {
      name: "Sales",
      ref: "D1",
      headerRow: true,
      columns: [{ name: "Month" }, { name: "Revenue" }],
      rows: [["Jan", 1]]
    });
    for (const structuredReferences of [true, false]) {
      expect(() =>
        Chart.createOptionsFromTable(other, table, { type: "line", structuredReferences })
      ).toThrow(/Table "Sales" is on sheet "Data", not "Other"/);
    }
  });
});
