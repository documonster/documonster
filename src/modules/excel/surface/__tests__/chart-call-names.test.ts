/**
 * `Chart.*` argument errors start with the public call — `Chart.createOptionsFromRows: …` — not the shared builder that
 * found the mistake (`createOptionsFromRows: …`, `buildChartModel: …`), and no Excel argument error names an internal
 * function.
 */
import { Chart, ExcelArgumentError, Row, Table, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.addMany(ws, [
    ["k", "v"],
    ["a", 1],
    ["b", 2]
  ]);
  return { wb, ws };
}

function chart() {
  const { ws } = sheet();
  return { ws, c: Chart.place(ws, "D1:H10", { type: "bar", series: [{ values: "S!$B$2:$B$3" }] }) };
}

function messageOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(ExcelArgumentError);
    return (error as Error).message;
  }
  throw new Error("did not throw");
}

const cases: [string, () => unknown][] = [
  [
    "place",
    () => Chart.place(sheet().ws, "D1:H10", { type: "bar", series: [{ values: 42 }] } as never)
  ],
  [
    "place",
    () =>
      Chart.place(sheet().ws, "D1:H10", {
        type: "pie",
        series: [{ values: "S!$B$2:$B$3", trendline: { type: "linear" } }]
      } as never)
  ],
  [
    "createOptionsFromRows",
    () =>
      Chart.createOptionsFromRows(sheet().ws, [{ a: 1 }], {
        type: "bar",
        categoryKey: "zz",
        valueKeys: ["a"]
      } as never)
  ],
  [
    "createOptionsFromRows",
    () =>
      Chart.createOptionsFromRows(sheet().ws, [{ a: 1 }], {
        type: "histogram",
        valueKeys: []
      } as never)
  ],
  [
    "createOptionsFromTable",
    () => {
      const { ws } = sheet();
      Table.add(ws, {
        name: "T_c",
        ref: "J1",
        columns: [{ name: "k" }, { name: "v" }],
        rows: [["a", 1]]
      });
      return Chart.createOptionsFromTable(ws, "T_c", {
        type: "bar",
        valueColumns: ["nope"]
      } as never);
    }
  ],
  [
    "createOptionsFromTable",
    () => Chart.createOptionsFromTable(sheet().ws, "Missing", { type: "bar" } as never)
  ],
  [
    "createSeriesFromColumns",
    () => Chart.createSeriesFromColumns(sheet().ws, { range: 42 } as never)
  ],
  ["setTitle", () => Chart.setTitle(chart().c, 42 as never)],
  ["setStyle", () => Chart.setStyle(chart().c, -3)],
  ["addSeries", () => Chart.addSeries(chart().c, null as never)],
  ["setSeriesValues", () => Chart.setSeriesValues(chart().c, 0, 42 as never)],
  ["setUserShapesXml", () => Chart.setUserShapesXml(chart().c, 42 as never)],
  ["copyTo", () => Chart.copyTo(chart().c, 42 as never)],
  ["clone", () => Chart.clone(chart().c, 42 as never)],
  ["mutate", () => Chart.mutate(chart().c, 42 as never)],
  ["updateSeries", () => Chart.updateSeries(chart().c, 0, { values: 42 } as never)],
  ["addSeriesFromOptions", () => Chart.addSeriesFromOptions(chart().c, { values: 42 } as never)],
  ["setSeriesName", () => Chart.setSeriesName(chart().c, 0, 42 as never)],
  ["setLegend", () => Chart.setLegend(chart().c, 42 as never)],
  ["setSpPr", () => Chart.setSpPr(chart().c, 42 as never)],
  ["toSvg", () => Chart.toSvg(chart().c, { width: -1 })],
  ["applyPreset", () => Chart.applyPreset("nope" as never, {} as never)],
  ["mutateEx", () => Chart.mutateEx(chart().c, null as never)],
  ["mutate", () => Chart.mutate(chart().c, () => {}, { preferRawPatch: "yes" } as never)],
  ["clone", () => Chart.clone(chart().c, { tl: 42 } as never)]
];

describe("Chart.* argument errors name the public call", () => {
  it("a mutator that is not a function is an argument error, not a TypeError from inside", () => {
    expect(() => Chart.mutate(chart().c, 42 as never)).toThrow(
      expect.objectContaining({ field: "mutator" })
    );
  });

  it.each(cases)("Chart.%s", (name, run) => {
    const message = messageOf(run);
    expect(message.startsWith(`Chart.${name}: `), message).toBe(true);
  });
});
