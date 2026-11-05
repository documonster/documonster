/**
 * The handle-free chart entries check what they are given at the boundary: a missing or wrong argument is an
 * `ExcelArgumentError` naming the public parameter and the call, never a `TypeError` from inside.
 */
import { ChartClassic, ChartEx, DrawingMl, ExcelArgumentError } from "@excel/chart/index";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function refusal(run: () => unknown): ExcelArgumentError {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(ExcelArgumentError);
    return err as ExcelArgumentError;
  }
  throw new Error("expected an ExcelArgumentError");
}

type Call = readonly [where: string, field: string, run: (bad: unknown) => unknown];

const workbook = Workbook.create();
const worksheet = Workbook.addWorksheet(workbook, "Sheet1");
Cell.setValue(worksheet, "A1", 1);

const ref = () => ({ formula: "Sheet1!$A$1" });

const CALLS: readonly Call[] = [
  ["ChartClassic.fillNumRef", "ref", bad => ChartClassic.fillNumRef(bad as never, workbook)],
  ["ChartClassic.fillNumRef", "workbook", bad => ChartClassic.fillNumRef(ref(), bad as never)],
  [
    "ChartClassic.fillNumRef",
    "options",
    bad => ChartClassic.fillNumRef(ref(), workbook, (bad ?? null) as never)
  ],
  [
    "ChartClassic.fillNumRef",
    "options.worksheet",
    bad => ChartClassic.fillNumRef(ref(), workbook, { worksheet: (bad ?? null) as never })
  ],
  ["ChartClassic.fillStrRef", "ref", bad => ChartClassic.fillStrRef(bad as never, workbook)],
  ["ChartClassic.fillStrRef", "workbook", bad => ChartClassic.fillStrRef(ref(), bad as never)],
  [
    "ChartClassic.fillCaches",
    "workbook",
    bad =>
      ChartClassic.fillCaches(
        ChartClassic.createModel({ type: "bar", series: [{ values: "Sheet1!$A$1" }] }),
        bad as never
      )
  ],
  ["ChartClassic.toColorsXml", "model", bad => ChartClassic.toColorsXml(bad as never)],
  ["ChartClassic.toStyleXml", "model", bad => ChartClassic.toStyleXml(bad as never)],
  ["ChartClassic.fromColorsXml", "rawXml", bad => ChartClassic.fromColorsXml(bad as never)],
  ["ChartClassic.fromStyleXml", "rawXml", bad => ChartClassic.fromStyleXml(bad as never)],
  [
    "ChartClassic.createOptionsFromTable",
    "worksheet",
    bad => ChartClassic.createOptionsFromTable(bad as never, "T", {} as never)
  ],
  [
    "ChartEx.createOptionsFromRows",
    "worksheet",
    bad => ChartEx.createOptionsFromRows(bad as never, [{ a: 1 }], { type: "funnel" } as never)
  ],
  ["ChartEx.toXml", "model", bad => ChartEx.toXml(bad as never)],
  ["DrawingMl.createSpPr", "props", bad => DrawingMl.createSpPr(bad as never)],
  ["DrawingMl.createTxPr", "props", bad => DrawingMl.createTxPr(bad as never)],
  ["DrawingMl.getSpPrFillColor", "spPr", bad => DrawingMl.getSpPrFillColor(bad as never)],
  ["DrawingMl.getSpPrGradient", "spPr", bad => DrawingMl.getSpPrGradient(bad as never)],
  ["DrawingMl.getSpPrLine", "spPr", bad => DrawingMl.getSpPrLine(bad as never)],
  ["DrawingMl.getSpPrPattern", "spPr", bad => DrawingMl.getSpPrPattern(bad as never)],
  ["DrawingMl.getTxPrColor", "txPr", bad => DrawingMl.getTxPrColor(bad as never)],
  ["DrawingMl.getTxPrFontSize", "txPr", bad => DrawingMl.getTxPrFontSize(bad as never)]
];

describe("chart entries refuse a wrong argument", () => {
  it.each(
    CALLS.flatMap(([where, field, run]) =>
      [undefined, 42].map(bad => [where, field, run, bad] as const)
    )
  )("%s: %s = %j", (where, field, run, bad) => {
    const err = refusal(() => run(bad));
    expect(err.field).toBe(field);
    expect(err.message.startsWith(`${where}: `)).toBe(true);
  });

  it("refuses a ChartEx model that has no chartSpace, and a non-boolean forceStructural", () => {
    expect(refusal(() => ChartEx.toXml({} as never)).field).toBe("model");
    const model = ChartEx.createModel({
      type: "funnel",
      categories: "Sheet1!$A$1:$A$3",
      series: [{ values: "Sheet1!$B$1:$B$3" }]
    });
    expect(refusal(() => ChartEx.toXml(model, { forceStructural: 1 as never })).field).toBe(
      "options.forceStructural"
    );
  });

  it("refuses a date1904 that is not a boolean before filling anything", () => {
    const r = ref() as { formula: string; numCache?: unknown };
    const err = refusal(() =>
      ChartClassic.fillNumRef(r as never, workbook, { date1904: "yes" as never })
    );
    expect(err.field).toBe("options.date1904");
    expect(r.numCache).toBeUndefined();
  });
});

describe("chart entries behave as before on valid arguments", () => {
  it("fills a reference, reads and writes the sidecars, and builds DrawingML", () => {
    const r = ref();
    ChartClassic.fillNumRef(r as never, workbook, { worksheet });
    expect(JSON.stringify(r)).toContain("1");
    const colors = ChartClassic.fromColorsXml(ChartClassic.toColorsXml({}));
    expect(typeof colors.rawXml).toBe("string");
    expect(ChartClassic.toStyleXml(ChartClassic.fromStyleXml("<cs:chartStyle/>"))).toBe(
      "<cs:chartStyle/>"
    );
    const spPr = DrawingMl.createSpPr({ fill: { solid: { rgb: "FF0000" } } } as never);
    expect(DrawingMl.getSpPrFillColor(spPr)).toEqual({ rgb: "FF0000" });
    expect(DrawingMl.getTxPrFontSize(DrawingMl.createTxPr({ size: 1200 }))).toBe(12);
  });
});
