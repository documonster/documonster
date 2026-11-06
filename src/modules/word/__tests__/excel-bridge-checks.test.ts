/**
 * `WordExcel`'s chart functions refuse a malformed chart or series with
 * `DocxArgumentError` naming the field, as `Build.chart` and
 * `WordExcel.toChartExXml` do. Before, `toChartSvg({})` / `toChartModel({})`
 * failed with `Cannot read properties of undefined (reading 'map')` and a
 * series without `categories` with `(reading 'length')`.
 *
 * The namespace uses one order (`toChart…`) and names its result: the
 * embedded workbook is returned as `.xlsx` bytes by `toChartXlsx`.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { WordExcel } from "../excel";
import type { Chart } from "../types";

const chart: Chart = {
  type: "bar",
  series: [{ name: "S", categories: ["a", "b"], values: [1, 2] }]
};

function refused(call: () => unknown, field: string, where: string): void {
  let error: unknown;
  try {
    call();
  } catch (e) {
    error = e;
  }
  expect(error, field).toBeInstanceOf(DocxArgumentError);
  expect((error as DocxArgumentError).field).toBe(field);
  expect((error as Error).message.startsWith(`${where}: `)).toBe(true);
}

describe("WordExcel chart functions", () => {
  it("are named in one order", () => {
    expect(typeof WordExcel.toChartSvg).toBe("function");
    expect(typeof WordExcel.toChartXlsx).toBe("function");
    expect(typeof WordExcel.toChartModel).toBe("function");
    expect(typeof WordExcel.toChartExXml).toBe("function");
    expect("chartToSvg" in WordExcel).toBe(false);
    expect("toChartWorkbook" in WordExcel).toBe(false);
  });

  it("render a valid chart", async () => {
    expect(WordExcel.toChartSvg(chart)).toMatch(/^<svg/);
    expect(WordExcel.toChartModel(chart).chart).toBeDefined();
    const xlsx = await WordExcel.toChartXlsx(chart.series);
    expect(xlsx.subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
  });

  it("refuse a chart without type or series", () => {
    refused(() => WordExcel.toChartSvg({} as never), "chart.type", "WordExcel.toChartSvg");
    refused(() => WordExcel.toChartModel({} as never), "chart.type", "WordExcel.toChartModel");
    refused(
      () => WordExcel.toChartSvg({ type: "bar" } as never),
      "chart.series",
      "WordExcel.toChartSvg"
    );
    refused(
      () => WordExcel.toChartModel({ type: "bar", series: [{ name: "S", values: [1] }] } as never),
      "chart.series[0].categories",
      "WordExcel.toChartModel"
    );
    refused(
      () => WordExcel.toChartSvg({ type: "bogus", series: [] } as never),
      "chart.type",
      "WordExcel.toChartSvg"
    );
  });

  it("toChartXlsx refuses a series without categories", async () => {
    const error = await WordExcel.toChartXlsx([{ name: "S", values: [1, 2] }] as never).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as DocxArgumentError).field).toBe("series[0].categories");
  });
});
