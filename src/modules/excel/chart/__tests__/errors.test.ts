import {
  ChartClassic,
  ChartEx,
  ChartOptionsError,
  ExcelError,
  isExcelError
} from "@excel/chart/index";
import * as ExcelEntry from "@excel/index";
import { BaseError } from "@utils/errors";
import { isXmlError } from "@xml/index";
import { describe, expect, it } from "vitest";

function thrownBy(run: () => unknown): unknown {
  try {
    run();
  } catch (err) {
    return err;
  }
  throw new Error("expected the call to throw");
}

describe("documonster/chart errors", () => {
  it("publishes the same classes as documonster/excel", () => {
    expect(ChartOptionsError).toBe(ExcelEntry.ChartOptionsError);
    expect(ExcelError).toBe(ExcelEntry.ExcelError);
    expect(isExcelError).toBe(ExcelEntry.isExcelError);
  });

  it("invalid options throw a ChartOptionsError, which is an ExcelError", () => {
    for (const run of [
      () => ChartClassic.createModel({} as never),
      () => ChartClassic.createModel({ type: "bar", series: [] }),
      () => ChartEx.createModel({ type: "histogram", series: [] }),
      () => ChartClassic.applyPreset("nope" as never, { series: [] })
    ]) {
      const error = thrownBy(run);
      expect(error).toBeInstanceOf(ChartOptionsError);
      expect(error).toBeInstanceOf(BaseError);
      expect(isExcelError(error)).toBe(true);
    }
  });

  it("malformed ChartEx XML throws documonster/xml's error", () => {
    expect(isXmlError(thrownBy(() => ChartEx.fromXml("<<<not xml")))).toBe(true);
  });

  it("isExcelError is false for unrelated errors and non-errors", () => {
    for (const value of [new BaseError("x"), new Error("x"), { name: "ChartOptionsError" }, null]) {
      expect(isExcelError(value)).toBe(false);
    }
  });
});
