/**
 * Number formats checked against real Excel.
 *
 * Each case in `excel-text-oracle.json` is `[value, format, result]`, where `result` is what Microsoft Excel
 * 16.113 for Mac returned for `=TEXT(value, format)`, with `#VALUE!` where Excel could not display the value.
 * They were collected by driving Excel through AppleScript, so they record behaviour rather than anyone's
 * reading of the documentation — several of them overturned conclusions that LibreOffice and SheetJS agreed on.
 *
 * Left out on purpose: AM/PM designators, which Excel's TEXT() takes from the system locale (the machine
 * measured was en-AU, which writes `pm` and does not know `A/P`).
 */
import { tryRenderNumberFormat } from "@utils/number-format-render";
import { describe, expect, it } from "vitest";

import cases from "./excel-text-oracle.json";

describe("number formats agree with Excel", () => {
  it.each(cases as [number, string, string][])("TEXT(%s, %j) is %j", (value, format, excel) => {
    expect(tryRenderNumberFormat(format, value) ?? "#VALUE!").toBe(excel);
  });
});
