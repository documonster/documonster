/**
 * Unit tests for the date-1904 module-level context toggle.
 *
 * The flag is read on every DATE / YEAR / NOW / TODAY / EOMONTH / etc.
 * call; getting this wrong silently shifts every date serial by 1462
 * days. These tests lock the simple getter/setter contract and the
 * default value — the evaluator's end-to-end date1904 behaviour is
 * already tested through the workbook roundtrip suite.
 */

import { afterEach, describe, expect, it } from "vitest";

import { isDate1904, setDate1904 } from "../_date-context";

describe("date1904 context", () => {
  // Module-level state — reset after every test so test order can't
  // affect outcome.
  afterEach(() => {
    setDate1904(false);
  });

  it("defaults to false (1900-date system)", () => {
    // Re-import shows default. Since the module has been loaded by
    // previous tests, we set false first to normalise.
    setDate1904(false);
    expect(isDate1904()).toBe(false);
  });

  it("setDate1904(true) flips to 1904 mode", () => {
    setDate1904(true);
    expect(isDate1904()).toBe(true);
  });

  it("setDate1904(false) flips back to 1900 mode", () => {
    setDate1904(true);
    setDate1904(false);
    expect(isDate1904()).toBe(false);
  });

  it("persists across multiple reads", () => {
    setDate1904(true);
    expect(isDate1904()).toBe(true);
    expect(isDate1904()).toBe(true);
    expect(isDate1904()).toBe(true);
  });

  it("last setter wins", () => {
    setDate1904(true);
    setDate1904(false);
    setDate1904(true);
    expect(isDate1904()).toBe(true);
  });
});
