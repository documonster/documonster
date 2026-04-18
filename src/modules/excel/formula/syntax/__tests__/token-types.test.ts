/**
 * Unit tests for the public helpers exported by `token-types.ts`.
 *
 * Specifically `stripFunctionPrefix`: every formula lookup in the
 * registry flows through it so cross-version XLSX imports resolve
 * their prefixed function tokens (`_XLFN.FILTER`, `_XLFN._XLWS.SORT`,
 * etc.) to canonical names.
 */

import { describe, expect, it } from "vitest";

import { stripFunctionPrefix } from "../token-types";

describe("stripFunctionPrefix", () => {
  it("returns plain names unchanged", () => {
    expect(stripFunctionPrefix("SUM")).toBe("SUM");
    expect(stripFunctionPrefix("VLOOKUP")).toBe("VLOOKUP");
  });

  it("strips _XLFN. prefix", () => {
    expect(stripFunctionPrefix("_XLFN.IFS")).toBe("IFS");
    expect(stripFunctionPrefix("_XLFN.FILTER")).toBe("FILTER");
  });

  it("strips _XLFN._XLWS. double prefix", () => {
    expect(stripFunctionPrefix("_XLFN._XLWS.SORT")).toBe("SORT");
    expect(stripFunctionPrefix("_XLFN._XLWS.UNIQUE")).toBe("UNIQUE");
  });

  it("preserves case of stripped name", () => {
    // The helper does NOT uppercase — callers are responsible for
    // case-normalising before registry lookup.
    expect(stripFunctionPrefix("_XLFN.sort")).toBe("sort");
  });

  it("leaves partial / malformed prefixes alone", () => {
    expect(stripFunctionPrefix("_XLF.SUM")).toBe("_XLF.SUM");
    expect(stripFunctionPrefix("XLFN.SUM")).toBe("XLFN.SUM");
    expect(stripFunctionPrefix("_xlfn.SUM")).toBe("_xlfn.SUM"); // lowercase prefix stays
  });

  it("empty input returns empty", () => {
    expect(stripFunctionPrefix("")).toBe("");
  });

  it("handles the pathological `_XLFN.` alone", () => {
    // A bare `_XLFN.` with nothing after it strips to empty string.
    expect(stripFunctionPrefix("_XLFN.")).toBe("");
  });

  it("double-prefix takes precedence over single (checked in order)", () => {
    // Belt-and-braces: verify the more specific `_XLFN._XLWS.` match
    // runs before the shorter `_XLFN.` match.
    expect(stripFunctionPrefix("_XLFN._XLWS.X")).toBe("X");
    expect(stripFunctionPrefix("_XLFN._XLWS.")).toBe("");
  });

  it("does not strip the prefix from a function that happens to start with `_XLFN`", () => {
    // The prefix must end with a period.
    expect(stripFunctionPrefix("_XLFNSUM")).toBe("_XLFNSUM");
  });
});
