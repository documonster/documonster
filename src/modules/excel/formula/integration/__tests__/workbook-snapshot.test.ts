/**
 * Unit tests for `workbook-snapshot` helpers — the pure string-key and
 * name-resolution functions that the evaluator relies on to address
 * cells and defined names consistently.
 */

import { describe, it, expect } from "vitest";

import {
  snapshotCellKey,
  formulaCellKey,
  spillCellKeyFromId,
  scopedNameKey,
  resolveDefinedName
} from "../workbook-snapshot";
import type { DefinedNameSnapshot } from "../workbook-snapshot";

describe("snapshotCellKey", () => {
  it("formats 'row:col'", () => {
    expect(snapshotCellKey(3, 7)).toBe("3:7");
  });

  it("distinguishes positions that look similar", () => {
    // Important: "12:3" vs "1:23" must be different
    expect(snapshotCellKey(12, 3)).toBe("12:3");
    expect(snapshotCellKey(1, 23)).toBe("1:23");
    expect(snapshotCellKey(12, 3)).not.toBe(snapshotCellKey(1, 23));
  });

  it("handles single-digit positions", () => {
    expect(snapshotCellKey(1, 1)).toBe("1:1");
  });

  it("handles large positions", () => {
    expect(snapshotCellKey(1048576, 16384)).toBe("1048576:16384");
  });
});

describe("formulaCellKey", () => {
  it("formats 'SheetName!row:col'", () => {
    expect(formulaCellKey("Data", 2, 3)).toBe("Data!2:3");
  });

  it("different sheets produce distinct keys", () => {
    expect(formulaCellKey("A", 1, 1)).not.toBe(formulaCellKey("B", 1, 1));
  });

  it("handles sheet names with spaces", () => {
    expect(formulaCellKey("My Sheet", 1, 1)).toBe("My Sheet!1:1");
  });
});

describe("spillCellKeyFromId", () => {
  it("uses worksheet ID for stable identity", () => {
    expect(spillCellKeyFromId(5, 2, 3)).toBe("ws:5!2:3");
  });

  it("different IDs distinct", () => {
    expect(spillCellKeyFromId(1, 1, 1)).not.toBe(spillCellKeyFromId(2, 1, 1));
  });
});

describe("scopedNameKey", () => {
  it("uses NUL separator for sheet-scoped names", () => {
    const k = scopedNameKey("Data", "MyName");
    expect(k).toBe("DATA\u0000MYNAME");
  });

  it("uppercases both sheet and name", () => {
    const k = scopedNameKey("data", "myname");
    expect(k).toBe("DATA\u0000MYNAME");
  });

  it("NUL separator avoids collision with literal sheet names", () => {
    // If we used ':' or '!' then a sheet like 'A:B' could collide.
    // NUL isn't allowed in Excel sheet/name literals, so keys are unique.
    expect(scopedNameKey("A", "B:C")).toBe("A\u0000B:C");
    expect(scopedNameKey("A:B", "C")).toBe("A:B\u0000C");
    expect(scopedNameKey("A", "B:C")).not.toBe(scopedNameKey("A:B", "C"));
  });
});

describe("resolveDefinedName", () => {
  function dn(name: string, range: string): DefinedNameSnapshot {
    return { name, ranges: [range] };
  }

  it("returns undefined when name not found", () => {
    const map = new Map<string, DefinedNameSnapshot>();
    expect(resolveDefinedName(map, "Unknown")).toBeUndefined();
  });

  it("finds workbook-scoped global name", () => {
    const map = new Map<string, DefinedNameSnapshot>([["TAXRATE", dn("TaxRate", "Sheet1!$A$1")]]);
    const r = resolveDefinedName(map, "TaxRate");
    expect(r?.name).toBe("TaxRate");
  });

  it("is case-insensitive for workbook names", () => {
    const map = new Map<string, DefinedNameSnapshot>([["TAXRATE", dn("TaxRate", "Sheet1!$A$1")]]);
    expect(resolveDefinedName(map, "taxrate")?.name).toBe("TaxRate");
    expect(resolveDefinedName(map, "TAXRATE")?.name).toBe("TaxRate");
  });

  it("sheet-scoped name takes precedence over workbook", () => {
    const map = new Map<string, DefinedNameSnapshot>([
      ["TAXRATE", dn("TaxRate", "Sheet1!$A$1")],
      [scopedNameKey("Report", "TaxRate"), dn("TaxRate", "Report!$B$1")]
    ]);
    const sheetScoped = resolveDefinedName(map, "TaxRate", "Report");
    expect(sheetScoped?.ranges[0]).toBe("Report!$B$1");
  });

  it("sheet-scoped lookup on unrelated sheet falls back to global", () => {
    const map = new Map<string, DefinedNameSnapshot>([
      ["TAXRATE", dn("TaxRate", "Sheet1!$A$1")],
      [scopedNameKey("Report", "TaxRate"), dn("TaxRate", "Report!$B$1")]
    ]);
    // Looking from a different sheet → finds workbook-level
    const r = resolveDefinedName(map, "TaxRate", "Other");
    expect(r?.ranges[0]).toBe("Sheet1!$A$1");
  });

  it("sheet scope is case-insensitive", () => {
    const map = new Map<string, DefinedNameSnapshot>([
      [scopedNameKey("Report", "TaxRate"), dn("TaxRate", "Report!$B$1")]
    ]);
    expect(resolveDefinedName(map, "TaxRate", "REPORT")?.ranges[0]).toBe("Report!$B$1");
    expect(resolveDefinedName(map, "taxrate", "report")?.ranges[0]).toBe("Report!$B$1");
  });
});
