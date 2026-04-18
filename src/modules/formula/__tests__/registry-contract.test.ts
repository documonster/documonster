/**
 * Verifies the registry contract that lets `Workbook.calculateFormulas()`
 * fail loudly when no engine is installed, and exercises the "no probe"
 * classification path for defined names.
 *
 * The global test setup (`src/test/setup-formula.ts`) installs both the
 * engine and the default syntax probe so the rest of the suite can run
 * without boilerplate. This file uninstalls both via
 * `uninstallFormulaEngine()` (symmetric reset) to exercise the cold-start
 * behaviour, then reinstalls so subsequent tests keep working.
 */

import { Workbook } from "@excel/workbook";
import { getDefaultSyntaxProbe } from "@formula/default-syntax-probe";
import {
  hasFormulaEngine,
  invokeFormulaEngine,
  tryInvokeFormulaEngine
} from "@formula/host-registry";
import { installFormulaEngine, uninstallFormulaEngine } from "@formula/install";
import { afterAll, beforeEach, describe, expect, it } from "vitest";

describe("formula registry contract", () => {
  beforeEach(() => {
    uninstallFormulaEngine();
  });

  afterAll(() => {
    // Restore so subsequent test files keep working.
    installFormulaEngine();
  });

  it("invokeFormulaEngine throws when no engine is installed", () => {
    const wb = new Workbook();
    expect(() => invokeFormulaEngine(wb)).toThrowError(/No formula engine is installed/);
  });

  it("Workbook.calculateFormulas() surfaces the registry error", () => {
    const wb = new Workbook();
    expect(() => wb.calculateFormulas()).toThrowError(/installFormulaEngine/);
  });

  it("tryInvokeFormulaEngine returns false (silent) when no engine is installed", () => {
    const wb = new Workbook();
    expect(tryInvokeFormulaEngine(wb)).toBe(false);
  });

  it("uninstallFormulaEngine is symmetric with installFormulaEngine", () => {
    // Precondition (set by beforeEach): both slots cleared.
    expect(hasFormulaEngine()).toBe(false);
    expect(getDefaultSyntaxProbe()).toBeNull();
  });

  it("installFormulaEngine wires both engine and default probe", () => {
    installFormulaEngine();

    expect(hasFormulaEngine()).toBe(true);
    expect(getDefaultSyntaxProbe()).not.toBeNull();

    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = { formula: "A1+A2" };

    wb.calculateFormulas();
    expect(ws.getCell("A3").result).toBe(3);
  });
});
