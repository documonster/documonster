/**
 * Cold-start classification tests for `DefinedNames`.
 *
 * The global test setup installs both the formula engine and the
 * default syntax probe, so the rest of the defined-names suite exercises
 * the strict-probe path. This file temporarily uninstalls both to verify
 * the conservative fallback behaviour documented in
 * `@formula/default-syntax-probe`:
 *
 *   Without a probe, any non-range, non-wrapper text classifies as
 *   **opaque** — we have no evidence it is a parseable formula, so we
 *   preserve `rawText` for round-trip and leave `formulaMap` empty.
 *
 * The previous implementation fell back to "non-empty == formula", which
 * meant the same XLSX produced different internal classification
 * depending on whether `installFormulaEngine()` had been called. This
 * suite locks in the stricter, deterministic behaviour.
 *
 * To keep the rest of the suite working, `installFormulaEngine()` is
 * restored in `afterAll`.
 */

import { DefinedNames } from "@excel/defined-names";
import { Workbook } from "@excel/workbook";
import {
  createFormulaSyntaxProbe,
  installFormulaEngine,
  uninstallFormulaEngine
} from "@formula/install";
import { afterAll, beforeEach, describe, expect, it } from "vitest";

describe("DefinedNames — no-probe (cold start) classification", () => {
  beforeEach(() => {
    uninstallFormulaEngine();
  });

  afterAll(() => {
    installFormulaEngine();
  });

  it("classifies pure cell reference as reference even without a probe", () => {
    const dn = new DefinedNames();
    dn.model = [{ name: "Single", ranges: [], rawText: "Sheet1!$A$1" }];

    expect(dn.matrixMap["Single"]).toBeDefined();
    expect(dn.formulaMap["Single"]).toBeUndefined();
    expect(dn.opaqueMap["Single"]).toBeUndefined();
  });

  it("classifies comma-separated range union as reference without a probe", () => {
    const dn = new DefinedNames();
    dn.model = [{ name: "Multi", ranges: [], rawText: "Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2" }];

    expect(dn.matrixMap["Multi"]).toBeDefined();
    expect(dn.formulaMap["Multi"]).toBeUndefined();
    expect(dn.opaqueMap["Multi"]).toBeUndefined();
  });

  it("classifies OFFSET(...) as opaque without a probe (not formula)", () => {
    // With a probe this would be classified as `formula`. Without one,
    // we have no evidence the expression is parseable so we preserve
    // the raw text verbatim.
    const dn = new DefinedNames();
    dn.model = [{ name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }];

    expect(dn.opaqueMap["Dyn"]).toBeDefined();
    expect(dn.opaqueMap["Dyn"].rawText).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
    expect(dn.formulaMap["Dyn"]).toBeUndefined();
    expect(dn.matrixMap["Dyn"]).toBeUndefined();
  });

  it("classifies bare identifier as opaque without a probe", () => {
    const dn = new DefinedNames();
    dn.model = [{ name: "Alias", ranges: [], rawText: "AnotherName" }];

    expect(dn.opaqueMap["Alias"]).toBeDefined();
    expect(dn.opaqueMap["Alias"].rawText).toBe("AnotherName");
    expect(dn.formulaMap["Alias"]).toBeUndefined();
  });

  it("classifies #REF! and array/string literals as opaque without a probe", () => {
    const dn = new DefinedNames();
    dn.model = [
      { name: "Err", ranges: [], rawText: "#REF!" },
      { name: "Str", ranges: [], rawText: '"hello"' },
      { name: "Arr", ranges: [], rawText: "{1,2;3,4}" }
    ];

    expect(dn.opaqueMap["Err"]).toBeDefined();
    expect(dn.opaqueMap["Str"]).toBeDefined();
    expect(dn.opaqueMap["Arr"]).toBeDefined();
  });

  it("classifies malformed paren-containing text as opaque without a probe", () => {
    // With or without a probe, `OFFSET(???` should not land in
    // `formulaMap`. Previously this was incorrectly classified as
    // formula by the "non-empty => formula" fallback.
    const dn = new DefinedNames();
    dn.model = [{ name: "Bad", ranges: [], rawText: "OFFSET(???" }];

    expect(dn.opaqueMap["Bad"]).toBeDefined();
    expect(dn.formulaMap["Bad"]).toBeUndefined();
  });

  it("opaque classification preserves rawText through model round-trip", () => {
    // Without a probe, OFFSET(...) is opaque. The rawText must survive
    // a model→getter round-trip so the underlying XLSX bytes are stable.
    const dn = new DefinedNames();
    dn.model = [{ name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }];

    const model = dn.model;
    const entry = model.find(m => m.name === "Dyn");
    expect(entry).toBeDefined();
    expect(entry!.kind).toBe("opaque");
    expect(entry!.rawText).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
  });

  // ---------------------------------------------------------------------------
  // Explicit per-instance probe overrides the (absent) default
  // ---------------------------------------------------------------------------

  it("explicit constructor probe overrides the missing default", () => {
    const probe = createFormulaSyntaxProbe();
    const dn = new DefinedNames(probe);
    dn.model = [{ name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }];

    expect(dn.formulaMap["Dyn"]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
    expect(dn.opaqueMap["Dyn"]).toBeUndefined();
  });

  it("Workbook({ formulaSyntaxProbe }) option threads through to DefinedNames", async () => {
    const probe = createFormulaSyntaxProbe();
    const wb = new Workbook({ formulaSyntaxProbe: probe });
    wb.definedNames.model = [{ name: "Dyn", ranges: [], rawText: "SUM(Sheet1!$A$1:$A$3)" }];

    expect(wb.definedNames.formulaMap["Dyn"]).toBe("SUM(Sheet1!$A$1:$A$3)");
  });

  // ---------------------------------------------------------------------------
  // Default probe resolution is lazy
  // ---------------------------------------------------------------------------

  it("probe installed after Workbook construction is used at model-set time", () => {
    const wb = new Workbook(); // constructed before installFormulaEngine
    installFormulaEngine(); // install *after* construction

    // Model assignment happens now — should see the newly-installed probe.
    wb.definedNames.model = [{ name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }];

    expect(wb.definedNames.formulaMap["Dyn"]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
    expect(wb.definedNames.opaqueMap["Dyn"]).toBeUndefined();
  });
});
