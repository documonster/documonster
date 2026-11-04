/**
 * Default-probe classification tests for `DefinedNames`.
 *
 * The toolkit has **no install step**: defined-name classification uses the
 * built-in tokenizer+parser probe (`defaultFormulaSyntaxProbe`, bound by every
 * public entry in `core/model-load.ts`) directly. This
 * suite verifies that classification works out of the box and that an explicit
 * per-instance probe still overrides the default.
 *
 * Range text classifies as `reference`; parseable expressions (e.g. `OFFSET`,
 * `SUM`) as `formula`; unparseable / literal text as `opaque` (preserving
 * `rawText` for round-trip).
 */

import { createDefinedNames, definedNamesModel } from "@excel/core/defined-names";
import { defaultFormulaSyntaxProbe, loadDefinedNamesModel } from "@excel/core/model-load";
import { getDefinedNames } from "@excel/core/workbook";
import { Workbook } from "@excel/index";
import { excelNameKey } from "@utils/name-key";
import { describe, expect, it } from "vitest";

describe("DefinedNames — default-probe classification (no install step)", () => {
  it("classifies pure cell reference as reference", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [{ name: "Single", ranges: [], rawText: "Sheet1!$A$1" }]);

    expect(dn.matrixMap[excelNameKey("Single")]).toBeDefined();
    expect(dn.formulaMap[excelNameKey("Single")]).toBeUndefined();
    expect(dn.opaqueMap[excelNameKey("Single")]).toBeUndefined();
  });

  it("classifies comma-separated range union as reference", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [
      { name: "Multi", ranges: [], rawText: "Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2" }
    ]);

    expect(dn.matrixMap[excelNameKey("Multi")]).toBeDefined();
    expect(dn.formulaMap[excelNameKey("Multi")]).toBeUndefined();
    expect(dn.opaqueMap[excelNameKey("Multi")]).toBeUndefined();
  });

  it("classifies parseable OFFSET(...) as formula out of the box", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [
      { name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }
    ]);

    expect(dn.formulaMap[excelNameKey("Dyn")]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
    expect(dn.opaqueMap[excelNameKey("Dyn")]).toBeUndefined();
  });

  it("classifies unparseable text as opaque (not a formula expression)", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [{ name: "Junk", ranges: [], rawText: "@@bad@@" }]);

    expect(dn.opaqueMap[excelNameKey("Junk")]).toBeDefined();
    expect(dn.opaqueMap[excelNameKey("Junk")].rawText).toBe("@@bad@@");
    expect(dn.formulaMap[excelNameKey("Junk")]).toBeUndefined();
  });

  it("classifies malformed paren-containing text as opaque", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [{ name: "Bad", ranges: [], rawText: "OFFSET(???" }]);

    expect(dn.opaqueMap[excelNameKey("Bad")]).toBeDefined();
    expect(dn.formulaMap[excelNameKey("Bad")]).toBeUndefined();
  });

  it("opaque classification preserves rawText through a model round-trip", () => {
    const dn = createDefinedNames();
    loadDefinedNamesModel(dn, [{ name: "Junk", ranges: [], rawText: "@@bad@@" }]);

    const model = definedNamesModel(dn);
    const entry = model.find(m => m.name === "Junk");
    expect(entry).toBeDefined();
    expect(entry!.kind).toBe("opaque");
    expect(entry!.rawText).toBe("@@bad@@");
  });

  // ---------------------------------------------------------------------------
  // Explicit per-instance probe overrides the default
  // ---------------------------------------------------------------------------

  it("explicit constructor probe overrides the default", () => {
    // A probe that rejects everything forces opaque classification even for
    // parseable text — proving the explicit probe wins over the default.
    const rejectAll = () => false;
    const dn = createDefinedNames(rejectAll);
    loadDefinedNamesModel(dn, [
      { name: "Dyn", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }
    ]);

    expect(dn.opaqueMap[excelNameKey("Dyn")]).toBeDefined();
    expect(dn.formulaMap[excelNameKey("Dyn")]).toBeUndefined();
  });

  it("the exported defaultFormulaSyntaxProbe reports parseability", () => {
    expect(defaultFormulaSyntaxProbe("SUM(A1:A3)")).toBe(true);
    expect(defaultFormulaSyntaxProbe("OFFSET(Sheet1!$A$1,0,0,3,1)")).toBe(true);
    expect(defaultFormulaSyntaxProbe("OFFSET(???")).toBe(false);
    expect(defaultFormulaSyntaxProbe("")).toBe(false);
  });

  it("Workbook({ formulaSyntaxProbe }) option threads through to DefinedNames", () => {
    const rejectAll = () => false;
    const wb = Workbook.create({ formulaSyntaxProbe: rejectAll });
    loadDefinedNamesModel(getDefinedNames(wb), [
      { name: "Dyn", ranges: [], rawText: "SUM(Sheet1!$A$1:$A$3)" }
    ]);

    // Custom probe rejects, so it stays opaque rather than formula.
    expect(getDefinedNames(wb).opaqueMap[excelNameKey("Dyn")]).toBeDefined();
    expect(getDefinedNames(wb).formulaMap[excelNameKey("Dyn")]).toBeUndefined();
  });

  it("default probe (no option) classifies SUM(...) as formula", () => {
    const wb = Workbook.create();
    loadDefinedNamesModel(getDefinedNames(wb), [
      { name: "Dyn", ranges: [], rawText: "SUM(Sheet1!$A$1:$A$3)" }
    ]);

    expect(getDefinedNames(wb).formulaMap[excelNameKey("Dyn")]).toBe("SUM(Sheet1!$A$1:$A$3)");
    expect(getDefinedNames(wb).opaqueMap[excelNameKey("Dyn")]).toBeUndefined();
  });
});
