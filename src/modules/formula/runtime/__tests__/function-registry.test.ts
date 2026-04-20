/**
 * Unit tests for the function registry — the single source of truth
 * for formula function lookup.
 *
 * The registry is accessed on every formula call via `lookupFunction`,
 * and it has to handle Excel's `_XLFN.` / `_XLWS.` prefix variants
 * transparently so cross-version-imported XLSX files still resolve
 * their function tokens.
 */

import { describe, expect, it } from "vitest";

import { lookupFunction } from "../function-registry";

describe("lookupFunction: core resolution", () => {
  it("resolves a plain function name", () => {
    const desc = lookupFunction("SUM");
    expect(desc).toBeDefined();
    expect(desc?.name).toBe("SUM");
  });

  it("is case-sensitive — registry stores uppercase only", () => {
    // The caller (evaluator) always uppercases names before lookup;
    // we don't silently lowercase for them.
    expect(lookupFunction("sum")).toBeUndefined();
    expect(lookupFunction("Sum")).toBeUndefined();
  });

  it("returns undefined for unknown names", () => {
    expect(lookupFunction("NOT_A_REAL_FUNCTION")).toBeUndefined();
  });

  it("returns undefined for empty string", () => {
    expect(lookupFunction("")).toBeUndefined();
  });
});

describe("lookupFunction: _XLFN prefix handling (R7)", () => {
  it("strips _XLFN. prefix before resolution", () => {
    const plain = lookupFunction("IFS");
    const prefixed = lookupFunction("_XLFN.IFS");
    expect(prefixed).toBe(plain);
  });

  it("strips _XLFN._XLWS. double prefix", () => {
    const plain = lookupFunction("SORT");
    const prefixed = lookupFunction("_XLFN._XLWS.SORT");
    expect(prefixed).toBe(plain);
  });

  it("prefix with unknown function still returns undefined", () => {
    expect(lookupFunction("_XLFN.NOT_REAL")).toBeUndefined();
  });

  it("prefix on core functions that predate the _XLFN era also works", () => {
    // Even canonical functions like SUM can arrive prefixed if an older
    // implementation over-annotated on export.
    const plain = lookupFunction("SUM");
    const prefixed = lookupFunction("_XLFN.SUM");
    expect(prefixed).toBe(plain);
  });

  it("malformed prefix returns undefined (not registered)", () => {
    expect(lookupFunction("_XLFN.")).toBeUndefined();
    expect(lookupFunction("_XLFN")).toBeUndefined();
  });
});

describe("lookupFunction: sample of each family", () => {
  // Smoke-test every major function family resolves. This catches
  // registration-wiring breaks that would otherwise only show up when
  // a user uses that specific function.
  //
  // Note: special forms (IF / IFS / CHOOSE / LAMBDA / LET / REDUCE /
  // MAP / SCAN / BYROW / BYCOL / MAKEARRAY / IFERROR / IFNA) are
  // handled by the evaluator directly and intentionally NOT in the
  // function registry — they need lazy argument evaluation which the
  // generic descriptor dispatch can't do.
  const FAMILIES: Record<string, string[]> = {
    math: [
      "SUM",
      "ABS",
      "ROUND",
      "POWER",
      "CEILING",
      "MOD",
      "COMBIN",
      "MMULT",
      "MDETERM",
      "MINVERSE",
      "MUNIT",
      "SERIESSUM"
    ],
    stats: [
      "AVERAGE",
      "MEDIAN",
      "STDEV",
      "VAR",
      "NORM.S.DIST",
      "CORREL",
      "PERCENTRANK",
      "PERCENTRANK.INC",
      "PERCENTRANK.EXC",
      "PROB"
    ],
    text: [
      "LEN",
      "LEFT",
      "UPPER",
      "LOWER",
      "TRIM",
      "CONCAT",
      "SUBSTITUTE",
      "TEXT",
      "REGEXTEST",
      "REGEXEXTRACT",
      "REGEXREPLACE",
      "VALUETOTEXT",
      "ARRAYTOTEXT",
      "ENCODEURL",
      "LEFTB",
      "RIGHTB",
      "MIDB",
      "LENB",
      "FINDB",
      "SEARCHB"
    ],
    date: ["DATE", "YEAR", "MONTH", "DAY", "TODAY", "NOW", "WEEKDAY", "EDATE", "DATEDIF"],
    lookup: ["VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX", "ADDRESS"],
    logical: ["AND", "OR", "NOT", "XOR"],
    financial: [
      "PMT",
      "FV",
      "PV",
      "NPER",
      "RATE",
      "IRR",
      "NPV",
      "SLN",
      "DB",
      "DDB",
      "SYD",
      "VDB",
      "ACCRINTM",
      "TBILLPRICE",
      "TBILLYIELD",
      "TBILLEQ",
      "PRICEMAT",
      "YIELDMAT"
    ],
    engineering: ["DEC2BIN", "HEX2DEC", "BITAND", "COMPLEX", "IMSUM", "BESSELJ"],
    database: [
      "DSUM",
      "DAVERAGE",
      "DCOUNT",
      "DCOUNTA",
      "DMAX",
      "DMIN",
      "DGET",
      "DSTDEV",
      "DSTDEVP",
      "DVAR",
      "DVARP"
    ],
    dynamic: ["SEQUENCE", "FILTER", "SORT", "UNIQUE", "RANDARRAY", "TAKE", "DROP"],
    info: ["ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "TYPE", "CELL", "INFO"]
  };

  for (const [family, names] of Object.entries(FAMILIES)) {
    it(`${family} family: all ${names.length} names resolve`, () => {
      for (const name of names) {
        const desc = lookupFunction(name);
        expect(desc, `lookupFunction("${name}") should resolve — ${family} family`).toBeDefined();
      }
    });
  }
});

describe("lookupFunction: alias equivalence", () => {
  // Aliases share the same underlying `invoke` implementation but may
  // have distinct `name` strings (registry stores them as separate
  // descriptor objects for traceability). Verify the underlying impl
  // is shared, not descriptor identity.
  function sameImpl(a: string, b: string) {
    const da = lookupFunction(a);
    const db = lookupFunction(b);
    expect(da, `${a} should resolve`).toBeDefined();
    expect(db, `${b} should resolve`).toBeDefined();
    expect(da!.invoke).toBe(db!.invoke);
  }

  it("MODE.SNGL is an alias of MODE", () => {
    sameImpl("MODE.SNGL", "MODE");
  });

  it("PERCENTILE.INC is an alias of PERCENTILE", () => {
    sameImpl("PERCENTILE.INC", "PERCENTILE");
  });

  it("QUARTILE.INC is an alias of QUARTILE", () => {
    sameImpl("QUARTILE.INC", "QUARTILE");
  });

  it("STDEV.S aliases STDEV", () => {
    sameImpl("STDEV.S", "STDEV");
  });

  it("VAR.S aliases VAR", () => {
    sameImpl("VAR.S", "VAR");
  });

  it("RANK.EQ aliases RANK", () => {
    sameImpl("RANK.EQ", "RANK");
  });

  it("CEILING.MATH / CEILING.PRECISE / ISO.CEILING use dedicated variant impls", () => {
    // Regression: Excel's CEILING.MATH and CEILING.PRECISE have different
    // semantics from classic CEILING (mixed-sign tolerance, |significance|,
    // mode argument). They used to delegate to fnCEILING, which silently
    // emitted #NUM! for negative num + positive sig — a soft divergence
    // from Excel. Each variant now has its own implementation; ISO.CEILING
    // aliases CEILING.PRECISE.
    const ceiling = lookupFunction("CEILING");
    const ceilingMath = lookupFunction("CEILING.MATH");
    const ceilingPrecise = lookupFunction("CEILING.PRECISE");
    const isoCeiling = lookupFunction("ISO.CEILING");
    expect(ceiling?.invoke).not.toBe(ceilingMath?.invoke);
    expect(ceiling?.invoke).not.toBe(ceilingPrecise?.invoke);
    expect(ceilingPrecise?.invoke).toBe(isoCeiling?.invoke);
  });

  it("FORECAST.LINEAR aliases FORECAST", () => {
    sameImpl("FORECAST.LINEAR", "FORECAST");
  });

  it("CONFIDENCE aliases CONFIDENCE.NORM (back-compat)", () => {
    sameImpl("CONFIDENCE", "CONFIDENCE.NORM");
  });

  it("B-variant text functions alias their non-B counterparts", () => {
    // In non-DBCS locales these are identical; Excel treats them as
    // aliases in practice for our target use case.
    sameImpl("LEFTB", "LEFT");
    sameImpl("RIGHTB", "RIGHT");
    sameImpl("MIDB", "MID");
    sameImpl("LENB", "LEN");
    sameImpl("FINDB", "FIND");
    sameImpl("SEARCHB", "SEARCH");
    sameImpl("REPLACEB", "REPLACE");
  });

  it("ZTEST / Z.TEST share impl", () => {
    sameImpl("ZTEST", "Z.TEST");
  });

  it("TTEST / T.TEST share impl", () => {
    sameImpl("TTEST", "T.TEST");
  });

  it("FTEST / F.TEST share impl", () => {
    sameImpl("FTEST", "F.TEST");
  });

  it("CHITEST / CHISQ.TEST share impl", () => {
    sameImpl("CHITEST", "CHISQ.TEST");
  });

  it("BINOMDIST / BINOM.DIST share impl", () => {
    sameImpl("BINOMDIST", "BINOM.DIST");
  });

  it("PERCENTRANK / PERCENTRANK.INC share impl", () => {
    sameImpl("PERCENTRANK", "PERCENTRANK.INC");
  });
});

describe("FunctionDescriptor shape", () => {
  it("has name, minArity, maxArity, and invoke", () => {
    const d = lookupFunction("SUM");
    expect(d).toBeDefined();
    expect(typeof d!.name).toBe("string");
    expect(typeof d!.minArity).toBe("number");
    expect(typeof d!.maxArity).toBe("number");
    expect(typeof d!.invoke).toBe("function");
    expect(d!.minArity).toBeLessThanOrEqual(d!.maxArity);
  });

  it("arity ranges match expected", () => {
    // SUM: 1..255; PI: 0..0; ROUND: 2..2
    expect(lookupFunction("PI")!.minArity).toBe(0);
    expect(lookupFunction("PI")!.maxArity).toBe(0);
    expect(lookupFunction("ROUND")!.minArity).toBe(2);
    expect(lookupFunction("ROUND")!.maxArity).toBe(2);
    expect(lookupFunction("SUM")!.minArity).toBe(1);
    expect(lookupFunction("SUM")!.maxArity).toBeGreaterThanOrEqual(2);
  });
});
