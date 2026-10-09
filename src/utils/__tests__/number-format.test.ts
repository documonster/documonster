import {
  isDateFmt,
  isGeneralFormat,
  numberFormatFacets,
  parseNumberFormat,
  splitFormatSections,
  tokenizeFormatSection
} from "@utils/number-format";
import type { FormatToken } from "@utils/number-format";
import { describe, expect, it } from "vitest";

/** The date/time parts a section resolves to, in order — the part of a token list a test cares about. */
function dateParts(section: string): string[] {
  return tokenizeFormatSection(section).flatMap((t: FormatToken) =>
    t.kind === "date" ? [t.part] : t.kind === "elapsed" ? [`[${t.unit}]`] : []
  );
}

describe("splitFormatSections", () => {
  it("does not split on an escaped semicolon", () => {
    expect(splitFormatSections("0\\;0;-0")).toEqual(["0\\;0", "-0"]);
  });

  it("does not split inside quotes or brackets", () => {
    expect(splitFormatSections('"a;b"0;[$;-409]0')).toEqual(['"a;b"0', "[$;-409]0"]);
  });

  it("does not treat an escaped quote as opening a literal", () => {
    expect(splitFormatSections('0\\";-0')).toEqual(['0\\"', "-0"]);
  });

  // `_` and `*` take the next character as their argument, exactly as `\\` does. The splitter used to
  // know only `\\`, so it disagreed with the tokenizer about these three formats.
  it.each([
    ["0_;0", ["0_;0"]],
    ["0*;0", ["0*;0"]],
    ['0_";0', ['0_"', "0"]],
    ["0_[;0", ["0_[", "0"]],
    ["0*];0", ["0*]", "0"]]
  ])("reads the argument of _ or * in %j as text", (fmt, sections) => {
    expect(splitFormatSections(fmt)).toEqual(sections);
  });
});

describe("tokenizeFormatSection", () => {
  describe("literal forms", () => {
    // Every way Excel has of making a character literal, applied to every letter that is otherwise a
    // code. Each scanner this module replaced missed at least one of these combinations.
    const codes = ["y", "m", "M", "d", "h", "s", "b", "E"];
    const wrappers: [string, (c: string) => string][] = [
      ["escaped", c => `\\${c}`],
      ["quoted", c => `"${c}"`],
      ["padding", c => `_${c}`],
      ["fill", c => `*${c}`],
      ["in an unknown bracket tag", c => `[${c}x]`]
    ];
    for (const [name, wrap] of wrappers) {
      it.each(codes)(`reads %s ${name} as no code`, code => {
        const tokens = tokenizeFormatSection(`0.0${wrap(code)}`);
        expect(tokens.some(t => t.kind === "date" || t.kind === "exponent")).toBe(false);
      });
    }

    it("keeps the text of quoted and escaped literals", () => {
      expect(tokenizeFormatSection('\\$0"M"')).toEqual([
        { kind: "literal", text: "$" },
        { kind: "digit", char: "0" },
        { kind: "literal", text: "M" }
      ]);
    });

    it("keeps pad and fill characters apart from literals", () => {
      expect(tokenizeFormatSection("_(*-0")).toEqual([
        { kind: "pad", char: "(" },
        { kind: "fill", char: "-" },
        { kind: "digit", char: "0" }
      ]);
    });

    it("reads an unterminated quote to the end of the section", () => {
      expect(tokenizeFormatSection('0"abc')).toEqual([
        { kind: "digit", char: "0" },
        { kind: "literal", text: "abc" }
      ]);
    });

    it("keeps a bare digit apart from a quoted or escaped one", () => {
      expect(tokenizeFormatSection('8"8"\\8')).toEqual([
        { kind: "numeral", text: "8" },
        { kind: "literal", text: "8" },
        { kind: "literal", text: "8" }
      ]);
    });

    it("leaves e and g as literals so an unquoted word stays a word", () => {
      expect(tokenizeFormatSection("zero").every(t => t.kind === "literal")).toBe(true);
    });
  });

  describe("bracket tags", () => {
    it.each([
      ["[Red]", { kind: "color", name: "Red" }],
      ["[Color12]", { kind: "color", name: "Color12" }],
      ["[>=100]", { kind: "condition", op: ">=", value: 100 }],
      ["[<>-1.5]", { kind: "condition", op: "<>", value: -1.5 }],
      ["[$€-407]", { kind: "locale", symbol: "€" }],
      ["[$-409]", { kind: "locale", symbol: "" }],
      ["[hh]", { kind: "elapsed", unit: "h", width: 2, fraction: 0 }],
      ["[MM]", { kind: "elapsed", unit: "m", width: 2, fraction: 0 }],
      ["[DBNum1]", { kind: "bracket", text: "DBNum1" }]
    ])("reads %s", (source, token) => {
      expect(tokenizeFormatSection(source)).toEqual([token]);
    });

    it("treats B1/B2 calendar switches as silent", () => {
      expect(tokenizeFormatSection("B2yyyy")[0]).toEqual({ kind: "bracket", text: "B2" });
    });
  });

  describe("month or minute", () => {
    it.each([
      ["yyyy-mm-dd hh:mm:ss", ["year", "month", "day", "hour", "minute", "second"]],
      ["mm:ss", ["minute", "second"]],
      ['h "at" mm', ["hour", "minute"]],
      ["hh dd mm", ["hour", "day", "minute"]],
      ["h mmm mm", ["hour", "month", "minute"]],
      ["dd mm", ["day", "month"]],
      ["ss mm", ["second", "minute"]],
      ["h:mm:ss mm", ["hour", "minute", "second", "month"]],
      ["mm dd ss", ["month", "day", "second"]],
      ["[h]:mm", ["[h]", "minute"]],
      ["mm:[ss]", ["month", "[s]"]],
      ["mmm h:m", ["month", "hour", "minute"]],
      ["MM/DD/YYYY", ["month", "day", "year"]],
      ["m", ["month"]]
    ])("resolves %s by position", (source, parts) => {
      expect(dateParts(source)).toEqual(parts);
    });
  });

  it("reads AM/PM and A/P in either case", () => {
    expect(tokenizeFormatSection("AM/PM a/p").filter(t => t.kind === "ampm")).toEqual([
      { kind: "ampm", short: false, lower: false },
      { kind: "ampm", short: true, lower: true }
    ]);
  });

  it("reads E+ and e- as exponents", () => {
    expect(tokenizeFormatSection("0E+0e-0").filter(t => t.kind === "exponent")).toEqual([
      { kind: "exponent", sign: "+", upper: true },
      { kind: "exponent", sign: "-", upper: false }
    ]);
  });

  it("reads General case-insensitively", () => {
    expect(tokenizeFormatSection("GENERAL")).toEqual([{ kind: "general" }]);
  });
});

describe("parseNumberFormat", () => {
  it("takes the fourth section as the text section", () => {
    const parsed = parseNumberFormat("0;-0;0;@");
    expect(parsed.numberSections.map(s => s.source)).toEqual(["0", "-0", "0"]);
    expect(parsed.textSection?.source).toBe("@");
  });

  it("takes a trailing @ section as the text section", () => {
    const parsed = parseNumberFormat("0.00;@");
    expect(parsed.numberSections.map(s => s.source)).toEqual(["0.00"]);
    expect(parsed.textSection?.source).toBe("@");
  });

  it("has no number section for a text-only format", () => {
    const parsed = parseNumberFormat('"Name: "@');
    expect(parsed.numberSections).toEqual([]);
    expect(parsed.textSection?.hasText).toBe(true);
  });

  it("records a section's condition", () => {
    expect(parseNumberFormat("[>=1000]0,K;0").sections[0].condition).toEqual({
      op: ">=",
      value: 1000
    });
  });
});

describe("numberFormatFacets", () => {
  it.each([
    ["yyyy-mm-dd", { date: true, time: false, elapsed: false }],
    ["h:mm AM/PM", { date: false, time: true, elapsed: false }],
    ["m/d/yy h:mm", { date: true, time: true, elapsed: false }],
    ["[h]:mm:ss", { date: false, time: true, elapsed: true }],
    ["0_m", { date: false, time: false, elapsed: false }],
    ["0*m", { date: false, time: false, elapsed: false }],
    ["0_h", { date: false, time: false, elapsed: false }],
    ["\\$0.0,,\\M", { date: false, time: false, elapsed: false }],
    ['0" hrs"', { date: false, time: false, elapsed: false }],
    // Only the first section speaks for a date.
    [";;;dd", { date: false, time: false, elapsed: false }]
  ])("reads %s", (fmt, expected) => {
    expect(numberFormatFacets(fmt)).toMatchObject(expected);
  });

  it("treats a missing format as General", () => {
    expect(numberFormatFacets(undefined).general).toBe(true);
    expect(numberFormatFacets("").general).toBe(true);
  });
});

describe("isDateFmt", () => {
  it.each([
    "DD/MM/YYYY",
    "YYYY",
    "HH:MM",
    "bbbb",
    "[h]:mm:ss",
    "[mm]:ss",
    "AM/PM h",
    "yyyy-mm-dd;@"
  ])("is true for %s", fmt => {
    expect(isDateFmt(fmt)).toBe(true);
  });

  it.each([
    "\\$0.0,,\\M",
    "0_m",
    "0*s",
    "0.0\\h",
    '#,##0" days"',
    "[Red]0",
    "[DBNum1]0",
    "@",
    "@yyyy",
    "General",
    "0.00E+00",
    "zero",
    // An elapsed tag alone is a quantity of time; read as a `Date` it would be misstated.
    "[s]",
    "[h]",
    "[ss].00"
  ])("is false for %s", fmt => {
    expect(isDateFmt(fmt)).toBe(false);
  });
});

describe("isGeneralFormat", () => {
  it.each(["General", "general", " General ", "[Red]General", "[$-409]General", ""])(
    "is true for %j",
    fmt => {
      expect(isGeneralFormat(fmt)).toBe(true);
    }
  );

  it.each(['General" kg"', "General;-General", "0", "[$€-407]General"])("is false for %j", fmt => {
    expect(isGeneralFormat(fmt)).toBe(false);
  });
});

describe("fractions of a second", () => {
  it.each([
    ["ss.00", { kind: "date", part: "second", width: 2, fraction: 2 }],
    ["s.0", { kind: "date", part: "second", width: 1, fraction: 1 }],
    ["[ss].000", { kind: "elapsed", unit: "s", width: 2, fraction: 3 }]
  ])("attaches the decimal places in %s to the seconds code", (fmt, token) => {
    expect(tokenizeFormatSection(fmt)).toEqual([token]);
  });

  it("leaves a point after any other code alone", () => {
    expect(tokenizeFormatSection("mm.00").map(t => t.kind)).toEqual([
      "date",
      "decimal",
      "digit",
      "digit"
    ]);
    expect(tokenizeFormatSection("ss.#").map(t => t.kind)).toEqual(["date", "decimal", "digit"]);
  });
});
