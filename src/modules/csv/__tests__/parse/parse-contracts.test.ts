/**
 * Contracts of the CSV parsers that the API review found broken: text after a closing quote
 * silently dropped, columns lost for their header's name, options ignored or misread,
 * `chunk` accepted by parsers that never call it, malformed records dropped by `parseRows`
 * without a word, and dynamic typing that rounded IDs and read date-only values in the
 * machine's timezone.
 */
import { Csv, CsvArgumentError, CsvError, isCsvError } from "@csv/index";
import type { CsvParseResultWithObjname } from "@csv/index";
import { describe, expect, expectTypeOf, it } from "vitest";

async function collect<T>(gen: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const row of gen) {
    out.push(row);
  }
  return out;
}

describe("text after a closing quote is a record error", () => {
  it("is reported and the record left out, in object mode", () => {
    const result = Csv.parse('h1,h2\n"a"b,c\nx,y', { headers: true });
    expect(result.rows).toEqual([{ h1: "x", h2: "y" }]);
    expect(result.errors).toEqual([
      {
        code: "ParseError",
        message: "Unexpected text after a closing quote",
        line: 2,
        raw: '"a"b,c'
      }
    ]);
    expect(result.invalidRows).toEqual([
      { row: ["a", "c"], reason: "Unexpected text after a closing quote" }
    ]);
  });

  it("reaches onSkip, also from Csv.parseRows and the stream", async () => {
    const skipped: string[] = [];
    const onSkip = (e: { code: string }) => skipped.push(e.code);
    Csv.parse('"a" b,c\nx,y', { onSkip });
    const rows = await collect(Csv.parseRows('"a" b,c\nx,y', { onSkip }));
    expect(rows).toEqual([["x", "y"]]);
    expect(skipped).toEqual(["ParseError", "ParseError"]);
  });

  it("is reported with info in array mode", () => {
    const result = Csv.parse('"a"b,c', { info: true });
    expect(result.rows).toEqual([]);
    expect(result.errors?.[0]?.code).toBe("ParseError");
  });

  it("does not add a field when the text runs to the line break", () => {
    const result = Csv.parse('"a"b\nc', { info: true });
    expect(result.invalidRows).toEqual([
      { row: ["a"], reason: "Unexpected text after a closing quote" }
    ]);
    expect(result.rows.map(r => r.record)).toEqual([["c"]]);
  });

  it("is still read as a quote with relaxQuotes", () => {
    // The mid-field quote is literal, so the field runs to the end unterminated: the record is
    // kept beside a MissingQuotes error where the result has `errors`.
    const result = Csv.parse('"a"b,c', { relaxQuotes: true, info: true });
    expect(result.rows.map(r => r.record)).toEqual([['a"b,c']]);
    expect(result.errors?.map(e => e.code)).toEqual(["MissingQuotes"]);
  });
});

describe("every header keeps its column", () => {
  const src = '"Total\nAmount",constructor,prototype,__proto__,"a\tb",ok\n1,2,3,4,5,6';
  const expected = ["Total\nAmount", "constructor", "prototype", "__proto__", "a\tb", "ok"];

  it("in Csv.parse", () => {
    const row = Csv.parse(src, { headers: true }).rows[0];
    expect(Object.keys(row)).toEqual(expected);
    expect(Object.values(row)).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(Object.getPrototypeOf(row)).toBeNull();
  });

  it("in Csv.parseRows", async () => {
    const [row] = await collect(Csv.parseRows(src, { headers: true }));
    expect(Object.keys(row as object)).toEqual(expected);
  });

  it("for a long header", () => {
    const long = "x".repeat(600);
    expect(Object.keys(Csv.parse(`${long},b\n1,2`, { headers: true }).rows[0])).toEqual([
      long,
      "b"
    ]);
  });

  it("round-trips through Csv.format", () => {
    const rows = [{ "Total\nAmount": 1, constructor: 2 }];
    const parsed = Csv.parse(Csv.format(rows), { headers: true }).rows;
    expect(parsed.map(r => ({ ...r }))).toEqual([{ "Total\nAmount": "1", constructor: "2" }]);
  });

  it("in Csv.getRowValuesByHeaders for a wide row", () => {
    const headers = Array.from({ length: 11 }, (_, i) => `c${i}`).concat("constructor");
    const row = headers.map((h, i) => [h, i] as [string, number]);
    expect(Csv.getRowValuesByHeaders(row, headers).at(-1)).toBe(11);
  });

  it("for an objname key", () => {
    const rows = Csv.parse("id,v\nconstructor,1\n__proto__,2", {
      headers: true,
      objname: "id"
    }).rows as Record<string, unknown>;
    expect(Object.keys(rows)).toEqual(["constructor", "__proto__"]);
  });
});

describe("parse options are checked", () => {
  it.each([
    ["headers as a number", { headers: 42 }],
    ["headers as a string", { headers: "x" }],
    ["a header that is not a string", { headers: ["a", 1] }],
    ["dynamicTyping as a string", { dynamicTyping: "x" }],
    ["a dynamicTyping column that is not boolean or function", { dynamicTyping: { a: 1 } }],
    ["castDate as a number", { castDate: 42 }],
    ["objname without headers", { objname: "a" }],
    ["objname as a number", { headers: true, objname: 42 }],
    ["a delimiter holding a line break", { delimiter: "\n" }],
    ["a lineEnding equal to the delimiter", { lineEnding: "," }],
    ["trim as a string", { trim: "yes" }],
    ["info as a string", { info: "yes" }],
    ["decimalSeparator as ;", { delimiter: ";", decimalSeparator: ";" }],
    ["decimalSeparator equal to the delimiter", { decimalSeparator: "," }]
  ])("%s", (_label, options) => {
    expect(() => Csv.parse("a,b\n1,2", options as never)).toThrow(CsvError);
  });

  it("refuses an objname that is not a header", () => {
    expect(() => Csv.parse("a,b\n1,2", { headers: true, objname: "zz" })).toThrow(CsvError);
  });

  it("lets a later objname key replace an earlier one, as documented", () => {
    expect(Csv.parse("id,v\nx,1\nx,2", { headers: true, objname: "id" }).rows).toEqual({
      x: { id: "x", v: "2" }
    });
  });

  it("types an objname result as an object of records, and a validate result as a result", async () => {
    const byId = Csv.parse("id,v\nx,1", { headers: true, objname: "id" });
    expectTypeOf(byId).toEqualTypeOf<CsvParseResultWithObjname<Record<string, unknown>>>();
    const byIdAsync = await Csv.parseAsync("id,v\nx,1", { headers: true, objname: "id" });
    expectTypeOf(byIdAsync).toEqualTypeOf<CsvParseResultWithObjname<Record<string, unknown>>>();
    const validated = Csv.parse("a\nb", { validate: () => true });
    expect(Array.isArray(validated)).toBe(false);
    expectTypeOf(validated.invalidRows).not.toBeAny();
  });
});

describe("chunk is refused where it would never be called", () => {
  it("by Csv.parse and Csv.parseAsync", async () => {
    expect(() => Csv.parse("a\nb", { chunk: () => undefined })).toThrow(CsvError);
    expect(() => Csv.parse("a\nb", { chunkSize: 2 })).toThrow(CsvError);
    await expect(Csv.parseAsync("a\nb", { chunk: () => undefined })).rejects.toThrow(CsvError);
  });

  it("but still works in Csv.parseRows", async () => {
    const chunks: unknown[] = [];
    await collect(
      Csv.parseRows("a\nb\nc", { chunkSize: 2, chunk: rows => void chunks.push(rows) })
    );
    expect(chunks).toEqual([[["a"], ["b"]], [["c"]]]);
  });
});

describe("a record left out because of an error always reaches onSkip", () => {
  it("in Csv.parse without skipRecordsWithError", () => {
    const skipped: unknown[] = [];
    const result = Csv.parse("a,b\n1\n3,4", {
      headers: true,
      onSkip: (e, r) => skipped.push([e.code, r])
    });
    expect(result.rows).toEqual([{ a: "3", b: "4" }]);
    expect(skipped).toEqual([["TooFewFields", ["1"]]]);
  });

  it("in Csv.parseRows, whose only report it is", async () => {
    const skipped: unknown[] = [];
    const rows = await collect(
      Csv.parseRows("a,b\n1\n3,4", { headers: true, onSkip: (e, r) => skipped.push([e.code, r]) })
    );
    expect(rows).toEqual([{ a: "3", b: "4" }]);
    expect(skipped).toEqual([["TooFewFields", ["1"]]]);
  });
});

describe("inputs are checked with CsvError", () => {
  // Bytes are an input now (`byte-input.test.ts`); what is no input at all is still refused.
  it.each([42, null, undefined, { a: 1 }])("Csv.parseAsync(%s)", async input => {
    await expect(Csv.parseAsync(input as never)).rejects.toThrow(CsvError);
  });

  it.each([42, null])("Csv.parseRows(%s)", async input => {
    await expect(collect(Csv.parseRows(input as never))).rejects.toThrow(CsvError);
  });

  it("onProgress that is not a function", async () => {
    await expect(Csv.parseAsync("a", { onProgress: 42 as never })).rejects.toThrow(CsvError);
  });

  it("detection helpers", () => {
    expect(() => Csv.detectDelimiter(42 as never)).toThrow(CsvError);
    expect(() => Csv.detectDelimiter("a;b", { delimitersToGuess: "x" as never })).toThrow(CsvError);
    expect(() => Csv.detectLinebreak(null as never)).toThrow(CsvError);
    expect(() => Csv.stripBom(42 as never)).toThrow(CsvError);
  });

  it("number helpers", () => {
    expect(() => Csv.formatNumber(1.5, ";" as never)).toThrow(CsvError);
    expect(() => Csv.formatNumber("1.5" as never, ",")).toThrow(CsvError);
    expect(() => Csv.parseNumber(15 as never, ".")).toThrow(CsvError);
    expect(() => Csv.parseNumber("1;5", ";" as never)).toThrow(CsvError);
    expect(Csv.parseNumber("", ".")).toBeNaN();
  });
});

describe("dynamic typing", () => {
  it("keeps integers no double holds exactly as text", () => {
    expect(
      Csv.parse(
        "id\n9007199254740993\n12345678901234567890\n9007199254740991\n1e400\n100000000000000000000",
        {
          headers: true,
          dynamicTyping: true
        }
      ).rows
    ).toEqual([
      { id: "9007199254740993" },
      { id: "12345678901234567890" },
      { id: 9007199254740991 },
      { id: "1e400" },
      { id: 1e20 }
    ]);
  });

  it("reads a date-only value as UTC midnight, wherever it runs", () => {
    const [row] = Csv.parse("d\n2020-01-02", { headers: true, castDate: true }).rows;
    expect((row.d as Date).toISOString()).toBe("2020-01-02T00:00:00.000Z");
    expect(Csv.format([row])).toBe("d\n2020-01-02T00:00:00.000Z");
  });

  it("keeps a year below 100 as written", () => {
    const [row] = Csv.parse("d\n0050-03-04", { headers: true, castDate: true }).rows;
    expect((row.d as Date).toISOString()).toBe("0050-03-04T00:00:00.000Z");
  });

  it("reads numbers with decimalSeparator ','", () => {
    const csv = Csv.format([[1.5, 1000.25, "1.000"]], { delimiter: ";", decimalSeparator: "," });
    expect(Csv.parse(csv, { delimiter: ";", decimalSeparator: ",", dynamicTyping: true })).toEqual([
      [1.5, 1000.25, "1.000"]
    ]);
  });

  it("applies decimalSeparator in the streaming parser and per column", async () => {
    expect(
      await collect(
        Csv.parseRows("a;b\n1,5;2,5", {
          delimiter: ";",
          headers: true,
          decimalSeparator: ",",
          dynamicTyping: { a: true }
        })
      )
    ).toEqual([{ a: 1.5, b: "2,5" }]);
  });
});

describe("r3: dynamicTyping reads only its own entries (TX-B1)", () => {
  const src = "__proto__,constructor,toString,a\n1,2,3,4";

  it("leaves columns named after Object.prototype members alone", () => {
    const [row] = Csv.parse(src, { headers: true, dynamicTyping: { a: true } }).rows;
    expect(Object.entries(row)).toEqual([
      ["__proto__", "1"],
      ["constructor", "2"],
      ["toString", "3"],
      ["a", 4]
    ]);
    expect(typeof row.constructor).toBe("string");
  });

  it("in Csv.applyDynamicTypingToRow, which also checks its config", () => {
    const row = Csv.applyDynamicTypingToRow(
      { constructor: "5", toString: "6", a: "7" },
      { a: true }
    );
    expect(row).toEqual({ constructor: "5", toString: "6", a: 7 });
    expect(() => Csv.applyDynamicTypingToRow({ a: "1" }, { a: 1 } as never)).toThrow(CsvError);
  });
});

describe("r3: objname is refused where it cannot be honoured (TX-B6)", () => {
  it("by Csv.parseRows and the parser stream", async () => {
    await expect(
      collect(Csv.parseRows("id,v\nx,1\nx,2", { headers: true, objname: "id" }))
    ).rejects.toThrow(CsvError);
    expect(() => Csv.createParserStream({ headers: true, objname: "id" })).toThrow(CsvError);
  });
});

describe("r3: a record error never vanishes (TX-C1)", () => {
  it("is thrown by the plain-array Csv.parse / parseAsync without onSkip", async () => {
    expect(() => Csv.parse('"a" b,c')).toThrow(CsvError);
    expect(() => Csv.parse('a\n"x')).toThrow(/MissingQuotes/);
    await expect(Csv.parseAsync('"a" b,c')).rejects.toThrow(CsvError);
  });

  it("reaches onSkip instead when one is given", () => {
    const skipped: string[] = [];
    expect(Csv.parse('a\n"x\n', { onSkip: e => void skipped.push(e.code) })).toEqual([["a"]]);
    expect(skipped).toEqual(["MissingQuotes"]);
  });

  it("is kept beside the record where the result has errors", () => {
    const result = Csv.parse('a\n"x', { info: true });
    expect(result.rows.map(r => r.record)).toEqual([["a"], ["x"]]);
    expect(result.errors?.map(e => e.code)).toEqual(["MissingQuotes"]);
  });

  it("is thrown by Csv.parseRows without onSkip", async () => {
    await expect(collect(Csv.parseRows("a,b\n1\n3,4", { headers: true }))).rejects.toThrow(
      CsvError
    );
  });
});

describe("X12: argument errors are CsvArgumentError with a field", () => {
  function caught(fn: () => unknown): CsvArgumentError {
    try {
      fn();
    } catch (error) {
      expect(error).toBeInstanceOf(CsvArgumentError);
      expect(error).toBeInstanceOf(CsvError);
      expect(isCsvError(error)).toBe(true);
      return error as CsvArgumentError;
    }
    throw new Error("expected an argument error");
  }

  it.each<[string, () => unknown, string]>([
    ["parse input", () => Csv.parse(42 as never), "input"],
    ["parse option", () => Csv.parse("a", { delimiter: "\n" }), "options.delimiter"],
    [
      "nested option",
      () => Csv.parse("a", { columnMismatch: { less: "x" } as never }),
      "options.columnMismatch.less"
    ],
    ["objname", () => Csv.parse("a\n1", { headers: true, objname: "zz" }), "options.objname"],
    ["chunk", () => Csv.parse("a", { chunkSize: 2 }), "options.chunkSize"],
    ["format row", () => Csv.format([["a"], "b"] as never), "data[1]"],
    ["format option", () => Csv.format([["a"]], { bom: "yes" as never }), "options.bom"],
    ["format cell", () => Csv.format([[Symbol("x")]] as never), "data"],
    ["forceQuote", () => Csv.forceQuote(1 as never), "value"],
    ["parseNumber", () => Csv.parseNumber("1", ";" as never), "decimalSeparator"],
    ["stripBom", () => Csv.stripBom(1 as never), "input"],
    ["getRowValues", () => Csv.getRowValues(1 as never), "row"],
    ["getRowValuesByHeaders", () => Csv.getRowValuesByHeaders([], 1 as never), "headers"],
    ["deduplicateHeaders", () => Csv.deduplicateHeaders(1 as never), "headers"],
    ["applyDynamicTypingToRow", () => Csv.applyDynamicTypingToRow(1 as never, true), "row"],
    ["generate", () => Csv.generate({ rows: -1 }), "options.rows"],
    ["generateData", () => Csv.generateData({ rows: -1 }), "options.rows"],
    [
      "parser stream",
      () => Csv.createParserStream({ headers: true, objname: "a" }),
      "options.objname"
    ]
  ])("%s", (_label, fn, field) => {
    const error = caught(fn);
    expect(error.field).toBe(field);
    expect(error.toJSON()).toMatchObject({ name: "CsvArgumentError", field });
  });

  it("is used for parseAsync input", async () => {
    await expect(Csv.parseAsync(42 as never)).rejects.toMatchObject({ field: "input" });
  });

  it("is not used for errors in the input data", () => {
    let error: unknown;
    try {
      Csv.parse('"a" b');
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(CsvError);
    expect(error).not.toBeInstanceOf(CsvArgumentError);
  });
});
