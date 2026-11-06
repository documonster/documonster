/**
 * Every parser and formatter checks its options before reading or writing a row: a value
 * with no reading is a `CsvError` naming the option. A formatter given `delimiter: ""` used
 * to write fields with nothing between them, and `quote` equal to the delimiter wrote a file
 * nothing could split again.
 */
import { Csv, CsvError } from "@csv/index";
import { describe, expect, it } from "vitest";

type Options = Record<string, unknown>;

function expectCsvError(call: () => unknown, option: string): void {
  let caught: unknown;
  try {
    call();
  } catch (err) {
    caught = err;
  }
  expect(caught).toBeInstanceOf(CsvError);
  expect((caught as Error).message).toContain(`${option} must`);
}

describe("Csv.format option checks", () => {
  const rows = [
    ["a", "b"],
    ["1", "2"]
  ];

  it.each<[string, Options]>([
    ["delimiter", { delimiter: "" }],
    ["delimiter", { delimiter: 1 }],
    ["quote", { quote: "" }],
    ["quote", { quote: "''" }],
    ["quote", { quote: "\n" }],
    ["quote", { quote: "," }],
    ["quote", { delimiter: ";", quote: ";" }],
    ["escape", { escape: "\\\\" }],
    ["escape", { escape: "," }],
    ["lineEnding", { lineEnding: "" }],
    ["lineEnding", { delimiter: "|", lineEnding: "|" }]
  ])("refuses %s in %j", (option, options) => {
    expectCsvError(() => Csv.format(rows, options as Parameters<typeof Csv.format>[1]), option);
    expectCsvError(
      () =>
        new Csv.FormatterStream(options as ConstructorParameters<typeof Csv.FormatterStream>[0]),
      option
    );
  });

  it("keeps accepting the options it has always written correctly", () => {
    expect(Csv.format(rows, { delimiter: "||" })).toBe("a||b\n1||2");
    expect(Csv.format(rows, { delimiter: "\t", quote: "'", escape: "\\" })).toBe("a\tb\n1\t2");
    expect(Csv.format([["x,y"]], { quote: false })).toBe("x,y");
    expect(Csv.format([['a"b']], { escape: "" })).toBe('"a""b"');
    expect(Csv.format(rows, { lineEnding: "\r\n" })).toBe("a,b\r\n1,2");
  });
});

describe("Csv.parse option checks", () => {
  it.each([42, null, {}, new Uint8Array(1)])("refuses input %j that is not text", bad => {
    expect(() => Csv.parse(bad as never)).toThrow(CsvError);
  });

  it.each<[string, Options]>([
    ["delimiter", { delimiter: 5 }],
    ["quote", { quote: '""' }],
    ["quote", { quote: "," }],
    ["quote", { delimiter: "||", quote: "|" }],
    ["escape", { escape: "\t", delimiter: "\t" }],
    ["delimitersToGuess[1]", { delimiter: "", delimitersToGuess: [",", ""] }],
    ["lineEnding", { lineEnding: 0 }],
    ["skipEmptyLines", { skipEmptyLines: "yes" }],
    ["maxRows", { maxRows: -1 }],
    ["toLine", { toLine: 1.5 }],
    ["skipLines", { skipLines: NaN }],
    ["skipRows", { skipRows: "2" }],
    ["maxRowBytes", { maxRowBytes: 0 }],
    ["chunkSize", { chunkSize: 0 }],
    ["columnMismatch.less", { columnMismatch: { less: "drop", more: "error" } }],
    ["columnMismatch.more", { columnMismatch: { less: "pad", more: "drop" } }],
    ["validate", { validate: true }]
  ])("refuses %s in %j", (option, options) => {
    expectCsvError(() => Csv.parse("a,b\n1,2", options as Parameters<typeof Csv.parse>[1]), option);
    expectCsvError(
      () => new Csv.ParserStream(options as ConstructorParameters<typeof Csv.ParserStream>[0]),
      option
    );
  });

  it("rejects parseAsync with the same error", async () => {
    await expect(Csv.parseAsync("a,b", { maxRows: NaN })).rejects.toBeInstanceOf(CsvError);
  });

  it("keeps multi-character delimiters, auto-detection and a disabled quote", () => {
    expect(Csv.parse("a||b", { delimiter: "||" })).toEqual([["a", "b"]]);
    expect(Csv.parse("a;b", { delimiter: "" })).toEqual([["a", "b"]]);
    expect(Csv.parse('a,"b', { quote: false })).toEqual([["a", '"b']]);
    expect(Csv.parse('a,"b', { quote: "" })).toEqual([["a", '"b']]);
    expect(Csv.parse("a,b\n1,2", { maxRows: 0, skipLines: 0, toLine: 2 })).toEqual([]);
  });
});

describe("Csv.generate option checks", () => {
  it.each<[string, Options]>([
    // `rows: NaN` never satisfied the stop condition, so this generated forever.
    ["rows", { rows: NaN }],
    ["rows", { rows: 2.5 }],
    ["columns", { columns: -1 }],
    ["delimiter", { delimiter: "" }],
    ["quote", { quote: "sometimes" }],
    ["floatPrecision", { floatPrecision: 101 }],
    ["seed", { seed: NaN }],
    ["until", { until: new Date(NaN) }]
  ])("refuses %s in %j", (option, options) => {
    expectCsvError(() => Csv.generate(options as Parameters<typeof Csv.generate>[0]), option);
  });

  it("throws CsvError for unlimited rows in the synchronous form", () => {
    expect(() => Csv.generate({ rows: -1 })).toThrow(CsvError);
  });

  it("keeps generating at the boundaries", () => {
    expect(Csv.generate({ rows: 0, columns: 0, seed: 1 }).data).toEqual([]);
    expect([...Csv.generateRows({ rows: 3, skipRows: 0, seed: 1, headers: false })]).toHaveLength(
      3
    );
  });
});

/** Refusals that changed when the checks moved to `@utils/schema`. */
describe("Csv option checks declared as schemas", () => {
  function refusal(call: () => unknown): { field: string; message: string } {
    try {
      call();
    } catch (err) {
      return { field: (err as { field: string }).field, message: (err as Error).message };
    }
    throw new Error("expected a refusal");
  }

  it("names the delimiter guess that holds a line break, not `delimiter`", () => {
    expect(refusal(() => Csv.parse("a", { delimiter: "", delimitersToGuess: ["\n"] }))).toEqual({
      field: "options.delimitersToGuess[0]",
      message:
        'Csv.parse: options.delimitersToGuess[0] must be a string without line breaks, got "\\n"'
    });
  });

  it("refuses an array as columnMismatch", () => {
    expect(refusal(() => Csv.parse("a", { columnMismatch: [] as never })).field).toBe(
      "options.columnMismatch"
    );
  });

  it("checks firstRowIsHeader", () => {
    expect(refusal(() => Csv.format([["a"]], { firstRowIsHeader: 1 as never })).message).toBe(
      "Csv.format: options.firstRowIsHeader must be a boolean, got 1"
    );
  });

  it("describes a refused decimalSeparator as every other option does", () => {
    expect(refusal(() => Csv.format([["a"]], { decimalSeparator: NaN as never })).message).toBe(
      'Csv.format: options.decimalSeparator must be "." or ",", got NaN'
    );
  });

  it("checks headers, bom and objectMode of the generators, and describes values", () => {
    expect(refusal(() => Csv.generate({ rows: 1, headers: [1] as never }))).toEqual({
      field: "options.headers[0]",
      message: "Csv.generate: options.headers[0] must be a string, got 1"
    });
    expect(refusal(() => Csv.generateData({ bom: 1 as never })).field).toBe("options.bom");
    expect(refusal(() => Csv.generateData({ objectMode: 1 as never })).field).toBe(
      "options.objectMode"
    );
    expect(refusal(() => Csv.generate({ seed: {} as never })).message).toBe(
      "Csv.generate: options.seed must be a finite number, got an object"
    );
    expect(refusal(() => Csv.generate(5 as never)).message).toBe(
      "Csv.generate: options must be an options object, got 5"
    );
  });

  it("refuses a row of applyDynamicTypingToRow that is not a column map, saying what it got", () => {
    expect(refusal(() => Csv.applyDynamicTypingToRow(5 as never, true))).toEqual({
      field: "row",
      message: "Csv.applyDynamicTypingToRow: row must be an object of column values, got 5"
    });
    expect(refusal(() => Csv.applyDynamicTypingToRow(["1"] as never, true)).field).toBe("row");
  });

  it("still accepts an undefined entry in quoteColumns", () => {
    expect(Csv.format([["a"]], { quoteColumns: [undefined as never] })).toBe("a");
  });
});

describe("Csv argument checks completed with schemas", () => {
  function refusal(call: () => unknown): { field: string; message: string } {
    try {
      call();
    } catch (err) {
      return { field: (err as { field: string }).field, message: (err as Error).message };
    }
    throw new Error("expected a refusal");
  }

  it("checks the quote of Csv.detectLinebreak", () => {
    expect(refusal(() => Csv.detectLinebreak("a\nb", { quote: 1 as never }))).toEqual({
      field: "options.quote",
      message: "Csv.detectLinebreak: options.quote must be a string, got 1"
    });
  });

  it("describes a refused row or header list by its value", () => {
    expect(refusal(() => Csv.getRowValues({} as never)).message).toBe(
      "Csv.getRowValues: row must be an array, got an object"
    );
    expect(refusal(() => Csv.deduplicateHeaders("a" as never)).message).toBe(
      'Csv.deduplicateHeaders: headers must be an array, got "a"'
    );
  });

  it("reports a wrong typeTransform.row result at the option", () => {
    expect(
      refusal(() => Csv.format([["a"]], { typeTransform: { row: () => 5 as never } }))
    ).toEqual({
      field: "options.typeTransform.row",
      message: "Csv.format: typeTransform.row must return an array or an object for data[0], got 5"
    });
  });
});
