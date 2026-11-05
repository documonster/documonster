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
