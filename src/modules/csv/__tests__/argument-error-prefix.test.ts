/**
 * An argument error names the public call the caller wrote, and `field` is the path from the
 * parameter name. CSV used to say `CSV parse:` / `CSV format:` / `CSV generate:` whichever
 * function was called.
 */
import { Csv, CsvArgumentError } from "@csv/index";
import { describe, expect, it } from "vitest";

async function caught(call: () => unknown): Promise<CsvArgumentError> {
  try {
    await call();
  } catch (error) {
    expect(error).toBeInstanceOf(CsvArgumentError);
    return error as CsvArgumentError;
  }
  throw new Error("expected a CsvArgumentError");
}

async function drain(rows: AsyncIterable<unknown>): Promise<void> {
  for await (const _ of rows) {
    // read every row
  }
}

describe("CSV argument errors name the public call", () => {
  it.each<[string, () => unknown, string, string]>([
    [
      "Csv.parse",
      () => Csv.parse("a", { delimiter: 5 as never }),
      "Csv.parse:",
      "options.delimiter"
    ],
    [
      "Csv.parseAsync",
      () => Csv.parseAsync("a", { delimiter: 5 as never }),
      "Csv.parseAsync:",
      "options.delimiter"
    ],
    ["Csv.parseAsync input", () => Csv.parseAsync(42 as never), "Csv.parseAsync:", "input"],
    [
      "Csv.parseAsync encoding",
      () => Csv.parseAsync((async function* () {})(), { encoding: "nope" }),
      "Csv.parseAsync:",
      "options.encoding"
    ],
    [
      "Csv.parseStream",
      () => drain(Csv.parseStream("a", { delimiter: 5 as never })),
      "Csv.parseStream:",
      "options.delimiter"
    ],
    [
      "Csv.parseStream objname",
      () => drain(Csv.parseStream("a", { headers: true, objname: "a" })),
      "Csv.parseStream:",
      "options.objname"
    ],
    [
      "new Csv.ParserStream",
      () => new Csv.ParserStream({ delimiter: 5 as never }),
      "new Csv.ParserStream:",
      "options.delimiter"
    ],
    [
      "new Csv.ParserStream encoding",
      () => new Csv.ParserStream({ encoding: "nope" }),
      "new Csv.ParserStream:",
      "options.encoding"
    ],
    ["Csv.format", () => Csv.format([[1]], { delimiter: "" }), "Csv.format:", "options.delimiter"],
    ["Csv.format data", () => Csv.format(42 as never), "Csv.format:", "data"],
    ["Csv.format cell", () => Csv.format([[() => 1]]), "Csv.format:", "data"],
    [
      "Csv.format decimalSeparator",
      () => Csv.format([[1]], { decimalSeparator: ";" as never }),
      "Csv.format:",
      "options.decimalSeparator"
    ],
    [
      "new Csv.FormatterStream",
      () => new Csv.FormatterStream({ delimiter: "" }),
      "new Csv.FormatterStream:",
      "options.delimiter"
    ],
    ["Csv.generate", () => Csv.generate({ rows: -2 }), "Csv.generate:", "options.rows"],
    [
      "Csv.generateRows",
      () => [...Csv.generateRows({ columns: "x" as never })],
      "Csv.generateRows:",
      "options.columns"
    ],
    ["Csv.generateData", () => Csv.generateData(42 as never), "Csv.generateData:", "options"],
    [
      "Csv.createGenerator",
      () => Csv.createGenerator(42 as never),
      "Csv.createGenerator:",
      "options"
    ],
    [
      "Csv.applyDynamicTyping without a value",
      () => (Csv.applyDynamicTyping as (...args: unknown[]) => unknown)(),
      "Csv.applyDynamicTyping:",
      "value"
    ],
    [
      "Csv.applyDynamicTyping with a number",
      () => Csv.applyDynamicTyping(42 as never),
      "Csv.applyDynamicTyping:",
      "value"
    ],
    [
      "Csv.applyDynamicTyping with a bad columnConfig",
      () => Csv.applyDynamicTyping("1", "yes" as never),
      "Csv.applyDynamicTyping:",
      "columnConfig"
    ],
    [
      "Csv.applyDynamicTypingToRow without dynamicTyping",
      () => (Csv.applyDynamicTypingToRow as (...args: unknown[]) => unknown)({ a: "1" }),
      "Csv.applyDynamicTypingToRow:",
      "dynamicTyping"
    ],
    [
      "Csv.applyDynamicTypingToRow with a non-string value",
      () => Csv.applyDynamicTypingToRow({ a: 1 } as never, true),
      "Csv.applyDynamicTypingToRow:",
      'row["a"]'
    ],
    [
      "new Csv.FormatterStream with null options",
      () => new Csv.FormatterStream(null as never),
      "new Csv.FormatterStream:",
      "options"
    ]
  ])("%s", async (_label, call, prefix, field) => {
    const error = await caught(call);
    expect(error.message.startsWith(prefix), error.message).toBe(true);
    expect(error.field).toBe(field);
  });

  it("a formatter stream names itself for a bad row", async () => {
    const stream = new Csv.FormatterStream();
    const error = await new Promise<Error>(resolve => {
      stream.on("error", resolve);
      stream.write(42 as never);
    });
    expect(error).toBeInstanceOf(CsvArgumentError);
    expect(error.message.startsWith("new Csv.FormatterStream:")).toBe(true);
  });

  it("dynamic typing still applies to valid input, and the row is untouched when refused", () => {
    expect(Csv.applyDynamicTyping("42")).toBe(42);
    expect(Csv.applyDynamicTyping("42", false)).toBe("42");
    const row = { a: "1", b: 2 } as never as Record<string, string>;
    expect(() => Csv.applyDynamicTypingToRow(row, true)).toThrow(CsvArgumentError);
    expect(row).toEqual({ a: "1", b: 2 });
    expect(Csv.applyDynamicTypingToRow({ a: "1" }, true)).toEqual({ a: 1 });
  });
});
