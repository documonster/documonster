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
      "Csv.parseRows",
      () => drain(Csv.parseRows("a", { delimiter: 5 as never })),
      "Csv.parseRows:",
      "options.delimiter"
    ],
    [
      "Csv.parseRows objname",
      () => drain(Csv.parseRows("a", { headers: true, objname: "a" })),
      "Csv.parseRows:",
      "options.objname"
    ],
    [
      "Csv.createParserStream",
      () => Csv.createParserStream({ delimiter: 5 as never }),
      "Csv.createParserStream:",
      "options.delimiter"
    ],
    [
      "Csv.createParserStream encoding",
      () => Csv.createParserStream({ encoding: "nope" }),
      "Csv.createParserStream:",
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
      "Csv.createFormatterStream",
      () => Csv.createFormatterStream({ delimiter: "" }),
      "Csv.createFormatterStream:",
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
    ]
  ])("%s", async (_label, call, prefix, field) => {
    const error = await caught(call);
    expect(error.message.startsWith(prefix), error.message).toBe(true);
    expect(error.field).toBe(field);
  });

  it("a formatter stream names itself for a bad row", async () => {
    const stream = Csv.createFormatterStream();
    const error = await new Promise<Error>(resolve => {
      stream.on("error", resolve);
      stream.write(42 as never);
    });
    expect(error).toBeInstanceOf(CsvArgumentError);
    expect(error.message.startsWith("Csv.createFormatterStream:")).toBe(true);
  });
});
