/** `Csv.ParserStream#transform` and `#validate` refuse a non-function as the module's argument error. */
import { Csv, CsvArgumentError } from "@csv/index";
import { describe, expect, it } from "vitest";

describe("Csv.ParserStream row hooks", () => {
  it.each([
    ["transform", "transformFunction"],
    ["validate", "validateFunction"]
  ] as const)("#%s refuses a non-function, naming %s", (method, field) => {
    const parser = new Csv.ParserStream();
    const call = (parser[method] as (fn: unknown) => unknown).bind(parser);
    try {
      expect(() => call(42)).toThrow(CsvArgumentError);
      expect(() => call(42)).toThrow(expect.objectContaining({ field }));
      expect(() => call(42)).toThrow(`Csv.ParserStream#${method}`);
    } finally {
      parser.destroy();
    }
  });
});
