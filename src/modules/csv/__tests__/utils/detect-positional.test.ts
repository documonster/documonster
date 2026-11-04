/**
 * `Csv.detectDelimiter` / `Csv.detectLinebreak` keep their deprecated positional forms beside the
 * options object, and a positional placeholder keeps the meaning it had there. Above all a `null`
 * quote: the positional form defaulted only `undefined`, so for `detectLinebreak` `null` meant
 * "no quote character" and a line break inside quotes counts.
 *
 * Expected values were observed against the last release (c3b2407c).
 */
import { Csv } from "@csv/index";
import { describe, expect, it } from "vitest";

function loose<T>(value: unknown): T {
  return value as T;
}

describe("Csv.detectLinebreak positional quote", () => {
  const quoted = '"a\r\nb"\nc';

  it("null disables quote handling, undefined defaults to a double quote", () => {
    expect(Csv.detectLinebreak(quoted, loose<string>(null))).toBe("\r\n");
    expect(Csv.detectLinebreak(quoted, loose<string>(undefined))).toBe("\n");
    expect(Csv.detectLinebreak(quoted, "'")).toBe("\r\n");
  });

  it("the options form is unaffected", () => {
    expect(Csv.detectLinebreak(quoted, {})).toBe("\n");
    expect(Csv.detectLinebreak(quoted, { quote: "'" })).toBe("\r\n");
  });
});

describe("Csv.detectDelimiter positional quote", () => {
  const colons = "a:b:c\n1:2:3\n";

  it("a null or undefined quote still reads the later positional arguments", () => {
    expect(Csv.detectDelimiter(colons, loose<string>(null), [":"])).toBe(":");
    expect(Csv.detectDelimiter(colons, undefined, [":"])).toBe(":");
    expect(Csv.detectDelimiter(colons, "'", [":"], "#", true)).toBe(":");
    expect(Csv.detectDelimiter("'a,b';c\n1;2", "'")).toBe(";");
    expect(Csv.detectDelimiter("a;b\n1;2", '"', loose<string[]>(null), loose<string>(null))).toBe(
      ";"
    );
  });

  it("a null quote over quoted fields", () => {
    expect(Csv.detectDelimiter('"a;b",c,d\n"1;2",3,4', loose<string>(null))).toBe(",");
  });

  it("the options form is unaffected", () => {
    expect(Csv.detectDelimiter(colons, { delimitersToGuess: [":"] })).toBe(":");
    expect(Csv.detectDelimiter("'a,b';c\n1;2", { quote: "'" })).toBe(";");
  });
});
