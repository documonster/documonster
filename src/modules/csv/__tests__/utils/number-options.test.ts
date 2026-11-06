/**
 * `Csv.formatNumber` / `Csv.parseNumber` take the separator as `{ decimalSeparator }`, and
 * `Csv.applyDynamicTypingToRow` takes `castDate` as `{ castDate }`; the positional forms stay as
 * deprecated overloads with the same results.
 *
 * Expected values were observed against the last release (v0.21.1).
 */
import { Csv } from "@csv/index";
import { describe, expect, it } from "vitest";

describe("Csv.formatNumber / Csv.parseNumber options", () => {
  it("the options form and the positional form agree", () => {
    expect(Csv.formatNumber(1234.5, { decimalSeparator: "," })).toBe("1234,5");
    expect(Csv.formatNumber(1234.5, ",")).toBe("1234,5");
    expect(Csv.formatNumber(1234.5, { decimalSeparator: "." })).toBe("1234.5");
    expect(Csv.formatNumber(1234.5, ".")).toBe("1234.5");

    expect(Csv.parseNumber(" 1234,5 ", { decimalSeparator: "," })).toBe(1234.5);
    expect(Csv.parseNumber(" 1234,5 ", ",")).toBe(1234.5);
    expect(Csv.parseNumber("1.5", { decimalSeparator: "," })).toBe(1.5);
    expect(Csv.parseNumber("1.5", ",")).toBe(1.5);
    expect(Csv.parseNumber("1,5", { decimalSeparator: "." })).toBeNaN();
    expect(Csv.parseNumber("1,5", ".")).toBeNaN();
  });

  it('defaults to "." with no options or an empty one, as the parser does', () => {
    expect(Csv.formatNumber(3.14)).toBe("3.14");
    expect(Csv.formatNumber(3.14, {})).toBe("3.14");
    expect(Csv.parseNumber("3.14")).toBe(3.14);
    expect(Csv.parseNumber("3,14", {})).toBeNaN();
  });

  it("names the field the way the caller wrote it", () => {
    expect(() => Csv.formatNumber(1, { decimalSeparator: ";" as never })).toThrow(
      expect.objectContaining({ field: "options.decimalSeparator" })
    );
    expect(() => Csv.formatNumber(1, ";" as never)).toThrow(
      expect.objectContaining({ field: "decimalSeparator" })
    );
  });
});

describe("Csv.applyDynamicTypingToRow options", () => {
  function row(): Record<string, string> {
    return { joined: "2024-03-01", note: "2024-03-01", n: "1,5" };
  }

  it("the options form and the positional castDate agree", () => {
    const positional = Csv.applyDynamicTypingToRow(row(), true, ["joined"]);
    const options = Csv.applyDynamicTypingToRow(row(), true, { castDate: ["joined"] });
    expect(options).toEqual(positional);
    expect(options.joined).toBeInstanceOf(Date);
    expect(options.note).toBe("2024-03-01");

    expect(Csv.applyDynamicTypingToRow(row(), false, { castDate: true })).toEqual(
      Csv.applyDynamicTypingToRow(row(), false, true)
    );
    expect(Csv.applyDynamicTypingToRow(row(), true, {})).toEqual(
      Csv.applyDynamicTypingToRow(row(), true)
    );
    expect(Csv.applyDynamicTypingToRow(row(), true, undefined)).toEqual(
      Csv.applyDynamicTypingToRow(row(), true)
    );
  });

  it("reads numbers with options.decimalSeparator, as the parser does", () => {
    expect(Csv.applyDynamicTypingToRow(row(), { n: true }).n).toBe("1,5");
    expect(Csv.applyDynamicTypingToRow(row(), { n: true }, { decimalSeparator: "," }).n).toBe(1.5);
    expect(
      Csv.parse("n\n1,5", {
        headers: true,
        delimiter: ";",
        dynamicTyping: true,
        decimalSeparator: ","
      }).rows
    ).toEqual([{ n: 1.5 }]);
  });

  it("names dynamicTyping by its parameter, not as an option", () => {
    const fieldOf = (fn: () => unknown) => {
      try {
        fn();
      } catch (error) {
        return (error as { field?: string }).field;
      }
      return undefined;
    };
    expect(fieldOf(() => Csv.applyDynamicTypingToRow(row(), "yes" as never))).toBe("dynamicTyping");
    expect(fieldOf(() => Csv.applyDynamicTypingToRow(row(), { n: 1 } as never))).toBe(
      'dynamicTyping["n"]'
    );
    expect(fieldOf(() => Csv.parse("a", { dynamicTyping: "yes" as never }))).toBe(
      "options.dynamicTyping"
    );
  });

  it("checks the options", () => {
    expect(() => Csv.applyDynamicTypingToRow(row(), true, { castDate: "x" as never })).toThrow(
      expect.objectContaining({ field: "options.castDate" })
    );
    expect(() =>
      Csv.applyDynamicTypingToRow(row(), true, { decimalSeparator: ";" as never })
    ).toThrow(expect.objectContaining({ field: "options.decimalSeparator" }));
  });
});
