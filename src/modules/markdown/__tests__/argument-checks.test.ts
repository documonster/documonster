/**
 * `Markdown.parse`, `Markdown.parseAll` and `Markdown.format` check their arguments: a value
 * with no reading is a `MarkdownError` naming the call and the field, not a `TypeError` from
 * inside the formatter or an alignment that silently became `"none"`.
 */
import { Markdown, MarkdownError, MarkdownParseError } from "@markdown/index";
import { describe, expect, it } from "vitest";

type FormatArgs = Parameters<typeof Markdown.format>;

function expectMarkdownError(call: () => unknown, field: string): void {
  let caught: unknown;
  try {
    call();
  } catch (err) {
    caught = err;
  }
  expect(caught).toBeInstanceOf(MarkdownError);
  expect(caught).not.toBeInstanceOf(MarkdownParseError);
  expect((caught as Error).message).toContain(`${field} must`);
}

describe("Markdown.format argument checks", () => {
  const headers = ["a", "b"];
  const rows = [[1, 2]];

  it.each<[string, unknown, unknown, unknown]>([
    // A parse result passed as the headers: this threw `TypeError` from deep inside.
    ["headers", { headers, rows, alignments: ["sideways"] }, undefined, undefined],
    ["headers[1]", ["a", 2], rows, {}],
    ["rows", headers, "1,2", {}],
    ["rows[0]", headers, [{ a: 1 }], {}],
    ["options", headers, rows, null],
    ["alignment", headers, rows, { alignment: "sideways" }],
    ["columns", headers, rows, { columns: "a" }],
    ["columns[0]", headers, rows, { columns: [5] }],
    ["columns[0].header", headers, rows, { columns: [{ alignment: "left" }] }],
    [
      "columns[1].alignment",
      headers,
      rows,
      { columns: ["A", { header: "B", alignment: "middle" }] }
    ],
    ["columns[0].minWidth", headers, rows, { columns: [{ header: "A", minWidth: -1 }] }],
    ["columns[0].minWidth", headers, rows, { columns: [{ header: "A", minWidth: NaN }] }],
    ["stringify", headers, rows, { stringify: "String" }],
    ["rows[0][1]", headers, rows, { stringify: (v: unknown) => (v === 2 ? 2 : String(v)) }]
  ])("refuses %s", (field, h, r, o) => {
    expectMarkdownError(
      () => Markdown.format(h as FormatArgs[0], r as FormatArgs[1], o as FormatArgs[2]),
      field
    );
  });

  it("formats valid boundary values as before", () => {
    expect(Markdown.format([], [])).toBe("");
    expect(
      Markdown.format(["a"], [[]], { alignment: "none", columns: [{ header: "A", minWidth: 0 }] })
    ).toBe("| A |\n|---|\n|   |\n");
    expect(Markdown.format(["a", "b"], [["x"]], { alignment: "right" })).toBe(
      "|   a |   b |\n|----:|----:|\n|   x |     |\n"
    );
  });
});

describe("Markdown.parse argument checks", () => {
  const table = "| a |\n| --- |\n| 1 |\n";

  it.each<[string, unknown, unknown]>([
    ["input", 42, {}],
    ["options", table, null],
    ["maxRows", table, { maxRows: -1 }],
    ["maxRows", table, { maxRows: 1.5 }],
    ["maxRows", table, { maxRows: NaN }]
  ])("refuses %s", (field, input, options) => {
    for (const call of [Markdown.parse, Markdown.parseAll]) {
      expectMarkdownError(
        () => call(input as string, options as Parameters<typeof Markdown.parse>[1]),
        field
      );
    }
  });

  it("still parses with maxRows: 0", () => {
    expect(Markdown.parse(table, { maxRows: 0 }).rows).toEqual([]);
  });
});

describe("Markdown.format column checks declared as a schema", () => {
  it("refuses an array as a column config at the column, not at its header", () => {
    let caught: { field?: string; message?: string } = {};
    try {
      Markdown.format(["a"], [["1"]], { columns: [[] as never] });
    } catch (err) {
      caught = err as typeof caught;
    }
    expect(caught.field).toBe("options.columns[0]");
    expect(caught.message).toBe(
      "Markdown.format: options.columns[0] must be a header string or { header, alignment?, minWidth? }, got an array"
    );
  });
});
