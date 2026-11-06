/**
 * Contracts of `Markdown.format` / `Markdown.parse` that the API review found broken:
 * `columns` cut down to `headers.length`, a `format` → `parse` round trip that could not be
 * made lossless (a literal `<br>` came back as a newline), tables read out of code blocks,
 * and boolean options that accepted any value.
 */
import {
  Markdown,
  MarkdownArgumentError,
  MarkdownError,
  MarkdownParseError
} from "@markdown/index";
import { describe, expect, it } from "vitest";

describe("Markdown.format — columns", () => {
  it("has as many columns as the longer of headers and columns", () => {
    expect(
      Markdown.format(["a"], [["1", "2"]], {
        columns: ["x", { header: "y", alignment: "right" }]
      })
    ).toBe("| x   |   y |\n|:----|----:|\n| 1   |   2 |\n");
  });

  it("works with columns alone", () => {
    expect(Markdown.format([], [["1"]], { columns: ["x"] })).toBe("| x   |\n|:----|\n| 1   |\n");
  });

  it("matches the module example's output", () => {
    expect(
      Markdown.format(
        ["Name", "Age"],
        [
          ["Alice", "30"],
          ["Bob", "25"]
        ]
      )
    ).toBe("| Name  | Age |\n|:------|:----|\n| Alice | 30  |\n| Bob   | 25  |\n");
  });
});

describe("format → parse round trip", () => {
  const headers = ["a", "b|c", "d\\e"];
  const rows = [
    ["x|y", "back\\slash", "<br>"],
    ["multi\nline", "<BR/> literal", "x<br />y"]
  ];

  it("is lossless with convertBr", () => {
    const md = Markdown.format(headers, rows);
    const parsed = Markdown.parse(md, { convertBr: true });
    expect(parsed.headers).toEqual(headers);
    expect(parsed.rows).toEqual(rows);
  });

  it("writes a literal <br> escaped, and parse reads it back as text", () => {
    expect(Markdown.format(["a"], [["<br>"]])).toBe("| a     |\n|:------|\n| \\<br> |\n");
    expect(Markdown.parse("| a |\n|-|\n| \\<br> |", { convertBr: true }).rows).toEqual([["<br>"]]);
  });

  it("leaves escapes as written with unescape: false", () => {
    expect(
      Markdown.parse("| a |\n|-|\n| \\<br> x<br>y |", { unescape: false, convertBr: true }).rows
    ).toEqual([["\\<br> x\ny"]]);
  });
});

describe("tables inside code blocks are code", () => {
  it.each([
    ["a backtick fence", "```\n| a |\n|-|\n| 1 |\n```\n\n| b |\n|-|\n| 2 |"],
    ["a tilde fence", "~~~md\n| a |\n|-|\n| 1 |\n~~~\n\n| b |\n|-|\n| 2 |"],
    ["an indented block", "    | a |\n    |-|\n    | 1 |\n\n| b |\n|-|\n| 2 |"]
  ])("in %s", (_label, md) => {
    expect(Markdown.parseAll(md).map(t => t.headers)).toEqual([["b"]]);
    expect(Markdown.parse(md).headers).toEqual(["b"]);
  });

  it("runs an unclosed fence to the end of the document", () => {
    expect(Markdown.parseAll("```\n| a |\n|-|\n| 1 |")).toEqual([]);
  });

  it("still reads a table indented less than four spaces", () => {
    expect(Markdown.parse("   | a |\n   |-|\n   | 1 |").rows).toEqual([["1"]]);
  });
});

describe("boolean options are checked", () => {
  it.each(["padding", "trailingNewline", "escapeContent"])("Markdown.format %s", name => {
    expect(() => Markdown.format(["a"], [["1"]], { [name]: "no" } as never)).toThrow(MarkdownError);
  });

  it.each(["trim", "unescape", "skipEmptyRows", "convertBr"])("Markdown.parse %s", name => {
    expect(() => Markdown.parse("| a |\n|-|", { [name]: 1 } as never)).toThrow(MarkdownError);
  });
});

describe("r3: Markdown.format keeps every cell (TX-B5)", () => {
  it("widens the table to the longest row, with empty header cells", () => {
    const md = Markdown.format(["a"], [["1", "2", "3"]]);
    expect(md).toBe("| a   |     |     |\n|:----|:----|:----|\n| 1   | 2   | 3   |\n");
    expect(Markdown.parse(md).rows).toEqual([["1", "2", "3"]]);
  });

  it("writes a table that has rows but no headers", () => {
    expect(Markdown.parse(Markdown.format([], [["1", "2"]])).rows).toEqual([["1", "2"]]);
  });
});

describe("r3: cell values with no text are refused", () => {
  it.each([
    ["a function", () => 1],
    ["a symbol", Symbol("x")]
  ])("%s", (_label, value) => {
    expect(() => Markdown.format(["a"], [[value]])).toThrow(MarkdownError);
  });

  it("a circular object", () => {
    const o: Record<string, unknown> = {};
    o.self = o;
    expect(() => Markdown.format(["a"], [[o]])).toThrow(MarkdownError);
  });

  it("keeps writing plain objects as JSON and bigint as digits", () => {
    expect(Markdown.format(["a", "b"], [[{ x: 1 }, 2n]])).toContain('| {"x":1} | 2   |');
  });
});

describe("X12: argument errors are MarkdownArgumentError with a field", () => {
  it.each<[string, () => unknown, string]>([
    ["parse input", () => Markdown.parse(42 as never), "input"],
    ["parse option", () => Markdown.parse("| a |\n|-|", { maxRows: -1 }), "options.maxRows"],
    ["format headers", () => Markdown.format(42 as never, []), "headers"],
    ["format row", () => Markdown.format(["a"], [1 as never]), "rows[0]"],
    [
      "format option",
      () => Markdown.format(["a"], [], { alignment: "x" as never }),
      "options.alignment"
    ],
    [
      "column",
      () => Markdown.format(["a"], [], { columns: [{ header: 1 as never }] }),
      "options.columns[0].header"
    ],
    [
      "stringify result",
      () => Markdown.format(["a"], [["x"]], { stringify: () => 1 as never }),
      "rows[0][0]"
    ],
    ["cell", () => Markdown.format(["a"], [[Symbol("x")]]), "rows"]
  ])("%s", (_label, fn, field) => {
    let error: unknown;
    try {
      fn();
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(MarkdownArgumentError);
    expect(error).toBeInstanceOf(MarkdownError);
    expect((error as MarkdownArgumentError).field).toBe(field);
    expect((error as MarkdownArgumentError).toJSON()).toMatchObject({ field });
  });

  it("is not used when the input has no table", () => {
    expect(() => Markdown.parse("hello")).toThrow(MarkdownParseError);
    expect(() => Markdown.parse("hello")).not.toThrow(MarkdownArgumentError);
  });
});
