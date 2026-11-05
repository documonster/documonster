/**
 * Contracts of `Csv.format` and the formatter stream that the API review found broken:
 * object rows written without a header row and in each row's own key order, `forceQuote` /
 * `preventQuote` serialised as JSON when written as cell values, and options or rows that
 * produced unreadable output or a raw `TypeError` instead of a `CsvError`.
 */
import { Csv, CsvError } from "@csv/index";
import { describe, expect, it } from "vitest";

function formatStream(
  options: ConstructorParameters<typeof Csv.FormatterStream>[0],
  rows: unknown[]
) {
  return new Promise<string>((resolve, reject) => {
    const stream = new Csv.FormatterStream(options);
    let out = "";
    stream.on("data", (chunk: string) => (out += chunk));
    stream.on("end", () => resolve(out));
    stream.on("error", reject);
    for (const row of rows) {
      stream.write(row as never);
    }
    stream.end();
  });
}

describe("Csv.format — keyed rows", () => {
  it("writes the first row's keys as the header row by default (README / TSDoc example)", () => {
    expect(
      Csv.format([
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 }
      ])
    ).toBe("name,age\nAlice,30\nBob,25");
  });

  it("reads every row by the first row's keys, whatever order it lists them in", () => {
    const rows = [
      { a: 1, b: 2 },
      { b: 3, a: 4 }
    ];
    expect(Csv.format(rows)).toBe("a,b\n1,2\n4,3");
    expect(Csv.format(rows, { headers: false })).toBe("1,2\n4,3");
    expect(Csv.format(rows, { headers: ["b", "a"] })).toBe("b,a\n2,1\n3,4");
  });

  it("does the same in the formatter stream", async () => {
    const rows = [
      { a: 1, b: 2 },
      { b: 3, a: 4 }
    ];
    expect(await formatStream({}, rows)).toBe("a,b\n1,2\n4,3");
    expect(await formatStream({ headers: false }, rows)).toBe("1,2\n4,3");
  });

  it("round-trips objects through Csv.parse", () => {
    const rows = [
      { x: "1", y: "a,b" },
      { y: "c", x: "2" }
    ];
    expect(Csv.parse(Csv.format(rows), { headers: true }).rows).toEqual([
      { x: "1", y: "a,b" },
      { x: "2", y: "c" }
    ]);
  });

  it("keeps an object row that follows array rows", () => {
    expect(Csv.format([["x", "y"], { a: 1, b: 2 }])).toBe("x,y\n1,2");
  });
});

describe("Csv.forceQuote / preventQuote as cell values", () => {
  it("quotes or leaves unquoted a value written straight into a row", () => {
    expect(Csv.format([[Csv.forceQuote("x"), "y"]])).toBe('"x",y');
    expect(Csv.format([[Csv.preventQuote("=1"), "y"]])).toBe("=1,y");
  });

  it("works in the formatter stream too", async () => {
    expect(await formatStream({}, [[Csv.forceQuote("x")]])).toBe('"x"');
  });

  it("refuses a value that is not a string", () => {
    expect(() => Csv.forceQuote(42 as never)).toThrow(CsvError);
    expect(() => Csv.preventQuote(undefined as never)).toThrow(CsvError);
  });

  it("honours the typeTransform form the TSDoc example shows", () => {
    expect(
      Csv.format([{ code: "007", note: "" }], {
        typeTransform: {
          string: (v, ctx) => (ctx.column === "code" || v === "" ? Csv.forceQuote(v) : v)
        }
      })
    ).toBe('code,note\n"007",""');
  });
});

describe("Csv.format — options that cannot be read back are refused", () => {
  it.each([
    ["a delimiter holding a line break", { delimiter: "\n" }],
    ["a delimiter holding CR", { delimiter: "\r" }],
    ["a lineEnding that overlaps the delimiter", { delimiter: ";", lineEnding: ";\n" }],
    ["a lineEnding holding the quote", { lineEnding: '"' }],
    ["headers that are not an array or a boolean", { headers: 42 }],
    ["a header that is not a string", { headers: ["a", 1] }],
    ["columns that are not an array", { columns: "abc" }],
    ["a column that is null", { columns: [null] }],
    ["a column that is a number", { columns: [42] }],
    ["quoteColumns that is a string", { quoteColumns: "x" }],
    ["typeTransform that is not an object", { typeTransform: 42 }],
    ["an unknown typeTransform key", { typeTransform: { strnig: () => "" } }],
    ["a typeTransform member that is not a function", { typeTransform: { string: "x" } }],
    ["a bom that is not a boolean", { bom: "yes" }],
    ["a decimalSeparator inside the delimiter", { delimiter: ",;", decimalSeparator: "," }]
  ])("%s", (_label, options) => {
    expect(() => Csv.format([["a"]], options as never)).toThrow(CsvError);
    expect(() => new Csv.FormatterStream(options as never)).toThrow(CsvError);
  });

  it("quotes a field holding a custom lineEnding", () => {
    const csv = Csv.format([["a|b", "c"]], { lineEnding: "|" });
    expect(csv).toBe('"a|b",c');
    expect(Csv.parse(csv, { lineEnding: "|" })).toEqual([["a|b", "c"]]);
  });
});

describe("Csv.format — rows and values are checked", () => {
  it.each([
    ["null data", null],
    ["a number", 42],
    ["a string", "abc"]
  ])("refuses %s as data", (_label, data) => {
    expect(() => Csv.format(data as never)).toThrow(CsvError);
  });

  it.each([
    ["a string row (was split into characters)", ["abc"]],
    ["a null row", [null, ["a"]]],
    ["a number row", [42]]
  ])("refuses %s", (_label, data) => {
    expect(() => Csv.format(data as never)).toThrow(CsvError);
  });

  it("refuses a string row in the stream as well", async () => {
    await expect(formatStream({}, ["abc"])).rejects.toThrow(CsvError);
  });

  it("refuses an Invalid Date", () => {
    expect(() => Csv.format([[new Date("x")]])).toThrow(CsvError);
  });

  it("refuses a typeTransform that returns something other than text", () => {
    expect(() => Csv.format([[1.5]], { typeTransform: { number: v => (v * 2) as never } })).toThrow(
      CsvError
    );
  });

  it("refuses a typeTransform.row result that is not a row", () => {
    expect(() => Csv.format([["a"]], { typeTransform: { row: () => "x" as never } })).toThrow(
      CsvError
    );
  });
});

describe("r3: a missing field is empty, not an Object.prototype member (TX-B2)", () => {
  it("in Csv.format", () => {
    expect(Csv.format([{ a: 1 }], { headers: ["__proto__", "constructor", "a"] })).toBe(
      "__proto__,constructor,a\n,,1"
    );
    expect(Csv.format([{}], { columns: ["constructor", "toString"] })).toBe(
      "constructor,toString\n,"
    );
  });

  it("still reads a getter of the row's class", () => {
    class Point {
      constructor(readonly x: number) {}
      get double(): number {
        return this.x * 2;
      }
    }
    expect(Csv.format([new Point(2) as never], { columns: ["x", "double"] })).toBe("x,double\n2,4");
  });

  it("does not quote a column because of Object.prototype", () => {
    expect(Csv.format([{ constructor: "a" }], { quoteColumns: { x: true } })).toBe(
      "constructor\na"
    );
  });
});

describe("r3: cell values with no CSV text are refused", () => {
  it.each([
    ["a function", () => 1],
    ["a symbol", Symbol("x")]
  ])("%s", (_label, value) => {
    expect(() => Csv.format([[value]] as never)).toThrow(CsvError);
  });

  it("a circular object", () => {
    const o: Record<string, unknown> = {};
    o.self = o;
    expect(() => Csv.format([[o]] as never)).toThrow(CsvError);
  });

  it("keeps writing plain objects and arrays as JSON, and bigint as digits", () => {
    expect(Csv.format([[{ a: 1 }, [1, 2], 2n]] as never)).toBe('"{""a"":1}","[1,2]",2');
  });
});
