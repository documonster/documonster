import { describe, expect, it } from "vitest";

import { Build, Document, DocxArgumentError, Io, Query } from "../index";

describe("Build.table takes its column widths in the options", () => {
  const rows = [Build.row([Build.cell("a"), Build.cell("b")])];

  it("as the deprecated third argument did", () => {
    const props = { layout: "fixed" as const };
    expect(Build.table(rows, { ...props, columnWidths: [1000, 2000] })).toEqual(
      Build.table(rows, props, [1000, 2000])
    );
    expect(Build.table(rows, { ...props, columnWidths: [1000, 2000] }).properties).toEqual(props);
  });

  it("and without them is what it was", () => {
    expect(Build.table(rows, { layout: "fixed" })).toEqual({
      type: "table",
      properties: { layout: "fixed" },
      columnWidths: undefined,
      rows
    });
    expect(() => Build.table(rows, { columnWidths: [-1] })).toThrow(/columnWidths\[0\]/);
  });
});

describe("Build.math* take their limits and position in options", () => {
  const x = [Build.mathRun("x")];
  const lo = [Build.mathRun("0")];
  const hi = [Build.mathRun("n")];

  it("as the deprecated positional arguments did", () => {
    expect(Build.mathSum(x, { sub: lo, sup: hi })).toEqual(Build.mathSum(x, lo, hi));
    expect(Build.mathIntegral(x, { sup: hi })).toEqual(Build.mathIntegral(x, undefined, hi));
    expect(Build.mathProduct(x, { sub: lo })).toEqual(Build.mathProduct(x, lo));
    expect(Build.mathNary("\u2210", x, { sub: lo, sup: hi })).toEqual(
      Build.mathNary("\u2210", x, lo, hi)
    );
    expect(Build.mathBar(x, { position: "bottom" })).toEqual(Build.mathBar(x, "bottom"));
    expect(Build.mathSum(x)).toEqual(Build.mathSum(x, {}));
    expect(Build.mathBar(x)).toEqual(Build.mathBar(x, {}));
  });

  it("checks the options with fields from the parameter", () => {
    const cases: [string, () => unknown][] = [
      ["options.sub", () => Build.mathSum(x, { sub: "x" as never })],
      ["options.sup", () => Build.mathNary("\u2211", x, { sup: 42 as never })],
      ["options.position", () => Build.mathBar(x, { position: "left" as never })],
      ["options", () => Build.mathSum(x, 42 as never)]
    ];
    for (const [field, call] of cases) {
      let error: unknown;
      try {
        call();
      } catch (e) {
        error = e;
      }
      expect(error, field).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, field).toBe(field);
    }
  });
});

describe("Build.underline takes its style in the properties", () => {
  it("as the deprecated style argument did", () => {
    expect(Build.underline("x", { underline: "double", bold: true })).toEqual(
      Build.underline("x", "double", { bold: true })
    );
    expect(Build.underline("x")).toEqual(Build.underline("x", undefined));
    expect(Build.underline("x", { bold: true }).properties?.underline).toBe("single");
  });
});

describe("Build.underline checks what it is given", () => {
  it.each([42, null, [], true])("refuses %j as the style", bad => {
    expect(() => Build.underline("x", bad as never)).toThrow(DocxArgumentError);
  });
});

describe("Io.read takes bytes", () => {
  it.each([42, null, {}, "PK"])("refuses %j by name", async bad => {
    await expect(Io.read(bad as never)).rejects.toMatchObject({
      name: "DocxArgumentError",
      field: "input"
    });
  });

  it("and an ArrayBuffer, which it always read", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "hi");
    const bytes = await Io.toBuffer(doc);
    const copy = bytes.slice().buffer;
    expect(Query.getText(await Io.read(copy))).toContain("hi");
  });
});

describe("the builders take readonly arrays", () => {
  it("as they only read them", () => {
    const runs = [Build.text("a"), Build.bold("b")] as const;
    const cells = [Build.cell("x")] as const;
    const rows = [Build.row(cells)] as const;
    const math = [Build.mathRun("x")] as const;
    const doc = Document.create();
    Document.addContent(doc, Build.paragraph(runs));
    Document.addContent(doc, Build.table(rows));
    Document.addContent(doc, Build.simpleTable([["a", "b"]] as const));
    Document.addContent(doc, Build.mathBlock([Build.mathFraction(math, math)]));
    Document.addBulletList(doc, ["one", [Build.text("two")]] as const);
    Document.addFootnote(doc, [Build.textParagraph("note")] as const);
    expect(Query.getText(doc)).toContain("ab");
    expect(runs).toHaveLength(2);
  });
});
