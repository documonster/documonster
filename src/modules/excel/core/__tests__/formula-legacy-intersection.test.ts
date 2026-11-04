/**
 * What Excel computes for formulas whose ranges it may intersect with the formula's row, in a row inside the ranges
 * (3) and one outside them (8), both as this library writes a formula set through the API and as another producer
 * stores it — a plain `<f>` with no dynamic-array metadata.
 *
 * Every expected value is Excel's (16.x, macOS): the workbooks were opened, recalculated and saved by Excel and the
 * cells read back. A formula written as an ordinary one is computed by Excel with implicit intersection, so this
 * library computes it the same way; one it writes as a dynamic array is computed over arrays.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Cell, DefinedNames, Table, Workbook } from "@excel";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { describe, expect, it } from "vitest";

const V = "#VALUE!";
const D0 = "#DIV/0!";

/** Formula, then Excel's value in row 3 and row 8: set through the API, then stored as a plain `<f>`. */
const CASES: readonly (readonly [string, unknown, unknown, unknown, unknown])[] = [
  ['SUM(IF(B1:B5="y",A1:A5))', 100, 100, 150, V],
  ['MEDIAN(IF(B1:B5="y",A1:A5))', 30, 30, 30, V],
  ["SUM(A1:A5*2)", 300, 300, 60, V],
  ["AVERAGE(A1:A5^2)", 1100, 1100, 900, V],
  ["SUM(LEN(B1:B5))", 5, 5, 1, V],
  ['SUMPRODUCT(A1:A5*(B1:B5="y"))', 100, 100, 100, 100],
  ['SUMPRODUCT(IF(B1:B5="y",A1:A5,0))', 100, 100, 150, V],
  ["COUNT(IF(A1:A5>20,A1:A5))", 3, 3, 5, 0],
  ["AND(A1:A5>5)", true, true, true, V],
  ["INDEX(A1:A5*3,4)", 120, 120, 120, 120],
  ["ROWS(A1:A5+0)", 5, 5, 5, 5],
  ["SUM(A1:A5)", 150, 150, 150, 150],
  ["MAX(A1:A5-A1)", 40, 40, 20, V],
  ["ABS(A1:A5)", 30, V, 30, V],
  ["A1:A5*1", 30, V, 30, V],
  ["SUM(--(A1:A5>15))", 4, 4, 1, V],
  ['TEXTJOIN(",",TRUE,IF(A1:A5>20,A1:A5,""))', "30,40,50", "30,40,50", "10,20,30,40,50", V],
  ['MIN(IF(B1:B5="x",A1:A5))', 10, 10, 0, V],
  ["SUM(Rng*2)", 60, V, 60, V],
  ["STDEV(A1:A5*1)", 15.811388300841896, 15.811388300841896, D0, V],
  ["SUM(Samples[Response]*1)", 22.5, 22.5, 2, V],
  ['MEDIAN(IF(Samples[Group]="A",Samples[Response]))', 3, 3, 0, V],
  ["SUM(T1!A1:A5*2)", 300, 300, 60, V],
  ["SUM(A1:A5/(A1:A5>0))", 150, 150, 30, V]
];

function column(index: number): string {
  return index < 22 ? String.fromCharCode(68 + index) : `A${String.fromCharCode(65 + index - 22)}`;
}

function build() {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  [10, 20, 30, 40, 50].forEach((value, index) => {
    Cell.setValue(sheet, `A${index + 1}`, value);
    Cell.setValue(sheet, `B${index + 1}`, ["x", "y", "y", "x", "y"][index]);
  });
  const t1 = Workbook.addWorksheet(workbook, "T1");
  [10, 20, 30, 40, 50].forEach((value, index) => Cell.setValue(t1, `A${index + 1}`, value));
  Table.add(t1, {
    name: "Samples",
    ref: "C1",
    columns: [{ name: "Group" }, { name: "Response" }],
    rows: [
      ["A", 1.5],
      ["B", 2],
      ["A", 7],
      ["A", 3],
      ["B", 9]
    ]
  });
  DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Rng", "S!$A$1:$A$5");
  CASES.forEach(([formula], index) => {
    for (const row of [3, 8]) {
      Cell.setValue(sheet, `${column(index)}${row}`, { formula });
    }
  });
  return workbook;
}

function results(workbook: Workbook.Handle): unknown[][] {
  const sheet = Workbook.getWorksheet(workbook, "S")!;
  return CASES.map(([formula], index) => [
    formula,
    ...[3, 8].map(row => {
      const value = Cell.getResult(sheet, `${column(index)}${row}`);
      return value !== null && typeof value === "object" && "error" in value ? value.error : value;
    })
  ]);
}

describe("implicit intersection, against Excel", () => {
  it("a formula set through the API computes what Excel computes for the file it is written to", () => {
    const workbook = build();
    calculateFormulas(workbook);
    expect(results(workbook)).toEqual(CASES.map(([formula, r3, r8]) => [formula, r3, r8]));
  });

  it("a formula a file stores as an ordinary one computes what Excel computes for it", async () => {
    const entries = await extractAll(await Workbook.toBuffer(build()));
    const legacy = await createZip(
      [...entries.entries()].map(([name, file]) => ({
        name,
        data:
          name === "xl/worksheets/sheet1.xml"
            ? new TextEncoder().encode(
                new TextDecoder()
                  .decode(file.data)
                  .replace(
                    /<c r="([A-Z]+\d+)"([^>]*?) cm="1"([^>]*)><f t="array" ref="\1">/g,
                    '<c r="$1"$2$3><f>'
                  )
                  .replace(/(<c [^>]*>)(<f>[^<]*<\/f>)<v>[^<]*<\/v>/g, "$1$2")
                  .replace(/(<c [^>]*?) t="(?:str|b|e)"([^>]*><f>)/g, "$1$2")
              )
            : file.data
      }))
    );
    const workbook = Workbook.create();
    await Workbook.read(workbook, legacy);
    calculateFormulas(workbook);
    expect(results(workbook)).toEqual(CASES.map(([formula, , , r3, r8]) => [formula, r3, r8]));
  });
});
