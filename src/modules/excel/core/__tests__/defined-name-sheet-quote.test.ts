import { DefinedNames, Row, Workbook, Worksheet } from "@excel/index";
/**
 * A sheet whose name holds an apostrophe, `it's`, is written `'it''s'!` in a reference. `colCache.decodeEx` returned
 * the quoted text as the sheet's name — `it''s`, the doubling left in — so `DefinedNames.add("'it''s'!$B$1")` covered
 * cells of a sheet that does not exist and listed `'it''''s'!$B$1`, and reading a file doubled the quotes again.
 */
import { colCache } from "@excel/utils/col-cache";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "it's");
  const names = Workbook.getDefinedNames(wb);
  DefinedNames.add(names, "'it''s'!$B$1", "M");
  DefinedNames.addFormula(names, "N", "'it''s'!$A$1");
  DefinedNames.addFormula(names, "F", "SUM('it''s'!$A$1:$A$3)");
  return { wb, ws };
}

function listed(wb: Workbook.Handle) {
  return Object.fromEntries(
    DefinedNames.list(Workbook.getDefinedNames(wb)).map(entry => [
      entry.name,
      entry.formulaExpression ?? entry.ranges
    ])
  );
}

describe("a defined name on a sheet named it's", () => {
  it("decodeEx gives the sheet's name, not its quoted spelling", () => {
    expect(colCache.decodeEx("'it''s'!$B$1")).toMatchObject({ sheetName: "it's", row: 1, col: 2 });
    expect(colCache.decodeEx("'it''s'!A1:B2")).toMatchObject({
      sheetName: "it's",
      top: 1,
      right: 2
    });
    expect(colCache.decodeEx("'My Sheet'!A1")).toMatchObject({ sheetName: "My Sheet" });
  });

  it("is listed as added, and names the sheet's cells", () => {
    const { wb } = book();
    expect(listed(wb)).toEqual({
      M: ["'it''s'!$B$1"],
      N: "'it''s'!$A$1",
      F: "SUM('it''s'!$A$1:$A$3)"
    });
    expect(DefinedNames.getNames(Workbook.getDefinedNames(wb), "'it''s'!B1")).toEqual(["M"]);
  });

  it.each(["xlsx", "xlsb"] as const)("reads back from %s as written", async format => {
    const { wb } = book();
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    // A formula that is a lone reference reads back as a reference — the same name either way.
    expect(listed(read)).toEqual({
      M: ["'it''s'!$B$1"],
      N: ["'it''s'!$A$1"],
      F: "SUM('it''s'!$A$1:$A$3)"
    });
  });

  it("moves with a row inserted above", () => {
    const { wb, ws } = book();
    Row.insert(ws, 1, []);
    expect(listed(wb)).toEqual({
      M: ["'it''s'!$B$2"],
      N: "'it''s'!$A$2",
      F: "SUM('it''s'!$A$2:$A$4)"
    });
  });

  it("follows the sheet when it is renamed, to and from a name with an apostrophe", () => {
    const { wb, ws } = book();
    Worksheet.setName(ws, "o'k 'x' y");
    expect(listed(wb)).toEqual({
      M: ["'o''k ''x'' y'!$B$1"],
      N: "'o''k ''x'' y'!$A$1",
      F: "SUM('o''k ''x'' y'!$A$1:$A$3)"
    });
    Worksheet.setName(ws, "Plain");
    expect(listed(wb)).toEqual({ M: ["Plain!$B$1"], N: "Plain!$A$1", F: "SUM(Plain!$A$1:$A$3)" });
  });
});
