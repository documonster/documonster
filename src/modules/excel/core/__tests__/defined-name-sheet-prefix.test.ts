/**
 * A defined name's text is a region only when what comes before its `!` is a sheet name as a formula writes one —
 * quoted (`'My Sheet'!`), or of the characters an unquoted name may hold. `@S!$A$1:$A$3` and `-S!$A$1` are formulas
 * whose operand is on `S`; they used to be read as regions on sheets named `@S` and `-S`, and written back quoted.
 */
import { DefinedNames, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

async function readBack(text: string) {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "S");
  Workbook.addWorksheet(wb, "My Sheet");
  DefinedNames.addFormula(Workbook.getDefinedNames(wb), "N", text);
  const read = await Workbook.read(await Workbook.toBuffer(wb));
  return DefinedNames.get(Workbook.getDefinedNames(read), "N");
}

describe("the sheet prefix of a defined name's region", () => {
  it.each(["@S!$A$1:$A$3", "-S!$A$1", "+S!$A$1:$B$2"])("%s is a formula", async text => {
    expect(await readBack(text)).toMatchObject({ formulaExpression: text, ranges: [text] });
  });

  it.each(["S!$A$1:$A$3", "'My Sheet'!$A$1", "S!$A$1,'My Sheet'!$B$2", "S!$A:$A"])(
    "%s is a region",
    async text => {
      const name = await readBack(text);
      expect(name?.ranges).toEqual(text.split(","));
      expect(name?.formulaExpression).toBeUndefined();
    }
  );
});
