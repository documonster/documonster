import { DefinedNames, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

/**
 * Defined names whose text is not a fixed set of cells survive a write in both containers.
 *
 * Both cases come from `many-defined-names.xlsx`, an Excel-saved workbook Excel refused after a round trip:
 * `_PAG2` is `!$A$73:$E$140` — a reference to "the sheet the name is used on" — and was written as `$A$73:$E$140`;
 * the workbook-scoped `_xlnm.Print_Area` and `_xlnm.Print_Titles`, `#REF!` since their sheet was deleted, belong to no
 * sheet and were dropped.
 */
const NAMES = [
  { name: "_PAG2", rawText: "!$A$73:$E$140" },
  { name: "_xlnm.Print_Area", rawText: "#REF!" },
  { name: "_xlnm.Print_Titles", rawText: "#REF!" }
];

function definitions(wb: Workbook.Handle): Map<string, string | undefined> {
  return new Map(
    DefinedNames.getModel(Workbook.getDefinedNames(wb))
      .filter(entry => entry.localSheetId === undefined)
      .map(entry => [
        entry.name,
        entry.rawText ?? entry.formulaExpression ?? entry.ranges.join(",")
      ])
  );
}

describe("defined names kept as text", () => {
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`${format}: keeps a sheetless reference and workbook-scoped print names`, async () => {
      const wb = Workbook.create();
      Workbook.addWorksheet(wb, "S");
      DefinedNames.setModel(
        Workbook.getDefinedNames(wb),
        NAMES.map(entry => ({ ...entry, ranges: [] }))
      );
      const again = await Workbook.read(
        Workbook.create(),
        await Workbook.toBuffer(wb, { format, unsupported: "ignore" })
      );
      const read = definitions(again);
      for (const { name, rawText } of NAMES) {
        expect(read.get(name), name).toBe(rawText);
      }
    });
  }
});
