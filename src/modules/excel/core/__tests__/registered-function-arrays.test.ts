/**
 * A registered function says whether it returns an array (`RegisterFunctionOptions.returnsArray`), and that
 * declaration is what decides whether a formula calling it is a dynamic array — over a built-in of the same name. The
 * invariant holds for it as for built-ins: the engine spills a formula exactly when the writers store it as a dynamic
 * array, and a formula that is not one shows the top-left value of an array its function returns, as Excel shows an
 * ordinary formula's.
 */
import { Archive } from "@archive";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, ExcelArgumentError, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const NUMBER = 1;
const ARRAY = 5;

/** `{1;2;3}`, as a registered function returns it. */
function column(): unknown {
  return {
    kind: ARRAY,
    rows: [1, 2, 3].map(value => [{ kind: NUMBER, value }]),
    height: 3,
    width: 1
  };
}

function book(name: string, formula: string, returnsArray?: boolean) {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Workbook.registerFunction(
    wb,
    name,
    column as never,
    returnsArray === undefined ? {} : { returnsArray }
  );
  Cell.setValue(ws, "B1", { formula });
  return { wb, ws };
}

const values = (ws: Parameters<typeof Cell.getResult>[0]) =>
  [1, 2, 3].map(row => {
    const value = Cell.getValue(ws, `B${row}`);
    return typeof value === "object" && value !== null && "formula" in value ? value.result : value;
  });

async function storedAsDynamicArray(wb: ReturnType<typeof Workbook.create>): Promise<boolean> {
  const files = await Archive.unpackZip(await Workbook.toBuffer(wb));
  const xml = new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
  return /<c r="B1"[^>]* cm="1"[^>]*><f t="array"/.test(xml);
}

const CASES: [string, string, boolean | undefined, (number | null)[], boolean][] = [
  // name, formula, returnsArray, values B1:B3, stored as a dynamic array
  ["CUSTOM", "CUSTOM()", true, [1, 2, 3], true],
  ["CUSTOM", "CUSTOM()*2", true, [2, 4, 6], true],
  ["CUSTOM", "CUSTOM()", undefined, [1, null, null], false],
  ["CUSTOM", "CUSTOM()", false, [1, null, null], false],
  ["CUSTOM", "SUM(CUSTOM())", true, [6, null, null], false],
  // A registered function's own declaration wins over the built-in it shadows.
  ["SEQUENCE", "SEQUENCE(3)", undefined, [1, null, null], false],
  ["SEQUENCE", "SEQUENCE(3)", true, [1, 2, 3], true],
  // And a declaration makes a function the built-ins call a value function an array one.
  ["TEXTJOIN", "TEXTJOIN()", true, [1, 2, 3], true]
];

describe("returnsArray decides, and the engine and the writers agree", () => {
  it.each(CASES)(
    "%s registered, %s with returnsArray %s",
    async (name, formula, returnsArray, expected, dynamic) => {
      const { wb, ws } = book(name, formula, returnsArray);
      calculateFormulas(wb);
      expect(values(ws)).toEqual(expected);
      expect(await storedAsDynamicArray(wb)).toBe(dynamic);
    }
  );

  it.each(["xlsx", "xlsb"] as const)(
    "and after a %s read-back, with the function registered again",
    async format => {
      for (const [name, formula, returnsArray, expected, dynamic] of CASES) {
        const { wb } = book(name, formula, returnsArray);
        calculateFormulas(wb);
        const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
        Workbook.registerFunction(
          read,
          name,
          column as never,
          returnsArray === undefined ? {} : { returnsArray }
        );
        const ws = Workbook.getWorksheet(read, "S")!;
        expect([
          formula,
          (Cell.getValue(ws, "B1") as { isDynamicArray?: boolean }).isDynamicArray === true
        ]).toEqual([formula, dynamic]);
        calculateFormulas(read);
        expect([formula, values(ws)]).toEqual([formula, expected]);
      }
    }
  );

  it("unregistering returns the decision to the built-in", async () => {
    const { wb, ws } = book("SEQUENCE", "SEQUENCE(3)");
    Workbook.unregisterFunction(wb, "SEQUENCE");
    calculateFormulas(wb);
    expect(values(ws)).toEqual([1, 2, 3]);
    expect(await storedAsDynamicArray(wb)).toBe(true);
  });
});

describe("RegisterFunctionOptions.returnsArray", () => {
  it("must be a boolean", () => {
    const wb = Workbook.create();
    expect(() =>
      Workbook.registerFunction(wb, "F", () => 1, { returnsArray: "yes" as never })
    ).toThrow(ExcelArgumentError);
    expect(() =>
      Workbook.registerFunction(wb, "F", () => 1, { returnsArray: "yes" as never })
    ).toThrow(/options\.returnsArray/);
  });
});
