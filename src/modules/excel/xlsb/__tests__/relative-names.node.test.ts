/**
 * A defined name with a relative or mixed reference survives a round trip in both containers, with its text intact.
 *
 * `formula-edges.excel.*` is Excel's own workbook: `Rel` is `S!A1` and `RelCol` is `S!$A1:$A3`. Both used to be read
 * as a reference with no cells and dropped on write, so a formula using either lost its definition.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { DefinedNames, Workbook } from "@excel";
import { iterateBiffRecords } from "@excel/xlsb/binary";
import { readWorkbookPart } from "@excel/xlsb/read/parts";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";
const BRT_NAME = 39;

const EXPECTED: Record<string, string> = {
  Rel: "S!A1",
  RelCol: "S!$A1:$A3",
  Glob: "S!$C$1",
  Uni: "S!$A$1,S!$B$2",
  WholeCol: "S!$A:$A",
  WholeRow: "S!$1:$1"
};

/** Each name's definition text, as the model would write it. */
function definitions(wb: Workbook.Handle): Map<string, string> {
  return new Map(
    DefinedNames.getModel(Workbook.getDefinedNames(wb)).map(entry => [
      entry.name,
      entry.kind === "opaque" ? entry.rawText! : (entry.formulaExpression ?? entry.ranges.join(","))
    ])
  );
}

/** Each `BrtName` record's payload in `xl/workbook.bin`, keyed by name. */
async function nameRecords(bytes: Uint8Array): Promise<Map<string, string>> {
  const workbook = (await extractAll(bytes)).get("xl/workbook.bin")!.data;
  const names = readWorkbookPart(workbook, "xl/workbook.bin").names.map(name => name.name);
  const payloads = [...iterateBiffRecords(workbook, "xl/workbook.bin")]
    .filter(record => record.id === BRT_NAME)
    .map(record => Array.from(record.payload, byte => byte.toString(16).padStart(2, "0")).join(""));
  return new Map(names.map((name, index) => [name, payloads[index]!]));
}

describe("relative defined names", () => {
  for (const source of ["formula-edges.excel.xlsx", "formula-edges.excel.xlsb"]) {
    for (const format of ["xlsx", "xlsb"] as const) {
      it(`${source} → ${format} keeps every reference name's text`, async () => {
        const wb = await Workbook.read(
          Workbook.create(),
          new Uint8Array(await readFile(`${DIRECTORY}/${source}`))
        );
        for (const [name, text] of Object.entries(EXPECTED)) {
          expect(definitions(wb).get(name), `read ${name}`).toBe(text);
        }
        const out = await Workbook.toBuffer(wb, { format, unsupported: "ignore" });
        const again = await Workbook.read(Workbook.create(), out);
        for (const [name, text] of Object.entries(EXPECTED)) {
          expect(definitions(again).get(name), `written ${name}`).toBe(text);
        }
      });
    }
  }

  it("writes Excel's own BrtName bytes for the relative names", async () => {
    const source = new Uint8Array(await readFile(`${DIRECTORY}/formula-edges.excel.xlsb`));
    const wb = await Workbook.read(Workbook.create(), source);
    const out = await Workbook.toBuffer(wb, { format: "xlsb", unsupported: "ignore" });
    const before = await nameRecords(source);
    const after = await nameRecords(out);
    for (const name of ["Rel", "RelCol"]) {
      expect(after.get(name), name).toBe(before.get(name));
    }
  });
});
