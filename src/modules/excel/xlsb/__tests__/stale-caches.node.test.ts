/**
 * A rewrite drops the caches that describe the source package's cells, in XLSB as it does in XLSX.
 *
 * `calcChain.bin` lists every formula cell in the order Excel last evaluated them. Excel refuses a workbook whose
 * chain names a cell that no longer holds a formula — or misses one that now does — so carrying the original
 * through a rewrite that changed any cell produced a file Excel would not open (all seven Excel-saved fixtures
 * below, after appending a value and a formula). The XLSX writer already declines to write `calcChain.xml`;
 * `fullCalcOnLoad` makes Excel rebuild the chain on open.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

const FIXTURES = [
  "dynamic-array.excel.xlsb",
  "eta-lambda.excel.xlsb",
  "formula-edges.excel.xlsb",
  "formula-matrix.excel.xlsb",
  "future-functions.excel.xlsb",
  "let-oracle.excel.xlsb",
  "table-this-row.excel.xlsb"
];

describe("XLSB stale caches", () => {
  it.each(FIXTURES)("%s: an edited rewrite carries no calcChain", async name => {
    const source = new Uint8Array(await readFile(`${DIRECTORY}/${name}`));
    expect((await extractAll(source)).has("xl/calcChain.bin")).toBe(true);

    const wb = await Workbook.read(Workbook.create(), source);
    const ws = Workbook.getWorksheets(wb)[0]!;
    Cell.setValue(ws, "A200", "x");
    Cell.setValue(ws, "B200", { formula: "1+1", result: 2 });
    const out = await extractAll(
      await Workbook.toBuffer(wb, { format: "xlsb", unsupported: "ignore" })
    );

    expect([...out.keys()].filter(path => /calcchain/i.test(path))).toEqual([]);
    const text = (path: string) => new TextDecoder().decode(out.get(path)!.data);
    expect(text("[Content_Types].xml")).not.toMatch(/calcchain/i);
    expect(text("xl/_rels/workbook.bin.rels")).not.toMatch(/calcChain/i);
  });

  it("reports the drop", async () => {
    const source = new Uint8Array(await readFile(`${DIRECTORY}/let-oracle.excel.xlsb`));
    const wb = await Workbook.read(Workbook.create(), source);
    const drops = (wb as unknown as { _opaqueDrops?: { path: string; reason: string }[] })
      ._opaqueDrops;
    expect(drops).toContainEqual(
      expect.objectContaining({ path: "xl/calcChain.bin", reason: "stale-cache" })
    );
  });
});
