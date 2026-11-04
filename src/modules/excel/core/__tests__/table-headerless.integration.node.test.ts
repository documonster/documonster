import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Table, Workbook } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, expect, it } from "vitest";

const excelTestDataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");

async function tableXml(buffer: Uint8Array | ArrayBuffer): Promise<string[]> {
  const files = await extractAll(new Uint8Array(buffer));
  return [...files.entries()]
    .filter(([path]) => /^xl\/tables\/table\d+\.xml$/.test(path))
    .map(([, entry]) => new TextDecoder().decode(entry.data));
}

describe("header-less tables", () => {
  it("are written without an autoFilter", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, {
      name: "T",
      ref: "A1",
      headerRow: false,
      columns: [{ name: "a" }, { name: "b" }],
      rows: [
        [1, 2],
        [3, 4]
      ]
    });
    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "header-less table" });
    const [xml] = await tableXml(buffer);
    expect(xml).toContain('headerRowCount="0"');
    expect(xml).not.toContain("<autoFilter");
  });

  it("stay without an autoFilter through read → write", async () => {
    const wb = Workbook.create();
    await Workbook.readFile(wb, excelTestDataPath("table-without-autofilter.xlsx"));
    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "table-without-autofilter rmw" });
    const [xml] = await tableXml(buffer);
    expect(xml).toContain('headerRowCount="0"');
    expect(xml).not.toContain("<autoFilter");
  });

  it("keep their autoFilter when the header row is turned back on", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const t = Table.add(ws, {
      name: "T",
      ref: "A1",
      headerRow: false,
      columns: [{ name: "a" }, { name: "b" }],
      rows: [[1, 2]]
    });
    Table.setHeaderRow(t, true);
    Table.commit(t);
    const [xml] = await tableXml(await Workbook.toBuffer(wb));
    expect(xml).toContain('<autoFilter ref="A1:B2"');
  });
});
