/**
 * A conditional-formatting rule's formulas survive XLSB, and a rule that has none is never written as one that tests
 * nothing — by either writer.
 *
 * The XLSB reader returned `formulae: []` for every comparison and expression rule. Written to XLSX that became a
 * `<cfRule type="cellIs">` with no `<formula>`; reading that file back and saving it threw
 * `Cannot read properties of undefined (reading 'forEach')` from the XLSX rule writer.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { ConditionalFormatting, ExcelUnsupportedFeatureError, Workbook, Worksheet } from "@excel";
import { getSheetModel } from "@excel/core/worksheet";
import { describe, expect, it } from "vitest";

type Block = { ref: string; rules: Record<string, unknown>[] };

function blocksOf(workbook: ReturnType<typeof Workbook.create>, sheet = "S"): Block[] {
  // The sheet's own model, not `Worksheet.getModel`'s copy: the tests edit it to reach the writers unchecked.
  return getSheetModel(Workbook.getWorksheet(workbook, sheet)!)
    .conditionalFormattings as unknown as Block[];
}

function sheetWith(rules: Record<string, unknown>[], ref = "B2:B4") {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Worksheet.addAoa(sheet, [
    ["H", "H"],
    [1, 1],
    [2, 5],
    [3, 9]
  ]);
  ConditionalFormatting.add(sheet, { ref, rules } as never);
  return workbook;
}

async function viaXlsb(workbook: ReturnType<typeof Workbook.create>) {
  const reopened = Workbook.create();
  await Workbook.read(reopened, await Workbook.toBuffer(workbook, { format: "xlsb" }));
  return reopened;
}

async function sheetXml(xlsx: Uint8Array): Promise<string> {
  const files = await extractAll(xlsx);
  return new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
}

describe("conditional-formatting formulas through XLSB", () => {
  it("brings a comparison's operands back, one and two", async () => {
    const reopened = await viaXlsb(
      sheetWith([
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: ["2"],
          style: { font: { bold: true } }
        },
        {
          type: "cellIs",
          operator: "between",
          formulae: ["1", "$A$4*2"],
          style: { font: { italic: true } }
        }
      ])
    );
    const rules = blocksOf(reopened)[0]!.rules;
    expect(rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan", formulae: ["2"] });
    expect(rules[1]).toMatchObject({
      type: "cellIs",
      operator: "between",
      formulae: ["1", "$A$4*2"]
    });
  });

  it("brings an expression back relative to the rule's range", async () => {
    // Stored once as offsets from B2 and applied by offset, so it has to come back as written against B2.
    const reopened = await viaXlsb(
      sheetWith([{ type: "expression", formulae: ["A2>1"], style: { font: { bold: true } } }])
    );
    expect(blocksOf(reopened)[0]!.rules[0]).toMatchObject({
      type: "expression",
      formulae: ["A2>1"]
    });
  });

  it("brings a containsText rule's search string back", async () => {
    const reopened = await viaXlsb(
      sheetWith([
        {
          type: "containsText",
          operator: "containsText",
          text: "ab",
          style: { font: { bold: true } }
        }
      ])
    );
    expect(blocksOf(reopened)[0]!.rules[0]).toMatchObject({ type: "containsText", text: "ab" });
  });

  it("survives XLSB → XLSX → read → XLSX with the formula in place", async () => {
    const reopened = await viaXlsb(
      sheetWith([
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: ["2"],
          style: { font: { bold: true } }
        }
      ])
    );
    const xlsx = await Workbook.toBuffer(reopened);
    expect(await sheetXml(xlsx)).toMatch(/<cfRule[^>]*type="cellIs"[^>]*>\s*<formula>2<\/formula>/);
    const again = Workbook.create();
    await Workbook.read(again, xlsx);
    const second = await Workbook.toBuffer(again);
    expect(await sheetXml(second)).toContain("<formula>2</formula>");
  });
});

describe("a comparison or expression rule with no formula", () => {
  const formulaLess = [
    { type: "cellIs", operator: "greaterThan", style: { font: { bold: true } } },
    { type: "expression", formulae: [], style: { font: { bold: true } } }
  ];

  it("is refused where it is added", () => {
    for (const rule of formulaLess) {
      expect(() => sheetWith([rule])).toThrow(/needs 1 formula/);
    }
  });

  // A model that did not come through `ConditionalFormatting.add` — one built by hand or carried over from elsewhere —
  // still reaches the writers, which keep refusing it.
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`is refused by the ${format} writer by default, and skipped under unsupported: "ignore"`, async () => {
      for (const rule of formulaLess) {
        const workbook = sheetWith([
          { type: "expression", formulae: ["A1"] },
          { type: "cellIs", operator: "lessThan", formulae: ["9"] }
        ]);
        blocksOf(workbook)[0]!.rules[0] = rule;
        await expect(Workbook.toBuffer(workbook, { format })).rejects.toBeInstanceOf(
          ExcelUnsupportedFeatureError
        );
        const written = await Workbook.toBuffer(workbook, { format, unsupported: "ignore" });
        const reopened = Workbook.create();
        await Workbook.read(reopened, written);
        const rules = blocksOf(reopened)[0]!.rules;
        expect(rules).toHaveLength(1);
        expect(rules[0]).toMatchObject({ type: "cellIs", operator: "lessThan", formulae: ["9"] });
      }
    });
  }

  it("read from an XLSX that has one, can be written again without throwing a TypeError", async () => {
    const workbook = sheetWith([{ type: "cellIs", operator: "lessThan", formulae: ["9"] }]);
    const xlsx = await Workbook.toBuffer(workbook);
    const files = await extractAll(xlsx);
    const entries = [...files.values()]
      .filter(file => file.type === "file")
      .map(file => {
        if (file.path !== "xl/worksheets/sheet1.xml") {
          return { name: file.path, data: file.data };
        }
        const xml = new TextDecoder().decode(file.data).replace(/<formula>9<\/formula>/, "");
        return { name: file.path, data: new TextEncoder().encode(xml) };
      });
    const damaged = Workbook.create();
    await Workbook.read(damaged, await createZip(entries));
    await expect(Workbook.toBuffer(damaged)).rejects.toBeInstanceOf(ExcelUnsupportedFeatureError);
    const written = await Workbook.toBuffer(damaged, { unsupported: "ignore" });
    expect(await sheetXml(written)).not.toContain("<cfRule");
  });
});
