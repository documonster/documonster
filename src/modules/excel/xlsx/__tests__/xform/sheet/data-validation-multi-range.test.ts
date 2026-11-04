/**
 * A validation over several ranges is one rule, its relative references measured from the top-left of all of them.
 *
 * The reader stored `B2:B4 D6:D8` as two entries that both kept the formula stated for `B2`, and an entry's formula is
 * relative to its own corner — so after a round trip `D6:D8` validated against the cells beside `B2`. Each entry is
 * now slid to its own corner on read, and the writers state the rule once again over the full sqref.
 */
import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import type { DataValidationRule } from "@excel";
import { XmlParseError } from "@xml";
import { describe, expect, it } from "vitest";

const SHEET = "xl/worksheets/sheet1.xml";

async function sheetXml(buffer: Uint8Array): Promise<string> {
  const file = (await extractAll(buffer)).get(SHEET)!;
  return new TextDecoder().decode(file.data);
}

async function withValidations(rules: string): Promise<Uint8Array> {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValidation(sheet, "Z1", { type: "custom", formulae: ["TRUE"] } as DataValidationRule);
  const archive = zip();
  for (const [name, file] of await extractAll(await Workbook.toBuffer(workbook))) {
    if (file.type === "directory") {
      continue;
    }
    archive.add(
      name,
      name === SHEET
        ? new TextDecoder()
            .decode(file.data)
            .replace(/<dataValidations.*<\/dataValidations>/s, rules)
        : file.data
    );
  }
  return archive.toBuffer();
}

const MULTI =
  '<dataValidations count="1"><dataValidation type="custom" sqref="B2:B4 D6:D8">' +
  "<formula1>A2&gt;$Z$1</formula1></dataValidation></dataValidations>";

describe("a multi-range data validation", () => {
  it("reads each range relative to its own corner", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, await withValidations(MULTI));
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    const formulae = (address: string) =>
      (Cell.getValidation(sheet, address) as { formulae?: unknown[] }).formulae;
    expect(formulae("B3")).toEqual(["A2>$Z$1"]);
    expect(formulae("D7")).toEqual(["C6>$Z$1"]);
  });

  it("writes one rule over the full sqref with the original formula", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, await withValidations(MULTI));
    const xml = await sheetXml(await Workbook.toBuffer(workbook));
    expect(xml.match(/<dataValidation /g)).toHaveLength(1);
    expect(xml).toContain('sqref="B2:B4 D6:D8"');
    expect(xml).toContain("<formula1>A2&gt;$Z$1</formula1>");
  });

  it("survives XLSB in both directions", async () => {
    const workbook = Workbook.create();
    await Workbook.read(workbook, await withValidations(MULTI));
    const xlsb = await Workbook.toBuffer(workbook, { format: "xlsb" });
    const reread = Workbook.create();
    await Workbook.read(reread, xlsb);
    const sheet = Workbook.getWorksheet(reread, "S")!;
    expect((Cell.getValidation(sheet, "D7") as { formulae?: unknown[] }).formulae).toEqual([
      "C6>$Z$1"
    ]);
    const xml = await sheetXml(await Workbook.toBuffer(reread));
    expect(xml).toContain('sqref="B2:B4 D6:D8"');
    expect(xml).toContain("<formula1>A2&gt;$Z$1</formula1>");
  });

  it("does not merge two cells whose identical text means different cells", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    const rule = { type: "custom", formulae: ["C1>0"] } as DataValidationRule;
    Cell.setValidation(sheet, "A1", rule);
    Cell.setValidation(sheet, "A2", rule);
    const xml = await sheetXml(await Workbook.toBuffer(workbook));
    expect(xml.match(/<dataValidation /g)).toHaveLength(2);
  });
});

// A damaged element. The reader closed a rule it never opened and read `formulae` off `undefined` — an untyped
// `TypeError` instead of the XML error the mangled tag is; a rule with no `sqref` is skipped, and the intact rule
// beside it kept.
describe("a damaged data validation", () => {
  it("reports a misspelled open tag as an XML error", async () => {
    const damaged = await withValidations(
      '<dataValidations count="1"><dataVaMidation type="whole" sqref="G1:H1">' +
        "<formula1>0</formula1><formula2>0</formula2></dataValidation></dataValidations>"
    );
    await expect(Workbook.read(Workbook.create(), damaged)).rejects.toBeInstanceOf(XmlParseError);
  });

  it("skips a rule with no sqref", async () => {
    const workbook = Workbook.create();
    await Workbook.read(
      workbook,
      await withValidations(
        '<dataValidations count="2"><dataValidation type="custom" sqref="B2">' +
          "<formula1>TRUE</formula1></dataValidation>" +
          '<dataValidation type="custom"><formula1>FALSE</formula1></dataValidation></dataValidations>'
      )
    );
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    expect((Cell.getValidation(sheet, "B2") as { formulae?: unknown[] }).formulae).toEqual([
      "TRUE"
    ]);
  });
});
