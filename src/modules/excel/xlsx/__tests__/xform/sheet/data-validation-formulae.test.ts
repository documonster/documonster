/**
 * A data validation bound is a formula, and a number is only one kind of formula.
 *
 * Excel writes `<formula1>A1</formula1>` and `<formula2>$B$1+1</formula2>` for a `whole` rule bounded by cells. The
 * reader parsed every `whole`/`decimal`/`textLength` bound as a number, so those came back as `NaN` — `null` once
 * serialised — and the rule was rewritten as `<formula1>NaN</formula1>`. Every rule type is covered against every
 * shape of bound: a literal, a reference, an expression, a defined name and a reference to another sheet.
 */
import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import type { DataValidationRule } from "@excel";
import { describe, expect, it } from "vitest";

const SHEET = "xl/worksheets/sheet1.xml";
const WHEN = new Date(Date.UTC(2020, 0, 15));

// `time` is not in the published union, but the reader returns it and both writers accept it.
type Typed = Exclude<DataValidationRule["type"], "any"> | "time";

/** The literal each type is bounded by, as the model holds it and as the file states it. */
const LITERALS: Record<Typed, { model: string | number | Date; xml: string }> = {
  whole: { model: 10, xml: "10" },
  decimal: { model: -2.5, xml: "-2.5" },
  textLength: { model: 5, xml: "5" },
  // A time bound is read as the file's text, as it always was.
  time: { model: "0.5", xml: "0.5" },
  date: { model: WHEN, xml: "43845" },
  list: { model: '"Yes,No"', xml: '"Yes,No"' },
  custom: { model: "TRUE", xml: "TRUE" }
};

const FORMULAE = {
  reference: "A1",
  expression: "$B$1+1",
  "defined name": "Limit",
  "other sheet": "'Other Sheet'!$A$1",
  function: "TODAY()"
};

async function sheetXml(buffer: Uint8Array): Promise<string> {
  for (const [name, file] of await extractAll(buffer)) {
    if (name === SHEET && file.type !== "directory") {
      return new TextDecoder().decode(file.data);
    }
  }
  throw new Error(`no ${SHEET}`);
}

function build(type: Typed, formulae: (string | number | Date)[]) {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Workbook.addWorksheet(workbook, "Other Sheet");
  Cell.setValidation(sheet, "C1", {
    type,
    ...(type === "list" || type === "custom" ? {} : { operator: "between" }),
    formulae
  } as DataValidationRule);
  return workbook;
}

async function readBack(buffer: Uint8Array) {
  const workbook = Workbook.create();
  await Workbook.read(workbook, buffer);
  return Cell.getValidation(Workbook.getWorksheet(workbook, "S")!, "C1") as {
    type: string;
    formulae?: unknown[];
  };
}

describe("data validation bounds", () => {
  const types = Object.keys(LITERALS) as Typed[];

  describe.each(types)("%s", type => {
    it("writes and reads a literal bound as the value it states", async () => {
      const { model, xml } = LITERALS[type];
      const buffer = await Workbook.toBuffer(build(type, [model, model]));
      expect(await sheetXml(buffer)).toContain(
        `<formula1>${xml.replace(/"/g, "&quot;")}</formula1>`
      );
      const rule = await readBack(buffer);
      expect(rule.type).toBe(type);
      expect(rule.formulae).toEqual([model, model]);
    });

    it.each(Object.entries(FORMULAE))(
      "writes and reads a %s bound as its text",
      async (_, text) => {
        const buffer = await Workbook.toBuffer(build(type, [text, text]));
        const escaped = text.replace(/'/g, "&apos;").replace(/"/g, "&quot;");
        const xml = await sheetXml(buffer);
        expect(
          xml.includes(`<formula1>${text}</formula1>`) ||
            xml.includes(`<formula1>${escaped}</formula1>`)
        ).toBe(true);
        expect((await readBack(buffer)).formulae).toEqual([text, text]);
      }
    );
  });

  it("reads the bounds Excel writes for a rule bounded by cells", async () => {
    // As Excel 16 saved it: a `whole` rule over K1:K5 with `formula1` a relative reference and `formula2` an
    // expression, beside a `decimal`, `date`, `time` and `textLength` rule bounded the same ways.
    const rules =
      '<dataValidations count="5">' +
      '<dataValidation type="whole" sqref="K1:K5"><formula1>A1</formula1><formula2>$B$1+1</formula2></dataValidation>' +
      '<dataValidation type="decimal" operator="greaterThan" sqref="L1"><formula1>MyRange</formula1></dataValidation>' +
      '<dataValidation type="date" sqref="L2"><formula1>TODAY()</formula1><formula2>43845</formula2></dataValidation>' +
      '<dataValidation type="time" operator="lessThan" sqref="L3"><formula1>0.75</formula1></dataValidation>' +
      '<dataValidation type="textLength" operator="lessThan" sqref="L4"><formula1>Other!$A$1*2</formula1></dataValidation>' +
      "</dataValidations>";
    const source = await Workbook.toBuffer(build("whole", [1, 2]));
    const archive = zip();
    for (const [name, file] of await extractAll(source)) {
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
    const workbook = Workbook.create();
    await Workbook.read(workbook, await archive.toBuffer());
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    const formulae = (address: string) =>
      (Cell.getValidation(sheet, address) as { formulae?: unknown[] }).formulae;
    expect(formulae("K3")).toEqual(["A1", "$B$1+1"]);
    expect(formulae("L1")).toEqual(["MyRange"]);
    expect(formulae("L2")).toEqual(["TODAY()", WHEN]);
    expect(formulae("L3")).toEqual(["0.75"]);
    expect(formulae("L4")).toEqual(["Other!$A$1*2"]);

    const xml = await sheetXml(await Workbook.toBuffer(workbook));
    expect(xml).toContain("<formula1>A1</formula1><formula2>$B$1+1</formula2>");
    expect(xml).toContain("<formula1>TODAY()</formula1><formula2>43845</formula2>");
    expect(xml).toContain("<formula1>0.75</formula1>");
    expect(xml).toContain("<formula1>Other!$A$1*2</formula1>");
  });

  it.each(["1/15/2020", "2020-01-15", "Jan 15 2020", "2020-01-15T00:00:00"])(
    "writes the date string %j as a date bound",
    async text => {
      const buffer = await Workbook.toBuffer(build("date", [text]));
      expect(await sheetXml(buffer)).toContain("<formula1>43845</formula1>");
      expect((await readBack(buffer)).formulae).toEqual([WHEN]);
    }
  );

  it.each(["'Other Sheet'!$A$1", "TODAY()", "A1", "MyName", "$B$1+1", "A:A", "DATE(2020,1,15)"])(
    "writes the date bound %j as formula text",
    async text => {
      const buffer = await Workbook.toBuffer(build("date", [text]));
      const xml = await sheetXml(buffer);
      expect(
        xml.includes(`<formula1>${text}</formula1>`) ||
          xml.includes(`<formula1>${text.replace(/'/g, "&apos;")}</formula1>`)
      ).toBe(true);
      expect((await readBack(buffer)).formulae).toEqual([text]);
    }
  );
});
