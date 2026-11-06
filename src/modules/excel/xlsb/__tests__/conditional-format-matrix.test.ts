/**
 * Every conditional-formatting rule type through both containers, and across from one to the other.
 *
 * The property asserted is the one a caller relies on: **reading an XLSB gives the model reading the XLSX gives**,
 * for the same workbook. Comparing the XLSB reader with the XLSB writer cannot catch a field both get wrong — that
 * is how `CFTextOper` stayed 0 for every text rule while a symmetry test passed — and comparing either with the
 * input catches only what the input happened to state. XLSX is the reference because its reader is the one with
 * Excel's own markup behind it.
 */

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { ConditionalFormatting, Row, Workbook, Worksheet } from "@excel";
import { getSheetModel } from "@excel/core/worksheet";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { Formula } from "@formula";
import { describe, expect, it } from "vitest";

type Rule = Parameters<typeof ConditionalFormatting.add>[1]["rules"][number];

const BOLD = { font: { bold: true } };

/**
 * A rule in a shape readers return but the public union does not describe — `{ type: "beginsWith" }`,
 * `{ type: "duplicateValues" }` (see `ContainsTextRuleType`). Writers accept them, so they are tested as input.
 */
const runtime = (rule: Record<string, unknown>): Rule => rule as unknown as Rule;

const RULES: readonly (readonly [string, string, Rule])[] = [
  ["cellIs greaterThan", "B2:B9", { type: "cellIs", operator: "greaterThan", formulae: ["5"] }],
  ["cellIs between", "B2:B9", { type: "cellIs", operator: "between", formulae: ["2", "7"] }],
  ["cellIs absolute", "B2:B9", { type: "cellIs", operator: "equal", formulae: ["$A$1"] }],
  ["cellIs relative", "C5:D9", { type: "cellIs", operator: "lessThan", formulae: ["E5"] }],
  ["cellIs string", "B2:B9", { type: "cellIs", operator: "equal", formulae: ['"x"'] }],
  ["expression", "C5:D9", { type: "expression", formulae: ["C5>D5"] }],
  [
    "expression with areas",
    "C5:D9",
    { type: "expression", formulae: ["AND($C5>0,C$5<10,SUM($A$1:A1)>0)"] }
  ],
  ["containsText", "C5:D9", { type: "containsText", operator: "containsText", text: "abc" }],
  [
    "notContainsText, as readers return it",
    "C5:D9",
    runtime({ type: "notContainsText", operator: "notContains", text: 'q"z' })
  ],
  [
    "notContainsText, operator spelling",
    "C5:D9",
    runtime({ type: "containsText", operator: "notContainsText", text: "q" })
  ],
  [
    "beginsWith, as readers return it",
    "C5:D9",
    runtime({ type: "beginsWith", operator: "beginsWith", text: "ab" })
  ],
  ["beginsWith, type only", "C5:D9", runtime({ type: "beginsWith", text: "ab" })],
  ["endsWith", "C5:D9", runtime({ type: "endsWith", operator: "endsWith", text: "yz" })],
  ["containsBlanks", "C5:D9", { type: "containsText", operator: "containsBlanks" }],
  ["containsBlanks as a type", "C5:D9", runtime({ type: "containsBlanks" })],
  ["notContainsBlanks", "C5:D9", { type: "containsText", operator: "notContainsBlanks" }],
  ["containsErrors", "C5:D9", { type: "containsText", operator: "containsErrors" }],
  ["notContainsErrors", "C5:D9", { type: "containsText", operator: "notContainsErrors" }],
  ["timePeriod", "C5:D9", { type: "timePeriod", timePeriod: "lastWeek" }],
  ["timePeriod today", "C5:D9", { type: "timePeriod", timePeriod: "today" }],
  ["top10", "B2:B9", { type: "top10", rank: 3, percent: false, bottom: false }],
  ["top10 bottom percent", "B2:B9", { type: "top10", rank: 20, percent: true, bottom: true }],
  // Stated explicitly: CT_CfRule defaults `aboveAverage` to true, so an XLSX that omits it reads back without the
  // field, while BrtBeginCFRule's template states it and the binary reader has always reported `true`.
  ["aboveAverage", "B2:B9", { type: "aboveAverage", aboveAverage: true }],
  ["belowAverage", "B2:B9", { type: "aboveAverage", aboveAverage: false }],
  [
    "equal or below average",
    "B2:B9",
    { type: "aboveAverage", aboveAverage: false, equalAverage: true }
  ],
  ["two deviations above", "B2:B9", { type: "aboveAverage", aboveAverage: true, stdDev: 2 }],
  ["duplicateValues", "B2:B9", runtime({ type: "duplicateValues" })],
  ["uniqueValues", "B2:B9", runtime({ type: "uniqueValues" })],
  [
    "colorScale",
    "B2:B9",
    {
      type: "colorScale",
      cfvo: [{ type: "min" }, { type: "max" }],
      color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }]
    }
  ],
  [
    "iconSet",
    "B2:B9",
    {
      type: "iconSet",
      iconSet: "3Arrows",
      cfvo: [
        { type: "percent", value: 0 },
        { type: "percent", value: 33 },
        { type: "percent", value: 67 }
      ]
    }
  ],
  ["multi-range sqref", "B2:B4 D6:D8", { type: "expression", formulae: ["B2>1"] }]
];

/** The graphical rules carry a style-less rule; every other one is given a format, as Excel's UI gives one. */
function styled(rule: Rule): Rule {
  return ["colorScale", "dataBar", "iconSet"].includes(rule.type) ? rule : { ...rule, style: BOLD };
}

function workbookWith(ref: string, rule: Rule) {
  const handle = Workbook.create();
  const sheet = Workbook.addWorksheet(handle, "S");
  for (let row = 1; row <= 10; row++) {
    Row.add(sheet, [row, row * 2, `ab${row}`, row % 3 ? "x" : null, row]);
  }
  ConditionalFormatting.add(sheet, { ref, rules: [{ ...styled(rule), priority: 1 }] });
  return handle;
}

/**
 * A workbook carrying `rule` without `ConditionalFormatting.add`'s checks — as a model built by hand reaches the
 * writers — so a writer's own refusal can be tested.
 */
function workbookWithUnchecked(ref: string, rule: Rule) {
  const handle = workbookWith(ref, { type: "expression", formulae: ["A1"] });
  const sheet = Workbook.getWorksheet(handle, "S")!;
  // The sheet's own model: `Worksheet.getModel` hands out a copy, which is what the public API must do.
  getSheetModel(sheet).conditionalFormattings![0]!.rules[0] = {
    ...styled(rule),
    priority: 1
  } as never;
  return handle;
}

async function readBack(bytes: Uint8Array, format: "xlsx" | "xlsb") {
  const handle = Workbook.create();
  await Workbook.read(handle, bytes, { format });
  return handle;
}

function formattings(handle: ReturnType<typeof Workbook.create>): unknown {
  // A round trip through JSON drops the `undefined` fields one reader sets and the other leaves out.
  //
  // Formulas are compared as the formula printer writes them, which normalises spacing. The redundant parentheses of
  // Excel's own `…>=(WEEKDAY(TODAY()))` survive the binary file now (`PtgParen`); the Excel-made fixtures in
  // `excel-corpus.node.test.ts` compare the text exactly.
  return JSON.parse(
    JSON.stringify(
      Worksheet.getModel(Workbook.listWorksheets(handle)[0]!).conditionalFormattings ?? [],
      (key, value: unknown) =>
        key === "formulae" && Array.isArray(value)
          ? value.map(formula =>
              typeof formula === "string" ? Formula.format(Formula.parse(formula)) : formula
            )
          : // A default is what an absent attribute means, and the binary records cannot tell "absent" from
            // "stated as the default" — so the XLSB reader leaves defaults out, as Excel's own XLSX does, while this
            // library's XLSX writer states `aboveAverage="1"` and reads it back. Only the defaults are dropped here.
            (key === "showValue" && value === true) ||
              (key === "reverse" && value === false) ||
              (key === "aboveAverage" && value === true)
            ? undefined
            : value
    )
  );
}

describe("conditional formatting: XLSB reads what XLSX reads", () => {
  it.each(RULES)("%s", async (_name, ref, rule) => {
    const viaXlsx = formattings(
      await readBack(await Workbook.toBuffer(workbookWith(ref, rule), { format: "xlsx" }), "xlsx")
    );
    const xlsbBytes = await Workbook.toBuffer(workbookWith(ref, rule), { format: "xlsb" });
    const viaXlsb = await readBack(xlsbBytes, "xlsb");
    expect(formattings(viaXlsb)).toEqual(viaXlsx);
    // And on across: what XLSB gave back, written as XLSX and read again, is still the same rule.
    const across = await readBack(await Workbook.toBuffer(viaXlsb, { format: "xlsx" }), "xlsx");
    expect(formattings(across)).toEqual(viaXlsx);
    // Nothing vanished on the way: one block, one rule.
    expect((viaXlsx as { rules: unknown[] }[])[0]?.rules).toHaveLength(1);
  });
});

describe("a data bar", () => {
  it("keeps its thresholds and its one colour", async () => {
    // Compared field by field rather than whole: the XLSX writer adds the Excel 2010 extension (`x14:dataBar`, with
    // its gradient, axis and negative-bar settings), which a `BrtCFRuleExt` would carry in XLSB and this writer does
    // not emit. What the classic `BrtBeginDatabar` holds must agree.
    const rule: Rule = {
      type: "dataBar",
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: "FF638EC6" }
    };
    const pick = (handle: ReturnType<typeof Workbook.create>) => {
      const [block] = formattings(handle) as { ref: string; rules: Record<string, unknown>[] }[];
      const { type, cfvo, color, priority } = block!.rules[0]!;
      return { ref: block!.ref, type, cfvo, color, priority };
    };
    const viaXlsx = pick(
      await readBack(
        await Workbook.toBuffer(workbookWith("B2:B9", rule), { format: "xlsx" }),
        "xlsx"
      )
    );
    const viaXlsb = await readBack(
      await Workbook.toBuffer(workbookWith("B2:B9", rule), { format: "xlsb" }),
      "xlsb"
    );
    expect(pick(viaXlsb)).toEqual(viaXlsx);
    expect(
      pick(await readBack(await Workbook.toBuffer(viaXlsb, { format: "xlsx" }), "xlsx"))
    ).toEqual(viaXlsx);
  });
});

describe("text conditions mean what they say", () => {
  const formulaOf = async (rule: Rule) => {
    const xlsb = await readBack(
      await Workbook.toBuffer(workbookWith("C5:D9", rule), { format: "xlsb" }),
      "xlsb"
    );
    const xlsx = await readBack(await Workbook.toBuffer(xlsb, { format: "xlsx" }), "xlsx");
    return (formattings(xlsx) as { rules: { formulae?: string[] }[] }[])[0]!.rules[0]!.formulae;
  };

  it("keeps a notContains rule inverted through XLSB and back to XLSX", async () => {
    // It came back as `containsText` and was written as `NOT(ISERROR(SEARCH(…)))` — the opposite condition.
    expect(
      await formulaOf(runtime({ type: "notContainsText", operator: "notContains", text: "ab" }))
    ).toEqual(['ISERROR(SEARCH("ab",C5))']);
    expect(
      await formulaOf(runtime({ type: "beginsWith", operator: "beginsWith", text: "ab" }))
    ).toEqual(['LEFT(C5,LEN("ab"))="ab"']);
    expect(
      await formulaOf(runtime({ type: "endsWith", operator: "endsWith", text: "yz" }))
    ).toEqual(['RIGHT(C5,LEN("yz"))="yz"']);
  });

  it("doubles a quote in the search string", async () => {
    expect(
      await formulaOf({ type: "containsText", operator: "containsText", text: 'q"z' })
    ).toEqual(['NOT(ISERROR(SEARCH("q""z",C5)))']);
  });
});

/** `iTemplate` and `iParam` of every `BrtBeginCFRule`, and the `sqrfx` count of every block. */
async function ruleFields(handle: ReturnType<typeof Workbook.create>) {
  const parts = await extractAll(await Workbook.toBuffer(handle, { format: "xlsb" }));
  const rules: { template: number; param: number }[] = [];
  const blocks: number[] = [];
  for (const entry of iterateInterpretableRecords(
    parts.get("xl/worksheets/sheet1.bin")!.data,
    "s"
  )) {
    const view = new DataView(entry.payload.buffer, entry.payload.byteOffset);
    const name = recordSpec(entry.id)?.name;
    if (name === "BrtBeginCFRule") {
      rules.push({ template: view.getUint32(4, true), param: view.getUint32(16, true) });
    } else if (name === "BrtBeginConditionalFormatting") {
      blocks.push(view.getInt32(8, true));
    }
  }
  return { rules, blocks };
}

describe("BrtBeginCFRule fields, against MS-XLSB", () => {
  it("writes CFTextOper as iParam of the text template (2.5.17)", async () => {
    const params: unknown[] = [];
    for (const operator of ["containsText", "notContains", "beginsWith", "endsWith"]) {
      const { rules } = await ruleFields(
        workbookWith("C5:D9", runtime({ type: "containsText", operator, text: "ab" }))
      );
      params.push(rules[0]);
    }
    expect(params).toEqual([
      { template: 0x08, param: 0 },
      { template: 0x08, param: 1 },
      { template: 0x08, param: 2 },
      { template: 0x08, param: 3 }
    ]);
  });

  it("writes CFDateOper, not CFTemp, as a time period's iParam (2.5.12)", async () => {
    const found: unknown[] = [];
    for (const timePeriod of ["today", "yesterday", "lastWeek", "thisMonth"] as const) {
      const { rules } = await ruleFields(workbookWith("C5:D9", { type: "timePeriod", timePeriod }));
      found.push(rules[0]);
    }
    expect(found).toEqual([
      { template: 0x0f, param: 0 },
      { template: 0x11, param: 1 },
      { template: 0x17, param: 4 },
      { template: 0x18, param: 9 }
    ]);
  });

  it("writes every range of a multi-range sqref (2.5.156)", async () => {
    const { blocks } = await ruleFields(
      workbookWith("B2:B4 D6:D8", { type: "expression", formulae: ["B2>1"] })
    );
    expect(blocks).toEqual([2]);
  });

  it("refuses a text rule with no search string, which strParam cannot be", async () => {
    expect(() => workbookWith("C5:D9", runtime({ type: "beginsWith" }))).toThrow(
      /text must be given/
    );
    await expect(
      Workbook.toBuffer(workbookWithUnchecked("C5:D9", runtime({ type: "beginsWith" })), {
        format: "xlsb"
      })
    ).rejects.toThrow(/search string/);
  });
});

describe("a whitespace-only formula is a missing formula in both containers", () => {
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`is refused by the ${format} writer`, async () => {
      expect(() => workbookWith("B2:B9", { type: "expression", formulae: [" "] })).toThrow(
        /expression rule needs 1 formula/
      );
      await expect(
        Workbook.toBuffer(workbookWithUnchecked("B2:B9", { type: "expression", formulae: [" "] }), {
          format
        })
      ).rejects.toThrow(/expression rule needs 1 formula/);
    });
  }
});

describe("the shapes readers have always returned, from Excel's own markup", () => {
  // Captured from the baseline build reading these `<cfRule>`s. Three text conditions keep XLSX's spelling of type
  // and operator; `text` is new — the attribute was not read — and is the only field added.
  const CASES: readonly (readonly [string, string, Record<string, unknown>])[] = [
    [
      "containsText",
      '<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="abc"><formula>NOT(ISERROR(SEARCH("abc",C5)))</formula></cfRule>',
      {
        type: "containsText",
        operator: "containsText",
        text: "abc",
        formulae: ['NOT(ISERROR(SEARCH("abc",C5)))']
      }
    ],
    [
      "notContainsText",
      '<cfRule type="notContainsText" dxfId="0" priority="1" operator="notContains" text="abc"><formula>ISERROR(SEARCH("abc",C5))</formula></cfRule>',
      {
        type: "notContainsText",
        operator: "notContains",
        text: "abc",
        formulae: ['ISERROR(SEARCH("abc",C5))']
      }
    ],
    [
      "beginsWith",
      '<cfRule type="beginsWith" dxfId="0" priority="1" operator="beginsWith" text="ab"><formula>LEFT(C5,LEN("ab"))="ab"</formula></cfRule>',
      {
        type: "beginsWith",
        operator: "beginsWith",
        text: "ab",
        formulae: ['LEFT(C5,LEN("ab"))="ab"']
      }
    ],
    [
      "endsWith",
      '<cfRule type="endsWith" dxfId="0" priority="1" operator="endsWith" text="yz"><formula>RIGHT(C5,LEN("yz"))="yz"</formula></cfRule>',
      { type: "endsWith", operator: "endsWith", text: "yz", formulae: ['RIGHT(C5,LEN("yz"))="yz"'] }
    ],
    [
      "containsBlanks",
      '<cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(C5))=0</formula></cfRule>',
      { type: "containsText", operator: "containsBlanks", formulae: ["LEN(TRIM(C5))=0"] }
    ],
    [
      "duplicateValues",
      '<cfRule type="duplicateValues" dxfId="0" priority="1"/>',
      { type: "duplicateValues" }
    ],
    [
      "uniqueValues",
      '<cfRule type="uniqueValues" dxfId="0" priority="1"/>',
      { type: "uniqueValues" }
    ]
  ];

  /** A workbook whose sheet carries exactly `rule` on C5:D9, as markup. */
  async function withRule(rule: string): Promise<Uint8Array> {
    const handle = Workbook.create();
    const sheet = Workbook.addWorksheet(handle, "S");
    ConditionalFormatting.add(sheet, {
      ref: "C5:D9",
      rules: [{ type: "expression", formulae: ["A1"], priority: 1, style: BOLD }]
    });
    const entries = await extractAll(await Workbook.toBuffer(handle));
    return createZip(
      [...entries]
        .filter(([, file]) => file.type !== "directory")
        .map(([name, file]) => ({
          name,
          data:
            name === "xl/worksheets/sheet1.xml"
              ? new TextEncoder().encode(
                  new TextDecoder().decode(file.data).replace(/<cfRule[\s\S]*<\/cfRule>/, rule)
                )
              : file.data
        }))
    );
  }

  const shapeOf = (handle: ReturnType<typeof Workbook.create>) => {
    const rule = {
      ...(formattings(handle) as { rules: Record<string, unknown>[] }[])[0]!.rules[0]!
    };
    // `style` is the formatting, compared elsewhere; `dxfId` and `ref` are what the binary reader has always kept
    // beside it.
    for (const key of ["style", "dxfId", "ref", "priority"]) {
      delete rule[key];
    }
    return rule;
  };

  it.each(CASES)("%s", async (_name, markup, expected) => {
    const fromXlsx = await readBack(await withRule(markup), "xlsx");
    expect(shapeOf(fromXlsx)).toEqual(expected);
    // Written to XLSB — which used to refuse three of these and write every text rule as "contains" — and read back.
    const fromXlsb = await readBack(await Workbook.toBuffer(fromXlsx, { format: "xlsb" }), "xlsb");
    expect(shapeOf(fromXlsb)).toEqual(expected);
    // And written to XLSX again, which used to drop five of them.
    const again = await readBack(await Workbook.toBuffer(fromXlsb, { format: "xlsx" }), "xlsx");
    expect(shapeOf(again)).toEqual(expected);
  });
});
