import { zip } from "@archive/create-archive";
import { Cell, Column, ExcelArgumentError, Workbook, Worksheet } from "@excel/index";
import { DISPLAY_LOCALES, shownCodeOfKeptId } from "@excel/utils/display-locale";
import type { ExcelDisplayLocale } from "@excel/utils/display-locale";
import { describe, expect, it } from "vitest";

/**
 * The codes Excel for Mac 16.115 wrote back for the currencies (5–8, 63–66), the accounting formats (41–44) and
 * the reserved 23–26 — saving a workbook of every id under each locale, with its own interface language. An id
 * missing from a locale's list is one Excel wrote no code for there: 23–26 as General, 63–66 as 5–8.
 */
const MEASURED: Readonly<Record<ExcelDisplayLocale, Readonly<Record<number, string>>>> = {
  "en-US": {
    5: '"$"#,##0_);\\("$"#,##0\\)',
    6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
    8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
  },
  "en-AU": {
    5: '"$"#,##0_);\\("$"#,##0\\)',
    6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
    8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
  },
  "en-GB": {
    5: '"£"#,##0_);\\("£"#,##0\\)',
    6: '"£"#,##0_);[Red]\\("£"#,##0\\)',
    7: '"£"#,##0.00_);\\("£"#,##0.00\\)',
    8: '"£"#,##0.00_);[Red]\\("£"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("£"* #,##0_);_("£"* \\(#,##0\\);_("£"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("£"* #,##0.00_);_("£"* \\(#,##0.00\\);_("£"* "-"??_);_(@_)'
  },
  "zh-CN": {
    5: '"¥"#,##0_);\\("¥"#,##0\\)',
    6: '"¥"#,##0_);[Red]\\("¥"#,##0\\)',
    7: '"¥"#,##0.00_);\\("¥"#,##0.00\\)',
    8: '"¥"#,##0.00_);[Red]\\("¥"#,##0.00\\)',
    23: "\\$#,##0_);\\(\\$#,##0\\)",
    24: "\\$#,##0_);[Red]\\(\\$#,##0\\)",
    25: "\\$#,##0.00_);\\(\\$#,##0.00\\)",
    26: "\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("¥"* #,##0_);_("¥"* \\(#,##0\\);_("¥"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("¥"* #,##0.00_);_("¥"* \\(#,##0.00\\);_("¥"* "-"??_);_(@_)'
  },
  "zh-TW": {
    5: '"$"#,##0_);\\("$"#,##0\\)',
    6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
    8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    23: '"US$"#,##0_);\\("US$"#,##0\\)',
    24: '"US$"#,##0_);[Red]\\("US$"#,##0\\)',
    25: '"US$"#,##0.00_);\\("US$"#,##0.00\\)',
    26: '"US$"#,##0.00_);[Red]\\("US$"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
  },
  "zh-HK": {
    5: '"HK$"#,##0_);\\("HK$"#,##0\\)',
    6: '"HK$"#,##0_);[Red]\\("HK$"#,##0\\)',
    7: '"HK$"#,##0.00_);\\("HK$"#,##0.00\\)',
    8: '"HK$"#,##0.00_);[Red]\\("HK$"#,##0.00\\)',
    23: '"US$"#,##0_);\\("US$"#,##0\\)',
    24: '"US$"#,##0_);[Red]\\("US$"#,##0\\)',
    25: '"US$"#,##0.00_);\\("US$"#,##0.00\\)',
    26: '"US$"#,##0.00_);[Red]\\("US$"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("HK$"* #,##0_);_("HK$"* \\(#,##0\\);_("HK$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("HK$"* #,##0.00_);_("HK$"* \\(#,##0.00\\);_("HK$"* "-"??_);_(@_)'
  },
  "ja-JP": {
    5: '"¥"#,##0;"¥"\\-#,##0',
    6: '"¥"#,##0;[Red]"¥"\\-#,##0',
    7: '"¥"#,##0.00;"¥"\\-#,##0.00',
    8: '"¥"#,##0.00;[Red]"¥"\\-#,##0.00',
    23: "\\$#,##0_);\\(\\$#,##0\\)",
    24: "\\$#,##0_);[Red]\\(\\$#,##0\\)",
    25: "\\$#,##0.00_);\\(\\$#,##0.00\\)",
    26: "\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)",
    41: '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ',
    42: '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ ',
    43: '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
    44: '_ "¥"* #,##0.00_ ;_ "¥"* \\-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ '
  },
  "ko-KR": {
    5: '"₩"#,##0_);\\("₩"#,##0\\)',
    6: '"₩"#,##0_);[Red]\\("₩"#,##0\\)',
    7: '"₩"#,##0.00_);\\("₩"#,##0.00\\)',
    8: '"₩"#,##0.00_);[Red]\\("₩"#,##0.00\\)',
    23: "\\$#,##0_);\\(\\$#,##0\\)",
    24: "\\$#,##0_);[Red]\\(\\$#,##0\\)",
    25: "\\$#,##0.00_);\\(\\$#,##0.00\\)",
    26: "\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("₩"* #,##0_);_("₩"* \\(#,##0\\);_("₩"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("₩"* #,##0.00_);_("₩"* \\(#,##0.00\\);_("₩"* "-"??_);_(@_)'
  },
  "th-TH": {
    5: '"฿"#,##0_);\\("฿"#,##0\\)',
    6: '"฿"#,##0_);[Red]\\("฿"#,##0\\)',
    7: '"฿"#,##0.00_);\\("฿"#,##0.00\\)',
    8: '"฿"#,##0.00_);[Red]\\("฿"#,##0.00\\)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("฿"* #,##0_);_("฿"* \\(#,##0\\);_("฿"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("฿"* #,##0.00_);_("฿"* \\(#,##0.00\\);_("฿"* "-"??_);_(@_)',
    63: 't"฿"#,##0_);\\(t"฿"#,##0\\)',
    64: 't"฿"#,##0_);[Red]\\(t"฿"#,##0\\)',
    65: 't"฿"#,##0.00_);\\(t"฿"#,##0.00\\)',
    66: 't"฿"#,##0.00_);[Red]\\(t"฿"#,##0.00\\)'
  }
};

describe("shownCodeOfKeptId", () => {
  it.each(DISPLAY_LOCALES)(
    "gives the currencies and accounting formats Excel writes in %s",
    locale => {
      const measured = MEASURED[locale];
      for (const id of [5, 6, 7, 8, 23, 24, 25, 26, 41, 42, 43, 44, 63, 64, 65, 66]) {
        const expected = measured[id] ?? (id <= 8 || id >= 41 ? undefined : "General");
        expect(shownCodeOfKeptId(id, locale), `${locale} ${id}`).toBe(
          expected ?? measured[id >= 63 ? id - 58 : id]
        );
      }
    }
  );

  it("is en-US's when no locale is named", () => {
    expect(shownCodeOfKeptId(7)).toBe(shownCodeOfKeptId(7, "en-US"));
    expect(shownCodeOfKeptId(30)).toBe("mm-dd-yy");
    expect(shownCodeOfKeptId(82)).toBeUndefined();
  });
});

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A row per id: 2023-03-15 18:00 in A, -1234.5 in B, both formatted with the id and nothing else. */
function workbookOfIds(ids: readonly number[]): Promise<Uint8Array> {
  const xfs = ids.map(id => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0"/>`).join("");
  const rows = ids
    .map(
      (_, i) =>
        `<row r="${i + 1}"><c r="A${i + 1}" s="${i + 1}"><v>45000.75</v></c><c r="B${i + 1}" s="${i + 1}"><v>-1234.5</v></c></row>`
    )
    .join("");
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="${MAIN}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="${ids.length + 1}"><xf numFmtId="0"/>${xfs}</cellXfs></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`
  };
  const archive = zip();
  for (const [name, text] of Object.entries(files)) {
    archive.add(name, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

// What a cell shows in each locale, as Excel for Mac or LibreOffice shows it in that locale (see the module's notes
// for where the two part, and which one this follows).
describe("a cell whose format is a code-less built-in, shown in a locale", () => {
  const CASES: readonly [ExcelDisplayLocale, number, string, string][] = [
    ["en-US", 7, "$45,000.75 ", "($1,234.50)"],
    ["en-GB", 5, "£45,001 ", "(£1,235)"],
    ["zh-CN", 7, "¥45,000.75 ", "(¥1,234.50)"],
    ["zh-CN", 27, "2023年3月", "########"],
    // A negative date under a time format shows the time of its day, as LibreOffice shows it.
    ["zh-CN", 34, "下午6时00分", "下午12时00分"],
    ["zh-TW", 27, "112/3/15", "########"],
    ["zh-TW", 23, "US$45,001 ", "(US$1,235)"],
    ["zh-HK", 28, "15日3月112年", "########"],
    ["ja-JP", 7, "¥45,000.75", "¥-1,234.50"],
    ["ja-JP", 27, "R5.3.15", "########"],
    ["ja-JP", 28, "令和5年3月15日", "########"],
    ["ko-KR", 31, "2023년 03월 15일", "########"],
    ["ko-KR", 5, "₩45,001 ", "(₩1,235)"],
    ["th-TH", 59, "๔๕๐๐๑", "-๑๒๓๕"],
    ["th-TH", 63, "฿๔๕,๐๐๑ ", "(฿๑,๒๓๕)"],
    ["th-TH", 71, "๑๕/๓/๒๕๖๖", "########"],
    ["th-TH", 81, "15/3/66", "########"],
    ["en-US", 71, "03-15-23", "########"],
    // What Excel and LibreOffice both show in English, one built-in along from the Thai form from 72 on.
    ["en-US", 73, "15-Mar-23", "########"],
    ["en-US", 75, "Mar-23", "########"],
    ["en-US", 77, "18:00:00", "12:00:00"],
    ["en-US", 79, "00:00", "00:00"],
    ["en-US", 81, "00:00.0", "00:00.0"]
  ];

  it("shows as Excel or LibreOffice in that locale shows it", async () => {
    const ids = [...new Set(CASES.map(([, id]) => id))];
    const ws = Workbook.getWorksheet(await Workbook.read(await workbookOfIds(ids)), "S")!;
    for (const [locale, id, positive, negative] of CASES) {
      const row = ids.indexOf(id) + 1;
      expect(Cell.getNumFmt(ws, { row, col: 1 }), `${id}`).toBeUndefined();
      expect(Cell.getDisplayText(ws, { row, col: 1 }, { locale }), `${locale} ${id}`).toBe(
        positive
      );
      expect(Cell.getDisplayText(ws, { row, col: 2 }, { locale }), `${locale} ${id} negative`).toBe(
        negative
      );
    }
  });
});

describe("the locale option of the display functions", () => {
  async function sheetOf(id: number) {
    return Workbook.getWorksheet(await Workbook.read(await workbookOfIds([id])), "S")!;
  }

  it("is refused by Cell.getDisplayText when it names no locale there is", async () => {
    const ws = await sheetOf(7);
    expect(() => Cell.getDisplayText(ws, "A1", { locale: "de-DE" as ExcelDisplayLocale })).toThrow(
      /Cell\.getDisplayText: options\.locale must be one of "en-US"/
    );
    // The deprecated position form still reads a column, not options.
    expect(Cell.getDisplayText(ws, 1, 1)).toBe("$45,000.75 ");
  });

  it("is what Worksheet.toJson shows a formatted value in", async () => {
    const ws = await sheetOf(7);
    expect(Worksheet.toJson(ws, { header: 1, raw: false, locale: "ko-KR" })).toEqual([
      ["₩45,000.75", "(₩1,234.50)"]
    ]);
  });

  it("is what Column.autoFit measures", async () => {
    const ws = await sheetOf(28);
    Column.autoFit(ws, "A");
    const english = Column.getWidth(ws, "A");
    Column.autoFit(ws, "A", { locale: "ja-JP" });
    // `03-15-23` against `令和5年3月15日`: wider glyphs, and more of them.
    expect(Column.getWidth(ws, "A")).toBeGreaterThan(english!);
    expect(() => Column.autoFit(ws, "A", { locale: "xx" as ExcelDisplayLocale })).toThrow(
      /Column\.autoFit: options\.locale/
    );
    // The deprecated forwards never took a locale; they refuse a wrong one as the module's argument error too.
    const wrong = { locale: "xx" as ExcelDisplayLocale };
    expect(() => Worksheet.autoFitColumn(ws, 1, wrong)).toThrow(ExcelArgumentError);
    expect(() => Worksheet.autoFitColumns(ws, wrong)).toThrow(ExcelArgumentError);
    expect(() => Worksheet.autoFitColumns(ws, 1, 1, wrong)).toThrow(ExcelArgumentError);
  });
});
