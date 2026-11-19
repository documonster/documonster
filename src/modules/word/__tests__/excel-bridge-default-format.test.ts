import { zip } from "@archive/create-archive";
import { Cell, Column, Workbook } from "@excel/index";
import { keepNumFmtId } from "@excel/utils/kept-num-fmt";
import { WordExcel } from "@word/excel";
import type { Paragraph, Table } from "@word/types";
import { describe, expect, it } from "vitest";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Default cell format `0.00` in Arial bold with a yellow fill; `cellXfs[1]` plain. A1 has no style, B1 is plain. */
async function workbookWithDefaultFormat(): Promise<Uint8Array> {
  const archive = zip();
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="${MAIN}"><fonts count="2"><font><sz val="10"/><name val="Verdana"/></font><font><b/><sz val="14"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="2" fontId="1" fillId="2" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>1.5</v></c><c r="B1" s="1"><v>2.5</v></c></row></sheetData></worksheet>`
  };
  for (const [name, text] of Object.entries(files)) {
    archive.add(name, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

describe("WordExcel.toDocx and the workbook's default format", () => {
  it("shows a cell with no style of its own in the default format, and takes the default font for Normal", async () => {
    const doc = WordExcel.toDocx(await Workbook.read(await workbookWithDefaultFormat()), {
      includeSheetHeadings: false
    });
    const normal = doc.styles?.find(style => style.styleId === "Normal");
    expect(normal?.runProperties).toMatchObject({ size: 20 });
    expect(JSON.stringify(normal?.runProperties)).toContain("Verdana");

    const table = doc.body.find(block => block.type === "table") as Table;
    const [a1, b1] = table.rows[0]!.cells;
    const run = (cell: typeof a1) =>
      (cell!.content[0] as Paragraph).children[0] as unknown as {
        properties?: { bold?: boolean };
        content: readonly { text: string }[];
      };
    expect(run(a1).content[0]!.text).toBe("1.50");
    expect(run(a1).properties?.bold).toBe(true);
    expect(a1!.properties?.shading).toBeDefined();
    expect(run(b1).content[0]!.text).toBe("2.5");
    expect(run(b1).properties?.bold).not.toBe(true);
    expect(b1!.properties?.shading).toBeUndefined();
  });
});

describe("WordExcel.toDocx and a format's fill", () => {
  const ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

  function firstParagraph(wb: Workbook.Handle, options?: { locale?: "zh-CN" }) {
    const doc = WordExcel.toDocx(wb, { includeSheetHeadings: false, ...options });
    const table = doc.body.find(block => block.type === "table") as Table;
    return table.rows[0]!.cells.map(cell => cell.content[0] as Paragraph);
  }

  // Word fills a width with a character through a tab stop's leader: a right stop at the cell's far edge puts what
  // follows the fill against it, as Excel does.
  it("is a right tab stop at the cell's edge, with the fill character as its leader", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Column.setWidth(ws, "A", 20);
    Cell.setValue(ws, "A1", -1234.5);
    Cell.setNumFmt(ws, "A1", ACCOUNTING);
    Cell.setValue(ws, "B1", 42);
    Cell.setNumFmt(ws, "B1", "0*-");
    Cell.setValue(ws, "C1", 5);
    Cell.setNumFmt(ws, "C1", "0*x");
    const [a1, b1, c1] = firstParagraph(wb);

    expect(a1!.properties?.tabs).toEqual([
      { type: "right", position: expect.any(Number), leader: "none" }
    ]);
    expect(a1!.properties!.tabs![0]!.position).toBeGreaterThan(1000);
    expect((a1!.children[0] as { content: unknown }).content).toEqual([
      { type: "text", text: " $" },
      { type: "tab" },
      { type: "text", text: "(1,234.50)" }
    ]);
    expect(b1!.properties?.tabs?.[0]?.leader).toBe("hyphen");
    // No leader draws an `x`: the text is shown as it is, without one.
    expect(c1!.properties?.tabs).toBeUndefined();
    expect((c1!.children[0] as { content: unknown }).content).toEqual([
      { type: "text", text: "5" }
    ]);
  });

  it("shows a currency the file names only by its id in the locale asked for", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 3.5);
    keepNumFmtId(Cell.getStyle(ws, "A1"), 7);
    const text = (p: Paragraph) => JSON.stringify(p.children);
    expect(text(firstParagraph(wb)[0]!)).toContain("$3.50");
    expect(text(firstParagraph(wb, { locale: "zh-CN" })[0]!)).toContain("¥3.50");
  });
});
