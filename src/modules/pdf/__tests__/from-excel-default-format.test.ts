import { zip } from "@archive/create-archive";
import { Cell, Workbook } from "@excel/index";
import { excelWorkbookToPdf } from "@pdf/excel-bridge";
import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * A workbook whose default cell format (`cellXfs[0]`) is `0.00` in Arial bold with a fill, and whose `cellXfs[1]` is
 * General in the first font with no fill. A1 has no style; B1 names `cellXfs[1]`; row 3 is not in the sheet.
 */
async function workbookWithDefaultFormat(): Promise<Uint8Array> {
  const archive = zip();
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="${MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="2" fontId="1" fillId="2" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>1.5</v></c><c r="B1" s="1"><v>2.5</v></c></row><row r="4"><c r="B4"><v>4</v></c></row></sheetData></worksheet>`
  };
  for (const [name, text] of Object.entries(files)) {
    archive.add(name, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

describe("Pdf.fromExcel and the workbook's default cell format", () => {
  it("shows a cell with no number format of its own in the default's", async () => {
    const wb = await Workbook.read(await workbookWithDefaultFormat());
    const text = (await Pdf.extract(await Pdf.fromExcel(wb))).pages[0].text;
    expect(text).toContain("1.50");
    expect(text).toContain("2.5");
  });

  // A cell with no style of its own is shown in the default format; one with a plain format of its own is not.
  it("draws a cell with no style of its own in the default's font and fill, and every cell it does not hold too", async () => {
    const wb = await Workbook.read(await workbookWithDefaultFormat());
    const sheet = (await excelWorkbookToPdf(wb)).sheets[0]! as {
      rows: Map<number, { cells: Map<number, { text: string; style?: Record<string, any> }> }>;
    };
    const a1 = sheet.rows.get(1)!.cells.get(1)!;
    expect(a1.text).toBe("1.50");
    expect(a1.style?.font).toMatchObject({ name: "Arial", bold: true, size: 14 });
    expect(a1.style?.fill?.pattern).toBe("solid");
    const b1 = sheet.rows.get(1)!.cells.get(2)!;
    expect(b1.style?.font?.bold).not.toBe(true);
    expect(b1.style?.fill?.pattern ?? "none").toBe("none");
    // Row 3 is not in the sheet, and A4 not in its row: both show the default's fill.
    expect(sheet.rows.get(3)?.cells.get(1)?.style?.fill?.pattern).toBe("solid");
    expect(sheet.rows.get(4)?.cells.get(1)?.style?.fill?.pattern).toBe("solid");
  });

  it("takes the workbook's default font for a cell with none, unless the caller names one", async () => {
    const wb = Workbook.create();
    Workbook.setDefaultFont(wb, { name: "Courier", size: 9 });
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", "x");
    const latin1 = (bytes: Uint8Array) => new TextDecoder("latin1").decode(bytes);
    expect(latin1(await Pdf.fromExcel(wb))).toContain("/BaseFont /Courier");
    expect(latin1(await Pdf.fromExcel(wb, { defaultFontFamily: "Times-Roman" }))).not.toContain(
      "/BaseFont /Courier"
    );
  });
});
