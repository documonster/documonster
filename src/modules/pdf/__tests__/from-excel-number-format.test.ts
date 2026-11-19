import { Cell, Column, Workbook } from "@excel/index";
import { keepNumFmtId } from "@excel/utils/kept-num-fmt";
import { PdfArgumentError } from "@pdf/errors";
import { excelWorkbookToPdf } from "@pdf/excel-bridge";
import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

const ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

describe("Pdf.fromExcel and number formats", () => {
  // `*x` repeats `x` until the text is as wide as the cell: only a renderer that knows the width can do it.
  it("fills a cell with its format's fill character, to the cell's width", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Column.setWidth(ws, "A", 20);
    Cell.setValue(ws, "A1", 42);
    Cell.setNumFmt(ws, "A1", "0*-");
    Cell.setValue(ws, "A2", -1234.5);
    Cell.setNumFmt(ws, "A2", ACCOUNTING);

    const sheet = (await excelWorkbookToPdf(wb)).sheets[0] as {
      rows: Map<number, { cells: Map<number, { text: string; fill?: unknown }> }>;
    };
    expect(sheet.rows.get(1)!.cells.get(1)).toMatchObject({
      text: "42",
      fill: { char: "-", at: 2 }
    });

    const text = (await Pdf.extract(await Pdf.fromExcel(wb))).pages[0]!.text;
    const dashes = /42(-+)/.exec(text)?.[1]?.length ?? 0;
    // A 20-character column holds far more than a few; it stops short of the cell's edge, never past it.
    expect(dashes).toBeGreaterThan(15);
    expect(dashes).toBeLessThan(40);
    // The currency sign stays at the left and the number at the right, with the fill between.
    expect(text).toMatch(/\$ {5,}\(1,234\.50\)/);
  });

  it("shows a currency the file names only by its id in the locale asked for", async () => {
    const wb = await Workbook.read(await Workbook.toBuffer(currencyWorkbook()));
    const latin = (await Pdf.extract(await Pdf.fromExcel(wb))).pages[0]!.text;
    const chinese = (await Pdf.extract(await Pdf.fromExcel(wb, { locale: "zh-CN" }))).pages[0]!
      .text;
    expect(latin).toContain("$3.50");
    expect(chinese).toContain("¥3.50");
  });

  it("refuses a locale there is none of, before it recalculates", async () => {
    let recalculated = false;
    const error = await Pdf.fromExcel(currencyWorkbook(), {
      locale: "de-DE" as "en-US",
      recalculate: () => {
        recalculated = true;
      }
    }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(PdfArgumentError);
    expect(String(error)).toMatch(/Pdf\.fromExcel: options\.locale must be one of/);
    expect(recalculated).toBe(false);
  });
});

/** A workbook whose A1 is formatted with built-in id 7, a currency with no code — as a read keeps it. */
function currencyWorkbook(): Workbook.Handle {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 3.5);
  // The id is kept the way a read keeps one: set aside on the style while `numFmt` is absent.
  keepNumFmtId(Cell.getStyle(ws, "A1"), 7);
  return wb;
}
