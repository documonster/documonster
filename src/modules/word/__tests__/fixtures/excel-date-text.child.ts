/**
 * Run by `excel-date-text.node.test.ts` under different `TZ` values: prints,
 * as JSON, how `WordExcel.toDocx` and `Pdf.fromExcel` render the same cells.
 */
import { Cell, Workbook } from "@excel/index";
import { Pdf } from "@pdf/index";
import { WordExcel } from "@word/excel";
import { Query } from "@word/index";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "S");
Cell.setValue(ws, "A1", new Date("2024-03-10T00:00:00Z"));
Cell.setNumFmt(ws, "A1", "yyyy-mm-dd");
Cell.setValue(ws, "B1", new Date("2024-12-31T18:30:00Z"));
Cell.setNumFmt(ws, "B1", "d mmm yyyy hh:mm");
Cell.setValue(ws, "C1", 0.25);
Cell.setNumFmt(ws, "C1", "0%");
// No number format: Excel's default date display.
Cell.setValue(ws, "D1", new Date("2024-03-10T00:00:00Z"));

const word = Query.getText(WordExcel.toDocx(wb, { includeSheetHeadings: false }));
const pdf = (await Pdf.extract(await Pdf.fromExcel(wb))).text;
process.stdout.write(JSON.stringify({ word, pdf }));
