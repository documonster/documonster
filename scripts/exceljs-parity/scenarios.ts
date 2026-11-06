/**
 * The scenarios `verify-exceljs-parity.ts` runs: each one twice, as exceljs code and as the port the exceljs guide
 * gives for it. Keep the two halves line for line — a scenario is only evidence when the documonster half is what the
 * guide says to write.
 */
import {
  Cell,
  Column,
  ConditionalFormatting,
  DefinedNames,
  Row,
  Stream,
  Table,
  Workbook,
  Worksheet
} from "documonster/excel";
import { ExcelCsv } from "documonster/excel/csv";

/* eslint-disable @typescript-eslint/no-explicit-any -- the exceljs halves are untyped */
type Any = any;

async function xlsx(wb: Any): Promise<Uint8Array> {
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } } as const;
const BORDER = {
  top: { style: "thin" },
  bottom: { style: "double", color: { argb: "FFFF0000" } }
} as const;

export const scenarios = [
  {
    name: "values by row and by column, 1-based",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getRow(1).values = [, "a", "b", 3];
      ws.getRow(2).values = ["x", "y"];
      ws.getColumn(5).values = [, 1, 2, 3];
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Row.setValues(ws, 1, [, "a", "b", 3], { base: 1 });
      Row.setValues(ws, 2, ["x", "y"], { base: 0 });
      Column.setValues(ws, 5, [, 1, 2, 3], { base: 1 });
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "cell, row and column styles",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getCell("A1").value = 1;
      ws.getCell("A1").font = { bold: true, color: { argb: "FF0000FF" } };
      ws.getCell("A1").fill = FILL;
      ws.getCell("A1").border = BORDER;
      ws.getCell("A1").alignment = { horizontal: "center", wrapText: true };
      ws.getCell("A1").numFmt = "0.00";
      ws.getCell("A2").value = "r";
      ws.getCell("B2").value = "r2";
      ws.getRow(2).font = { italic: true };
      ws.getCell("C1").value = "c";
      ws.getColumn(3).fill = FILL;
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setFont(ws, "A1", { bold: true, color: { argb: "FF0000FF" } });
      Cell.setFill(ws, "A1", FILL);
      Cell.setBorder(ws, "A1", BORDER);
      Cell.setAlignment(ws, "A1", { horizontal: "center", wrapText: true });
      Cell.setNumFmt(ws, "A1", "0.00");
      Cell.setValue(ws, "A2", "r");
      Cell.setValue(ws, "B2", "r2");
      Row.setFont(ws, 2, { italic: true });
      Cell.setValue(ws, "C1", "c");
      Column.setFill(ws, 3, FILL);
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "row.splice → Row.spliceCells",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getRow(1).values = [, "a", "b", "c", "d"];
      ws.getCell("B1").font = { bold: true };
      ws.getCell("C1").note = "keep me";
      ws.getRow(1).splice(1, 1);
      ws.getRow(1).splice(2, 0, "x", "y");
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Row.setValues(ws, 1, [, "a", "b", "c", "d"], { base: 1 });
      Cell.setFont(ws, "B1", { bold: true });
      Cell.setNote(ws, "C1", "keep me");
      Row.spliceCells(ws, 1, 1, 1);
      Row.spliceCells(ws, 1, 2, 0, "x", "y");
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "merges, with and without the master's style",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getCell("A1").value = "m";
      ws.getCell("A1").fill = FILL;
      ws.mergeCells("A1:B2");
      ws.getCell("D1").value = "n";
      ws.getCell("D1").fill = FILL;
      ws.mergeCellsWithoutStyle("D1:E2");
      ws.mergeCells(4, 1, 5, 2);
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", "m");
      Cell.setFill(ws, "A1", FILL);
      Worksheet.merge(ws, "A1:B2");
      Cell.setValue(ws, "D1", "n");
      Cell.setFill(ws, "D1", FILL);
      Worksheet.mergeWithoutStyle(ws, "D1:E2");
      Worksheet.merge(ws, 4, 1, 5, 2);
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "notes, plain and rich",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getCell("A1").value = 1;
      ws.getCell("A1").note = "plain";
      ws.getCell("B1").note = {
        texts: [{ text: "rich", font: { bold: true } }],
        editAs: "absolute"
      };
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setNote(ws, "A1", "plain");
      Cell.setNote(ws, "B1", {
        texts: [{ text: "rich", font: { bold: true } }],
        editAs: "absolute"
      });
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "data validation and conditional formatting",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getCell("A1").dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ['"a,b,c"'],
        showErrorMessage: true,
        errorTitle: "No",
        error: "Pick one"
      };
      ws.getCell("B1").dataValidation = { type: "whole", operator: "between", formulae: [1, 10] };
      ws.addConditionalFormatting({
        ref: "C1:C5",
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: ["5"],
            style: { font: { bold: true } }
          }
        ]
      });
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValidation(ws, "A1", {
        type: "list",
        allowBlank: true,
        formulae: ['"a,b,c"'],
        showErrorMessage: true,
        errorTitle: "No",
        error: "Pick one"
      });
      Cell.setValidation(ws, "B1", { type: "whole", operator: "between", formulae: [1, 10] });
      ConditionalFormatting.add(ws, {
        ref: "C1:C5",
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: ["5"],
            priority: 1,
            style: { font: { bold: true } }
          }
        ]
      });
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "tables and defined names",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.addTable({
        name: "Sales",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Item" }, { name: "Qty" }],
        rows: [
          ["a", 1],
          ["b", 2]
        ]
      });
      wb.definedNames.add("S!$B$2:$B$3", "qty");
      ws.getCell("E1").name = "single";
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Table.add(ws, {
        name: "Sales",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Item" }, { name: "Qty" }],
        rows: [
          ["a", 1],
          ["b", 2]
        ]
      });
      DefinedNames.add(Workbook.getDefinedNames(wb), "S!$B$2:$B$3", "qty");
      Cell.setNames(ws, "E1", ["single"]);
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "page setup, views, auto filter, sheet state",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S", { pageSetup: { paperSize: 9, orientation: "landscape" } });
      ws.getCell("A1").value = "h";
      ws.pageSetup.fitToPage = true;
      ws.pageSetup.printArea = "A1:C10";
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = "A1:C1";
      wb.addWorksheet("Hidden", { state: "hidden" });
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S", {
        pageSetup: { paperSize: 9, orientation: "landscape" }
      });
      Cell.setValue(ws, "A1", "h");
      ws.pageSetup.fitToPage = true;
      ws.pageSetup.printArea = "A1:C10";
      ws.views = [{ state: "frozen", ySplit: 1 }];
      Worksheet.setAutoFilter(ws, "A1:C1");
      Workbook.addWorksheet(wb, "Hidden", { state: "hidden" });
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "column definitions, keyed rows, row and column settings",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.columns = [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 25 }
      ];
      ws.addRow({ id: 1, name: "one" });
      ws.addRow([2, "two"]);
      ws.getRow(3).height = 30;
      ws.getRow(2).hidden = true;
      ws.getRow(3).outlineLevel = 1;
      ws.getColumn(2).hidden = true;
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Column.setDefinitions(ws, [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 25 }
      ]);
      Row.add(ws, { id: 1, name: "one" });
      Row.add(ws, [2, "two"]);
      Row.setHeight(ws, 3, 30);
      Row.setHidden(ws, 2, true);
      Row.setOutlineLevel(ws, 3, 1);
      Column.setHidden(ws, 2, true);
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "inserting and splicing rows",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.addRows([["r1"], ["r2"], ["r3"], ["r4"]]);
      ws.getCell("A2").font = { bold: true };
      ws.insertRow(2, ["new"], "i");
      ws.spliceRows(4, 1);
      ws.duplicateRow(1, 1, true);
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Row.addMany(ws, [["r1"], ["r2"], ["r3"], ["r4"]]);
      Cell.setFont(ws, "A2", { bold: true });
      Row.insert(ws, 2, ["new"], { copyStyle: "above" });
      Row.splice(ws, 4, 1);
      Row.duplicate(ws, 1, 1, { insert: true });
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "values: dates, hyperlinks, rich text, formulas, errors, booleans",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.getCell("A1").value = new Date(Date.UTC(2024, 0, 15, 12, 30));
      ws.getCell("A2").value = { text: "Docs", hyperlink: "https://example.com" };
      ws.getCell("A3").value = {
        richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }]
      };
      ws.getCell("A4").value = { formula: "A5*2", result: 6 };
      ws.getCell("A5").value = 3;
      ws.getCell("A6").value = { error: "#N/A" };
      ws.getCell("A7").value = true;
      return xlsx(wb);
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", new Date(Date.UTC(2024, 0, 15, 12, 30)));
      Cell.setValue(ws, "A2", { text: "Docs", hyperlink: "https://example.com" });
      Cell.setValue(ws, "A3", {
        richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }]
      });
      Cell.setValue(ws, "A4", { formula: "A5*2", result: 6 });
      Cell.setValue(ws, "A5", 3);
      Cell.setValue(ws, "A6", { error: "#N/A" });
      Cell.setValue(ws, "A7", true);
      return Workbook.toBuffer(wb);
    }
  },
  {
    name: "streaming writer",
    async exceljs(ExcelJS: Any) {
      const { PassThrough } = await import("node:stream");
      const pass = new PassThrough();
      const chunks: Buffer[] = [];
      pass.on("data", (c: Buffer) => chunks.push(c));
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: pass, useStyles: true });
      const ws = wb.addWorksheet("S");
      ws.columns = [{ header: "A", key: "a", width: 12 }];
      const row = ws.addRow({ a: 1 });
      row.getCell(1).font = { bold: true };
      ws.mergeCells("C2:D2");
      row.commit();
      ws.addRow({ a: 2 }).commit();
      ws.commit();
      await wb.commit();
      return new Uint8Array(Buffer.concat(chunks));
    },
    async documonster() {
      const wb = new Stream.WorkbookWriter({ useStyles: true });
      const ws = Workbook.addWorksheet(wb, "S");
      Column.setDefinitions(ws, [{ header: "A", key: "a", width: 12 }]);
      const row = ws.addRow({ a: 1 });
      Stream.setCellFont(Stream.getRowCell(row, 1), { bold: true });
      Worksheet.merge(ws, "C2:D2");
      Stream.commitRow(row);
      Stream.commitRow(ws.addRow({ a: 2 }));
      ws.commit();
      return wb.toBuffer();
    }
  }
];

/** What a CSV read put in the sheet, as the same plain rows from either library. */
function rows(values: unknown[][]): unknown[][] {
  return values.map(r => r.map(v => (v instanceof Date ? v.toISOString() : v)));
}

const CSV = 'name,qty,when,ok\nwidget,3,2024-01-15T10:00:00,true\n"a, b",4.5,2024-01-15,false\n';

/** A timestamp that states its zone: exceljs kept it as text, documonster reads the instant (the guide's §17). */
const ZONED = "at\n2024-01-15T10:00:00Z\n2024-01-15T10:00:00+02:00\n";

export const csvScenarios = [
  {
    name: "CSV read",
    async exceljs(ExcelJS: Any, readable: (text: string) => unknown) {
      const wb = new ExcelJS.Workbook();
      const ws = await wb.csv.read(readable(CSV));
      const out: unknown[][] = [];
      ws.eachRow((row: Any) => out.push(row.values.slice(1)));
      return rows(out);
    },
    async documonster() {
      const wb = await ExcelCsv.read(CSV);
      return rows(Worksheet.getValues(Workbook.listWorksheets(wb)[0]!, { base: 0 }) as unknown[][]);
    }
  },
  {
    name: "CSV read: zoned timestamps, with exceljs's date formats (the guide's way to keep them as text)",
    async exceljs(ExcelJS: Any, readable: (text: string) => unknown) {
      const wb = new ExcelJS.Workbook();
      const ws = await wb.csv.read(readable(ZONED));
      const out: unknown[][] = [];
      ws.eachRow((row: Any) => out.push(row.values.slice(1)));
      return rows(out);
    },
    async documonster() {
      const wb = await ExcelCsv.read(ZONED, {
        dateFormats: ["MM-DD-YYYY", "YYYY-MM-DD[T]HH:mm:ss", "YYYY-MM-DD"]
      });
      return rows(Worksheet.getValues(Workbook.listWorksheets(wb)[0]!, { base: 0 }) as unknown[][]);
    }
  },
  {
    name: "CSV write",
    async exceljs(ExcelJS: Any) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("S");
      ws.addRows([
        ["name", "qty"],
        ["a, b", 1.5],
        ['say "hi"', 2]
      ]);
      return (await wb.csv.writeBuffer()).toString();
    },
    async documonster() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Row.addMany(ws, [
        ["name", "qty"],
        ["a, b", 1.5],
        ['say "hi"', 2]
      ]);
      return new TextDecoder().decode(await ExcelCsv.toBuffer(wb));
    }
  }
];

/* eslint-enable @typescript-eslint/no-explicit-any */
