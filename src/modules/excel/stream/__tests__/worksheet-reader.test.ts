import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { cellFont, cellGetValue, cellNumFmt, cellOwnStyle, cellSetValue } from "@excel/core/cell";
import type { RowData } from "@excel/core/row";
import type { WorkbookData } from "@excel/core/workbook-core";
import { rowGetCell } from "@excel/core/worksheet";
import { ExcelStreamStateError } from "@excel/errors";
import { Cell, Row, Workbook, Worksheet } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { describe, it, expect } from "vitest";

// =============================================================================
// Helpers
// =============================================================================

/**
 * Create a minimal XLSX buffer from a builder callback.
 * Avoids filesystem I/O — fast + cross-platform.
 */
async function buildXlsxBuffer(builder: (wb: WorkbookData) => void): Promise<Uint8Array> {
  const wb = Workbook.create();
  builder(wb);
  return Workbook.toBuffer(wb);
}

/**
 * Rewrite one sheet's XML inside an XLSX buffer. Used to produce shapes our writer
 * never emits but other producers do — e.g. `<row>`/`<c>` without the optional `r`.
 */
async function patchSheetXml(
  buffer: Uint8Array,
  patch: (xml: string) => string,
  path = "xl/worksheets/sheet1.xml"
): Promise<Uint8Array> {
  const entries = await extractAll(buffer);
  const archive = zip();
  for (const [name, file] of entries) {
    if (file.type === "directory") {
      continue;
    }
    archive.add(name, name === path ? patch(new TextDecoder().decode(file.data)) : file.data);
  }
  return archive.bytes();
}

/** Result of streaming a single worksheet */
interface SheetResult {
  name: string;
  id: number | string;
  sheetNo: number;
  rows: RowData[];
  dimensions: unknown;
  columns: unknown[];
}

/**
 * Read an XLSX buffer with the streaming WorkbookReader and collect all
 * worksheets + their rows. This consumes everything inside the reader's
 * iteration context (the only correct way to use the async iterator API).
 */
async function readAllSheets(
  buffer: Uint8Array,
  readerOptions?: Record<string, string>
): Promise<SheetResult[]> {
  const reader = new WorkbookReader(buffer, {
    worksheets: "emit",
    sharedStrings: "cache",
    styles: "cache",
    hyperlinks: "cache",
    entries: "emit",
    ...readerOptions
  });

  const sheets: SheetResult[] = [];
  for await (const ws of reader) {
    const rows: RowData[] = [];
    for await (const row of ws) {
      rows.push(row as RowData);
    }
    sheets.push({
      name: ws.name,
      id: ws.id,
      sheetNo: (ws as any).sheetNo,
      rows,
      dimensions: (ws as any).dimensions,
      columns: (ws as any).columns ?? []
    });
  }
  return sheets;
}

/**
 * Convenience: read a single-sheet buffer and return its rows + metadata.
 */
async function readFirstSheet(
  buffer: Uint8Array,
  readerOptions?: Record<string, string>
): Promise<SheetResult> {
  const sheets = await readAllSheets(buffer, readerOptions);
  if (sheets.length === 0) {
    throw new Error("No worksheet found in buffer");
  }
  return sheets[0];
}

// =============================================================================
// Tests
// =============================================================================

describe("WorksheetReader", () => {
  // ===========================================================================
  // Lifecycle
  // ===========================================================================

  describe("lifecycle", () => {
    it("destroy() throws ExcelStreamStateError", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 1);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        expect(() => ws.destroy()).toThrow(ExcelStreamStateError);
        // Must consume rows to avoid hanging
        for await (const _row of ws) {
          /* drain */
        }
      }
    });
  });

  // ===========================================================================
  // Properties
  // ===========================================================================

  describe("properties", () => {
    it("has correct name from workbook metadata", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "MySheet"), "A1", 1);
      });
      const sheet = await readFirstSheet(buffer);
      expect(sheet.name).toBe("MySheet");
    });

    it("exposes sheet id and sheetNo", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "First"), "A1", 1);
      });
      const sheet = await readFirstSheet(buffer);
      expect(sheet.id).toBeDefined();
      expect(sheet.sheetNo).toBeTypeOf("number");
    });
  });

  // ===========================================================================
  // Dimensions
  // ===========================================================================

  describe("dimensions", () => {
    it("reflects the parsed data extent", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 1);
        Cell.setValue(ws, "C3", 2);
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(2);
      expect(sheet.dimensions).toBeDefined();
    });
  });

  // ===========================================================================
  // Columns
  // ===========================================================================

  describe("columns", () => {
    it("returns column definitions after parsing", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Worksheet.setColumns(ws, [
          { header: "ID", key: "id", width: 10 },
          { header: "Name", key: "name", width: 32 }
        ]);
        Worksheet.addRow(ws, { id: 1, name: "Alice" });
      });
      const sheet = await readFirstSheet(buffer);

      // Columns should be populated from the <cols> XML element
      expect(sheet.columns.length).toBeGreaterThanOrEqual(2);
    });

    it("getColumn() by number creates columns on demand", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 1);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        // Getting column 5 should create columns 1-5
        const col = ws.getColumn(5);
        expect(col).toBeDefined();
        expect(ws.columns.length).toBeGreaterThanOrEqual(5);

        for await (const _row of ws) {
          /* drain */
        }
      }
    });

    it("getColumn() by letter resolves correctly", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 1);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        const colC = ws.getColumn("C");
        expect(colC).toBeDefined();
        expect(colC.number).toBe(3);

        for await (const _row of ws) {
          /* drain */
        }
      }
    });

    it("column key management works correctly", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 1);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        const col = ws.getColumn(1);
        ws.setColumnKey("myKey", col);
        expect(ws.getColumnKey("myKey")).toBe(col);

        // getColumn by key should resolve
        expect(ws.getColumn("myKey")).toBe(col);

        // eachColumnKey iterates
        const keys: string[] = [];
        ws.eachColumnKey((_column, key) => keys.push(key));
        expect(keys).toContain("myKey");

        // delete
        ws.deleteColumnKey("myKey");
        expect(ws.getColumnKey("myKey")).toBeUndefined();

        for await (const _row of ws) {
          /* drain */
        }
      }
    });
  });

  // ===========================================================================
  // Async Iteration — Cell Value Types
  // ===========================================================================

  describe("cell value types", () => {
    it("parses numbers", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 42);
        Cell.setValue(ws, "B1", 3.14);
        Cell.setValue(ws, "C1", -100);
        Cell.setValue(ws, "D1", 0);
      });
      const sheet = await readFirstSheet(buffer);
      const row = sheet.rows[0];

      expect(cellGetValue(rowGetCell(row, 1))).toBe(42);
      expect(cellGetValue(rowGetCell(row, 2))).toBe(3.14);
      expect(cellGetValue(rowGetCell(row, 3))).toBe(-100);
      expect(cellGetValue(rowGetCell(row, 4))).toBe(0);
    });

    it("parses strings via shared strings", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "hello");
        Cell.setValue(ws, "B1", "world");
        Cell.setValue(ws, "A2", "hello"); // duplicate — same shared string
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe("hello");
      expect(cellGetValue(rowGetCell(sheet.rows[0], 2))).toBe("world");
      expect(cellGetValue(rowGetCell(sheet.rows[1], 1))).toBe("hello");
    });

    it("parses booleans", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", true);
        Cell.setValue(ws, "B1", false);
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe(true);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 2))).toBe(false);
    });

    it("parses dates when numFmt indicates a date format", async () => {
      const testDate = new Date(2024, 0, 15); // Jan 15, 2024
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", testDate);
        Cell.setStyle(ws, "A1", { numFmt: "yyyy-mm-dd" });
      });
      const sheet = await readFirstSheet(buffer);
      const cellValue = cellGetValue(rowGetCell(sheet.rows[0], 1));

      expect(cellValue).toBeInstanceOf(Date);
      expect((cellValue as Date).getFullYear()).toBe(2024);
    });

    it("parses formulas with numeric results", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 10);
        Cell.setValue(ws, "A2", { formula: "A1*2", result: 20 });
      });
      const sheet = await readFirstSheet(buffer);
      const val = cellGetValue(rowGetCell(sheet.rows[1], 1)) as { formula: string; result: number };

      expect(val.formula).toBe("A1*2");
      expect(val.result).toBe(20);
    });

    it("parses formulas with string results", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", {
          formula: 'CONCATENATE("Hello", " ", "World")',
          result: "Hello World"
        });
      });
      const sheet = await readFirstSheet(buffer);
      const val = cellGetValue(rowGetCell(sheet.rows[0], 1)) as { formula: string; result: string };

      expect(val.formula).toContain("CONCATENATE");
      expect(val.result).toBe("Hello World");
    });

    it("parses error values", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", { error: "#DIV/0!" as any });
      });
      const sheet = await readFirstSheet(buffer);
      const val = cellGetValue(rowGetCell(sheet.rows[0], 1)) as { error: string };

      expect(val.error).toBe("#DIV/0!");
    });

    it("parses rich text via shared strings", async () => {
      const richText = {
        richText: [
          { font: { bold: true }, text: "Bold " },
          { font: { italic: true }, text: "Italic" }
        ]
      };
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", richText);
      });
      const sheet = await readFirstSheet(buffer);
      const val = cellGetValue(rowGetCell(sheet.rows[0], 1)) as {
        richText: Array<{ text: string }>;
      };

      expect(val).toHaveProperty("richText");
      expect(val.richText.length).toBe(2);
      expect(val.richText[0].text).toBe("Bold ");
      expect(val.richText[1].text).toBe("Italic");
    });
  });

  // ===========================================================================
  // Row Properties
  // ===========================================================================

  describe("row properties", () => {
    it("parses row height", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        const row = Worksheet.getRow(ws, 1);
        row.height = 30;
        cellSetValue(rowGetCell(row, 1), "tall row");
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].height).toBe(30);
    });

    it("parses row number correctly for sparse rows", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "row1");
        Cell.setValue(ws, "A5", "row5"); // rows 2-4 are empty
      });
      const sheet = await readFirstSheet(buffer);

      // Only non-empty rows are emitted
      expect(sheet.rows.length).toBe(2);
      expect(sheet.rows[0].number).toBe(1);
      expect(sheet.rows[1].number).toBe(5);
    });
  });

  // ===========================================================================
  // Cell shapes the streaming reader used to get wrong — read by `CellXform` now
  // ===========================================================================

  describe("cell values agree with the full reader", () => {
    const SHEET_DATA = [
      "<sheetData>",
      '<row r="1">',
      '<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain</t></r></is></c>',
      '<c r="B1" t="b"><f>1=1</f><v>1</v></c>',
      '<c r="C1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
      '<c r="D1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>',
      '<c r="E1" t="d"><v>2020-01-15T00:00:00</v></c>',
      '<c r="F1" t="inlineStr"><is><t>plain _x0041_</t></is></c>',
      "</row>",
      '<row r="3"><c r="A3"><v>1</v></c><c r="B3"><f t="shared" ref="B3:B5" si="0">A3*2</f><v>2</v></c></row>',
      '<row r="4"><c r="A4"><v>2</v></c><c r="B4"><f t="shared" si="0"/><v>4</v></c></row>',
      '<row r="5"><c r="A5"><v>3</v></c><c r="B5"><f t="shared" si="0"/><v>6</v></c></row>',
      "</sheetData>"
    ].join("");

    async function buildFixture(): Promise<Uint8Array> {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 0);
      });
      return patchSheetXml(buffer, xml =>
        xml.replace(/<sheetData>.*<\/sheetData>|<sheetData\/>/s, SHEET_DATA)
      );
    }

    it("reads each shape correctly when streaming", async () => {
      const sheet = await readFirstSheet(await buildFixture());
      const [row1, row3, row4, row5] = sheet.rows;
      const value = (row: RowData, col: number) => cellGetValue(rowGetCell(row, col));

      expect(value(row1, 1)).toEqual({
        richText: [{ text: "Bold", font: { bold: true } }, { text: " plain" }]
      });
      expect(value(row1, 2)).toMatchObject({ formula: "1=1", result: true });
      expect(value(row1, 3)).toMatchObject({ formula: "1/0", result: { error: "#DIV/0!" } });
      expect(value(row1, 4)).toMatchObject({ formula: '"a"&"b"', result: "ab" });
      expect(value(row1, 5)).toEqual(new Date(Date.UTC(2020, 0, 15)));
      expect(value(row1, 6)).toBe("plain A");

      expect(value(row3, 2)).toMatchObject({ formula: "A3*2", result: 2 });
      expect(value(row4, 2)).toMatchObject({ formula: "A4*2", sharedFormula: "B3", result: 4 });
      expect(value(row5, 2)).toMatchObject({ formula: "A5*2", sharedFormula: "B3", result: 6 });
    });

    it("matches the full reader cell for cell", async () => {
      const fixture = await buildFixture();
      const sheet = await readFirstSheet(fixture);
      const wb = Workbook.create();
      await Workbook.read(wb, fixture);
      const ws = Workbook.getWorksheet(wb, "Sheet1")!;
      for (const row of sheet.rows) {
        for (let col = 1; col <= row.cells.length; col++) {
          const streamed = cellGetValue(rowGetCell(row, col));
          const loaded = Cell.getValue(ws, row.number, col);
          if (streamed && typeof streamed === "object" && "sharedFormula" in streamed) {
            // A streamed shared-formula cell also carries its slid formula text, which a loaded one derives on demand.
            expect(streamed).toMatchObject(loaded as object);
            expect(Cell.getFormula(ws, row.number, col)).toBe(
              (streamed as { formula: string }).formula
            );
          } else {
            expect(streamed).toEqual(loaded);
          }
        }
      }
    });
  });

  describe("hyperlinks after the sheet", () => {
    it("are on worksheet.hyperlinks once its rows are read, with external targets from the reader", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", { text: "ext", hyperlink: "https://example.com/" });
        Cell.setValue(ws, "B2", { text: "int", hyperlink: "#Sheet1!A1" });
      });
      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        hyperlinks: "cache"
      });
      for await (const ws of reader) {
        const rows: RowData[] = [];
        for await (const row of ws) {
          rows.push(row as RowData);
        }
        // `<hyperlinks>` follows `<sheetData>`, so a row carries only the display text…
        expect(cellGetValue(rowGetCell(rows[0], 1))).toBe("ext");
        // …and the links arrive with the end of the sheet.
        const { A1, B2 } = ws.hyperlinks!;
        expect(B2.target).toBe("#Sheet1!A1");
        expect(reader.getHyperlinkTarget(ws.sheetNo, A1.rId!)).toBe("https://example.com/");
      }
    });
  });

  // ===========================================================================
  // Omitted `r` attributes (ECMA-376 §18.3.1.73 / §18.3.1.4: both optional)
  // ===========================================================================

  describe("omitted r attributes", () => {
    /**
     * Rows 1, (2), 5, (6) and cells A, (B), D, (E): every omitted address must
     * follow the previous one, not the count seen so far.
     */
    async function buildWithoutSomeRefs(): Promise<Uint8Array> {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        for (const r of [1, 2, 5, 6]) {
          for (const c of ["A", "B", "D", "E"]) {
            Cell.setValue(ws, `${c}${r}`, `${c}${r}`);
          }
        }
        Cell.setStyle(ws, "E6", { font: { bold: true } });
      });
      return patchSheetXml(buffer, xml =>
        xml.replace(/<row r="(2|6)"/g, "<row").replace(/<c r="[BE]\d+"/g, "<c")
      );
    }

    function expectAddresses(rows: { number: number; get: (col: number) => unknown }[]) {
      expect(rows.map(r => r.number)).toEqual([1, 2, 5, 6]);
      for (const row of rows) {
        for (const [col, letter] of [
          [1, "A"],
          [2, "B"],
          [4, "D"],
          [5, "E"]
        ] as const) {
          expect(row.get(col)).toBe(`${letter}${row.number}`);
        }
      }
    }

    it("are derived from the previous row and cell when streaming", async () => {
      const patched = await buildWithoutSomeRefs();
      const xml = new TextDecoder().decode(
        (await extractAll(patched)).get("xl/worksheets/sheet1.xml")!.data
      );
      expect(xml).toContain("<row ");
      expect(xml).not.toContain('r="E6"');

      const sheet = await readFirstSheet(patched);
      expectAddresses(
        sheet.rows.map(row => ({
          number: row.number,
          get: (col: number) => cellGetValue(rowGetCell(row, col))
        }))
      );
      expect(cellFont(rowGetCell(sheet.rows[3], 5))?.bold).toBe(true);
    });

    it("does not file a malformed reference under the next column", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 1);
        Cell.setValue(ws, "B1", 2);
      });
      const patched = await patchSheetXml(buffer, xml =>
        xml.replace('<c r="B1"', '<c r="garbage"')
      );
      await expect(readFirstSheet(patched)).rejects.toThrow(/garbage/);
    });

    it("reads an absolute reference by the column it names", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "a");
        Cell.setValue(ws, "C1", "c");
      });
      const patched = await patchSheetXml(buffer, xml => xml.replace('<c r="C1"', '<c r="$C$1"'));
      const sheet = await readFirstSheet(patched);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 3))).toBe("c");
    });

    it("agree with the full reader", async () => {
      const patched = await buildWithoutSomeRefs();
      const wb = Workbook.create();
      await Workbook.read(wb, patched);
      const ws = Workbook.getWorksheet(wb, "Sheet1")!;
      const rows: { number: number; get: (col: number) => unknown }[] = [];
      Worksheet.eachRow(ws, (_row, number) => {
        rows.push({ number, get: col => Cell.getValue(ws, `${"ABCDE"[col - 1]}${number}`) });
      });
      expectAddresses(rows);
    });
  });

  // ===========================================================================
  // Styles
  // ===========================================================================

  describe("styles", () => {
    it("parses cell styles when styles are cached", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 42);
        Cell.setStyle(ws, "A1", { font: { bold: true, size: 14 } });
        Cell.setStyle(ws, "A1", { numFmt: "0.00%" });
      });
      const sheet = await readFirstSheet(buffer);
      const cell = rowGetCell(sheet.rows[0], 1);

      expect(cellGetValue(cell)).toBe(42);
      expect(cellFont(cell)).toBeDefined();
      expect(cellFont(cell)!.bold).toBe(true);
      expect(cellNumFmt(cell)).toBe("0.00%");
    });

    it("lets a cell's style inherited from its row be changed without touching the row", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Row.setFont(ws, 1, { bold: true });
        Cell.setValue(ws, "A1", 1);
        Cell.setValue(ws, "B1", 2);
      });
      const patched = await patchSheetXml(buffer, xml =>
        xml.replace(/(<c r="[AB]1") s="\d+"/g, "$1")
      );
      const sheet = await readFirstSheet(patched);
      const row = sheet.rows[0];
      const a1 = rowGetCell(row, 1);
      expect(cellFont(a1)?.bold).toBe(true);
      // The documented way to write through a facet: it copies a shared one first.
      cellOwnStyle(a1).font!.italic = true;
      expect(cellFont(a1)?.italic).toBe(true);
      expect(cellFont(rowGetCell(row, 2))?.bold).toBe(true);
      expect(cellFont(rowGetCell(row, 2))?.italic).toBeUndefined();
    });

    it("gives each cell its own mutable style, as the streaming handle always has", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        for (const ref of ["A1", "B1"]) {
          Cell.setValue(ws, ref, 1);
          Cell.setStyle(ws, ref, { font: { bold: true } });
        }
      });
      const sheet = await readFirstSheet(buffer);
      const a1 = rowGetCell(sheet.rows[0], 1);
      const b1 = rowGetCell(sheet.rows[0], 2);
      a1.style.font!.bold = false;
      expect(cellFont(b1)?.bold).toBe(true);
    });
  });

  // ===========================================================================
  // Multi-column + Sparse Data
  // ===========================================================================

  describe("multi-column data", () => {
    it("correctly maps cells to columns across multiple rows", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Worksheet.addRow(ws, [1, "Alice", true]);
        Worksheet.addRow(ws, [2, "Bob", false]);
        Worksheet.addRow(ws, [3, "Charlie", true]);
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(3);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe(1);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 2))).toBe("Alice");
      expect(cellGetValue(rowGetCell(sheet.rows[0], 3))).toBe(true);
      expect(cellGetValue(rowGetCell(sheet.rows[2], 1))).toBe(3);
      expect(cellGetValue(rowGetCell(sheet.rows[2], 2))).toBe("Charlie");
    });

    it("handles sparse rows (cells with gaps)", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 1);
        Cell.setValue(ws, "D1", 4); // B1, C1 are empty
        Cell.setValue(ws, "A2", "start");
        Cell.setValue(ws, "F2", "end"); // B2-E2 are empty
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe(1);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 4))).toBe(4);
      expect(cellGetValue(rowGetCell(sheet.rows[1], 1))).toBe("start");
      expect(cellGetValue(rowGetCell(sheet.rows[1], 6))).toBe("end");
    });
  });

  // ===========================================================================
  // Multiple Worksheets
  // ===========================================================================

  describe("multiple worksheets", () => {
    it("reads all worksheets in order with correct names and data", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "First"), "A1", "sheet1");
        Cell.setValue(Workbook.addWorksheet(wb, "Second"), "A1", "sheet2");
        Cell.setValue(Workbook.addWorksheet(wb, "Third"), "A1", "sheet3");
      });

      const sheets = await readAllSheets(buffer);

      expect(sheets.length).toBe(3);
      expect(sheets[0].name).toBe("First");
      expect(cellGetValue(rowGetCell(sheets[0].rows[0], 1))).toBe("sheet1");
      expect(sheets[1].name).toBe("Second");
      expect(cellGetValue(rowGetCell(sheets[1].rows[0], 1))).toBe("sheet2");
      expect(sheets[2].name).toBe("Third");
      expect(cellGetValue(rowGetCell(sheets[2].rows[0], 1))).toBe("sheet3");
    });
  });

  // ===========================================================================
  // read() Event-Based API
  // ===========================================================================

  describe("read() event API", () => {
    it("emits 'row' events and 'finished' at the end", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "one");
        Cell.setValue(ws, "A2", "two");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const result = await new Promise<{ rows: RowData[]; finished: boolean }>(
        (resolve, reject) => {
          const rows: RowData[] = [];
          let finished = false;

          reader.on("worksheet", worksheet => {
            worksheet.on("row", (row: RowData) => rows.push(row));
            worksheet.on("finished", () => {
              finished = true;
              resolve({ rows, finished });
            });
          });
          reader.on("error", reject);
          reader.read();
        }
      );

      expect(result.finished).toBe(true);
      expect(result.rows.length).toBe(2);
      expect(cellGetValue(rowGetCell(result.rows[0], 1))).toBe("one");
      expect(cellGetValue(rowGetCell(result.rows[1], 1))).toBe("two");
    });
  });

  // ===========================================================================
  // Hyperlinks
  // ===========================================================================

  describe("hyperlinks", () => {
    it("caches hyperlinks when hyperlinks option is 'cache'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", {
          text: "Google",
          hyperlink: "https://www.google.com"
        });
        Cell.setValue(ws, "A2", "no link");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache",
        hyperlinks: "cache",
        entries: "emit"
      });

      for await (const ws of reader) {
        const rows: RowData[] = [];
        for await (const row of ws) {
          rows.push(row as RowData);
        }
        expect(rows.length).toBe(2);

        // hyperlinks should be cached on the WorksheetReader
        const hyperlinks = (ws as any).hyperlinks as
          | Record<string, { ref: string; rId: string }>
          | undefined;
        expect(hyperlinks).toBeDefined();
        expect(hyperlinks!["A1"]).toBeDefined();
        expect(hyperlinks!["A1"].rId).toBeTypeOf("string");
      }
    });
  });

  // ===========================================================================
  // Large Dataset
  // ===========================================================================

  describe("large dataset", () => {
    it("handles 1000 rows without issues", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        for (let i = 1; i <= 1000; i++) {
          Worksheet.addRow(ws, [i, `name-${i}`, i * 1.5]);
        }
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(1000);
      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe(1);
      expect(cellGetValue(rowGetCell(sheet.rows[999], 1))).toBe(1000);
      expect(cellGetValue(rowGetCell(sheet.rows[999], 2))).toBe("name-1000");
      expect(cellGetValue(rowGetCell(sheet.rows[999], 3))).toBe(1500);
    });
  });

  // ===========================================================================
  // Edge Cases
  // ===========================================================================

  describe("edge cases", () => {
    it("handles an empty worksheet", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Workbook.addWorksheet(wb, "Empty");
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(0);
    });

    it("handles XML special characters in string values", async () => {
      const specialStr = "xml chars: & < > \" '";
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", specialStr);
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe(specialStr);
    });

    it("handles OOXML _xHHHH_ escape patterns in strings", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "_x000D_");
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe("_x000D_");
    });

    it("handles mixed value types in a single row", async () => {
      const now = new Date(2024, 5, 15);
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", 42);
        Cell.setValue(ws, "B1", "text");
        Cell.setValue(ws, "C1", true);
        Cell.setValue(ws, "D1", now);
        Cell.setStyle(ws, "D1", { numFmt: "yyyy-mm-dd" });
      });
      const sheet = await readFirstSheet(buffer);
      const row = sheet.rows[0];

      expect(cellGetValue(rowGetCell(row, 1))).toBe(42);
      expect(cellGetValue(rowGetCell(row, 2))).toBe("text");
      expect(cellGetValue(rowGetCell(row, 3))).toBe(true);
      expect(cellGetValue(rowGetCell(row, 4))).toBeInstanceOf(Date);
    });

    it("handles unicode and CJK characters", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "你好世界");
        Cell.setValue(ws, "B1", "日本語テスト");
        Cell.setValue(ws, "C1", "한국어");
        Cell.setValue(ws, "D1", "Ñoño Año");
      });
      const sheet = await readFirstSheet(buffer);

      expect(cellGetValue(rowGetCell(sheet.rows[0], 1))).toBe("你好世界");
      expect(cellGetValue(rowGetCell(sheet.rows[0], 2))).toBe("日本語テスト");
      expect(cellGetValue(rowGetCell(sheet.rows[0], 3))).toBe("한국어");
      expect(cellGetValue(rowGetCell(sheet.rows[0], 4))).toBe("Ñoño Año");
    });
  });
});
