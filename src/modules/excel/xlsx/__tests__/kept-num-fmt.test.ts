import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { cellGetValue, type CellData } from "@excel/core/cell";
import { Cell, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { shownCodeOfKeptId } from "@excel/utils/display-locale";
import { keptNumFmtId } from "@excel/utils/kept-num-fmt";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

/**
 * A built-in number format with no one format code — 5–8 are currencies, 41–44 accounting, 27–36 and 50–81 East Asian
 * and Thai formats, each shown in the reader's locale — used to be dropped on read: the cell read as General and was
 * written back as General. It is kept by its id instead, written back while `numFmt` is absent, and reported (see `kept-num-fmt.ts`).
 */

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** cellXfs: 0 General, 1 currency (7), 2 East Asian date (30), 3 a custom id with no `<numFmt>`, 4 `0.00`. */
const STYLES = `<?xml version="1.0"?><styleSheet xmlns="${MAIN}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="2"><xf numFmtId="0"/><xf numFmtId="7"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0"/><xf numFmtId="7"/><xf numFmtId="30"/><xf numFmtId="170"/><xf numFmtId="2"/></cellXfs><cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Money" xfId="1"/></cellStyles></styleSheet>`;

async function workbookWith(sheet: string, styles = STYLES): Promise<Uint8Array> {
  const archive = zip();
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="${MAIN}">${sheet}</worksheet>`
  };
  for (const [name, text] of Object.entries(files)) {
    archive.add(name, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

const SHEET =
  '<cols><col min="3" max="3" width="9" style="2"/></cols><sheetData>' +
  '<row r="1"><c r="A1" s="1"><v>12.5</v></c><c r="B1" s="3"><v>1</v></c><c r="D1" s="4"><v>2</v></c></row>' +
  '<row r="2" s="1" customFormat="1"><c r="A2"><v>1</v></c></row>' +
  "</sheetData>";

/** The `numFmtId` each written cell's `s` points at, by address, and the written cellXfs. */
async function writtenFormats(bytes: Uint8Array): Promise<Record<string, number>> {
  const parts = await extractAll(bytes);
  const text = (name: string) => new TextDecoder().decode(parts.get(name)!.data);
  const xfs = [...text("xl/styles.xml").matchAll(/<cellXfs[^>]*>(.*?)<\/cellXfs>/gs)][0]![1]!;
  const ids = [...xfs.matchAll(/<xf numFmtId="(\d+)"/g)].map(m => Number(m[1]));
  const formats: Record<string, number> = {};
  for (const [, r, s] of text("xl/worksheets/sheet1.xml").matchAll(
    /<c r="(\w+)"(?: s="(\d+)")?/g
  )) {
    formats[r!] = ids[Number(s ?? 0)]!;
  }
  return formats;
}

/** The `numFmtId` of each written named cell style's base xf, by name. */
async function namedStyleFormats(bytes: Uint8Array): Promise<Record<string, number>> {
  const styles = new TextDecoder().decode((await extractAll(bytes)).get("xl/styles.xml")!.data);
  const xfs = /<cellStyleXfs[^>]*>(.*?)<\/cellStyleXfs>/s.exec(styles)![1]!;
  const ids = [...xfs.matchAll(/<xf numFmtId="(\d+)"/g)].map(m => Number(m[1]));
  const formats: Record<string, number> = {};
  for (const [, name, xfId] of styles.matchAll(/<cellStyle name="([^"]+)" xfId="(\d+)"/g)) {
    if (name !== "Normal") {
      formats[name!] = ids[Number(xfId)]!;
    }
  }
  return formats;
}

describe("a built-in number format with no code", () => {
  it("reads as no numFmt, is reported where it is, and is written back by its id", async () => {
    const bytes = await workbookWith(SHEET);
    const { workbook, unrecognizedValues } = await Workbook.readWithDiagnostics(bytes);
    const ws = Workbook.getWorksheet(workbook, "S")!;
    expect(Cell.getNumFmt(ws, "A1")).toBeUndefined();
    expect(unrecognizedValues).toEqual([
      { path: 'cellStyles["Money"].numFmt', value: "7" },
      { sheet: "S", path: "C:C.style.numFmt", value: "30" },
      { sheet: "S", path: "A1.style.numFmt", value: "7" },
      { sheet: "S", path: "2:2.style.numFmt", value: "7" }
    ]);
    expect(Workbook.listUnrecognizedValues(workbook)).toEqual(unrecognizedValues);

    const formats = await writtenFormats(await Workbook.toBuffer(workbook));
    expect(formats).toMatchObject({ A1: 7, D1: 2, A2: 0 });
  });

  // Excel shows a kept id in the form of its locale; display takes the English one, as every other built-in code here
  // is (`mm-dd-yy` for 14). The id is still what the cell holds and what is written.
  it("is shown as Excel in an English locale shows it, and is still written by its id", async () => {
    const bytes = await workbookWith(SHEET);
    for (const format of ["xlsx", "xlsb"] as const) {
      const source = await Workbook.read(bytes);
      const wb =
        format === "xlsx"
          ? source
          : await Workbook.read(await Workbook.toBuffer(source, { format }), { format });
      const ws = Workbook.getWorksheet(wb, "S")!;
      Cell.setValue(ws, "E2", -1234.5); // takes row 2's format, the currency
      Cell.setValue(ws, "C3", 45000);
      expect(Cell.getDisplayText(ws, "A1"), format).toBe("$12.50 ");
      expect(Cell.getDisplayText(ws, "E2"), format).toBe("($1,234.50)");
      expect(Cell.getDisplayText(ws, "C3"), format).toBe("03-15-23");
      expect(Cell.getNumFmt(ws, "A1"), format).toBeUndefined();
    }
    const wb = await Workbook.read(bytes);
    Cell.getDisplayText(Workbook.getWorksheet(wb, "S")!, "A1");
    expect((await writtenFormats(await Workbook.toBuffer(wb))).A1).toBe(7);
  });

  // What Excel wrote back for each id, saving a workbook of every kept one (an English locale).
  it("is shown in the form Excel writes it back as", () => {
    expect(shownCodeOfKeptId(7)).toBe('"$"#,##0.00_);\\("$"#,##0.00\\)');
    expect(shownCodeOfKeptId(65)).toBe(shownCodeOfKeptId(7));
    expect(shownCodeOfKeptId(23)).toBe("General");
    expect(shownCodeOfKeptId(30)).toBe("mm-dd-yy");
    expect(shownCodeOfKeptId(33)).toBe("h:mm:ss");
    expect(shownCodeOfKeptId(62)).toBe("#,##0.00");
    expect(shownCodeOfKeptId(78)).toBe("m/d/yy h:mm");
    expect(shownCodeOfKeptId(42)).toBe('_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)');
    expect(shownCodeOfKeptId(81)).toBe("mm:ss.0");
    expect(shownCodeOfKeptId(82)).toBeUndefined();
  });

  // An accounting id is the locale's, as a currency is: a file that names it bare — with no `<numFmt>` saying which
  // locale's form it means — keeps the id. It used to read as the Chinese code, whatever the file's locale.
  it("is what an accounting format named only by its id reads as", async () => {
    // cellXfs 5 and 6: the accounting format 42, and 47.
    const styles = STYLES.replace('<cellXfs count="5">', '<cellXfs count="7">').replace(
      "</cellXfs>",
      '<xf numFmtId="42"/><xf numFmtId="47"/></cellXfs>'
    );
    const bytes = await workbookWith(
      '<sheetData><row r="1"><c r="A1" s="5"><v>-1234.5</v></c><c r="B1" s="6"><v>45000.75</v></c></row></sheetData>',
      styles
    );
    const { workbook, unrecognizedValues } = await Workbook.readWithDiagnostics(bytes);
    const ws = Workbook.getWorksheet(workbook, "S")!;
    expect(Cell.getNumFmt(ws, "A1")).toBeUndefined();
    expect(unrecognizedValues).toContainEqual({ sheet: "S", path: "A1.style.numFmt", value: "42" });
    expect(Cell.getDisplayText(ws, "A1")).toBe(" $(1,235)");
    expect(Cell.getNumFmt(ws, "B1")).toBe("mm:ss.0");
    expect(Cell.getDisplayText(ws, "B1")).toBe("00:00.0");
    expect((await writtenFormats(await Workbook.toBuffer(workbook))).A1).toBe(42);
  });

  it("is the format of a cell created later in its column or row", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "C5", 3);
    Cell.setValue(ws, "D2", 4);
    expect(await writtenFormats(await Workbook.toBuffer(wb))).toMatchObject({ C5: 30, D2: 7 });
  });

  it("gives way to a numFmt that is set", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setNumFmt(ws, "A1", "0.00");
    expect(Workbook.listUnrecognizedValues(wb).map(v => v.path)).not.toContain("A1.style.numFmt");
    expect((await writtenFormats(await Workbook.toBuffer(wb))).A1).toBe(2);
  });

  it("does not share an xf with General", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", 1);
    expect(await writtenFormats(await Workbook.toBuffer(wb))).toMatchObject({ A1: 7, E1: 0 });
  });

  it("is dropped by discardUnrecognizedValues, as the filter chooses", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    expect(
      Workbook.discardUnrecognizedValues(wb, { filter: v => v.path === "A1.style.numFmt" })
    ).toBe(1);
    expect(Workbook.listUnrecognizedValues(wb).map(v => v.path)).toEqual([
      'cellStyles["Money"].numFmt',
      "C:C.style.numFmt",
      "2:2.style.numFmt"
    ]);
    expect((await writtenFormats(await Workbook.toBuffer(wb))).A1).toBe(0);
  });

  it("is reported for a range of styled blanks, and written back", async () => {
    const bytes = await workbookWith(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row>' +
        '<row r="5"><c r="A5" s="1"/><c r="B5" s="1"/></row><row r="6"><c r="A6" s="1"/><c r="B6" s="1"/></row>' +
        "</sheetData>"
    );
    const { workbook, unrecognizedValues } = await Workbook.readWithDiagnostics(bytes, {
      blankCells: "collapse"
    });
    expect(unrecognizedValues).toContainEqual({
      sheet: "S",
      path: "A5:B6.style.numFmt",
      value: "7"
    });
    expect(await writtenFormats(await Workbook.toBuffer(workbook))).toMatchObject({
      A5: 7,
      B6: 7
    });
  });

  it("is kept on the style the streaming reader yields", async () => {
    const reader = new WorkbookReader(await workbookWith(SHEET), {
      worksheets: "emit",
      styles: "cache"
    });
    const kept: Record<string, number | undefined> = {};
    for await (const ws of reader) {
      for await (const row of ws) {
        for (const cell of (row as { cells: ({ address: string; style: object } | undefined)[] })
          .cells) {
          if (cell) {
            kept[cell.address] = keptNumFmtId(cell.style);
          }
        }
      }
    }
    expect(kept).toMatchObject({ A1: 7, D1: undefined });
  });

  it("survives a round trip through XLSB", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    const xlsb = await Workbook.toBuffer(wb, { format: "xlsb" });
    const { workbook, unrecognizedValues } = await Workbook.readWithDiagnostics(xlsb, {
      format: "xlsb"
    });
    expect(unrecognizedValues).toEqual(
      expect.arrayContaining([
        { path: 'cellStyles["Money"].numFmt', value: "7" },
        { sheet: "S", path: "C:C.style.numFmt", value: "30" },
        { sheet: "S", path: "A1.style.numFmt", value: "7" },
        { sheet: "S", path: "2:2.style.numFmt", value: "7" }
      ])
    );
    const written = await Workbook.toBuffer(workbook);
    expect(await writtenFormats(written)).toMatchObject({ A1: 7, D1: 2 });
    expect(await namedStyleFormats(written)).toEqual({ Money: 7 });
  });

  it("is written back for a named cell style, and dropped from it by discardUnrecognizedValues", async () => {
    const wb = await Workbook.read(await workbookWith(SHEET));
    expect(await namedStyleFormats(await Workbook.toBuffer(wb))).toEqual({ Money: 7 });
    expect(
      Workbook.discardUnrecognizedValues(wb, { filter: v => v.path.startsWith("cellStyles") })
    ).toBe(1);
    expect(await namedStyleFormats(await Workbook.toBuffer(wb))).toEqual({ Money: 0 });
  });

  it("makes a number a date when it shows a date in every locale, in both containers", async () => {
    const bytes = await workbookWith(
      '<sheetData><row r="1"><c r="A1" s="2"><v>45000</v></c><c r="B1" s="1"><v>45000</v></c>' +
        '<c r="C1" s="2"><f>45000</f><v>45000</v></c></row></sheetData>'
    );
    const date = new Date(Date.UTC(2023, 2, 15));
    for (const ws of [
      Workbook.getWorksheet(await Workbook.read(bytes), "S")!,
      Workbook.getWorksheet(
        await Workbook.read(
          await Workbook.toBuffer(await Workbook.read(bytes), { format: "xlsb" }),
          { format: "xlsb" }
        ),
        "S"
      )!
    ]) {
      expect(Cell.getValue(ws, "A1")).toEqual(date);
      expect(Cell.getValue(ws, "B1")).toBe(45000);
      expect(Cell.getValue(ws, "C1")).toEqual({ formula: "45000", result: date });
    }
    const written = await Workbook.toBuffer(await Workbook.read(bytes));
    expect(await writtenFormats(written)).toMatchObject({ A1: 30, C1: 30 });

    // The streaming reader of either container reads it so too.
    for (const format of ["xlsx", "xlsb"] as const) {
      const input =
        format === "xlsx"
          ? bytes
          : await Workbook.toBuffer(await Workbook.read(bytes), { format: "xlsb" });
      const values: Record<string, unknown> = {};
      for await (const ws of new WorkbookReader(input, { worksheets: "emit", styles: "cache" })) {
        for await (const row of ws) {
          for (const cell of (row as { cells: (CellData | undefined)[] }).cells) {
            if (cell) {
              values[cell.address] = cellGetValue(cell);
            }
          }
        }
      }
      expect(values, format).toMatchObject({ A1: date, B1: 45000 });
    }
  });

  // The Thai formats 71–81 are dates and times in the one locale that defines them (`ว/ด/ปปปป`, `ช:นน`), and Excel
  // shows them as dates and times in every other.
  it("makes a number a date under a Thai date or time id", async () => {
    const styles = STYLES.replace('<cellXfs count="5">', '<cellXfs count="7">').replace(
      "</cellXfs>",
      '<xf numFmtId="71"/><xf numFmtId="75"/></cellXfs>'
    );
    const ws = Workbook.getWorksheet(
      await Workbook.read(
        await workbookWith(
          '<sheetData><row r="1"><c r="A1" s="5"><v>45000</v></c><c r="B1" s="6"><v>0.75</v></c></row></sheetData>',
          styles
        )
      ),
      "S"
    )!;
    expect(Cell.getValue(ws, "A1")).toEqual(new Date(Date.UTC(2023, 2, 15)));
    expect(Cell.getValue(ws, "B1")).toBeInstanceOf(Date);
    expect(Cell.getDisplayText(ws, "A1", { locale: "th-TH" })).toBe("๑๕/๓/๒๕๖๖");
  });
});

describe("a custom number format id with no <numFmt>", () => {
  // It names no format in the file, and a writer numbers its custom formats afresh: written back, it would name
  // whichever format took that number. So it reads as General, is reported as lost, and is not written back.
  it("reads as General, is reported as lost, and is not kept", async () => {
    const { workbook, unrecognizedValues, lost } = await Workbook.readWithDiagnostics(
      await workbookWith(SHEET)
    );
    const ws = Workbook.getWorksheet(workbook, "S")!;
    expect(Cell.getNumFmt(ws, "B1")).toBeUndefined();
    expect(unrecognizedValues.map(v => v.path)).not.toContain("B1.style.numFmt");
    expect(lost).toEqual([
      "xl/styles.xml: cellXfs[3] names number format 170, which the styles part does not define: read as General"
    ]);
  });

  it.each(["xlsx", "xlsb"] as const)(
    "does not come to name a format the writer numbers as it, through %s",
    async format => {
      // `<numFmt numFmtId="164">` is gone: cellXfs[3] names 164, and the writer gives a new format that number.
      const wb = await Workbook.read(
        await workbookWith(
          '<sheetData><row r="1"><c r="A1" s="3"><v>1.25</v></c><c r="B1"><v>2</v></c></row></sheetData>',
          STYLES.replace('numFmtId="170"', 'numFmtId="164"')
        )
      );
      Cell.setNumFmt(Workbook.getWorksheet(wb, "S")!, "B1", "0.000");
      const back = Workbook.getWorksheet(
        await Workbook.read(await Workbook.toBuffer(wb, { format }), { format }),
        "S"
      )!;
      expect(Cell.getNumFmt(back, "B1")).toBe("0.000");
      expect(Cell.getNumFmt(back, "A1")).not.toBe("0.000");
      expect(Cell.getDisplayText(back, "A1")).toBe("1.25");
    }
  );

  it("is reported as lost by the XLSB reader too", async () => {
    const wb = Workbook.create();
    Cell.setNumFmt(Workbook.addWorksheet(wb, "S"), "A1", "0.000");
    Cell.setValue(Workbook.getWorksheet(wb, "S")!, "A1", 1.25);
    const parts = await extractAll(await Workbook.toBuffer(wb, { format: "xlsb" }));
    // Renumber the part's one `BrtFmt` from 164 to 165, leaving the cell format that names 164 dangling.
    const styles = parts.get("xl/styles.bin")!.data;
    let renumbered = false;
    for (const record of iterateInterpretableRecords(styles, "xl/styles.bin")) {
      if (recordSpec(record.id)?.name === "BrtFmt") {
        expect(record.payload[0] + 256 * record.payload[1]).toBe(164);
        record.payload[0] = 165;
        renumbered = true;
      }
    }
    expect(renumbered).toBe(true);
    const archive = zip();
    for (const [name, file] of parts) {
      if (file.type !== "directory") {
        archive.add(name, file.data);
      }
    }
    const report = await Workbook.readWithDiagnostics(await archive.toBuffer(), { format: "xlsb" });
    expect(report.lost).toContainEqual(
      expect.stringMatching(/^xl\/styles\.bin: cellXfs\[\d+\] names number format 164, which/)
    );
    expect(Cell.getNumFmt(Workbook.getWorksheet(report.workbook, "S")!, "A1")).toBeUndefined();
  });
});
