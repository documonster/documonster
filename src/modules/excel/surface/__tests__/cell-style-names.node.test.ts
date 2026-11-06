/**
 * Excel compares cell style names without regard to case: given `Brand Title` and `brand title` it keeps one style and
 * points both cells at it (tmp/office-tests/r4-style). So the library keys named cell styles the same way — a name
 * equal to an existing one ignoring case is that style — in the workbook, both readers, both writers and the
 * streaming writer.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { cellSetStyle, cellSetValue } from "@excel/core/cell";
import { Cell, Stream, Workbook, ExcelWorksheetNameError } from "@excel/index";
import { describe, expect, it } from "vitest";

async function stylesXml(bytes: Uint8Array): Promise<string> {
  return new TextDecoder().decode((await extractAll(bytes)).get("xl/styles.xml")!.data);
}

function cellStyleNames(xml: string): string[] {
  return [...xml.matchAll(/<cellStyle [^>]*name="([^"]*)"/g)].map(m => m[1]);
}

function cellXfIds(xml: string): string[] {
  const cellXfs = xml.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/s)![1];
  return [...cellXfs.matchAll(/<xf [^>]*xfId="(\d+)"/g)].map(m => m[1]);
}

describe("named cell styles compare without regard to case", () => {
  it("adding a name that differs only in case replaces the style and keeps the first spelling", () => {
    const wb = Workbook.create();
    Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
    const entry = Workbook.addCellStyle(wb, "brand title", { font: { italic: true } });
    expect(entry).toEqual({ name: "Brand Title", font: { italic: true } });
    expect(Workbook.listCellStyles(wb)).toEqual([{ name: "Brand Title", font: { italic: true } }]);
    expect(Workbook.getCellStyle(wb, "BRAND TITLE")?.name).toBe("Brand Title");
    expect(Workbook.removeCellStyle(wb, "brand TITLE")).toBe(true);
    expect(Workbook.listCellStyles(wb)).toEqual([]);
  });

  it("refuses the built-in Normal in any case", () => {
    const wb = Workbook.create();
    for (const name of ["Normal", "normal", "NORMAL"]) {
      expect(() => Workbook.addCellStyle(wb, name, {}), name).toThrow(ExcelWorksheetNameError);
    }
  });

  it("useBuiltInCellStyle replaces a custom style of the same name in another case", () => {
    const wb = Workbook.create();
    Workbook.addCellStyle(wb, "heading 1", { font: { italic: true } });
    expect(Workbook.useBuiltInCellStyle(wb, "Heading1")).toBe("heading 1");
    expect(Workbook.listCellStyles(wb)).toHaveLength(1);
    expect(Workbook.listCellStyles(wb)[0].builtinId).toBe(16);
  });

  it("Cell.applyStyle takes the name in any case", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
    Cell.applyStyle(ws, "A1", "brand title");
    expect(Cell.findStyle(ws, "A1")?.styleName).toBe("Brand Title");
  });

  it.each(["xlsx", "xlsb"] as const)(
    "a cell whose styleName differs in case from the style is written pointing at it (%s)",
    async format => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
      Cell.setValue(ws, "A1", "a");
      Cell.setValue(ws, "A2", "b");
      Cell.applyStyle(ws, "A1", "Brand Title");
      Cell.setStyle(ws, "A2", { styleName: "BRAND TITLE" });
      const bytes = await Workbook.toBuffer(wb, { format });
      if (format === "xlsx") {
        const xml = await stylesXml(bytes);
        expect(cellStyleNames(xml)).toEqual(["Normal", "Brand Title"]);
        const ids = cellXfIds(xml);
        expect(new Set(ids.slice(1))).toEqual(new Set(["1"]));
      }
      const back = await Workbook.read(bytes);
      expect(Workbook.listCellStyles(back).map(s => s.name)).toEqual(["Brand Title"]);
      // XLSB does not record which named style a cell uses (`xlsb/write/types.ts`), in either case.
      if (format === "xlsx") {
        const sheet = Workbook.getWorksheet(back, "S")!;
        expect(Cell.findStyle(sheet, "A1")?.styleName).toBe("Brand Title");
        expect(Cell.findStyle(sheet, "A2")?.styleName).toBe("Brand Title");
      }
    }
  );

  it.each(["xlsx", "xlsb"] as const)(
    "re-adding a name in another case writes one style with the later definition (%s)",
    async format => {
      const wb = Workbook.create();
      Workbook.addWorksheet(wb, "S");
      Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
      Workbook.addCellStyle(wb, "brand title", { font: { italic: true } });
      const back = await Workbook.read(await Workbook.toBuffer(wb, { format }));
      expect(Workbook.listCellStyles(back)).toEqual([
        expect.objectContaining({
          name: "Brand Title",
          font: expect.objectContaining({ italic: true })
        })
      ]);
      expect(Workbook.getCellStyle(back, "Brand Title")?.font?.bold).toBeFalsy();
    }
  );

  it("an XLSB file carrying two styles that differ only in case is read with the first", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
    // Same length as "brand title", so the record can be patched in place.
    Workbook.addCellStyle(wb, "Zzzzzzzzzzz", { font: { italic: true } });
    const entries = await extractAll(await Workbook.toBuffer(wb, { format: "xlsb" }));
    const utf16 = (text: string) =>
      new Uint8Array(new Uint16Array([...text].map(c => c.charCodeAt(0))).buffer);
    const from = utf16("Zzzzzzzzzzz");
    const to = utf16("brand title");
    const bytes = await createZip(
      [...entries]
        .filter(([, file]) => file.type !== "directory")
        .map(([name, file]) => {
          if (name !== "xl/styles.bin") {
            return { name, data: file.data };
          }
          const data = new Uint8Array(file.data);
          const at = Buffer.from(data).indexOf(Buffer.from(from));
          expect(at).toBeGreaterThan(0);
          data.set(to, at);
          return { name, data };
        })
    );
    const back = await Workbook.read(bytes);
    expect(Workbook.listCellStyles(back)).toEqual([
      expect.objectContaining({
        name: "Brand Title",
        font: expect.objectContaining({ bold: true })
      })
    ]);
  });

  it("a file carrying two styles that differ only in case is read as Excel reads it: the first wins", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Workbook.addCellStyle(wb, "Brand Title", { font: { bold: true } });
    Workbook.addCellStyle(wb, "Zzz", { font: { italic: true } });
    Cell.setValue(ws, "A1", "a");
    Cell.setValue(ws, "A2", "b");
    Cell.applyStyle(ws, "A1", "Brand Title");
    Cell.applyStyle(ws, "A2", "Zzz");
    const entries = await extractAll(await Workbook.toBuffer(wb));
    const bytes = await createZip(
      [...entries]
        .filter(([, file]) => file.type !== "directory")
        .map(([name, file]) => ({
          name,
          data:
            name === "xl/styles.xml"
              ? new TextEncoder().encode(
                  new TextDecoder().decode(file.data).replace('name="Zzz"', 'name="brand title"')
                )
              : file.data
        }))
    );
    const back = await Workbook.read(bytes);
    expect(Workbook.listCellStyles(back)).toEqual([
      expect.objectContaining({
        name: "Brand Title",
        font: expect.objectContaining({ bold: true })
      })
    ]);
    const sheet = Workbook.getWorksheet(back, "S")!;
    expect(Cell.findStyle(sheet, "A2")?.styleName).toBe("Brand Title");
    const xml = await stylesXml(await Workbook.toBuffer(back));
    expect(cellStyleNames(xml)).toEqual(["Normal", "Brand Title"]);
  });

  it("the streaming writer keeps one style per name ignoring case", async () => {
    const chunks: Uint8Array[] = [];
    const { Writable } = await import("node:stream");
    const sink = new Writable({
      write(chunk, _enc, cb) {
        chunks.push(chunk);
        cb();
      }
    });
    const writer = new Stream.WorkbookWriter({ stream: sink, useStyles: true });
    writer.addCellStyle("Brand Title", { font: { bold: true } });
    expect(writer.addCellStyle("brand title", { font: { italic: true } }).name).toBe("Brand Title");
    expect(() => writer.addCellStyle("normal", {})).toThrow(ExcelWorksheetNameError);
    const sheet = writer.addWorksheet("S");
    const cell = sheet.getCell("A1");
    cellSetValue(cell, "x");
    cellSetStyle(cell, { styleName: "BRAND TITLE" });
    sheet.commit();
    await writer.commit();
    const xml = await stylesXml(Buffer.concat(chunks));
    expect(cellStyleNames(xml)).toEqual(["Normal", "Brand Title"]);
    // Redefining allocates a fresh style xf; the cell points at whichever the one `<cellStyle>` names.
    const styleXf = xml.match(
      /<cellStyle [^>]*name="Brand Title"[^>]*xfId="(\d+)"|<cellStyle [^>]*xfId="(\d+)"[^>]*name="Brand Title"/
    )!;
    expect(new Set(cellXfIds(xml).slice(1))).toEqual(new Set([styleXf[1] ?? styleXf[2]]));
  });
});
