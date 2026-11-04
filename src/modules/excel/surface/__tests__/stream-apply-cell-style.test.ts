/**
 * `Stream.applyCellStyle(c, name)` gives a streamed cell a named style the writer registered, checked as
 * `Cell.applyStyle` checks one: an unknown name throws, the cell records the style's own spelling, and a committed
 * row refuses the edit. Setting `cell.style.styleName` by hand was the only way, and checked nothing.
 */
import { Archive } from "@archive";
import { ExcelArgumentError, ExcelStreamStateError, Stream } from "@excel";
import { describe, expect, it } from "vitest";

async function written(
  build: (wb: Stream.WorkbookWriter) => Promise<void> | void
): Promise<string> {
  const chunks: Uint8Array[] = [];
  const stream = new WritableStream<Uint8Array>({
    write(chunk) {
      chunks.push(chunk);
    }
  });
  const wb = new Stream.WorkbookWriter({ stream, useStyles: true });
  await build(wb);
  await wb.commit();
  const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, at);
    at += chunk.length;
  }
  return new TextDecoder().decode((await Archive.unpackZip(bytes)).get("xl/styles.xml")!.data);
}

describe("Stream.applyCellStyle", () => {
  it("applies a style registered with addCellStyle, by the name in any case", async () => {
    const xml = await written(wb => {
      wb.addCellStyle("Totals", { font: { bold: true } });
      const ws = wb.addWorksheet("S");
      const c = Stream.getRowCell(ws.getRow(1), 1);
      Stream.setCellValue(c, 1);
      Stream.applyCellStyle(c, "TOTALS");
      expect(c.style.styleName).toBe("Totals");
      ws.commit();
    });
    expect(xml).toMatch(/<cellStyle name="Totals" xfId="1"/);
    expect(xml).toMatch(/<xf [^>]*xfId="1"/);
  });

  it("a built-in style added with useBuiltInCellStyle", async () => {
    await written(wb => {
      const name = wb.useBuiltInCellStyle("Good");
      const c = Stream.getRowCell(wb.addWorksheet("S").getRow(1), 1);
      Stream.applyCellStyle(c, name);
      expect(c.style.styleName).toBe(name);
    });
  });

  it("refuses an unknown name, an empty one, and a committed row — changing nothing", async () => {
    await written(wb => {
      const ws = wb.addWorksheet("S");
      const row = ws.getRow(1);
      const c = Stream.getRowCell(row, 1);
      expect(() => Stream.applyCellStyle(c, "Nope")).toThrow(ExcelArgumentError);
      expect(() => Stream.applyCellStyle(c, "Nope")).toThrow(
        expect.objectContaining({
          field: "name",
          message: expect.stringMatching(/^Stream\.applyCellStyle: /)
        })
      );
      expect(() => Stream.applyCellStyle(c, "")).toThrow(ExcelArgumentError);
      expect(c.style.styleName).toBeUndefined();
      wb.addCellStyle("Late", { font: { italic: true } });
      Stream.commitRow(row);
      expect(() => Stream.applyCellStyle(c, "Late")).toThrow(ExcelStreamStateError);
    });
  });
});
