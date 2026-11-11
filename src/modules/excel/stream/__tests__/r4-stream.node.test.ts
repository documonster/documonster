/**
 * Round-4 review fixes in the streaming API: column settings refuse once the columns are written, the row reader
 * indexes as `Row.getValues` does, the handle setters cover what the document model's do, a refused row adds nothing,
 * and a damaged package is reported as every reader reports it.
 */
import { Readable, Writable } from "node:stream";

import { Archive, ZipCrc32MismatchError } from "@archive";
import {
  ExcelArgumentError,
  ExcelFileError,
  ExcelStreamStateError,
  ExcelInvalidValueTypeError,
  Stream,
  Workbook,
  Worksheet,
  Column,
  Row,
  isExcelError
} from "@excel";
import { describe, expect, it } from "vitest";

function writer(options: { useStyles?: boolean } = {}) {
  const chunks: Uint8Array[] = [];
  const stream = new Writable({
    write(chunk, _encoding, callback) {
      chunks.push(chunk);
      callback();
    }
  });
  const wb = new Stream.WorkbookWriter({ stream, ...options });
  return { wb, ws: wb.addWorksheet("S"), bytes: () => Buffer.concat(chunks) };
}

describe("R4-E5 a streamed sheet's columns refuse once written", () => {
  it("Stream.setColumnWidth before the first committed row is written; after, it refuses", async () => {
    const { wb, ws, bytes } = writer();
    Stream.setColumnWidth(ws.getColumn(2), 30);
    const row = ws.addRow(["a", "b"]);
    Stream.commitRow(row);
    expect(() => Stream.setColumnWidth(ws.getColumn(1), 40)).toThrow(ExcelStreamStateError);
    expect(() => Stream.setColumnHidden(ws.getColumn(1), true)).toThrow(ExcelStreamStateError);
    expect(() => Stream.setColumnOutlineLevel(ws.getColumn(1), 1)).toThrow(ExcelStreamStateError);
    ws.commit();
    await wb.commit();
    const read = Workbook.listWorksheets(await Workbook.read(bytes()))[0]!;
    expect(Column.getWidth(read, 2)).toBe(30);
    expect(Column.getWidth(read, 1)).toBeUndefined();
  });

  it("Stream.setColumnWidth(undefined) clears a width, as Column.setWidth does", async () => {
    const { wb, ws, bytes } = writer();
    Stream.setColumnWidth(ws.getColumn(1), 30);
    Stream.setColumnWidth(ws.getColumn(1), undefined);
    Stream.setColumnHidden(ws.getColumn(2), true);
    Stream.setColumnOutlineLevel(ws.getColumn(3), 2);
    expect(() => Stream.setColumnOutlineLevel(ws.getColumn(3), 9)).toThrow(ExcelArgumentError);
    Stream.commitRow(ws.addRow(["a", "b", "c"]));
    ws.commit();
    await wb.commit();
    const read = Workbook.listWorksheets(await Workbook.read(bytes()))[0]!;
    expect(Column.getWidth(read, 1)).toBeUndefined();
    expect(Column.isHidden(read, 2)).toBe(true);
    expect(Column.getOutlineLevel(read, 3)).toBe(2);
  });
});

describe("R4-S1 Stream.getRowValues indexes as Row.getValues does", () => {
  it("is 0-based by default, column-numbered with { base: 1 }", () => {
    const { ws } = writer();
    const row = ws.addRow(["Ada", 42]);
    expect(Stream.getRowValues(row)).toEqual(["Ada", 42]);
    expect(Stream.getRowValues(row, { base: 1 })).toEqual([undefined, "Ada", 42]);
    expect(() => Stream.getRowValues(row, { base: 2 } as never)).toThrow(
      expect.objectContaining({ field: "options.base" })
    );
    // The deprecated name keeps its 1-based reading.
    expect(Stream.rowValues(row)).toEqual([undefined, "Ada", 42]);
  });
});

describe("R4-G1 the streamed handle setters cover the document model's", () => {
  it("number format, protection and outline level, checked and refused once committed", async () => {
    const { wb, ws, bytes } = writer({ useStyles: true });
    const row = ws.addRow([1, 2]);
    Stream.setRowNumFmt(row, "0.00");
    Stream.setRowProtection(row, { locked: false });
    Stream.setRowOutlineLevel(row, 1);
    Stream.setCellProtection(Stream.getRowCell(row, 2), { hidden: true });
    expect(() => Stream.setRowOutlineLevel(row, 8)).toThrow(ExcelArgumentError);
    expect(() => Stream.setRowNumFmt(row, 42 as never)).toThrow(ExcelArgumentError);
    Stream.commitRow(row);
    expect(() => Stream.setRowNumFmt(row, "0")).toThrow(ExcelStreamStateError);
    expect(() => Stream.setCellProtection(Stream.getRowCell(row, 1), {})).toThrow(
      ExcelStreamStateError
    );
    ws.commit();
    await wb.commit();
    const read = Workbook.listWorksheets(await Workbook.read(bytes()))[0]!;
    expect(Row.getOutlineLevel(read, 1)).toBe(1);
    expect(Row.getNumFmt(read, 1)).toBe("0.00");
    expect(Row.getProtection(read, 1)).toMatchObject({ locked: false });
  });
});

describe("WorksheetWriter#addRows adds nothing when a value is refused", () => {
  it("checks every row first, naming the row and the value", () => {
    const { ws } = writer();
    let error: unknown;
    try {
      ws.addRows([[1], [2, Number.NaN]]);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelInvalidValueTypeError);
    expect((error as ExcelInvalidValueTypeError).field).toBe("values[1][1]");
    expect(ws.findRow(1)).toBeUndefined();
    expect(() => ws.addRow([Number.NaN])).toThrow(expect.objectContaining({ field: "value[0]" }));
    expect(ws.findRow(1)).toBeUndefined();
  });
});

describe("WorksheetWriter#addConditionalFormatting stores its ref canonical", () => {
  it("as ConditionalFormatting.add does", () => {
    const { ws } = writer();
    ws.addConditionalFormatting({
      ref: "a1:b2",
      rules: [{ type: "expression", formulae: ["TRUE"], style: {} }]
    });
    expect(
      (ws as unknown as { conditionalFormatting: { ref: string }[] }).conditionalFormatting[0]!.ref
    ).toBe("A1:B2");
  });
});

/**
 * The package with one byte of `partName`'s stored data flipped, so its CRC-32 no longer matches. The
 * byte is the `0` of `version="1.0"`, which becomes `1.1`: the part still parses, so it is the CRC-32
 * check — not the XML parser — that has to find the damage.
 */
async function damaged(bytes: Uint8Array, partName: string): Promise<Uint8Array> {
  const parts = await Archive.unpackZip(bytes);
  const out = await Archive.packZip(new Map([...parts].map(([k, v]) => [k, v.data])), { level: 0 });
  const name = new TextEncoder().encode(partName);
  for (let i = 0; i < out.length - 30; i++) {
    if (out[i] === 0x50 && out[i + 1] === 0x4b && out[i + 2] === 3 && out[i + 3] === 4) {
      const n = out[i + 26]! | (out[i + 27]! << 8);
      const extra = out[i + 28]! | (out[i + 29]! << 8);
      if (
        new TextDecoder().decode(out.subarray(i + 30, i + 30 + n)) ===
        new TextDecoder().decode(name)
      ) {
        out[i + 30 + n + extra + 17]! ^= 0x01;
        return out;
      }
    }
  }
  throw new Error(`no part ${partName}`);
}

describe("X3 a damaged package is an ExcelFileError in every Excel reader", () => {
  async function sample(): Promise<Uint8Array> {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.addMany(ws, [
      ["a", "b"],
      [1, 2]
    ]);
    return Workbook.toBuffer(wb);
  }

  it.each(["xl/workbook.xml", "xl/worksheets/sheet1.xml"])(
    "Workbook.read, Workbook.readStream and Stream.WorkbookReader on a damaged %s",
    async part => {
      const bad = await damaged(await sample(), part);
      const errors: unknown[] = [];
      await Workbook.read(bad).catch(e => errors.push(e));
      await Workbook.readStream(Readable.from([bad])).catch(e => errors.push(e));
      try {
        for await (const sheet of new Stream.WorkbookReader(Readable.from([bad]))) {
          for await (const _row of sheet) {
            // drain
          }
        }
      } catch (e) {
        errors.push(e);
      }
      expect(errors).toHaveLength(3);
      for (const error of errors) {
        expect(error).toBeInstanceOf(ExcelFileError);
        expect(isExcelError(error)).toBe(true);
        expect((error as Error).cause).toBeInstanceOf(ZipCrc32MismatchError);
      }
    }
  );

  it("XLSB too", async () => {
    const wb = Workbook.create();
    Row.add(Workbook.addWorksheet(wb, "S"), [1]);
    const bad = await damaged(await Workbook.toBuffer(wb, { format: "xlsb" }), "xl/workbook.bin");
    await expect(Workbook.read(bad)).rejects.toBeInstanceOf(ExcelFileError);
  });
});

describe("R4-C4 Workbook.readStream reads XLSB", () => {
  it("with format: 'xlsb', as its error says", async () => {
    const wb = Workbook.create();
    Row.add(Workbook.addWorksheet(wb, "S"), [1, 2]);
    const xlsb = await Workbook.toBuffer(wb, { format: "xlsb" });
    const read = await Workbook.readStream(Readable.from([xlsb]), { format: "xlsb" });
    expect(Row.getValues(Workbook.getWorksheet(read, "S")!, 1)).toEqual([1, 2]);
    await expect(Workbook.readStream(Readable.from([xlsb]))).rejects.toThrow(
      /Workbook\.readStream cannot/
    );
    expect(Worksheet.getName(Workbook.listWorksheets(read)[0]!)).toBe("S");
  });
});
