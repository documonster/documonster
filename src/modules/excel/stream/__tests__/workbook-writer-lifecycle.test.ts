/**
 * Stream.WorkbookWriter — the `commit` / `toBuffer` lifecycle.
 */

import { ExcelStreamStateError } from "@excel/errors";
import { Cell, Stream, Workbook } from "@excel/index";
import { Writable } from "@stream";
import { describe, it, expect } from "vitest";

function collectingStream(): { stream: InstanceType<typeof Writable>; bytes: () => Uint8Array } {
  const chunks: Uint8Array[] = [];
  const stream = new Writable({
    write(chunk: Uint8Array, _encoding: string, callback: () => void) {
      chunks.push(chunk);
      callback();
    }
  });
  return {
    stream,
    bytes: () => {
      const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
      let at = 0;
      for (const c of chunks) {
        out.set(c, at);
        at += c.length;
      }
      return out;
    }
  };
}

function writeRows(wb: InstanceType<typeof Stream.WorkbookWriter>): void {
  const ws = wb.addWorksheet("Data");
  for (let i = 1; i <= 50; i++) {
    Stream.commitRow(ws.addRow([i, `row ${i}`]));
  }
}

async function readBack(bytes: Uint8Array) {
  const wb = await Workbook.read(bytes);
  return Workbook.getWorksheet(wb, "Data")!;
}

describe("Stream.WorkbookWriter — commit / toBuffer", () => {
  it.each(["xlsx", "xlsb"] as const)("toBuffer returns a readable %s file", async format => {
    const wb = new Stream.WorkbookWriter({ format });
    writeRows(wb);
    const bytes = await wb.toBuffer();
    const ws = await readBack(bytes);
    expect(Cell.getValue(ws, "A50")).toBe(50);
    expect(Cell.getValue(ws, "B1")).toBe("row 1");
  });

  it("toBuffer after commit returns the same bytes, and may be repeated", async () => {
    const wb = new Stream.WorkbookWriter();
    writeRows(wb);
    await wb.commit();
    const first = await wb.toBuffer();
    expect(await wb.toBuffer()).toEqual(first);
    expect(Cell.getValue(await readBack(first), "A2")).toBe(2);
  });

  it("toBuffer waits for a commit already running", async () => {
    const wb = new Stream.WorkbookWriter();
    writeRows(wb);
    const committing = wb.commit();
    const bytes = await wb.toBuffer();
    await committing;
    expect(Cell.getValue(await readBack(bytes), "A1")).toBe(1);
  });

  it("refuses a second commit", async () => {
    const wb = new Stream.WorkbookWriter();
    writeRows(wb);
    await wb.commit();
    await expect(wb.commit()).rejects.toThrow(ExcelStreamStateError);
  });

  it("refuses toBuffer when given a stream, which commit ends", async () => {
    const { stream, bytes } = collectingStream();
    const wb = new Stream.WorkbookWriter({ stream });
    writeRows(wb);
    await expect(wb.toBuffer()).rejects.toBeInstanceOf(ExcelStreamStateError);
    await wb.commit();
    expect(Cell.getValue(await readBack(bytes()), "A3")).toBe(3);
    await expect(wb.toBuffer()).rejects.toThrow(/stream or file/);
  });
});
