/**
 * XLSX stream vs non-stream parity tests (shared)
 *
 * Goals:
 * - Output parity: `xlsx.write(stream)` output bytes === `xlsx.writeBuffer()` bytes.
 * - Input parity: parsing from stream (`xlsx.read(stream)`) matches parsing from buffer (`xlsx.load(buf)`).
 *
 * This is about determinism/correctness, not performance.
 */

import { cellSetValue, cellGetValue } from "@excel/cell";
import { Workbook } from "@excel/index";
import { StreamBuf } from "@excel/utils/stream-buf";
import type { WorkbookData } from "@excel/workbook-core";
import { getCell } from "@excel/worksheet";
import { describe, it, expect, beforeAll } from "vitest";

interface StreamLike {
  on: (event: string, listener: (...args: any[]) => void) => any;
  end: (chunk?: any) => any;
}

interface PassThroughCtor {
  new (): StreamLike;
}

interface ParityTestContext {
  Workbook: { create: () => WorkbookData };
  PassThrough: PassThroughCtor;
}

function uint8ArrayEquals(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) {
    return true;
  }
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

function buildSmallWorkbook(wb: any): void {
  const ws1 = Workbook.addWorksheet(wb, "Sheet1");
  cellSetValue(getCell(ws1, "A1"), "hello");
  cellSetValue(getCell(ws1, "B1"), 42);
  cellSetValue(getCell(ws1, "C1"), 3.14);

  cellSetValue(getCell(ws1, "A2"), "world");
  cellSetValue(getCell(ws1, "B2"), -7);
  cellSetValue(getCell(ws1, "C2"), 0);

  const ws2 = Workbook.addWorksheet(wb, "Second");
  cellSetValue(getCell(ws2, "A1"), "x");
  cellSetValue(getCell(ws2, "B1"), 1);
  cellSetValue(getCell(ws2, "A2"), "y");
  cellSetValue(getCell(ws2, "B2"), 2);
}

export function createXlsxStreamParityTests(getContext: () => ParityTestContext) {
  describe("XLSX stream/non-stream parity", () => {
    let ctx: ParityTestContext;

    beforeAll(() => {
      ctx = getContext();
    });

    it("write(stream) output bytes === writeBuffer() bytes", async () => {
      const wb = ctx.Workbook.create();
      buildSmallWorkbook(wb);

      const options = { zip: { level: 6, modTime: new Date(Date.UTC(2000, 0, 1, 0, 0, 0)) } };

      const bufferNonStream = await Workbook.toXlsxBuffer(wb, options);

      // In browser builds, `xlsx.write()` expects a StreamBuf-like sink with
      // synchronous `write()`/`end()`; using StreamBuf keeps behavior consistent
      // across Node and browser.
      const stream = new StreamBuf();
      await Workbook.writeXlsxStream(wb, stream as any, options);
      const bufferStream = stream.read() || new Uint8Array(0);

      expect(bufferStream.length).toBe(bufferNonStream.length);
      expect(uint8ArrayEquals(bufferStream, bufferNonStream)).toBe(true);
    }, 30000);

    it("read(stream) model matches load(buffer)", async () => {
      const wb = ctx.Workbook.create();
      buildSmallWorkbook(wb);

      const data = await Workbook.toXlsxBuffer(wb, {
        zip: { level: 6, modTime: new Date(Date.UTC(2000, 0, 1, 0, 0, 0)) }
      });

      const wbFromBuffer = ctx.Workbook.create();
      await Workbook.loadXlsx(wbFromBuffer, data);

      const wbFromStream = ctx.Workbook.create();
      const stream = new ctx.PassThrough();
      stream.end(data);
      await Workbook.readXlsxStream(wbFromStream, stream as any);

      const sheet1a = Workbook.getWorksheet(wbFromBuffer, "Sheet1")!;
      const sheet1b = Workbook.getWorksheet(wbFromStream, "Sheet1")!;

      expect(cellGetValue(getCell(sheet1a, "A1"))).toBe("hello");
      expect(cellGetValue(getCell(sheet1b, "A1"))).toBe("hello");
      expect(cellGetValue(getCell(sheet1b, "B1"))).toBe(42);
      expect(cellGetValue(getCell(sheet1b, "C1"))).toBe(3.14);

      const secondA = Workbook.getWorksheet(wbFromBuffer, "Second")!;
      const secondB = Workbook.getWorksheet(wbFromStream, "Second")!;

      expect(cellGetValue(getCell(secondA, "A2"))).toBe("y");
      expect(cellGetValue(getCell(secondB, "A2"))).toBe("y");
      expect(cellGetValue(getCell(secondB, "B2"))).toBe(2);
    }, 30000);
  });
}
