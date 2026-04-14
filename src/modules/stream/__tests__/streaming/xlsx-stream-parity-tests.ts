/**
 * XLSX stream vs non-stream parity tests (shared)
 *
 * Goals:
 * - Output parity: `xlsx.write(stream)` output bytes === `xlsx.writeBuffer()` bytes.
 * - Input parity: parsing from stream (`xlsx.read(stream)`) matches parsing from buffer (`xlsx.load(buf)`).
 *
 * This is about determinism/correctness, not performance.
 */

import { StreamBuf } from "@excel/utils/stream-buf";
import { describe, it, expect, beforeAll } from "vitest";

interface WorkbookLike {
  addWorksheet: (name?: string) => any;
  getWorksheet: (id?: any) => any;
  xlsx: {
    write: (stream: any, options?: any) => Promise<any>;
    writeBuffer: (options?: any) => Promise<Uint8Array>;
    load: (data: Uint8Array, options?: any) => Promise<any>;
    read: (stream: any, options?: any) => Promise<any>;
  };
}

interface StreamLike {
  on: (event: string, listener: (...args: any[]) => void) => any;
  end: (chunk?: any) => any;
}

interface PassThroughCtor {
  new (): StreamLike;
}

interface ParityTestContext {
  Workbook: new () => WorkbookLike;
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

function buildSmallWorkbook(wb: WorkbookLike): void {
  const ws1 = wb.addWorksheet("Sheet1");
  ws1.getCell("A1").value = "hello";
  ws1.getCell("B1").value = 42;
  ws1.getCell("C1").value = 3.14;

  ws1.getCell("A2").value = "world";
  ws1.getCell("B2").value = -7;
  ws1.getCell("C2").value = 0;

  const ws2 = wb.addWorksheet("Second");
  ws2.getCell("A1").value = "x";
  ws2.getCell("B1").value = 1;
  ws2.getCell("A2").value = "y";
  ws2.getCell("B2").value = 2;
}

export function createXlsxStreamParityTests(getContext: () => ParityTestContext) {
  describe("XLSX stream/non-stream parity", () => {
    let ctx: ParityTestContext;

    beforeAll(() => {
      ctx = getContext();
    });

    it("write(stream) output bytes === writeBuffer() bytes", async () => {
      const wb = new ctx.Workbook();
      buildSmallWorkbook(wb);

      const options = { zip: { level: 6, modTime: new Date(Date.UTC(2000, 0, 1, 0, 0, 0)) } };

      const bufferNonStream = await wb.xlsx.writeBuffer(options);

      // In browser builds, `xlsx.write()` expects a StreamBuf-like sink with
      // synchronous `write()`/`end()`; using StreamBuf keeps behavior consistent
      // across Node and browser.
      const stream = new StreamBuf();
      await wb.xlsx.write(stream as any, options);
      const bufferStream = stream.read() || new Uint8Array(0);

      expect(bufferStream.length).toBe(bufferNonStream.length);
      expect(uint8ArrayEquals(bufferStream, bufferNonStream)).toBe(true);
    }, 30000);

    it("read(stream) model matches load(buffer)", async () => {
      const wb = new ctx.Workbook();
      buildSmallWorkbook(wb);

      const data = await wb.xlsx.writeBuffer({
        zip: { level: 6, modTime: new Date(Date.UTC(2000, 0, 1, 0, 0, 0)) }
      });

      const wbFromBuffer = new ctx.Workbook();
      await wbFromBuffer.xlsx.load(data);

      const wbFromStream = new ctx.Workbook();
      const stream = new ctx.PassThrough();
      stream.end(data);
      await wbFromStream.xlsx.read(stream as any);

      const sheet1a = wbFromBuffer.getWorksheet("Sheet1");
      const sheet1b = wbFromStream.getWorksheet("Sheet1");

      expect(sheet1a.getCell("A1").value).toBe("hello");
      expect(sheet1b.getCell("A1").value).toBe("hello");
      expect(sheet1b.getCell("B1").value).toBe(42);
      expect(sheet1b.getCell("C1").value).toBe(3.14);

      const secondA = wbFromBuffer.getWorksheet("Second");
      const secondB = wbFromStream.getWorksheet("Second");

      expect(secondA.getCell("A2").value).toBe("y");
      expect(secondB.getCell("A2").value).toBe("y");
      expect(secondB.getCell("B2").value).toBe(2);
    }, 30000);
  });
}
