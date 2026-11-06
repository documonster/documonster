/**
 * `ExcelCsv.read` / `addWorksheet` take the inputs `ExcelCsvInput` names, and `ExcelCsvInput` names what they
 * take: a Web `ReadableStream` of bytes (also one that is not async-iterable, as in browsers
 * without `ReadableStream` iteration) and any async iterable of text or bytes.
 */
import { Cell, Workbook } from "@excel";
import { ExcelCsv, type ExcelCsvInput } from "@excel/bridge/csv";
import { describe, expect, expectTypeOf, it } from "vitest";

const bytes = new TextEncoder().encode("a,b\n1,2\n");

function webStream(iterable: boolean): ReadableStream<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes);
      controller.close();
    }
  });
  if (!iterable) {
    // A browser `ReadableStream` without async iteration: only `getReader()`
    Object.defineProperty(stream, Symbol.asyncIterator, { value: undefined });
  }
  return stream;
}

async function* chunks(): AsyncGenerator<string> {
  yield "a,b\n";
  yield "1,2\n";
}

describe("ExcelCsv.read takes what ExcelCsvInput names", () => {
  it.each([
    ["an async-iterable ReadableStream", () => webStream(true)],
    ["a ReadableStream with only getReader()", () => webStream(false)],
    ["an async generator of strings", () => chunks()]
  ])("%s", async (_name, make) => {
    const wb = await ExcelCsv.read(make() as ExcelCsvInput);
    const ws = Workbook.getWorksheets(wb)[0]!;
    expect(Cell.getValue(ws, "A1")).toBe("a");
    expect(Cell.getValue(ws, "B2")).toBe(2);
  });

  it("ExcelCsvInput includes them", () => {
    expectTypeOf<ReadableStream<Uint8Array>>().toMatchTypeOf<ExcelCsvInput>();
    expectTypeOf<AsyncIterable<Uint8Array | string>>().toMatchTypeOf<ExcelCsvInput>();
    expectTypeOf<DataView>().toMatchTypeOf<ExcelCsvInput>();
    expectTypeOf<Uint16Array>().toMatchTypeOf<ExcelCsvInput>();
  });

  it("reads any ArrayBufferView as the bytes it views", async () => {
    const framed = new Uint8Array(bytes.length + 4);
    framed.set(bytes, 2);
    const wb = await ExcelCsv.read(new DataView(framed.buffer, 2, bytes.length));
    const ws = Workbook.getWorksheets(wb)[0]!;
    expect(Cell.getValue(ws, "A1")).toBe("a");
    expect(Cell.getValue(ws, "B2")).toBe(2);
  });
});
