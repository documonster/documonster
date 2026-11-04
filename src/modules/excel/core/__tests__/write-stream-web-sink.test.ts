/**
 * `Workbook.writeStream` takes a Web `WritableStream` — the sink a browser, Deno or a `Response` body gives — as well
 * as a Node writable, in both containers; anything that is neither is an `ExcelArgumentError` on `stream`. It used to
 * crash on a Web stream with `this.pipedStream.write is not a function`.
 */
import { Cell, ExcelArgumentError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", "hello");
  return wb;
}

function collector() {
  const chunks: Uint8Array[] = [];
  let closed = false;
  const stream = new WritableStream<Uint8Array>({
    write(chunk) {
      chunks.push(chunk);
    },
    close() {
      closed = true;
    }
  });
  const bytes = (): Uint8Array => {
    const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    for (const chunk of chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  };
  return { stream, bytes, isClosed: () => closed };
}

describe("Workbook.writeStream into a Web WritableStream", () => {
  it.each(["xlsx", "xlsb"] as const)(
    "%s: writes a package that reads back, and closes the stream",
    async format => {
      const { stream, bytes, isClosed } = collector();
      await Workbook.writeStream(book(), stream, { format });
      expect(isClosed()).toBe(true);
      const read = await Workbook.read(bytes());
      expect(Cell.getValue(Workbook.getWorksheet(read, "S")!, "A1")).toBe("hello");
      expect(stream.locked).toBe(false);
    }
  );

  it("a sink that fails rejects the call", async () => {
    const stream = new WritableStream<Uint8Array>({
      write() {
        throw new Error("disk full");
      }
    });
    await expect(Workbook.writeStream(book(), stream)).rejects.toThrow("disk full");
  });

  it("anything else is refused on `stream`", async () => {
    for (const stream of [undefined, null, 42, "out.xlsx", {}, { write() {} }]) {
      await expect(Workbook.writeStream(book(), stream as never)).rejects.toThrow(
        ExcelArgumentError
      );
      await expect(Workbook.writeStream(book(), stream as never)).rejects.toThrow(
        expect.objectContaining({
          field: "stream",
          message: expect.stringMatching(/^Workbook\.writeStream: stream/)
        })
      );
    }
  });
});
