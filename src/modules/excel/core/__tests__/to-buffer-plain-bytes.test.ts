import { toBuffer as toBufferBrowser } from "@excel/core/xlsx-io.browser";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, expectTypeOf, it } from "vitest";

describe("Workbook.toBuffer resolves to a plain Uint8Array over an ArrayBuffer", () => {
  function sample(): Workbook.Handle {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    return wb;
  }

  it("on the Node entry", async () => {
    const bytes = await Workbook.toBuffer(sample());
    expectTypeOf(bytes).toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expect(bytes.constructor).toBe(Uint8Array);
    expect(new Blob([bytes]).size).toBe(bytes.length);
  });

  it("on the browser entry, so `new Blob([bytes])` type-checks there too", async () => {
    const bytes = await toBufferBrowser(sample());
    expectTypeOf(bytes).toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expect(bytes.constructor).toBe(Uint8Array);
    expect(new Blob([bytes]).size).toBe(bytes.length);
  });
});
