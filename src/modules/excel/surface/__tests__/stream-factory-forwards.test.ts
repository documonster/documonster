import { Stream, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("Workbook.createStreamWriter / createStreamReader (deprecated)", () => {
  it("build the classes their replacement constructors build", () => {
    expect(Workbook.createStreamWriter()).toBeInstanceOf(Stream.WorkbookWriter);
    expect(Workbook.createStreamReader(new Uint8Array(0))).toBeInstanceOf(Stream.WorkbookReader);
  });

  it("refuse what the constructor refuses, with the same error", () => {
    const viaFactory = captured(() => Workbook.createStreamReader(42 as never));
    const viaClass = captured(() => new Stream.WorkbookReader(42 as never));
    expect(viaFactory?.constructor).toBe(viaClass?.constructor);
    expect(viaFactory?.message).toBe(viaClass?.message);
  });
});

function captured(fn: () => unknown): Error | undefined {
  try {
    fn();
  } catch (error) {
    return error as Error;
  }
  return undefined;
}
