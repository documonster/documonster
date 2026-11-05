/**
 * The streaming writer and reader refuse an options value that is not an object when they are constructed, with an
 * `ExcelArgumentError` naming `options` — `null` used to reach a field read and throw a `TypeError`.
 */
import { ExcelArgumentError, Stream } from "@excel/index";
import { describe, expect, it } from "vitest";

function refusal(run: () => unknown): ExcelArgumentError {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(ExcelArgumentError);
    return err as ExcelArgumentError;
  }
  throw new Error("expected an ExcelArgumentError");
}

describe("streaming options", () => {
  it.each([null, 42, "x"])("new Stream.WorkbookWriter(%j) is refused", options => {
    const err = refusal(() => new Stream.WorkbookWriter(options as never));
    expect(err.field).toBe("options");
    expect(err.message).toMatch(/^WorkbookWriter: options must be an object/);
  });

  it.each([null, 42])("new Stream.WorkbookReader(bytes, %j) is refused", options => {
    const err = refusal(() => new Stream.WorkbookReader(new Uint8Array(1), options as never));
    expect(err.field).toBe("options");
    expect(err.message).toMatch(/^WorkbookReader: options must be an object/);
  });

  it("names a wrong reader option by its path in options", () => {
    const password = refusal(
      () => new Stream.WorkbookReader(new Uint8Array(1), { password: 42 } as never)
    );
    expect(password.field).toBe("options.password");
    expect(password.message).toMatch(/^WorkbookReader: options\.password must be a string/);
    const choice = refusal(
      () => new Stream.WorkbookReader(new Uint8Array(1), { worksheets: "emitt" } as never)
    );
    expect(choice.field).toBe("options.worksheets");
  });

  it("still constructs both without options", () => {
    expect(new Stream.WorkbookWriter()).toBeInstanceOf(Stream.WorkbookWriter);
    expect(new Stream.WorkbookReader(new Uint8Array(1))).toBeInstanceOf(Stream.WorkbookReader);
  });
});
