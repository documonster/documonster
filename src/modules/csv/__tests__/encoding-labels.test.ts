/**
 * An `encoding` label `TextDecoder` does not know used to escape as a bare `RangeError`; it is
 * a `CsvError` naming the option.
 */
import { Csv, CsvError } from "@csv/index";
import { describe, expect, it } from "vitest";

const BOGUS = "no-such-encoding";

async function* chunks(bytes: Uint8Array): AsyncIterable<Uint8Array> {
  yield bytes;
}

describe("CSV encoding labels", () => {
  const bytes = new TextEncoder().encode("a,b\n1,2");

  it("parseAsync refuses an unknown label with CsvError", async () => {
    await expect(Csv.parseAsync(chunks(bytes), { encoding: BOGUS })).rejects.toThrow(CsvError);
    await expect(Csv.parseAsync(chunks(bytes), { encoding: BOGUS })).rejects.toThrow(/encoding/);
  });

  it("the parser stream refuses it when it is created", () => {
    expect(() => Csv.createParserStream({ encoding: BOGUS })).toThrow(CsvError);
  });

  it("keeps decoding with a real label", async () => {
    const latin = Uint8Array.of(0x61, 0x2c, 0xe9);
    expect(await Csv.parseAsync(chunks(latin), { encoding: "windows-1252" })).toEqual([["a", "é"]]);
  });
});
