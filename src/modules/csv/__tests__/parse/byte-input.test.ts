/**
 * The async parsers take a whole input in bytes — a `Uint8Array` (any `ArrayBufferView`), an
 * `ArrayBuffer` or a `Blob` (a `File`) — as `ExcelCsv.read` does, decoding it as one chunk of the
 * byte stream it is equivalent to. They took a byte stream but refused the same bytes whole.
 */
import { Csv, CsvArgumentError, type CsvAsyncInput } from "@csv/index";
import { describe, expect, expectTypeOf, it } from "vitest";

const text = "name,qty\nwidget,2\ngadget,5\n";
const bytes = new TextEncoder().encode(text);

async function* chunksOf(value: Uint8Array): AsyncGenerator<Uint8Array> {
  yield value;
}

async function rows(input: CsvAsyncInput): Promise<unknown[]> {
  const out: unknown[] = [];
  for await (const row of Csv.parseRows(input)) {
    out.push(row);
  }
  return out;
}

describe("Csv.parseAsync / Csv.parseRows take bytes and a Blob", () => {
  const expected = Csv.parse(text);

  it.each([
    ["Uint8Array", () => bytes],
    ["ArrayBuffer", () => bytes.slice().buffer],
    ["DataView", () => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)],
    ["Blob", () => new Blob([bytes])]
  ] as const)("%s", async (_name, make) => {
    expect(await Csv.parseAsync(make() as CsvAsyncInput)).toEqual(expected);
    expect(await rows(make() as CsvAsyncInput)).toEqual(expected);
  });

  it("decodes as the same bytes streamed would: encoding, BOM", async () => {
    const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...bytes]);
    expect(await Csv.parseAsync(withBom, { headers: true })).toEqual(
      await Csv.parseAsync(chunksOf(withBom), { headers: true })
    );
    const latin1 = new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x2c, 0x31]); // "café,1"
    expect(await Csv.parseAsync(latin1, { encoding: "latin1" })).toEqual([["café", "1"]]);
  });

  it("reports progress over the bytes", async () => {
    const seen: Array<number | undefined> = [];
    await Csv.parseAsync(bytes, { onProgress: p => seen.push(p.bytesProcessed) });
    expect(seen.at(-1)).toBe(bytes.length);
  });

  it("still refuses what is no input at all", async () => {
    await expect(Csv.parseAsync(42 as never)).rejects.toThrow(CsvArgumentError);
    await expect(Csv.parseAsync(null as never)).rejects.toThrow(CsvArgumentError);
  });

  it("is in the type", () => {
    expectTypeOf<Uint8Array>().toMatchTypeOf<CsvAsyncInput>();
    expectTypeOf<ArrayBuffer>().toMatchTypeOf<CsvAsyncInput>();
    expectTypeOf<Blob>().toMatchTypeOf<CsvAsyncInput>();
  });
});
