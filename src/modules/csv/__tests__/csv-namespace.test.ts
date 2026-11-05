import type { CsvColumnConfig, CsvRecordWithInfo, ColumnConfig, RecordWithInfo } from "@csv/index";
import { Csv } from "@csv/index";
import { parseCsvWithProgress } from "@csv/parse/async";
import { CsvFormatterStream, CsvParserStream } from "@csv/stream/index";
import { csvGenerateAsync } from "@csv/utils/generate";
import { describe, expect, it } from "vitest";

describe("Csv namespace", () => {
  it("parseAsync reports progress through options.onProgress", async () => {
    const seen: { rowsProcessed: number; bytesProcessed?: number }[] = [];
    const rows = await Csv.parseAsync("a,b\n1,2\n3,4", { onProgress: info => seen.push(info) });
    expect(rows).toEqual([
      ["a", "b"],
      ["1", "2"],
      ["3", "4"]
    ]);
    expect(seen.at(-1)).toEqual({ rowsProcessed: 3, bytesProcessed: 11 });
  });

  it("onProgress counts data rows only, after byte-only calls while a stream is read", async () => {
    const seen: { rowsProcessed: number; bytesProcessed?: number }[] = [];
    await Csv.parseAsync(
      (async function* () {
        yield "h1,h2\n";
        yield "1,2\n";
      })(),
      { headers: true, onProgress: info => seen.push(info) }
    );
    expect(seen).toEqual([
      { rowsProcessed: 0, bytesProcessed: 6 },
      { rowsProcessed: 0, bytesProcessed: 10 },
      { rowsProcessed: 1, bytesProcessed: 10 }
    ]);
  });

  it("keeps parseWithProgress as a deprecated forward", () => {
    expect(Csv.parseWithProgress).toBe(parseCsvWithProgress);
  });

  it("generateRowsAsync is the async row generator; generateAsync forwards to it", () => {
    expect(Csv.generateRowsAsync).toBe(csvGenerateAsync);
    expect(Csv.generateAsync).toBe(csvGenerateAsync);
  });

  it("publishes the stream classes, built with their constructors", () => {
    expect(Csv.ParserStream).toBe(CsvParserStream);
    expect(Csv.FormatterStream).toBe(CsvFormatterStream);
    expect(new Csv.ParserStream()).toBeInstanceOf(CsvParserStream);
    expect(new Csv.FormatterStream()).toBeInstanceOf(CsvFormatterStream);
  });

  it("keeps the deprecated factories, which build the same classes", async () => {
    expect(Csv.createParserStream({ headers: true })).toBeInstanceOf(Csv.ParserStream);
    expect(Csv.createFormatterStream()).toBeInstanceOf(Csv.FormatterStream);
    expect(() => Csv.createParserStream({ delimiter: 5 as never })).toThrow(
      /^new Csv\.ParserStream: /
    );

    const rows = await new Promise<unknown[]>((resolve, reject) => {
      const out: unknown[] = [];
      const parser = Csv.createParserStream({ headers: true });
      parser.on("data", row => out.push(row));
      parser.on("end", () => resolve(out));
      parser.on("error", reject);
      parser.end("a,b\n1,2\n");
    });
    expect(rows).toEqual([{ a: "1", b: "2" }]);
  });

  it("detectDelimiter takes an options object and still accepts the positional form", () => {
    expect(Csv.detectDelimiter("a:b\n1:2", { delimitersToGuess: [":"] })).toBe(":");
    expect(Csv.detectDelimiter("a:b\n1:2", '"', [":"])).toBe(":");
    expect(Csv.detectDelimiter("a;b\n1;2")).toBe(";");
  });

  it("detectLinebreak takes an options object and still accepts a positional quote", () => {
    expect(Csv.detectLinebreak("'a\nb',c\r\nd", { quote: "'" })).toBe("\r\n");
    expect(Csv.detectLinebreak("'a\nb',c\r\nd", "'")).toBe("\r\n");
  });

  it("parseAsync with onProgress reproduces parseWithProgress exactly", async () => {
    const input = (): AsyncIterable<string> =>
      (async function* () {
        yield "name,age\nA";
        yield "lice,30\nBob,4";
        yield "0\n";
      })();
    for (const options of [{}, { headers: true }, { headers: true, info: true }] as const) {
      const viaForward: unknown[] = [];
      const viaOption: unknown[] = [];
      const a = await Csv.parseWithProgress(input(), options, info => viaForward.push(info));
      const b = await Csv.parseAsync(input(), {
        ...options,
        onProgress: info => viaOption.push(info)
      });
      expect(b).toEqual(a);
      expect(viaOption).toEqual(viaForward);
    }
  });

  it("the positional detect forms keep their defaults for undefined and empty arguments", () => {
    expect(Csv.detectDelimiter("a|b\n1|2", undefined, ["|"])).toBe("|");
    expect(Csv.detectDelimiter("a;b\n1;2", undefined)).toBe(";");
    expect(Csv.detectLinebreak('"a\r\nb"\nc', undefined)).toBe("\n");
    expect(Csv.detectLinebreak('"a\r\nb"\nc', { quote: '"' })).toBe("\n");
    expect(Csv.detectDelimiter("#x;y\na,b\n1,2", '"', undefined, "#")).toBe(",");
    expect(Csv.detectDelimiter("a,b\n\n\n1,2", '"', undefined, undefined, true)).toBe(",");
  });

  it("the stream classes stay public values for instanceof and subclassing", () => {
    class Tagged extends Csv.ParserStream {}
    expect(new Tagged()).toBeInstanceOf(Csv.ParserStream);
  });

  it("the unprefixed type names are aliases of the Csv-prefixed ones", () => {
    const column: ColumnConfig = { key: "a", header: "A" } satisfies CsvColumnConfig;
    const record: RecordWithInfo<string[]> = {
      record: ["x"],
      info: { index: 0, line: 1, offset: 0, quoted: [false] }
    } satisfies CsvRecordWithInfo<string[]>;
    expect(column.header).toBe("A");
    expect(record.record).toEqual(["x"]);
  });
});
