/**
 * A malformed file is not a wrong argument. A position in a file's content that is outside the grid, or not where it
 * is written — a `<c r>` in another row, column `XFE`, row `0`, a merge from `C0` — is reported by every reader as the
 * reader's own parse error, with the part named in the message and the position error as `cause`:
 *
 * - XLSX, buffered and streamed (`Workbook.read`, `readStream`, `readFile`) and `Stream.WorkbookReader`: `XlsxParseError`;
 * - XLSB: `XlsbParseError`;
 * - `ExcelCsv` / `ExcelMarkdown`: `ExcelFileError` naming the row.
 *
 * The position errors themselves are argument errors (`ExcelArgumentError`), which is what they mean at an API call;
 * a read used to let them through unchanged.
 */
import { Readable } from "node:stream";

import { Archive } from "@archive";
import {
  Cell,
  ColumnOutOfBoundsError,
  ExcelArgumentError,
  ExcelError,
  ExcelFileError,
  InvalidAddressError,
  isExcelError,
  RowOutOfBoundsError,
  Stream,
  Workbook,
  Worksheet,
  XlsbParseError,
  XlsxParseError
} from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { ExcelMarkdown } from "@excel/bridge/markdown";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const SHEET = "xl/worksheets/sheet1.xml";

type ErrorClass = abstract new (...args: never[]) => Error;

async function source(format: "xlsx" | "xlsb" = "xlsx"): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "B5", 2);
  Worksheet.merge(ws, "C1:D2");
  return Workbook.toBuffer(wb, { format });
}

async function withPart(
  bytes: Uint8Array,
  path: string,
  edit: (data: Uint8Array) => Uint8Array
): Promise<Uint8Array> {
  const files = await Archive.unpackZip(bytes);
  const entries = new Map<string, Uint8Array>();
  for (const [name, entry] of files) {
    entries.set(name, name === path ? edit(entry.data) : entry.data);
  }
  return Archive.packZip(entries);
}

async function xlsxWith(edit: (xml: string) => string): Promise<Uint8Array> {
  return withPart(await source(), SHEET, data =>
    new TextEncoder().encode(edit(new TextDecoder().decode(data)))
  );
}

async function failure(run: () => Promise<unknown>): Promise<Error> {
  try {
    await run();
  } catch (error) {
    return error as Error;
  }
  throw new Error("the read succeeded");
}

async function streamed(bytes: Uint8Array): Promise<void> {
  for await (const event of new Stream.WorkbookReader(bytes as never, {
    worksheets: "emit"
  }).parse()) {
    if ((event as { eventType: string }).eventType === "worksheet") {
      for await (const _row of (event as { value: AsyncIterable<unknown> }).value) {
        // Reading the rows is what fails.
      }
    }
  }
}

function expectFileError(error: Error, type: ErrorClass, part: string, cause: ErrorClass): void {
  expect(error).toBeInstanceOf(type);
  expect(error).not.toBeInstanceOf(ExcelArgumentError);
  expect(error).toBeInstanceOf(ExcelError);
  expect(isExcelError(error)).toBe(true);
  expect(error.cause).toBeInstanceOf(cause);
  expect(error.message).toContain(part);
  expect(error.message).toContain((error.cause as Error).message);
}

const XLSX_CASES: [string, (xml: string) => string, ErrorClass, boolean][] = [
  [
    "a cell written in another row",
    xml => xml.replace('r="B5"', 'r="B7"'),
    InvalidAddressError,
    true
  ],
  ["a column past XFD", xml => xml.replace('r="B5"', 'r="XFE5"'), ColumnOutOfBoundsError, true],
  [
    "row 0",
    xml => xml.replace('<row r="5"', '<row r="0"').replace('r="B5"', 'r="B0"'),
    RowOutOfBoundsError,
    true
  ],
  [
    "a row past 1048576",
    xml => xml.replace('<row r="5"', '<row r="1048577"').replace('r="B5"', 'r="B1048577"'),
    RowOutOfBoundsError,
    true
  ],
  // The streaming reader reads no merges.
  [
    "a merge from row 0",
    xml => xml.replace('ref="C1:D2"', 'ref="C0:D2"'),
    RowOutOfBoundsError,
    false
  ],
  [
    "a merge to column XFE",
    xml => xml.replace('ref="C1:D2"', 'ref="C1:XFE2"'),
    ColumnOutOfBoundsError,
    false
  ]
];

describe("XLSX: a bad position in a part is an XlsxParseError naming the part", () => {
  describe.each(XLSX_CASES)("%s", (_name, edit, cause, inRows) => {
    it("Workbook.read", async () => {
      const bytes = await xlsxWith(edit);
      expectFileError(await failure(() => Workbook.read(bytes)), XlsxParseError, SHEET, cause);
    });

    it("Workbook.readStream", async () => {
      const bytes = await xlsxWith(edit);
      expectFileError(
        await failure(() => Workbook.readStream(Readable.from([bytes]))),
        XlsxParseError,
        SHEET,
        cause
      );
    });

    it.runIf(inRows)("Stream.WorkbookReader", async () => {
      const bytes = await xlsxWith(edit);
      expectFileError(await failure(() => streamed(bytes)), XlsxParseError, SHEET, cause);
    });
  });

  it("the same position at an API call is still an argument error", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "S");
    expect(() => Cell.getValue(ws, "B0")).toThrow(ExcelArgumentError);
    expect(() => Worksheet.merge(ws, "C1:XFE2")).toThrow(ExcelArgumentError);
  });
});

/**
 * The XLSB reader decodes positions with its own checks (`checkSheetRange` in `xlsb/binary.ts`), which already raise
 * `XlsbParseError` naming the part and record — the shape the XLSX readers now share.
 */
describe("XLSB: a bad position in a part is an XlsbParseError naming the part", () => {
  it("a merge past the last row", async () => {
    const bytes = await withPart(await source("xlsb"), "xl/worksheets/sheet1.bin", data => {
      const copy = data.slice();
      for (const record of iterateInterpretableRecords(copy, "sheet")) {
        if (recordSpec(record.id)?.name === "BrtMergeCell") {
          // `rwLast` — the second 32-bit field — past row 1048576.
          new DataView(record.payload.buffer, record.payload.byteOffset).setUint32(
            4,
            0x200000,
            true
          );
        }
      }
      return copy;
    });
    const error = await failure(() => Workbook.read(bytes));
    expect(error).toBeInstanceOf(XlsbParseError);
    expect(error).not.toBeInstanceOf(ExcelArgumentError);
    expect(error.message).toMatch(/^xl\/worksheets\/sheet1\.bin: BrtMergeCell names row 2097153/);
  });
});

describe("CSV and Markdown: a row wider than a sheet is an ExcelFileError naming the row", () => {
  const wide = (cells: number) => Array.from({ length: cells }, (_, i) => `c${i}`);

  it("ExcelCsv.read and addWorksheet", async () => {
    const csv = `a\n${wide(16385).join(",")}`;
    for (const run of [
      () => ExcelCsv.read(csv),
      () => ExcelCsv.read(new TextEncoder().encode(csv)),
      () => ExcelCsv.read(Readable.from([csv])),
      () => ExcelCsv.addWorksheet(Workbook.create(), csv)
    ]) {
      const error = await failure(run);
      expectFileError(error, ExcelFileError, "row 2", ColumnOutOfBoundsError);
      expect(error.message).toContain("CSV");
    }
  });

  it("ExcelMarkdown.read", async () => {
    const md = `| ${wide(16385).join(" | ")} |\n|${" --- |".repeat(16385)}\n`;
    const error = await failure(async () => ExcelMarkdown.read(md));
    expectFileError(error, ExcelFileError, "row 1", ColumnOutOfBoundsError);
    expect(error.message).toContain("Markdown");
  });
});
